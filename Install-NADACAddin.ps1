#Requires -Version 5.1
<#
.SYNOPSIS
    Installs the Pharmacy Tools Excel Add-In (NADAC Lookup) v2.3.3

.DESCRIPTION
    Builds an Excel .xlam add-in containing:
      - modNADAC          : NADAC Lookup main module (v2.3.3)
      - frmDatasetManager : Full Dataset Manager UserForm (v2.3.3)
    Then injects a custom ribbon tab and pharmacy-bottle icon.
    Registers the add-in so it loads automatically in Excel.

    v2.3.3 (2026-04-09):
      - Fix: m_LastHttpStatus moved to module-level declarations
        (was incorrectly placed inside HttpPost function body)

    v2.3.2 (2026-04-09):
      - Distribution UUID auto-discovered from CMS metastore at runtime.
        Hardcoded UUIDs now serve only as a fast-path fallback; stale UUIDs
        no longer cause 400 errors.
      - POST 400 auto-retry: on a 400 response the add-in re-queries the
        metastore for a fresh UUID and retries the batch once before
        falling back to GET.

    v2.3.1 hotfix (2026-04-09):
      - 2026 dataset and distribution UUIDs hardcoded (POST now works for 2026)

    v2.3 changes vs v2.2:
      - Dataset Manager replaced with a proper UserForm grid showing
        all built-in and custom years with Edit / Delete / Reset
      - Re-run colour reset: all rows cleared to white before each run;
        salmon-red re-applied only to current-run failures
      - About screen updated

.NOTES
    Run as the end user (not as Administrator).
    Excel must be installed.  PowerShell 5.1+.
#>

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

Write-Host ''
Write-Host '  =====================================================' -ForegroundColor Cyan
Write-Host '   Pharmacy Tools Add-In  --  Installer  v2.3.3'        -ForegroundColor Cyan
Write-Host '  =====================================================' -ForegroundColor Cyan
Write-Host ''

# -----------------------------------------------------------------------------
#  PRE-FLIGHT
# -----------------------------------------------------------------------------
$RegVBOM = 'HKCU:\Software\Microsoft\Office'
$OfficeVersions = Get-ChildItem $RegVBOM -ErrorAction SilentlyContinue |
    Where-Object { $_.PSChildName -match '^\d+\.\d+$' } |
    Sort-Object { [double]$_.PSChildName } -Descending

if (-not $OfficeVersions) {
    Write-Host '  ERROR: Microsoft Office registry keys not found.' -ForegroundColor Red
    Write-Host '  Please ensure Excel is installed before running this script.' -ForegroundColor Red
    exit 1
}

$OfficeVer  = $OfficeVersions[0].PSChildName
$VBOMPath   = "$RegVBOM\$OfficeVer\Excel\Security"
$VBOMName   = 'AccessVBOM'

Write-Host "  Detected Office version : $OfficeVer" -ForegroundColor DarkGray

# Check for running Excel instances
$xlProcs = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
if ($xlProcs) {
    Write-Host ''
    Write-Host '  WARNING: Excel is currently open.' -ForegroundColor Yellow
    Write-Host '  Please save your work and close Excel, then press Enter to continue.'
    Read-Host '  Press Enter when Excel is closed'
    $xlProcs = Get-Process -Name EXCEL -ErrorAction SilentlyContinue
    if ($xlProcs) {
        Write-Host '  Excel is still running. Proceeding anyway -- the add-in may not load' -ForegroundColor Yellow
        Write-Host '  correctly until Excel is restarted.' -ForegroundColor Yellow
    }
}

# Enable AccessVBOM (trust VBA project object model)
$OldVBOM = $null
try {
    $OldVBOM = (Get-ItemProperty -Path $VBOMPath -Name $VBOMName -ErrorAction Stop).$VBOMName
} catch { $OldVBOM = $null }
Set-ItemProperty -Path $VBOMPath -Name $VBOMName -Value 1 -Type DWord -Force
Write-Host "  AccessVBOM enabled for Office $OfficeVer" -ForegroundColor DarkGray

# Install path
$AddinDir  = Join-Path $env:APPDATA 'PharmacyTools'
$AddinPath = Join-Path $AddinDir    'NADACLookup.xlam'
if (-not (Test-Path $AddinDir)) { New-Item -ItemType Directory -Path $AddinDir -Force | Out-Null }

# -----------------------------------------------------------------------------
#  EMBEDDED ASSETS
# -----------------------------------------------------------------------------
$IconB64 = 'iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAABV0lEQVR4nM1XMQ7CMAxMKv6RkQ906wMYkPITlg5IHZgYKjGw8JNKDDygWz/QsS8pA2oVJXZquylwEyTO+eI4rqMVA8ZWI8VuaGpN5SQZUh1LhEQNpI45QtCJmHOMULKGTMQ5Vw5HtoVzbA3ErWMGEscQYrwZxWgtfC7X105CKEk2DKCyGImx1Xi/nVDC8vwg35KhqXXmD6QMvQ+Im30EQ1PrUilRjUgiYMI+L4Kxvmuja6DcCerAEpbKMzbv7376z4qAsdW4zwvVdy26209k4HBDYkRHAIWfCl9YIABLFgivy3H+fbg+2c6V8nIg1eeXgzkCHOfuzv0xSiRcZJDzb0YiuIZbVkII4C1YEjGFmZuEEEgNSSoYW40+v3Yn3Qmso6HUgL5roz2Ay88uxakxC1hSvAax6JIaR3cBRdhkT2lwf96W/+fDhELIgehplkLI6sepVAgnX95X9Nb4jUrLGgAAAABJRU5ErkJggg=='

$VbaCode = @'
Option Explicit

' --- Version ------------------------------------------------------------------
Public  Const ADDIN_VERSION    As String = "2.3.3"
Private Const ADDIN_BUILD_DATE As String = "2026-04-09"

' Module-level: tracks last HTTP status so callers can detect 400 vs network error
Private m_LastHttpStatus As Long

' --- API constants ------------------------------------------------------------
Private Const API_BASE    As String = "https://data.medicaid.gov/api/1/datastore/query/"
Private Const SEARCH_BASE As String = "https://data.medicaid.gov/api/1/search"
Private Const BATCH_SIZE  As Long   = 15
Private Const API_LIMIT   As Long   = 5000
Private Const RATE_MS     As Long   = 1200

' --- Registry key (Public so frmDatasetManager can access it) -----------------
Public Const REG_KEY As String = "HKCU\Software\PharmacyToolsAddin\Datasets"

' --- Hardcoded annual dataset IDs ---------------------------------------------
' Key = year string, Value = CMS dataset UUID
Public Function KnownYearIDs() As Object
    Dim d As Object : Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    d("2026") = "fbb83258-11c7-47f5-8b18-5f8e79f7e704"
    d("2025") = "f38d0706-1239-442c-a3cc-40ef1b686ac0"
    d("2024") = "99315a95-37ac-4eee-946a-3c523b4c481e"
    d("2023") = "4a00010a-132b-4e4d-a611-543c9521280f"
    d("2022") = "dfa2ab14-06c2-457a-9e36-5cb6d80f8d93"
    d("2021") = "d5eaf378-dcef-5779-83de-acdd8347d68e"
    d("2019") = "76a1984a-6d69-5e4d-86c8-65eb31f0506d"
    d("2013") = "1fe73992-cbfd-5109-97bc-dee8b33fdcff"
    Set KnownYearIDs = d
End Function

' --- Hardcoded distribution UUIDs (for POST /api/1/datastore/query/{distUUID}) -
' Key = Dataset ID UUID, Value = Distribution UUID
' 2024, 2025, 2026 confirmed; other years use GET fallback.
Public Function KnownDistUUIDs() As Object
    Dim d As Object : Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    d("fbb83258-11c7-47f5-8b18-5f8e79f7e704") = "8b801945-2507-5057-8f97-eb7586889f3d"  ' 2026
    d("f38d0706-1239-442c-a3cc-40ef1b686ac0") = "ae004d7f-5799-5de3-91ec-f1247f1a5452"  ' 2025
    d("99315a95-37ac-4eee-946a-3c523b4c481e") = "b3b205f4-e788-5ec6-a342-d889111a6c2e"  ' 2024
    Set KnownDistUUIDs = d
End Function

' --- Load user-defined / override datasets from registry ----------------------
' Registry value "UserDatasets" under REG_KEY holds a pipe-delimited list:
'   year:datasetID:distributionUUID|year:datasetID:distributionUUID|...
' Entries whose year matches a built-in take precedence (override).
Public Sub LoadUserDatasets(known As Object, knownDist As Object)
    Dim raw As String : raw = GetRegValue(REG_KEY, "UserDatasets")
    If raw = "" Then Exit Sub
    Dim entries() As String : entries = Split(raw, "|")
    Dim i As Long
    For i = 0 To UBound(entries)
        Dim e As String : e = Trim(entries(i))
        If e <> "" Then
            Dim p() As String : p = Split(e, ":")
            If UBound(p) >= 1 Then
                Dim yr   As String : yr   = Trim(p(0))
                Dim dsID As String : dsID = Trim(p(1))
                Dim dtID As String : dtID = ""
                If UBound(p) >= 2 Then dtID = Trim(p(2))
                If yr <> "" And dsID <> "" Then
                    known(yr) = dsID
                    If dtID <> "" Then knownDist(dsID) = dtID
                End If
            End If
        End If
    Next i
End Sub

' --- Ribbon callbacks ---------------------------------------------------------
Public Sub RunNADACLookup(control As IRibbonControl)
    RunLookup
End Sub

Public Sub RunManageDatasets(control As IRibbonControl)
    ShowDatasetManager
End Sub

Public Sub RunAbout(control As IRibbonControl)
    ShowAbout
End Sub

' --- Main lookup --------------------------------------------------------------
Public Sub RunLookup()
    On Error GoTo ErrHandler
    LogRotate
    LogLine "=== NADAC Lookup v" & ADDIN_VERSION & " started ==="

    Dim wb As Workbook : Set wb = ActiveWorkbook
    If wb Is Nothing Then
        MsgBox "Please open a workbook first.", vbExclamation, "NADAC Lookup"
        Exit Sub
    End If

    Dim ws As Worksheet : Set ws = wb.ActiveSheet
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data rows found on the active sheet.", vbExclamation, "NADAC Lookup"
        Exit Sub
    End If

    ' -- Detect or ask for NDC column ------------------------------------------
    Dim ndcCol As Long
    ndcCol = FindHeaderCol(ws, Array("NDC", "NDC Code", "NationalDrugCode", "National Drug Code"))
    If ndcCol > 0 Then
        If MsgBox("NDC column detected:  " & ColLtr(ndcCol) & _
                  "  (""" & ws.Cells(1, ndcCol).Value & """)" & vbCrLf & vbCrLf & _
                  "Use this column?", vbYesNo + vbQuestion, "NADAC Lookup") = vbNo Then
            ndcCol = 0
        End If
    End If
    If ndcCol = 0 Then
        Dim si As String
        si = InputBox("Enter the column letter containing NDC codes  (e.g.  M):", "NADAC Lookup")
        If si = "" Then Exit Sub
        On Error Resume Next
        ndcCol = ws.Range(UCase(Trim(si)) & "1").Column
        On Error GoTo ErrHandler
        If ndcCol = 0 Then MsgBox "Invalid column.", vbCritical, "NADAC Lookup" : Exit Sub
    End If

    ' -- Detect or ask for Date column -----------------------------------------
    Dim dateCol As Long
    dateCol = FindHeaderCol(ws, Array("Date", "Dispense Date", "DispenseDate", _
                                      "Fill Date", "FillDate", "Rx Date", "RxDate"))
    If dateCol = 0 Then
        Dim di As String
        di = InputBox("Enter the column letter for the dispense date  (e.g.  A):", "NADAC Lookup", "A")
        If di = "" Then Exit Sub
        On Error Resume Next
        dateCol = ws.Range(UCase(Trim(di)) & "1").Column
        On Error GoTo ErrHandler
        If dateCol = 0 Then MsgBox "Invalid column.", vbCritical, "NADAC Lookup" : Exit Sub
    End If

    ' -- Confirm ---------------------------------------------------------------
    If MsgBox("NADAC Lookup  --  Ready to run." & vbCrLf & vbCrLf & _
              "  NDC column   :  " & ColLtr(ndcCol) & "  (""" & ws.Cells(1, ndcCol).Value & """)" & vbCrLf & _
              "  Date column  :  " & ColLtr(dateCol) & "  (""" & ws.Cells(1, dateCol).Value & """)" & vbCrLf & _
              "  Data rows    :  " & (lastRow - 1) & vbCrLf & vbCrLf & _
              "Columns to be added / updated:" & vbCrLf & _
              "   NADAC  |  NADAC Effective Date" & vbCrLf & vbCrLf & _
              "Each row is matched to the correct annual NADAC dataset" & vbCrLf & _
              "based on the year of its dispense date." & vbCrLf & vbCrLf & _
              "Click OK to begin.", vbOKCancel + vbInformation, "NADAC Lookup") = vbCancel Then
        Exit Sub
    End If

    ' -- Pass 1: scan rows -- build per-year NDC sets and max-date map ---------
    Dim yearNDCs  As Object : Set yearNDCs  = CreateObject("Scripting.Dictionary")
    Dim yearMaxDt As Object : Set yearMaxDt = CreateObject("Scripting.Dictionary")
    yearNDCs.CompareMode = 1 : yearMaxDt.CompareMode = 1

    Dim rowCN() As String : ReDim rowCN(2 To lastRow)
    Dim rowYR() As String : ReDim rowYR(2 To lastRow)

    Dim r As Long
    For r = 2 To lastRow
        Dim rawV As String : rawV = Trim(CStr(ws.Cells(r, ndcCol).Value))
        Dim cn   As String : cn   = CleanNDC(rawV)
        rowCN(r) = cn

        On Error Resume Next
        Dim dt   As Date    : dt   = CDate(ws.Cells(r, dateCol).Value)
        Dim dtOK As Boolean : dtOK = (Err.Number = 0)
        Err.Clear : On Error GoTo ErrHandler

        If dtOK And cn <> "" Then
            Dim yr As String : yr = CStr(Year(dt))
            rowYR(r) = yr
            If Not yearNDCs.Exists(yr) Then
                Dim nd As Object : Set nd = CreateObject("Scripting.Dictionary")
                nd.CompareMode = 0 : yearNDCs.Add yr, nd
            End If
            Dim ndSet As Object : Set ndSet = yearNDCs(yr)
            If Not ndSet.Exists(cn) Then ndSet(cn) = ""
            If Not yearMaxDt.Exists(yr) Then
                yearMaxDt(yr) = dt
            ElseIf dt > CDate(yearMaxDt(yr)) Then
                yearMaxDt(yr) = dt
            End If
        End If
    Next r

    If yearNDCs.Count = 0 Then
        MsgBox "No valid NDC / date pairs found.", vbExclamation, "NADAC Lookup"
        Exit Sub
    End If

    ' -- Output columns DECLARED here (assigned after Pass 1 to avoid column shift) -
    Dim colNPU As Long
    Dim colED  As Long

    ' -- Resolve dataset IDs for all required years ----------------------------
    Dim known     As Object : Set known     = KnownYearIDs()
    Dim knownDist As Object : Set knownDist = KnownDistUUIDs()
    LoadUserDatasets known, knownDist

    Dim dynCache As Object : Set dynCache = CreateObject("Scripting.Dictionary")
    dynCache.CompareMode = 1

    ' Count total batches for progress bar
    Dim totalBatches As Long : totalBatches = 0
    Dim yk As Variant
    For Each yk In yearNDCs.Keys
        Dim cnt As Long : cnt = yearNDCs(yk).Count
        totalBatches = totalBatches + Int((cnt + BATCH_SIZE - 1) / BATCH_SIZE)
    Next yk

    Dim apiCache   As Object  : Set apiCache   = CreateObject("Scripting.Dictionary")
    Dim batchsDone As Long    : batchsDone = 0
    Dim truncWarn  As Boolean : truncWarn  = False
    Dim missYears  As String  : missYears  = ""

    Application.ScreenUpdating = False
    Application.Calculation    = xlCalculationManual

    ' -- Iterate years ascending -----------------------------------------------
    Dim sortedYears() As String : sortedYears = SortedKeys(yearNDCs)
    Dim yi As Long
    For yi = 0 To UBound(sortedYears)
        yr = sortedYears(yi)

        ' Resolve Dataset ID
        Dim dsID As String : dsID = ""
        If known.Exists(yr) Then
            dsID = CStr(known(yr))
        ElseIf dynCache.Exists(yr) Then
            dsID = CStr(dynCache(yr))
        Else
            Application.StatusBar = "NADAC Lookup  |  Discovering " & yr & " dataset..."
            DoEvents : PauseMs RATE_MS
            dsID = DiscoverDatasetID(yr)
            If dsID <> "" Then
                dynCache(yr) = dsID
                LogLine "  Discovered " & yr & " dataset: " & dsID
            Else
                missYears = missYears & yr & ", "
                cnt = yearNDCs(yr).Count
                batchsDone = batchsDone + Int((cnt + BATCH_SIZE - 1) / BATCH_SIZE)
                GoTo NextYear
            End If
        End If

        ' Resolve Distribution UUID -------------------------------------------------
        ' Always query the CMS metastore for the live distribution UUID so that
        ' dataset refreshes (which rotate the UUID) are handled automatically.
        ' The hardcoded KnownDistUUIDs table is a session-start fast-path only.
        Dim distUUID As String : distUUID = ""
        If knownDist.Exists(dsID) Then distUUID = CStr(knownDist(dsID))
        Application.StatusBar = "NADAC Lookup  |  Verifying " & yr & " dataset UUID..."
        DoEvents
        Dim freshUUID As String : freshUUID = FetchDistUUID(dsID)
        If freshUUID <> "" Then
            If freshUUID <> distUUID And distUUID <> "" Then
                LogLine "  UUID rotated for " & yr & ": " & distUUID & " -> " & freshUUID
            ElseIf distUUID = "" Then
                LogLine "  UUID discovered for " & yr & ": " & freshUUID
            End If
            distUUID = freshUUID
            knownDist(dsID) = freshUUID   ' update session cache
        End If

        ' Batch-query this year's NDCs
        Dim allNDCs() As Variant : allNDCs = yearNDCs(yr).Keys
        cnt = yearNDCs(yr).Count
        Dim maxDt As Date : maxDt = CDate(yearMaxDt(yr))
        LogLine "  Year=" & yr & "  dsID=" & dsID & "  distUUID=" & _
                IIf(distUUID <> "", distUUID, "(none)") & _
                "  NDCs=" & cnt & "  maxDt=" & Format(maxDt, "yyyy-mm-dd")

        Dim bi As Long
        For bi = 0 To Int((cnt + BATCH_SIZE - 1) / BATCH_SIZE) - 1
            Dim bStart As Long : bStart = bi * BATCH_SIZE
            Dim bEnd   As Long : bEnd   = bStart + BATCH_SIZE - 1
            If bEnd >= cnt Then bEnd = cnt - 1

            Dim bNDCs() As String : ReDim bNDCs(0 To bEnd - bStart)
            Dim j As Long
            For j = 0 To bEnd - bStart
                bNDCs(j) = CStr(allNDCs(bStart + j))
            Next j

            batchsDone = batchsDone + 1
            Dim pct As Long : pct = CLng(CDbl(batchsDone) / CDbl(totalBatches) * 100)
            Application.StatusBar = "NADAC Lookup  |  " & yr & "  Batch " & (bi + 1) & _
                                    "/" & Int((cnt + BATCH_SIZE - 1) / BATCH_SIZE) & _
                                    "  [" & ProgBar(batchsDone - 1, totalBatches, 22) & "]  " & pct & "%"
            DoEvents

            Dim resp As String : resp = ""
            If distUUID <> "" Then
                Dim body As String : body = BuildPostBody(bNDCs, maxDt)
                LogLine "  POST batch " & (bi + 1) & "  NDCs=" & Join(bNDCs, ",")
                resp = HttpPost(API_BASE & distUUID, body)
                ' Auto-retry once on 400: re-query metastore for a fresh UUID
                If resp = "" And m_LastHttpStatus = 400 Then
                    LogLine "  POST 400 -- re-fetching UUID for " & yr
                    Application.StatusBar = "NADAC Lookup  |  " & yr & " UUID refresh..."
                    DoEvents
                    Dim retryUUID As String : retryUUID = FetchDistUUID(dsID)
                    If retryUUID <> "" And retryUUID <> distUUID Then
                        LogLine "  Retrying with UUID: " & retryUUID
                        distUUID = retryUUID
                        knownDist(dsID) = retryUUID
                        PauseMs RATE_MS
                        resp = HttpPost(API_BASE & distUUID, body)
                    End If
                End If
                If resp <> "" Then
                    Dim arrStr As String : arrStr = ExtractResultsArray(resp)
                    If arrStr <> "" And arrStr <> "[]" Then
                        ParseIntoCache arrStr, apiCache
                        If RecordCount(arrStr) >= API_LIMIT Then truncWarn = True
                    End If
                End If
            Else
                ' Fallback: GET per NDC for years without a distribution UUID
                Dim k As Long
                For k = 0 To UBound(bNDCs)
                    LogLine "  GET fallback  ndc=" & bNDCs(k)
                    resp = HttpGet(BuildFallbackURL(dsID, bNDCs(k), maxDt))
                    If resp <> "" And resp <> "[]" Then ParseIntoCache resp, apiCache
                    If k < UBound(bNDCs) Then PauseMs RATE_MS
                Next k
            End If

            If bi < Int((cnt + BATCH_SIZE - 1) / BATCH_SIZE) - 1 Then PauseMs RATE_MS
        Next bi
        PauseMs RATE_MS

NextYear:
    Next yi

    ' -- Assign output columns (AFTER Pass 1 to avoid column-shift bug) --------
    colNPU = EnsureCol(ws, "NADAC")
    colED  = EnsureCol(ws, "NADAC Effective Date")

    ' -- Pass 2: write results -------------------------------------------------
    Application.StatusBar = "NADAC Lookup  |  Writing results..."
    DoEvents

    ' v2.3: Reset ALL data-row interior colours so re-runs start clean.
    ' Red is then re-applied only to rows that fail THIS run.
    For r = 2 To lastRow
        ws.Rows(r).Interior.ColorIndex = xlNone
    Next r

    Dim written  As Long : written  = 0
    Dim noRate   As Long : noRate   = 0
    Dim notFound As Long : notFound = 0
    Dim skipped  As Long : skipped  = 0

    For r = 2 To lastRow
        cn = rowCN(r) : yr = rowYR(r)

        If cn = "" Or yr = "" Then
            ws.Cells(r, colNPU).Value = "Invalid date/NDC"
            skipped = skipped + 1
            GoTo WriteNext
        End If

        If Len(missYears) > 0 And InStr(missYears, yr & ",") > 0 Then
            ws.Cells(r, colNPU).Value = "Dataset not found (" & yr & ")"
            skipped = skipped + 1
            GoTo WriteNext
        End If

        On Error Resume Next
        Dim dispDt As Date : dispDt = CDate(ws.Cells(r, dateCol).Value)
        If Err.Number <> 0 Then
            ws.Cells(r, colNPU).Value = "Invalid date"
            skipped = skipped + 1
            Err.Clear : On Error GoTo ErrHandler
            GoTo WriteNext
        End If
        Err.Clear : On Error GoTo ErrHandler

        If apiCache.Exists(cn) Then
            Dim best As String
            best = BestRecord(CStr(apiCache(cn)), dispDt)
            If best <> "" Then
                Dim fNPU  As String : fNPU  = JField(best, "nadac_per_unit")
                Dim fPU   As String : fPU   = JField(best, "per_unit_indicator")
                Dim fDate As String : fDate = JField(best, "effective_date")
                If fNPU <> "" Then
                    ws.Cells(r, colNPU).Value        = CDbl(fNPU)
                    ws.Cells(r, colNPU).NumberFormat = "$#,##0.0000"
                    If Len(fDate) >= 10 Then
                        ws.Cells(r, colED).Value        = CDate(Left(fDate, 10))
                        ws.Cells(r, colED).NumberFormat = "mm/dd/yyyy"
                    End If
                    written = written + 1
                Else
                    ws.Cells(r, colNPU).Value = "Parse error"
                    noRate = noRate + 1
                End If
            Else
                ws.Cells(r, colNPU).Value = "No rate for date"
                noRate = noRate + 1
            End If
        Else
            ws.Cells(r, colNPU).Value  = "NDC not in NADAC"
            ws.Rows(r).Interior.Color  = RGB(255, 199, 199)   ' salmon red
            notFound = notFound + 1
        End If

WriteNext:
    Next r

    Application.ScreenUpdating = True
    Application.Calculation    = xlCalculationAutomatic
    Application.StatusBar      = False

    LogLine "  Results: written=" & written & " noRate=" & noRate & _
            " notFound=" & notFound & " skipped=" & skipped

    ' -- Summary ---------------------------------------------------------------
    Dim msg As String
    msg = "NADAC Lookup complete." & vbCrLf & vbCrLf & _
          "   Rows updated               :  " & written  & vbCrLf & _
          "   No matching rate for date  :  " & noRate   & vbCrLf & _
          "   NDC not in NADAC           :  " & notFound & vbCrLf & _
          "   Skipped (invalid/missing)  :  " & skipped
    If Len(missYears) > 0 Then
        msg = msg & vbCrLf & vbCrLf & _
              "Warning: could not locate NADAC dataset(s) for year(s):" & vbCrLf & _
              "   " & Left(missYears, Len(missYears) - 2) & vbCrLf & _
              "Use Pharmacy Tools > Manage Datasets to add them."
    End If
    If truncWarn Then
        msg = msg & vbCrLf & vbCrLf & _
              "Note: one or more batches hit the row limit (" & API_LIMIT & ")." & vbCrLf & _
              "Some older rates may be incomplete."
    End If
    MsgBox msg, vbInformation, "NADAC Lookup"
    Exit Sub

ErrHandler:
    Application.ScreenUpdating = True
    Application.Calculation    = xlCalculationAutomatic
    Application.StatusBar      = False
    LogLine "ERROR " & Err.Number & ": " & Err.Description
    MsgBox "Error " & Err.Number & ": " & Err.Description, vbCritical, "NADAC Lookup"
End Sub

' --- About screen -------------------------------------------------------------
Public Sub ShowAbout()
    Dim nl As String : nl = vbCrLf
    Dim s  As String
    s = "  Pharmacy Tools  --  NADAC Lookup"                                  & nl
    s = s & "  Version " & ADDIN_VERSION & "   Build " & ADDIN_BUILD_DATE     & nl
    s = s & "  " & String(46, "-")                                             & nl & nl
    s = s & "  Tools in this version:"                                         & nl & nl
    s = s & "    NADAC Lookup"                                                 & nl
    s = s & "    Queries the CMS Medicaid NADAC API and"                       & nl
    s = s & "    populates NADAC and NADAC Effective Date for each"             & nl
    s = s & "    NADAC Effective Date.  Rows are matched to"                   & nl
    s = s & "    the correct annual dataset by dispense-date"                  & nl
    s = s & "    year.  Supports 2013 - present."                              & nl & nl
    s = s & "    Manage Datasets"                                              & nl
    s = s & "    View, add, and edit annual NADAC dataset"                     & nl
    s = s & "    entries.  Built-in years (2013-2025) are"                     & nl
    s = s & "    shown and can be overridden if CMS changes"                   & nl
    s = s & "    a UUID.  Custom years can be added or"                        & nl
    s = s & "    deleted without reinstalling."                                & nl & nl
    s = s & "  " & String(46, "-")                                             & nl & nl
    s = s & "  Developed by:"                                                  & nl
    s = s & "    Jeremiah Maudlin"                                             & nl
    s = s & "    Omega Integrated Management"                                  & nl
    s = s & "    jmaudlin@omegaim.com"                                         & nl
    s = s & "    +1 (641) 660-6367"                                            & nl & nl
    s = s & "  Data source:  data.medicaid.gov"                                & nl
    s = s & "  Log file:     %USERPROFILE%\Desktop\NADAC_Lookup_Log.txt"
    MsgBox s, vbInformation, "About Pharmacy Tools"
End Sub

' --- Discover dataset ID for a year not in the hardcoded table ----------------
Private Function DiscoverDatasetID(yr As String) As String
    Dim title As String : title = "NADAC (National Average Drug Acquisition Cost) " & yr
    Dim url   As String : url   = SEARCH_BASE & "?fulltext=" & UrlEnc(title) & "&page-size=10"
    Dim resp  As String : resp  = HttpGet(url)
    If resp = "" Then DiscoverDatasetID = "" : Exit Function

    Dim pos  As Long   : pos  = 1
    Dim tTag As String : tTag = """title"":"""
    Dim iTag As String : iTag = """identifier"":"""
    Do While pos <= Len(resp)
        Dim tp As Long : tp = InStr(pos, resp, tTag, vbBinaryCompare)
        If tp = 0 Then Exit Do
        Dim ts As Long : ts = tp + Len(tTag)
        Dim te As Long : te = ts
        Do While te <= Len(resp) : If Mid(resp,te,1) = """" Then Exit Do : te=te+1 : Loop
        If LCase(Trim(Mid(resp,ts,te-ts))) = LCase(Trim(title)) Then
            Dim ip As Long : ip = InStr(tp, resp, iTag, vbBinaryCompare)
            If ip > 0 Then
                Dim is2 As Long : is2 = ip + Len(iTag)
                Dim ie  As Long : ie  = is2
                Do While ie <= Len(resp) : If Mid(resp,ie,1)="""" Then Exit Do : ie=ie+1 : Loop
                DiscoverDatasetID = Mid(resp, is2, ie - is2) : Exit Function
            End If
        End If
        pos = te + 1
    Loop
    DiscoverDatasetID = ""
End Function

' --- Build JSON POST body for one batch ---------------------------------------
Private Function BuildPostBody(ndcs() As String, maxDt As Date) As String
    Dim conds As String : conds = ""
    Dim i As Long
    For i = 0 To UBound(ndcs)
        If i > 0 Then conds = conds & ","
        conds = conds & "{""property"":""ndc"",""value"":""" & ndcs(i) & """,""operator"":""=""}"
    Next i
    Dim dtS As String : dtS = Format(maxDt, "yyyy-mm-dd") & "T23:59:59"
    BuildPostBody = "{" & _
        """conditions"":[" & _
            "{""groupOperator"":""or"",""conditions"":[" & conds & "]}," & _
            "{""property"":""effective_date"",""value"":""" & dtS & """,""operator"":""<=""}" & _
        "]," & _
        """limit"":" & API_LIMIT & "," & _
        """sort"":[{""property"":""ndc""},{""property"":""effective_date"",""order"":""desc""}]" & _
        "}"
End Function

' --- Extract the inner array from {"results":[...]} --------------------------
Private Function ExtractResultsArray(resp As String) As String
    Dim tag As String : tag = """results"":["
    Dim p   As Long   : p   = InStr(1, resp, tag, vbBinaryCompare)
    If p = 0 Then
        If Left(Trim(resp), 1) = "[" Then ExtractResultsArray = resp Else ExtractResultsArray = ""
        Exit Function
    End If
    Dim s   As Long : s   = p + Len(tag) - 1   ' points to [
    Dim dep As Long : dep = 0
    Dim pos As Long
    For pos = s To Len(resp)
        Select Case Mid(resp, pos, 1)
            Case "[" : dep = dep + 1
            Case "]"
                dep = dep - 1
                If dep = 0 Then
                    ExtractResultsArray = Mid(resp, s, pos - s + 1)
                    Exit Function
                End If
        End Select
    Next pos
    ExtractResultsArray = ""
End Function

' --- Fallback GET URL (years without a confirmed distribution UUID) ------------
Private Function BuildFallbackURL(dsID As String, ndc As String, maxDt As Date) As String
    Dim w As String
    w = "ndc='" & ndc & "' and effective_date<='" & Format(maxDt, "yyyy-mm-dd") & "T23:59:59'"
    BuildFallbackURL = "https://data.medicaid.gov/resource/" & dsID & ".json" & _
                       "?$where=" & UrlEnc(w) & _
                       "&$limit=" & API_LIMIT & _
                       "&$order=" & UrlEnc("effective_date DESC")
End Function

' --- HTTP POST via WinHttp ----------------------------------------------------
Private Function HttpPost(url As String, body As String) As String
    On Error GoTo Fail
    m_LastHttpStatus = 0
    Dim h As Object : Set h = CreateObject("WinHttp.WinHttpRequest.5.1")
    h.Open "POST", url, False
    h.SetTimeouts 15000, 15000, 90000, 90000
    h.SetRequestHeader "Content-Type", "application/json"
    h.SetRequestHeader "Accept", "application/json"
    h.SetRequestHeader "User-Agent", "Excel-NADAC-Addin/" & ADDIN_VERSION
    h.Send body
    m_LastHttpStatus = h.Status
    LogLine "  POST " & h.Status & " " & h.StatusText & " (" & Len(h.ResponseText) & " bytes)"
    If h.Status = 200 Then HttpPost = h.ResponseText Else HttpPost = ""
    Exit Function
Fail: LogLine "  POST ERROR: " & Err.Number & " " & Err.Description : HttpPost = ""
End Function


' --- Auto-discover the current distribution UUID from the CMS DKAN metastore --
' Called once per year before batching so UUID rotations are handled silently.
' Returns "" on any error (caller falls back to hardcoded or GET).
Private Function FetchDistUUID(dsID As String) As String
    On Error GoTo Fail
    Dim url As String
    url = "https://data.medicaid.gov/api/1/metastore/schemas/dataset/items/" & _
          dsID & "?show-reference-ids=true"
    Dim resp As String : resp = HttpGet(url)
    If resp = "" Then FetchDistUUID = "" : Exit Function
    ' Response contains: "distribution":[{"identifier":"UUID","data":{...}}]
    ' Extract the first distribution identifier value.
    Dim tag As String : tag = """distribution"":[{""identifier"":"""
    Dim p   As Long   : p   = InStr(1, resp, tag, vbBinaryCompare)
    If p = 0 Then FetchDistUUID = "" : Exit Function
    Dim s As Long : s = p + Len(tag)
    Dim e As Long : e = s
    Do While e <= Len(resp)
        If Mid(resp, e, 1) = """" Then Exit Do
        e = e + 1
    Loop
    Dim uuid As String : uuid = Mid(resp, s, e - s)
    ' Sanity check: valid UUIDs are 36 chars with hyphens
    If Len(uuid) = 36 And InStr(uuid, "-") > 0 Then
        FetchDistUUID = uuid
    Else
        FetchDistUUID = ""
    End If
    Exit Function
Fail:
    FetchDistUUID = ""
End Function

' --- HTTP GET via WinHttp -----------------------------------------------------
Private Function HttpGet(url As String) As String
    On Error GoTo Fail
    Dim h As Object : Set h = CreateObject("WinHttp.WinHttpRequest.5.1")
    h.Open "GET", url, False
    h.SetTimeouts 15000, 15000, 90000, 90000
    h.SetRequestHeader "Accept", "application/json"
    h.SetRequestHeader "User-Agent", "Excel-NADAC-Addin/" & ADDIN_VERSION
    h.Send
    LogLine "  GET " & h.Status & " " & h.StatusText & " (" & Len(h.ResponseText) & " bytes)"
    If h.Status = 200 Then HttpGet = h.ResponseText Else HttpGet = ""
    Exit Function
Fail: LogLine "  GET ERROR: " & Err.Number & " " & Err.Description : HttpGet = ""
End Function

' --- Parse JSON array into NDC-keyed cache ------------------------------------
Private Sub ParseIntoCache(jsonArr As String, cache As Object)
    Dim pos As Long : pos = 1
    Dim dep As Long : dep = 0
    Dim oS  As Long : oS  = 0
    Do While pos <= Len(jsonArr)
        Select Case Mid(jsonArr, pos, 1)
            Case "{"
                dep = dep + 1 : If dep = 1 Then oS = pos
            Case "}"
                If dep = 1 And oS > 0 Then
                    Dim obj As String : obj = Mid(jsonArr, oS, pos - oS + 1)
                    Dim k   As String : k   = JField(obj, "ndc")
                    If k <> "" Then
                        If cache.Exists(k) Then
                            Dim ex As String : ex = CStr(cache(k))
                            cache(k) = Left(ex, Len(ex) - 1) & "," & obj & "]"
                        Else
                            cache(k) = "[" & obj & "]"
                        End If
                    End If
                    oS = 0
                End If
                dep = dep - 1
        End Select
        pos = pos + 1
    Loop
End Sub

' --- Best record: latest effective_date <= dispDt ----------------------------
Private Function BestRecord(jsonArr As String, dispDt As Date) As String
    Dim pos As Long : pos = 1
    Dim dep As Long : dep = 0
    Dim oS  As Long : oS  = 0
    Do While pos <= Len(jsonArr)
        Select Case Mid(jsonArr, pos, 1)
            Case "{" : dep = dep + 1 : If dep = 1 Then oS = pos
            Case "}"
                If dep = 1 And oS > 0 Then
                    Dim obj As String : obj = Mid(jsonArr, oS, pos - oS + 1)
                    Dim es  As String : es  = JField(obj, "effective_date")
                    If Len(es) >= 10 Then
                        On Error Resume Next
                        Dim ed As Date : ed = CDate(Left(es, 10))
                        If Err.Number = 0 And ed <= dispDt Then
                            BestRecord = obj : Exit Function
                        End If
                        Err.Clear : On Error GoTo 0
                    End If
                    oS = 0
                End If
                dep = dep - 1
        End Select
        pos = pos + 1
    Loop
    BestRecord = ""
End Function

' --- Count top-level JSON objects --------------------------------------------
Private Function RecordCount(jsonArr As String) As Long
    Dim pos As Long : pos = 1
    Dim dep As Long : dep = 0
    Dim cnt As Long : cnt = 0
    Do While pos <= Len(jsonArr)
        Select Case Mid(jsonArr, pos, 1)
            Case "{" : dep = dep + 1
            Case "}" : If dep = 1 Then cnt = cnt + 1 : dep = dep - 1
        End Select
        pos = pos + 1
    Loop
    RecordCount = cnt
End Function

' --- Minimal JSON field extractor --------------------------------------------
Private Function JField(obj As String, fname As String) As String
    Dim p As Long
    p = InStr(1, obj, """" & fname & """", vbBinaryCompare)
    If p = 0 Then JField = "" : Exit Function
    p = p + Len(fname) + 2
    Dim c As String
    Do While p <= Len(obj)
        c = Mid(obj, p, 1)
        If c <> " " And c <> ":" Then Exit Do
        p = p + 1
    Loop
    If p > Len(obj) Then JField = "" : Exit Function
    c = Mid(obj, p, 1)
    If c = """" Then
        p = p + 1 : Dim sb As String : sb = ""
        Do While p <= Len(obj)
            c = Mid(obj, p, 1)
            If c = "\" Then p = p + 1 : sb = sb & Mid(obj, p, 1) : GoTo NextChar
            If c = """" Then Exit Do
            sb = sb & c
NextChar:   p = p + 1
        Loop
        JField = sb
    ElseIf c = "n" Then
        JField = ""
    Else
        Dim e As Long : e = p
        Do While e <= Len(obj)
            c = Mid(obj, e, 1)
            If c="," Or c="}" Or c="]" Or c=" " Or c=vbCr Or c=vbLf Then Exit Do
            e = e + 1
        Loop
        JField = Mid(obj, p, e - p)
    End If
End Function

' --- Registry helpers (Public for form access) --------------------------------
Public Function GetRegValue(regKey As String, valName As String) As String
    On Error Resume Next
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    GetRegValue = CStr(sh.RegRead(regKey & "\" & valName))
    If Err.Number <> 0 Then GetRegValue = ""
    Err.Clear
End Function

Public Sub SetRegValue(regKey As String, valName As String, val As String)
    On Error Resume Next
    Dim sh As Object : Set sh = CreateObject("WScript.Shell")
    sh.RegWrite regKey & "\" & valName, val, "REG_SZ"
End Sub

' --- Sort dictionary keys ascending ------------------------------------------
Public Function SortedKeys(d As Object) As String()
    Dim keys() As Variant : keys = d.Keys
    Dim n As Long : n = UBound(keys)
    Dim i As Long, j As Long, tmp As Variant
    For i = 0 To n-1 : For j = i+1 To n
        If CStr(keys(j)) < CStr(keys(i)) Then tmp=keys(i):keys(i)=keys(j):keys(j)=tmp
    Next j : Next i
    Dim result() As String : ReDim result(0 To n)
    For i = 0 To n : result(i) = CStr(keys(i)) : Next i
    SortedKeys = result
End Function

' --- Strip hyphens/spaces; zero-pad to 11 digits -----------------------------
Private Function CleanNDC(raw As String) As String
    Dim s As String : s = Replace(Replace(Trim(raw), "-", ""), " ", "")
    Do While Len(s) < 11 And Len(s) > 0 : s = "0" & s : Loop
    CleanNDC = s
End Function

' --- URL-encode ---------------------------------------------------------------
Private Function UrlEnc(s As String) As String
    Dim i As Long, c As String, result As String : result = ""
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        Select Case c
            Case "A" To "Z","a" To "z","0" To "9","-","_",".","~" : result = result & c
            Case Else : result = result & "%" & Right("0" & Hex(Asc(c)), 2)
        End Select
    Next i
    UrlEnc = result
End Function

' --- Column number to letter(s) ----------------------------------------------
Private Function ColLtr(col As Long) As String
    Dim n As Long : n = col : Dim r2 As Long : Dim res As String : res = ""
    Do : r2 = ((n-1) Mod 26) : res = Chr(65+r2) & res : n = (n-r2-1)\26 : Loop While n > 0
    ColLtr = res
End Function

' --- Find column by header (case-insensitive) --------------------------------
Private Function FindHeaderCol(ws As Worksheet, candidates As Variant) As Long
    Dim lc As Long : lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    Dim c As Long : Dim h As Variant
    For c = 1 To lc
        Dim v As String : v = LCase(Trim(CStr(ws.Cells(1, c).Value)))
        For Each h In candidates
            If v = LCase(Trim(CStr(h))) Then FindHeaderCol = c : Exit Function
        Next h
    Next c
    FindHeaderCol = 0
End Function

' --- Find or append a bold-header column -------------------------------------
Private Function EnsureCol(ws As Worksheet, hdr As String) As Long
    Dim f As Long : f = FindHeaderCol(ws, Array(hdr))
    If f > 0 Then EnsureCol = f : Exit Function
    Dim lc As Long : lc = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column + 1
    ws.Cells(1, lc).Value = hdr : ws.Cells(1, lc).Font.Bold = True
    EnsureCol = lc
End Function

' --- ASCII block progress bar ------------------------------------------------
Private Function ProgBar(curr As Long, tot As Long, w As Long) As String
    If tot = 0 Then ProgBar = String(w, "-") : Exit Function
    Dim f As Long : f = CLng(CDbl(curr) / CDbl(tot) * CDbl(w))
    If f > w Then f = w
    ProgBar = String(f, "#") & String(w - f, ".")
End Function

' --- Millisecond pause -------------------------------------------------------
Private Sub PauseMs(ms As Long)
    Dim t0 As Double : t0 = Timer
    Dim wt As Double : wt = ms / 1000#
    Dim el As Double
    Do : DoEvents : el = Timer - t0
        If el < 0 Then el = el + 86400
        If el >= wt Then Exit Do
    Loop
End Sub

' --- Log file on Desktop (5 MB rotation) -------------------------------------

Public Sub LogLine(msg As String)
    On Error Resume Next
    Dim path As String : path = Environ("USERPROFILE") & "\Desktop\NADAC_Lookup_Log.txt"
    Dim fn As Integer  : fn   = FreeFile
    Open path For Append As #fn
    Print #fn, Format(Now, "yyyy-mm-dd hh:mm:ss") & "  " & msg
    Close #fn
End Sub

Public Sub LogRotate()
    On Error Resume Next
    Dim path As String : path = Environ("USERPROFILE") & "\Desktop\NADAC_Lookup_Log.txt"
    If Len(Dir(path)) = 0 Then Exit Sub
    Dim fso As Object : Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.GetFile(path).Size > 5242880 Then    ' 5 MB max
        Dim bak As String : bak = Environ("USERPROFILE") & "\Desktop\NADAC_Lookup_Log_bak.txt"
        If Len(Dir(bak)) > 0 Then Kill bak
        Name path As bak
    End If
End Sub
' --- Dataset Manager (MsgBox/InputBox -- no UserForm needed) ------------------
Public Sub ShowDatasetManager()
    Dim nl As String : nl = vbCrLf
    Dim keepGoing As Boolean : keepGoing = True
    Do While keepGoing
        Dim known     As Object : Set known     = KnownYearIDs()
        Dim knownDist As Object : Set knownDist = KnownDistUUIDs()
        Dim raw       As String : raw = GetRegValue(REG_KEY, "UserDatasets")
        Dim uDict     As Object : Set uDict = CreateObject("Scripting.Dictionary")
        uDict.CompareMode = 1
        If raw <> "" Then
            Dim ents() As String : ents = Split(raw, "|")
            Dim i As Long
            For i = 0 To UBound(ents)
                Dim e As String : e = Trim(ents(i))
                If e <> "" Then
                    Dim ep() As String : ep = Split(e, ":")
                    If UBound(ep) >= 1 Then
                        Dim uy As String : uy = Trim(ep(0))
                        Dim uv As String : uv = Trim(ep(1)) & ":" & IIf(UBound(ep) >= 2, Trim(ep(2)), "")
                        If uy <> "" Then uDict(uy) = uv
                    End If
                End If
            Next i
        End If
        Dim yStr As String : yStr = ""
        Dim yr As Variant
        For Each yr In known.Keys : yStr = yStr & CStr(yr) & "|" : Next yr
        For Each yr In uDict.Keys
            If Not known.Exists(CStr(yr)) And InStr(yStr, CStr(yr) & "|") = 0 Then
                yStr = yStr & CStr(yr) & "|"
            End If
        Next yr
        If yStr = "" Then Exit Do
        If Right(yStr, 1) = "|" Then yStr = Left(yStr, Len(yStr) - 1)
        Dim ya() As String : ya = Split(yStr, "|")
        Dim ii As Long, jj As Long, tmp As String
        For ii = 0 To UBound(ya) - 1
            For jj = ii + 1 To UBound(ya)
                If CStr(ya(jj)) > CStr(ya(ii)) Then
                    tmp = ya(ii) : ya(ii) = ya(jj) : ya(jj) = tmp
                End If
            Next jj
        Next ii
        Dim disp As String
        disp = "MANAGE NADAC DATASETS  --  Pharmacy Tools" & nl & String(60, "-") & nl & nl
        Dim rr As Long
        For rr = 0 To UBound(ya)
            Dim yrS As String : yrS = Trim(ya(rr))
            If yrS = "" Then GoTo SkipD
            Dim dsID As String : dsID = ""
            Dim distID As String : distID = ""
            Dim eType As String : eType = ""
            If known.Exists(yrS) Then
                If uDict.Exists(yrS) Then
                    eType = "Built-in*"
                    Dim op() As String : op = Split(CStr(uDict(yrS)), ":")
                    dsID = Trim(op(0))
                    If UBound(op) >= 1 Then distID = Trim(op(1))
                    If dsID = "" Then dsID = CStr(known(yrS))
                Else
                    eType = "Built-in"
                    dsID = CStr(known(yrS))
                    If knownDist.Exists(dsID) Then distID = CStr(knownDist(dsID))
                End If
            Else
                eType = "Custom"
                Dim cp() As String : cp = Split(CStr(uDict(yrS)), ":")
                dsID = Trim(cp(0))
                If UBound(cp) >= 1 Then distID = Trim(cp(1))
            End If
            disp = disp & "  " & yrS & "  [" & eType & "]" & nl
            disp = disp & "    Dataset : " & dsID & nl
            disp = disp & "    Dist    : " & IIf(distID <> "", distID, "(none -- GET fallback)") & nl & nl
SkipD:
        Next rr
        disp = disp & String(60, "-") & nl & nl
        disp = disp & "  YES    = Add a new year" & nl
        disp = disp & "  NO     = Edit or Delete/Reset an existing year" & nl
        disp = disp & "  CANCEL = Close"
        Dim ch As VbMsgBoxResult
        ch = MsgBox(disp, vbYesNoCancel + vbInformation, "Manage NADAC Datasets")
        If ch = vbCancel Then
            keepGoing = False
        ElseIf ch = vbYes Then
            DSAddWizard
        ElseIf ch = vbNo Then
            DSEditOrDelete known, knownDist, uDict
        End If
    Loop
End Sub

Private Sub DSEditOrDelete(known As Object, knownDist As Object, uDict As Object)
    Dim nl As String : nl = vbCrLf
    Dim tyr As String
    tyr = Trim(InputBox("Enter the year to edit or delete:", "Edit / Delete Year"))
    If tyr = "" Then Exit Sub
    Dim dsID As String : dsID = ""
    Dim distID As String : distID = ""
    Dim eType As String : eType = ""
    If known.Exists(tyr) Then
        If uDict.Exists(tyr) Then
            eType = "Built-in*"
            Dim op2() As String : op2 = Split(CStr(uDict(tyr)), ":")
            dsID = Trim(op2(0))
            If dsID = "" Then dsID = CStr(known(tyr))
            If UBound(op2) >= 1 Then distID = Trim(op2(1))
        Else
            eType = "Built-in"
            dsID = CStr(known(tyr))
            If knownDist.Exists(dsID) Then distID = CStr(knownDist(dsID))
        End If
    ElseIf uDict.Exists(tyr) Then
        eType = "Custom"
        Dim cp2() As String : cp2 = Split(CStr(uDict(tyr)), ":")
        dsID = Trim(cp2(0))
        If UBound(cp2) >= 1 Then distID = Trim(cp2(1))
    Else
        MsgBox "Year " & tyr & " not found.", vbExclamation, "Not Found"
        Exit Sub
    End If
    Dim info As String
    info = "Year " & tyr & "  [" & eType & "]" & nl & nl & _
           "Dataset ID : " & dsID & nl & _
           "Dist UUID  : " & IIf(distID <> "", distID, "(none)") & nl & nl
    If eType = "Built-in" Then
        If MsgBox(info & "YES = Edit (stores a registry override)" & nl & "CANCEL = Back", _
                  vbYesCancel + vbInformation, "Edit " & tyr) = vbYes Then
            DSEditYear tyr, dsID, distID
        End If
    ElseIf eType = "Built-in*" Then
        Dim r2 As VbMsgBoxResult
        r2 = MsgBox(info & "YES = Edit the override" & nl & _
                    "NO  = Reset to factory defaults" & nl & "CANCEL = Back", _
                    vbYesNoCancel + vbInformation, "Modify " & tyr)
        If r2 = vbYes Then DSEditYear tyr, dsID, distID
        If r2 = vbNo  Then DSResetYear tyr
    Else
        Dim r3 As VbMsgBoxResult
        r3 = MsgBox(info & "YES = Edit UUIDs" & nl & _
                    "NO  = Delete this custom year" & nl & "CANCEL = Back", _
                    vbYesNoCancel + vbInformation, "Modify " & tyr)
        If r3 = vbYes Then DSEditYear tyr, dsID, distID
        If r3 = vbNo  Then DSDeleteYear tyr
    End If
End Sub

Private Sub DSEditYear(yr As String, curDs As String, curDist As String)
    Dim newDs As String
    newDs = Trim(InputBox("Dataset ID for " & yr & ":" & vbCrLf & vbCrLf & _
                          "Current: " & curDs & vbCrLf & "(Blank = cancel)", "Edit Dataset ID", curDs))
    If newDs = "" Then Exit Sub
    Dim newDt As String
    newDt = Trim(InputBox("Distribution UUID for " & yr & ":" & vbCrLf & vbCrLf & _
                          "Current: " & IIf(curDist <> "", curDist, "(none)") & vbCrLf & _
                          "Leave blank for GET fallback.", "Edit Distribution UUID", curDist))
    Dim ex As String : ex = GetRegValue(REG_KEY, "UserDatasets")
    ex = DSRemoveYear(ex, yr)
    Dim ne As String : ne = yr & ":" & newDs & ":" & newDt
    SetRegValue REG_KEY, "UserDatasets", IIf(ex = "", ne, ex & "|" & ne)
    MsgBox "Year " & yr & " saved.", vbInformation, "Saved"
End Sub

Private Sub DSResetYear(yr As String)
    If MsgBox("Reset " & yr & " to factory defaults?", vbYesNo + vbQuestion, "Reset") = vbNo Then Exit Sub
    Dim ex As String : ex = GetRegValue(REG_KEY, "UserDatasets")
    SetRegValue REG_KEY, "UserDatasets", DSRemoveYear(ex, yr)
    MsgBox "Year " & yr & " reset to defaults.", vbInformation, "Reset"
End Sub

Private Sub DSDeleteYear(yr As String)
    If MsgBox("Delete custom year " & yr & "?  Cannot be undone.", _
              vbYesNo + vbQuestion, "Delete") = vbNo Then Exit Sub
    Dim ex As String : ex = GetRegValue(REG_KEY, "UserDatasets")
    SetRegValue REG_KEY, "UserDatasets", DSRemoveYear(ex, yr)
    MsgBox "Year " & yr & " deleted.", vbInformation, "Deleted"
End Sub

Private Function DSRemoveYear(raw As String, yr As String) As String
    If raw = "" Then DSRemoveYear = "" : Exit Function
    Dim parts() As String : parts = Split(raw, "|")
    Dim result  As String : result = ""
    Dim i As Long
    For i = 0 To UBound(parts)
        Dim e As String : e = Trim(parts(i))
        If e <> "" Then
            Dim ep() As String : ep = Split(e, ":")
            If UBound(ep) >= 0 Then
                If LCase(Trim(ep(0))) <> LCase(Trim(yr)) Then
                    result = IIf(result = "", e, result & "|" & e)
                End If
            End If
        End If
    Next i
    DSRemoveYear = result
End Function

Private Sub DSAddWizard()
    Dim nl As String : nl = vbCrLf
    MsgBox "ADD DATASET WIZARD  |  Step 1 of 4" & nl & String(40, "-") & nl & nl & _
           "Adds a new annual NADAC dataset entry." & nl & nl & _
           "You need two UUIDs from data.medicaid.gov:" & nl & nl & _
           "  1. Dataset ID   (from the dataset page URL)" & nl & _
           "  2. Distribution UUID  (from the API tab)", vbInformation, "Add Dataset -- 1 of 4"
    Dim newYear As String
    newYear = Trim(InputBox("ADD DATASET WIZARD  |  Step 2 of 4" & nl & _
                            "Enter the 4-digit year:", "Add Dataset -- 2 of 4"))
    If newYear = "" Then Exit Sub
    If Not IsNumeric(newYear) Or Len(newYear) <> 4 Then
        MsgBox "Please enter a valid 4-digit year.", vbExclamation, "Add Dataset" : Exit Sub
    End If
    Dim newDsID As String
    newDsID = Trim(InputBox("ADD DATASET WIZARD  |  Step 3 of 4" & nl & _
        "Dataset ID for " & newYear & nl & nl & _
        "1. Go to data.medicaid.gov" & nl & _
        "2. Search: NADAC ... " & newYear & nl & _
        "3. Copy the UUID from the URL  (.../dataset/XXXX)", "Add Dataset -- 3 of 4"))
    If newDsID = "" Then Exit Sub
    Dim newDist As String
    newDist = Trim(InputBox("ADD DATASET WIZARD  |  Step 4 of 4" & nl & _
        "Distribution UUID for " & newYear & nl & nl & _
        "On the dataset page:" & nl & _
        "1. Click the API tab" & nl & _
        "2. Expand  POST /api/1/datastore/query/{distributionId}" & nl & _
        "3. Click Try it out -- copy the UUID shown" & nl & nl & _
        "(Leave blank to use GET fallback.)", "Add Dataset -- 4 of 4"))
    If MsgBox("CONFIRM  --  Save new entry?" & nl & nl & _
              "Year            : " & newYear & nl & _
              "Dataset ID      : " & newDsID & nl & _
              "Distribution UUID: " & IIf(newDist <> "", newDist, "(none)"), _
              vbYesNo + vbQuestion, "Add Dataset -- Confirm") = vbNo Then Exit Sub
    Dim ex As String : ex = GetRegValue(REG_KEY, "UserDatasets")
    ex = DSRemoveYear(ex, newYear)
    Dim ne As String : ne = newYear & ":" & newDsID & ":" & newDist
    SetRegValue REG_KEY, "UserDatasets", IIf(ex = "", ne, ex & "|" & ne)
    MsgBox "Year " & newYear & " added.", vbInformation, "Saved"
End Sub
'@


$RibbonXml = @'
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui">
  <ribbon>
    <tabs>
      <tab id="tabPharmacyTools" label="Pharmacy Tools">
        <group id="grpNADAC" label="NADAC">
          <button id="btnNADACLookup"
                  label="NADAC Lookup"
                  image="rIdImg1"
                  size="large"
                  onAction="RunNADACLookup"
                  screentip="NADAC Lookup"
                  supertip="Query the CMS Medicaid NADAC API and populate NADAC and NADAC Effective Date for each row." />
        </group>
        <group id="grpTools" label="Tools">
          <button id="btnManageDatasets"
                  label="Manage Datasets"
                  imageMso="TableInsert"
                  size="normal"
                  onAction="RunManageDatasets"
                  screentip="Manage NADAC Datasets"
                  supertip="View, add, and edit annual NADAC dataset entries. Built-in years (2013-2025) can be overridden. Custom years can be added or deleted." />
          <button id="btnAbout"
                  label="About"
                  imageMso="Help"
                  size="normal"
                  onAction="RunAbout"
                  screentip="About Pharmacy Tools"
                  supertip="Version information, tool list, and support contact." />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
'@


# -----------------------------------------------------------------------------
#  STEP 1  --  Build XLAM via Excel COM
# -----------------------------------------------------------------------------
Write-Host '  Step 1/3 : Building XLAM ...' -ForegroundColor White

$Xl  = $null
$Wb  = $null
try {
    $Xl = New-Object -ComObject Excel.Application
    $Xl.Visible              = $false
    $Xl.DisplayAlerts        = $false
    $Xl.AutomationSecurity   = 1        # msoAutomationSecurityLow

    $Wb = $Xl.Workbooks.Add()

    # -- Main VBA module -------------------------------------------------------
    $Mod = $Wb.VBProject.VBComponents.Add(1)    # 1 = vbext_ct_StdModule
    $Mod.Name = 'modNADAC'
    $Mod.CodeModule.AddFromString($VbaCode)
    Write-Host '    modNADAC added.' -ForegroundColor DarkGray

    # No UserForm -- Dataset Manager uses MsgBox/InputBox in modNADAC

    # -- Save as XLAM ----------------------------------------------------------
    $TmpPath = Join-Path $env:TEMP 'NADACLookup_tmp.xlam'
    if (Test-Path $TmpPath) { Remove-Item $TmpPath -Force }
    $Wb.SaveAs($TmpPath, 55)    # 55 = xlOpenXMLAddIn (.xlam)
    Write-Host '    Saved to temp XLAM.' -ForegroundColor DarkGray

} finally {
    if ($Wb)  { try { $Wb.Close($false) }  catch {} }
    if ($Xl)  { try { $Xl.Quit() }         catch {} }
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($Xl) | Out-Null
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
Write-Host '  Step 1/3 : XLAM built.' -ForegroundColor Green


# -----------------------------------------------------------------------------
#  STEP 2  --  Inject ribbon XML and icon into XLAM zip
# -----------------------------------------------------------------------------
Write-Host '  Step 2/3 : Injecting ribbon and icon ...' -ForegroundColor White

Add-Type -AssemblyName System.IO.Compression.FileSystem

$ZipPath = $TmpPath + '.zip'
Copy-Item $TmpPath $ZipPath -Force

$Zip = [System.IO.Compression.ZipFile]::Open($ZipPath, 'Update')

# Helper: remove existing entry if present
function Remove-ZipEntry($zip, $entryName) {
    $existing = $zip.Entries | Where-Object { $_.FullName -eq $entryName }
    if ($existing) { $existing.Delete() }
}

# -- customUI/customUI14.xml ---------------------------------------------------
Remove-ZipEntry $Zip 'customUI/customUI14.xml'
$cuEntry = $Zip.CreateEntry('customUI/customUI14.xml')
$sw = [System.IO.StreamWriter]::new($cuEntry.Open(), [System.Text.Encoding]::UTF8)
$sw.Write($RibbonXml)
$sw.Close()

# -- image injected above into customUI/images/ -----------------------------
# -- customUI/images/image1.png  (pharmacy bottle icon) ----------------------
Remove-ZipEntry $Zip 'customUI/images/image1.png'
$imgBytes = [Convert]::FromBase64String($IconB64)
$imgEntry = $Zip.CreateEntry('customUI/images/image1.png')
$imgStream = $imgEntry.Open()
$imgStream.Write($imgBytes, 0, $imgBytes.Length)
$imgStream.Close()

# -- customUI/_rels/customUI14.xml.rels  (image relationship) -----------------
Remove-ZipEntry $Zip 'customUI/_rels/customUI14.xml.rels'
$cuRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rIdImg1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="images/image1.png"/></Relationships>'
$cuRelsE = $Zip.CreateEntry('customUI/_rels/customUI14.xml.rels')
$swCR = [System.IO.StreamWriter]::new($cuRelsE.Open(), [System.Text.Encoding]::UTF8)
$swCR.Write($cuRels)
$swCR.Close()

# -- _rels/.rels  (add customUI relationship) ----------------------------------
$relsEntry = $Zip.Entries | Where-Object { $_.FullName -eq '_rels/.rels' }
$relsXml   = ''
if ($relsEntry) {
    $sr      = [System.IO.StreamReader]::new($relsEntry.Open())
    $relsXml = $sr.ReadToEnd()
    $sr.Close()
    $relsEntry.Delete()
}
if ($relsXml -eq '') {
    $relsXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>'
}
# Remove any stale customUI relationship
$relsXml = $relsXml -replace '<Relationship[^/]*/>\s*(?=<)', {
    $m = $args[0].Value
    if ($m -match 'customUI') { '' } else { $m }
}
$relsXml = $relsXml -replace '(<Relationship[^>]*customUI[^>]*/>\s*)', ''
# Inject fresh one before </Relationships>
$cuRel = '<Relationship Id="rId100" Type="http://schemas.microsoft.com/office/2007/relationships/ui/extensibility" Target="customUI/customUI14.xml"/>'
$relsXml = $relsXml -replace '</Relationships>', "$cuRel</Relationships>"

$newRels = $Zip.CreateEntry('_rels/.rels')
$sw2 = [System.IO.StreamWriter]::new($newRels.Open(), [System.Text.Encoding]::UTF8)
$sw2.Write($relsXml)
$sw2.Close()

# (no workbook.xml.rels image entry needed -- icon is in customUI/images/)

# -- [Content_Types].xml  (register customUI and png types) --------------------
$ctEntry = $Zip.Entries | Where-Object { $_.FullName -eq '[Content_Types].xml' }
$ctXml   = ''
if ($ctEntry) {
    $sr5  = [System.IO.StreamReader]::new($ctEntry.Open())
    $ctXml = $sr5.ReadToEnd()
    $sr5.Close()
    $ctEntry.Delete()
}
if ($ctXml -notmatch 'customUI14\.xml') {
    $cuType = '<Override PartName="/customUI/customUI14.xml" ContentType="application/xml"/>'
    $ctXml  = $ctXml -replace '</Types>', "$cuType</Types>"
}
if ($ctXml -notmatch 'Extension="png"') {
    $pngType = '<Default Extension="png" ContentType="image/png"/>'
    $ctXml   = $ctXml -replace '</Types>', "$pngType</Types>"
}
$newCt = $Zip.CreateEntry('[Content_Types].xml')
$sw6 = [System.IO.StreamWriter]::new($newCt.Open(), [System.Text.Encoding]::UTF8)
$sw6.Write($ctXml)
$sw6.Close()

$Zip.Dispose()

# Rename back to .xlam and move to install dir
if (Test-Path $AddinPath) { Remove-Item $AddinPath -Force }
Move-Item $ZipPath $AddinPath -Force
if (Test-Path $TmpPath) { Remove-Item $TmpPath -Force }

Write-Host '  Step 2/3 : Ribbon and icon injected.' -ForegroundColor Green


# -----------------------------------------------------------------------------
#  STEP 3  --  Register add-in via Excel registry (OPEN key)
# -----------------------------------------------------------------------------
Write-Host '  Step 3/3 : Registering add-in ...' -ForegroundColor White

$XlAddinsKey = "HKCU:\Software\Microsoft\Office\$OfficeVer\Excel\Add-in Manager"
if (-not (Test-Path $XlAddinsKey)) {
    New-Item -Path $XlAddinsKey -Force | Out-Null
}
Set-ItemProperty -Path $XlAddinsKey -Name $AddinPath -Value '' -Type String -Force

# Write the OPEN key so Excel loads it automatically
$OpenKey = "HKCU:\Software\Microsoft\Office\$OfficeVer\Excel\Options"
if (-not (Test-Path $OpenKey)) {
    New-Item -Path $OpenKey -Force | Out-Null
}
# Find or assign an OPEN slot (OPEN, OPEN1, OPEN2, ...)
$OpenProps  = Get-ItemProperty -Path $OpenKey -ErrorAction SilentlyContinue
$ExistSlot  = $null
$NextSlot   = 'OPEN'
$slotNum    = 0
do {
    $slotName = if ($slotNum -eq 0) { 'OPEN' } else { "OPEN$slotNum" }
    $slotVal  = $OpenProps.$slotName
    if ($slotVal -eq $null) {
        if ($ExistSlot -eq $null) { $NextSlot = $slotName }
        break
    }
    if ($slotVal -like "*NADACLookup.xlam*") {
        $ExistSlot = $slotName
        break
    }
    $slotNum++
} while ($slotNum -le 20)

$RegSlot = if ($ExistSlot) { $ExistSlot } else { $NextSlot }
Set-ItemProperty -Path $OpenKey -Name $RegSlot -Value "/R `"$AddinPath`"" -Type String -Force
Write-Host "    Registered in OPEN slot: $RegSlot" -ForegroundColor DarkGray

Write-Host '  Step 3/3 : Add-in registered.' -ForegroundColor Green


# -----------------------------------------------------------------------------
#  RESTORE AccessVBOM
# -----------------------------------------------------------------------------
try {
    if ($OldVBOM -ne $null) {
        Set-ItemProperty -Path $VBOMPath -Name $VBOMName -Value $OldVBOM -Type DWord -Force
    } else {
        Remove-ItemProperty -Path $VBOMPath -Name $VBOMName -ErrorAction SilentlyContinue
    }
} catch {}

# -----------------------------------------------------------------------------
#  DONE
# -----------------------------------------------------------------------------
Write-Host ''
Write-Host '  =====================================================' -ForegroundColor Green
Write-Host '   Pharmacy Tools Add-In installed successfully!'        -ForegroundColor Green
Write-Host '   Version 2.3.3'                                        -ForegroundColor Green
Write-Host '  =====================================================' -ForegroundColor Green
Write-Host ''
Write-Host '  What is new in v2.3.3:' -ForegroundColor White
Write-Host '    - Distribution UUID auto-discovered from CMS metastore'
Write-Host '      at runtime -- UUID rotations are handled silently'
Write-Host '    - POST 400 auto-retry: stale UUID triggers a live'
Write-Host '      UUID refresh and one retry before GET fallback'
Write-Host '    - 2026 dataset and distribution UUIDs added'
Write-Host '    - Fix: m_LastHttpStatus declared at module level'
Write-Host ''
Write-Host '  Next steps:' -ForegroundColor White
Write-Host '    1. Open Excel'
Write-Host '    2. Open your pharmacy spreadsheet'
Write-Host '    3. Click the  Pharmacy Tools  tab in the ribbon'
Write-Host '    4. Click  NADAC Lookup  to price your rows'
Write-Host '    5. Click  Manage Datasets  to view or edit dataset config'
Write-Host ''
Write-Host '  Add-in installed to:' -ForegroundColor DarkGray
Write-Host "    $AddinPath"          -ForegroundColor DarkGray
Write-Host ''
Write-Host '  To uninstall:' -ForegroundColor DarkGray
Write-Host '    Excel > File > Options > Add-ins > Manage: Excel Add-ins > Go' -ForegroundColor DarkGray
Write-Host "    Uncheck NADACLookup, then delete: $AddinPath" -ForegroundColor DarkGray
Write-Host ''