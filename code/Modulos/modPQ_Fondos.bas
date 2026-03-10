'==========================
' modPQ_Fondos (producci?n)
'==========================

Private Const KEEP_PQ_QUERIES As Boolean = True

' DEBUG (renombrado)
Private Const DEBUG_RENAME As Boolean = False
Private Const DEBUG_RENAME_MSGBOX As Boolean = False   ' Muestra MsgBox con el log de renombrado

' Graficos / resumen
Private Const BUILD_GRAFICOS As Boolean = True

' Tiempo total del proceso (Timer)
Private mT0Total As Double

' Log de etapas para resumen
Private mStageLog As String

' Log de debug (renombrado)
Private mDbg As String

'======================
' Estado Application
'======================
Private mAppFrozen As Boolean
Private mPrevScreenUpdating As Boolean
Private mPrevEnableEvents As Boolean
Private mPrevDisplayAlerts As Boolean
Private mPrevCalculation As XlCalculation
Private mPrevStatusBar As Variant

Private Sub SafeApp(ByVal freeze As Boolean)
    On Error Resume Next
    With Application
        If freeze Then
            If Not mAppFrozen Then
                mPrevScreenUpdating = .ScreenUpdating
                mPrevEnableEvents = .EnableEvents
                mPrevDisplayAlerts = .DisplayAlerts
                mPrevCalculation = .Calculation
                mPrevStatusBar = .StatusBar
                mAppFrozen = True
            End If
            .ScreenUpdating = False
            .EnableEvents = False
            .DisplayAlerts = False
            .Calculation = xlCalculationManual
        Else
            If mAppFrozen Then
                .ScreenUpdating = mPrevScreenUpdating
                .EnableEvents = mPrevEnableEvents
                .DisplayAlerts = mPrevDisplayAlerts
                .Calculation = mPrevCalculation
                .StatusBar = mPrevStatusBar
                mAppFrozen = False
            Else
                .StatusBar = False
            End If
        End If
    End With
End Sub

'======================
' DEBUG helpers
'======================
Private Sub DbgReset()
    mDbg = vbNullString
End Sub

Private Sub DbgAdd(ByVal s As String)
    If Not DEBUG_RENAME Then Exit Sub
    If Len(mDbg) = 0 Then
        mDbg = s
    Else
        mDbg = mDbg & vbCrLf & s
    End If
End Sub

Private Sub DbgShow(Optional ByVal titulo As String = "DEBUG Fondos")
    If Not DEBUG_RENAME Then Exit Sub
    If Not DEBUG_RENAME_MSGBOX Then Exit Sub
    If Len(mDbg) = 0 Then Exit Sub
    MsgBox mDbg, vbInformation, titulo
End Sub

Private Function BoolTxt(ByVal b As Boolean) As String
    If b Then BoolTxt = "SI" Else BoolTxt = "NO"
End Function

Private Function TryGetMultiUserEditing(ByVal wb As Workbook) As String
    On Error GoTo fin
    If wb.MultiUserEditing Then
        TryGetMultiUserEditing = "SI"
    Else
        TryGetMultiUserEditing = "NO"
    End If
    Exit Function
fin:
    TryGetMultiUserEditing = "(no disponible)"
End Function

Private Sub DebugWorkbookStatus(ByVal wb As Workbook)
    DbgAdd "Libro: " & wb.name
    DbgAdd "ReadOnly: " & BoolTxt(wb.ReadOnly)
    DbgAdd "ProtectStructure: " & BoolTxt(wb.ProtectStructure)
    DbgAdd "ProtectWindows: " & BoolTxt(wb.ProtectWindows)
    DbgAdd "MultiUserEditing: " & TryGetMultiUserEditing(wb)
End Sub

Private Function SheetInfo(ByVal wb As Workbook, ByVal sheetName As String) As String
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        SheetInfo = "(no existe)"
    Else
        SheetInfo = "Existe. Visible=" & CStr(ws.Visible) & " Len=" & Len(ws.name)
    End If
End Function

Private Function TableExists(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.name, tableName, vbTextCompare) = 0 Then
                TableExists = True
                Exit Function
            End If
        Next lo
    Next ws
    TableExists = False
End Function

Private Sub DebugListHojas(Optional ByVal maxItems As Long = 60)
    Dim ws As Worksheet, k As Long
    DbgAdd "Hojas (hasta " & CStr(maxItems) & "):"
    k = 0
    For Each ws In ThisWorkbook.Worksheets
        k = k + 1
        If k > maxItems Then
            DbgAdd "  ... (más hojas omitidas)"
            Exit For
        End If
        DbgAdd "  - " & ws.name & " | Visible=" & CStr(ws.Visible) & " | Len=" & Len(ws.name)
    Next ws
End Sub

'======================
' Tiempo (Timer) con control de medianoche
'======================
Private Function ElapsedSec(ByVal t0 As Double) As Double
    Dim t As Double
    t = Timer
    If t < t0 Then t = t + 86400#
    ElapsedSec = t - t0
End Function

Private Function FormatElapsed(ByVal secs As Double) As String
    Dim s As Long, hh As Long, mm As Long, ss As Long
    If secs < 0 Then secs = 0
    s = CLng(secs)
    hh = s \ 3600
    mm = (s \ 60) Mod 60
    ss = s Mod 60

    If hh > 0 Then
        FormatElapsed = Format$(hh, "00") & ":" & Format$(mm, "00") & ":" & Format$(ss, "00")
    Else
        FormatElapsed = Format$(mm, "00") & ":" & Format$(ss, "00")
    End If
End Function

Private Sub StatusStage(ByVal stageLabel As String, ByVal tStage0 As Double)
    If mT0Total <= 0 Then
        Application.StatusBar = "Cargando " & stageLabel & "... " & FormatElapsed(ElapsedSec(tStage0))
    Else
        Application.StatusBar = "Cargando " & stageLabel & "... " & FormatElapsed(ElapsedSec(tStage0)) & " | Total " & FormatElapsed(ElapsedSec(mT0Total))
    End If
End Sub

Private Sub AppendStageLog(ByVal stageLabel As String, ByVal secStage As Double)
    Dim line As String
    line = stageLabel & ": " & FormatElapsed(secStage) & " (" & Format(secStage, "0.0") & " s)"
    If Len(mStageLog) = 0 Then
        mStageLog = line
    Else
        mStageLog = mStageLog & vbCrLf & line
    End If
End Sub

'======================
' Hojas y utilidades
'======================
Private Function EnsureSheet(ByVal nm As String) As Worksheet
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If sh Is Nothing Then
        Set sh = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        sh.name = nm
    End If
    Set EnsureSheet = sh
End Function

Private Sub ClearSheetButKeepName(ByVal sh As Worksheet)
    Dim lo As ListObject, qt As QueryTable, pt As PivotTable, co As ChartObject
    On Error Resume Next
    For Each pt In sh.PivotTables
        pt.TableRange2.Clear
    Next pt
    For Each co In sh.ChartObjects
        co.Delete
    Next co
    For Each lo In sh.ListObjects
        lo.Delete
    Next lo
    For Each qt In sh.QueryTables
        qt.Delete
    Next qt
    sh.Cells.Clear
    On Error GoTo 0
End Sub

Private Sub DeleteAllTablesByName(ByVal wb As Workbook, ByVal tableName As String)
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.name, tableName, vbTextCompare) = 0 Then
                On Error Resume Next
                lo.Delete
                On Error GoTo 0
            End If
        Next lo
    Next ws
End Sub

Private Function TableNameExists(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.name, tableName, vbTextCompare) = 0 Then
                TableNameExists = True
                Exit Function
            End If
        Next lo
    Next ws
    TableNameExists = False
End Function

Private Sub SetTableNameSafe(ByVal wb As Workbook, ByVal lo As ListObject, ByVal desiredName As String)
    Dim nm As String, k As Long

    nm = desiredName
    If Len(Trim$(nm)) = 0 Then Exit Sub

    On Error Resume Next
    lo.name = nm
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    Err.Clear
    On Error GoTo 0

    For k = 2 To 50
        nm = desiredName & "_" & CStr(k)
        If Not TableNameExists(wb, nm) Then
            On Error Resume Next
            lo.name = nm
            On Error GoTo 0
            Exit Sub
        End If
    Next k
End Sub

Private Sub DeleteSheetIfExists(ByVal wb As Workbook, ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0
    If Not ws Is Nothing Then
        On Error Resume Next
        ws.Delete
        On Error GoTo 0
    End If
End Sub

Private Sub DeleteLegacyGraficoSheets()
    Dim i As Long
    Dim ws As Worksheet
    Dim nm As String

    For i = ThisWorkbook.Worksheets.Count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        nm = UCase$(ws.name)

        If nm = "AUX_WORK" Or nm = "CHARTS_WORK" Then
            On Error Resume Next
            ws.Delete
            On Error GoTo 0

        ElseIf Left$(nm, 4) = "AUX_" Then
            If InStr(1, nm, "_SUS_", vbTextCompare) > 0 Or InStr(1, nm, "_RES_", vbTextCompare) > 0 Then
                On Error Resume Next
                ws.Delete
                On Error GoTo 0
            End If

        ElseIf InStr(1, nm, "_GRAFICOS_", vbTextCompare) > 0 Then
            If InStr(1, nm, "_SUS_", vbTextCompare) > 0 Or InStr(1, nm, "_RES_", vbTextCompare) > 0 Then
                On Error Resume Next
                ws.Delete
                On Error GoTo 0
            End If
        End If
    Next i
End Sub

Private Function TryDeleteSheetVerbose(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        DbgAdd "DeleteSheetIfExists: '" & sheetName & "' no existe."
        TryDeleteSheetVerbose = True
        Exit Function
    End If

    DbgAdd "DeleteSheetIfExists: intentando borrar '" & ws.name & "' (Visible=" & CStr(ws.Visible) & ")"
    On Error GoTo EH
    ws.Delete
    DbgAdd "DeleteSheetIfExists: borrada OK '" & sheetName & "'."
    TryDeleteSheetVerbose = True
    Exit Function
EH:
    DbgAdd "DeleteSheetIfExists: ERROR al borrar '" & sheetName & "' | " & CStr(Err.Number) & " | " & Err.Description
    TryDeleteSheetVerbose = False
End Function

'======================
' Nombre seguro de hoja + liberar nombre
'======================
Private Function SanitizeSheetName(ByVal desired As String) As String
    Dim nm As String
    nm = desired

    nm = Replace(nm, "[", "(")
    nm = Replace(nm, "]", ")")
    nm = Replace(nm, ":", " - ")
    nm = Replace(nm, "\", " - ")
    nm = Replace(nm, "/", " - ")
    nm = Replace(nm, "?", " - ")
    nm = Replace(nm, "*", " - ")

    nm = Trim$(nm)
    If Len(nm) = 0 Then nm = "Hoja"

    If Len(nm) > 31 Then nm = Left$(nm, 31)
    SanitizeSheetName = nm
End Function

Private Sub FreeSheetName(ByVal wb As Workbook, ByVal safeName As String, Optional ByVal exceptSheet As Worksheet)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(safeName)
    On Error GoTo 0

    If ws Is Nothing Then
        DbgAdd "FreeSheetName: '" & safeName & "' no existe, no hay colisión."
        Exit Sub
    End If

    If Not exceptSheet Is Nothing Then
        If ws Is exceptSheet Then
            DbgAdd "FreeSheetName: '" & safeName & "' es la misma hoja destino, no se toca."
            Exit Sub
        End If
    End If

    DbgAdd "FreeSheetName: colisión detectada. Ya existe hoja '" & ws.name & "'. Se intentará renombrar a OLD."

    Dim base As String, tmp As String, k As Long
    base = Left$(safeName, 20)
    If Len(base) = 0 Then base = "OLD"

    For k = 1 To 50
        tmp = base & "_OLD_" & Format$(k, "00")
        tmp = SanitizeSheetName(tmp)
        On Error Resume Next
        ws.name = tmp
        If Err.Number = 0 Then
            On Error GoTo 0
            DbgAdd "FreeSheetName: liberado '" & safeName & "' renombrando existente a '" & tmp & "'."
            Exit Sub
        End If
        DbgAdd "FreeSheetName: intento " & CStr(k) & " falló | " & CStr(Err.Number) & " | " & Err.Description
        Err.Clear
        On Error GoTo 0
    Next k

    DbgAdd "FreeSheetName: NO se pudo liberar '" & safeName & "'."
End Sub

Private Sub RenameSheetExact(ByVal sh As Worksheet, ByVal desired As String)
    Dim nm As String
    nm = SanitizeSheetName(desired)

    DbgAdd "RenameSheetExact: destino '" & desired & "' | Sanitizado '" & nm & "' | Len=" & Len(nm)
    DbgAdd "RenameSheetExact: hoja actual '" & sh.name & "'"

    FreeSheetName sh.parent, nm, sh

    On Error GoTo fallback
    sh.name = nm
    DbgAdd "RenameSheetExact: OK. Resultado '" & sh.name & "'"
    Exit Sub

fallback:
    DbgAdd "RenameSheetExact: ERROR directo | " & CStr(Err.Number) & " | " & Err.Description
    Err.Clear
    RenameSheetSafe sh, nm
    DbgAdd "RenameSheetExact: fallback RenameSheetSafe. Resultado '" & sh.name & "'"
End Sub

Private Sub RenameSheetSafe(ByVal sh As Worksheet, ByVal desired As String)
    Dim nm As String
    nm = SanitizeSheetName(desired)

    Dim base As String
    base = nm

    Dim k As Long
    k = 0

    On Error GoTo exists
    sh.name = nm
    Exit Sub

exists:
    Err.Clear
    Do
        k = k + 1
        Dim suf As String
        suf = "_" & CStr(k)

        Dim maxBase As Long
        maxBase = 31 - Len(suf)
        If maxBase < 1 Then maxBase = 1

        nm = Left$(base, maxBase) & suf
        On Error GoTo exists
        sh.name = nm
        Exit Sub
    Loop
End Sub

'======================
' Texto + quitar alerta (CUC / NUMERO DOCUMENTO / N OP / NRO OPERACION BANCO)
'======================
Private Sub ForceTextIdentityColumns(ByVal lo As ListObject)
    ' CUC
    ForceTextColumnByName lo, "CUC"
    IgnoreNumberAsTextByName lo, "CUC"

    ' NUMERO DOCUMENTO (variantes)
    ForceTextColumnByName lo, "NUMERO DE DOCUMENTO"
    IgnoreNumberAsTextByName lo, "NUMERO DE DOCUMENTO"
    ForceTextColumnByName lo, "NÚMERO DE DOCUMENTO"
    IgnoreNumberAsTextByName lo, "NÚMERO DE DOCUMENTO"
    ForceTextColumnByName lo, "NUMERO DOCUMENTO"
    IgnoreNumberAsTextByName lo, "NUMERO DOCUMENTO"
    ForceTextColumnByName lo, "NÚMERO DOCUMENTO"
    IgnoreNumberAsTextByName lo, "NÚMERO DOCUMENTO"

    ' N OP (variantes)
    ForceTextColumnByName lo, "N OP"
    IgnoreNumberAsTextByName lo, "N OP"
    ForceTextColumnByName lo, "NRO OP"
    IgnoreNumberAsTextByName lo, "NRO OP"

    ' NRO OPERACION BANCO (variantes)
    ForceTextColumnByName lo, "NRO OPERACIÓN BANCO"
    IgnoreNumberAsTextByName lo, "NRO OPERACIÓN BANCO"
    ForceTextColumnByName lo, "NRO OPERACION BANCO"
    IgnoreNumberAsTextByName lo, "NRO OPERACION BANCO"
End Sub

Private Function StripDiacriticsUpper(ByVal s As String) As String
    Dim t As String
    t = s

    t = Replace(t, "Á", "A")
    t = Replace(t, "À", "A")
    t = Replace(t, "Â", "A")
    t = Replace(t, "Ä", "A")

    t = Replace(t, "É", "E")
    t = Replace(t, "È", "E")
    t = Replace(t, "Ê", "E")
    t = Replace(t, "Ë", "E")

    t = Replace(t, "Í", "I")
    t = Replace(t, "Ì", "I")
    t = Replace(t, "Î", "I")
    t = Replace(t, "Ï", "I")

    t = Replace(t, "Ó", "O")
    t = Replace(t, "Ò", "O")
    t = Replace(t, "Ô", "O")
    t = Replace(t, "Ö", "O")

    t = Replace(t, "Ú", "U")
    t = Replace(t, "Ù", "U")
    t = Replace(t, "Û", "U")
    t = Replace(t, "Ü", "U")

    t = Replace(t, "Ñ", "N")

    StripDiacriticsUpper = t
End Function

Private Function CanonColName(ByVal s As String) As String
    Dim t As String
    t = UCase$(Trim$(s))
    t = Replace(t, Chr$(160), " ")
    t = StripDiacriticsUpper(t)
    t = Replace(t, "°", "")
    t = Replace(t, "º", "")
    t = Replace(t, " ", "")
    CanonColName = t
End Function

Private Function FindListColumnByName(ByVal lo As ListObject, ByVal colName As String) As ListColumn
    Dim lc As ListColumn
    Dim want As String
    want = CanonColName(colName)

    For Each lc In lo.ListColumns
        If CanonColName(lc.name) = want Then
            Set FindListColumnByName = lc
            Exit Function
        End If
    Next lc
    Set FindListColumnByName = Nothing
End Function

Private Sub ForceTextColumnByName(ByVal lo As ListObject, ByVal colName As String)
    On Error GoTo fin
    If lo Is Nothing Then Exit Sub

    Dim lc As ListColumn
    Set lc = FindListColumnByName(lo, colName)
    If lc Is Nothing Then Exit Sub

    lc.Range.NumberFormat = "@"
fin:
End Sub

Private Sub IgnoreNumberAsTextByName(ByVal lo As ListObject, ByVal colName As String)
    On Error GoTo fin
    If lo Is Nothing Then Exit Sub

    Dim lc As ListColumn
    Set lc = FindListColumnByName(lo, colName)
    If lc Is Nothing Then Exit Sub

    On Error Resume Next
    If Not lc.DataBodyRange Is Nothing Then
        lc.DataBodyRange.Errors(xlNumberAsText).Ignore = True
    End If
    lc.Range.Errors(xlNumberAsText).Ignore = True
    On Error GoTo 0
fin:
End Sub

'======================
' Fechas para sufijo (tomadas de la data cargada)
'======================
Private Function LastDayOfMonth(ByVal d As Date) As Date
    LastDayOfMonth = DateSerial(Year(d), Month(d) + 1, 0)
End Function

Private Function FirstDayOfMonth(ByVal d As Date) As Date
    FirstDayOfMonth = DateSerial(Year(d), Month(d), 1)
End Function

Private Function TryCoerceExcelDate(ByVal v As Variant, ByRef outD As Date) As Boolean
    On Error GoTo fin
    If IsError(v) Or IsEmpty(v) Then GoTo fin

    If IsDate(v) Then
        outD = CDate(v)
        TryCoerceExcelDate = True
        Exit Function
    End If

    If IsNumeric(v) Then
        Dim n As Double
        n = CDbl(v)
        If n > 0# And n < 60000# Then
            outD = DateSerial(1899, 12, 30) + n
            TryCoerceExcelDate = True
            Exit Function
        End If
    End If

fin:
    TryCoerceExcelDate = False
End Function

Private Function GetMinMaxDateFromLO(ByVal lo As ListObject, ByVal colName As String, ByRef outMin As Date, ByRef outMax As Date) As Boolean
    On Error GoTo fin
    GetMinMaxDateFromLO = False
    If lo Is Nothing Then Exit Function

    Dim lc As ListColumn
    Set lc = FindListColumnByName(lo, colName)
    If lc Is Nothing Then Exit Function
    If lc.DataBodyRange Is Nothing Then Exit Function

    Dim c As Range
    Dim d As Date, gotAny As Boolean
    gotAny = False

    For Each c In lc.DataBodyRange.Cells
        If TryCoerceExcelDate(c.Value2, d) Then
            If Not gotAny Then
                outMin = d
                outMax = d
                gotAny = True
            Else
                If d < outMin Then outMin = d
                If d > outMax Then outMax = d
            End If
        End If
    Next c

    GetMinMaxDateFromLO = gotAny
    Exit Function

fin:
    GetMinMaxDateFromLO = False
End Function

'======================
' Mes abreviado ES
'======================
Private Function MesAbrevES(ByVal d As Date) As String
    Dim m As Long
    m = Month(d)
    Select Case m
        Case 1: MesAbrevES = "ENE"
        Case 2: MesAbrevES = "FEB"
        Case 3: MesAbrevES = "MAR"
        Case 4: MesAbrevES = "ABR"
        Case 5: MesAbrevES = "MAY"
        Case 6: MesAbrevES = "JUN"
        Case 7: MesAbrevES = "JUL"
        Case 8: MesAbrevES = "AGO"
        Case 9: MesAbrevES = "SEP"
        Case 10: MesAbrevES = "OCT"
        Case 11: MesAbrevES = "NOV"
        Case 12: MesAbrevES = "DIC"
        Case Else: MesAbrevES = "MES"
    End Select
End Function

'======================
' Power Query M (RAW / MAIN / ALERTAS)
'======================
Private Function M_RAW_SUS(ByVal rutaArchivo As String) As String
    Dim m As String
    m = _
"let" & vbCrLf & _
"    Source = Excel.Workbook(File.Contents(""" & rutaArchivo & """), null, true)," & vbCrLf & _
"    Data_SUS = Source{[Item=""SUS"",Kind=""Sheet""]}[Data]," & vbCrLf & _
"    Promoted = Table.PromoteHeaders(Data_SUS, [PromoteAllScalars=true])" & vbCrLf & _
"in" & vbCrLf & _
"    Promoted"
    M_RAW_SUS = m
End Function

Private Function M_RAW_RES(ByVal rutaArchivo As String) As String
    Dim m As String
    m = _
"let" & vbCrLf & _
"    Source = Excel.Workbook(File.Contents(""" & rutaArchivo & """), null, true)," & vbCrLf & _
"    Data_RES = Source{[Item=""RES"",Kind=""Sheet""]}[Data]," & vbCrLf & _
"    Promoted = Table.PromoteHeaders(Data_RES, [PromoteAllScalars=true])" & vbCrLf & _
"in" & vbCrLf & _
"    Promoted"
    M_RAW_RES = m
End Function

Private Function M_SUS(ByVal mesesSel As Long) As String
    Dim m As String
    m = _
"let" & vbCrLf & _
"    Source = #" & Chr$(34) & "Consulta - RAW_SUS" & Chr$(34) & "," & vbCrLf & _
"    Changed = Table.TransformColumnTypes(Source, {{""FECHA PROCESO"", type date}})," & vbCrLf & _
"    Cut = Table.SelectRows(Changed, each [FECHA PROCESO] >= Date.AddMonths(Date.From(DateTime.LocalNow()), -" & CStr(mesesSel) & "))" & vbCrLf & _
"in" & vbCrLf & _
"    Cut"
    M_SUS = m
End Function

Private Function M_RES(ByVal mesesSel As Long) As String
    Dim m As String
    m = _
"let" & vbCrLf & _
"    Source = #" & Chr$(34) & "Consulta - RAW_RES" & Chr$(34) & "," & vbCrLf & _
"    Changed = Table.TransformColumnTypes(Source, {{""FECHA PROCESO"", type date}})," & vbCrLf & _
"    Cut = Table.SelectRows(Changed, each [FECHA PROCESO] >= Date.AddMonths(Date.From(DateTime.LocalNow()), -" & CStr(mesesSel) & "))" & vbCrLf & _
"in" & vbCrLf & _
"    Cut"
    M_RES = m
End Function

Private Function M_ALERTAS(ByVal baseName As String) As String
    Dim m As String
    m = _
"let" & vbCrLf & _
"    Source = #" & Chr$(34) & "Consulta - " & baseName & Chr$(34) & "," & vbCrLf & _
"    Changed = Table.TransformColumnTypes(Source, {{""SUMA_MONTOS"", type number}, {""DESVIACION_MEDIA_%"", type number}})," & vbCrLf & _
"    Keep = Table.SelectRows(Changed, each [NIVEL_RIESGO] <> null and Text.Length(Text.From([NIVEL_RIESGO])) > 0)" & vbCrLf & _
"in" & vbCrLf & _
"    Keep"
    M_ALERTAS = m
End Function

'======================
' Conexiones PQ
'======================
Private Function EnsurePQConnection(ByVal queryName As String) As WorkbookConnection
    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim conn As WorkbookConnection
    On Error Resume Next
    Set conn = wb.Connections("Consulta - " & queryName)
    If conn Is Nothing Then Set conn = wb.Connections("Query - " & queryName)
    If conn Is Nothing Then Set conn = wb.Connections("PQ_" & queryName)
    If conn Is Nothing Then Set conn = wb.Connections(queryName)
    On Error GoTo 0

    If conn Is Nothing Then
        Set conn = wb.Connections.Add2( _
            name:="Consulta - " & queryName, _
            Description:="", _
            ConnectionString:="OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & queryName & ";Extended Properties=""""", _
            CommandText:=queryName, _
            lCmdtype:=xlCmdSql, _
            CreateModelConnection:=False, _
            ImportRelationships:=False)
    End If

    Set EnsurePQConnection = conn
End Function

'======================
' Tablas para conexiones + refresh robusto
'======================
Private Function EnsureTableForConnection(ByVal sh As Worksheet, ByVal conn As WorkbookConnection, ByVal loName As String) As ListObject
    Dim lo As ListObject
    On Error Resume Next
    Set lo = sh.ListObjects(loName)
    On Error GoTo 0

    If lo Is Nothing Then
        Dim qt As QueryTable
        Set qt = sh.QueryTables.Add(Connection:=conn, Destination:=sh.Range("A1"))
        qt.name = "QT_" & loName
        qt.BackgroundQuery = False
        qt.Refresh BackgroundQuery:=False

        Set lo = sh.ListObjects.Add(xlSrcRange, qt.ResultRange, , xlYes)
        lo.name = loName
    End If

    Set EnsureTableForConnection = lo
End Function

Private Sub RefreshListObject(ByVal lo As ListObject, ByVal conn As WorkbookConnection)
    On Error Resume Next

    If Not lo.QueryTable Is Nothing Then
        lo.QueryTable.BackgroundQuery = False
        lo.QueryTable.Refresh BackgroundQuery:=False
        Exit Sub
    End If

    If Not conn Is Nothing Then
        If conn.Type = xlConnectionTypeOLEDB Then conn.OLEDBConnection.BackgroundQuery = False
        conn.Refresh
    End If
End Sub

Private Function HasImportPlaceholder(ByVal lo As ListObject) As Boolean
    On Error GoTo fin
    HasImportPlaceholder = False
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function

    Dim r As Range
    Set r = lo.DataBodyRange.Cells(1, 1)

    Dim s As String
    s = CStr(r.Value2)

    If InStr(1, s, "Importando", vbTextCompare) > 0 Then
        HasImportPlaceholder = True
    End If
    Exit Function
fin:
    HasImportPlaceholder = False
End Function

Private Function FreezeListObject(ByVal lo As ListObject) As ListObject
    On Error GoTo fin
    If lo Is Nothing Then
        Set FreezeListObject = lo
        Exit Function
    End If

    Dim rng As Range
    Set rng = lo.Range

    Dim sh As Worksheet
    Set sh = lo.parent

    Dim loName As String
    loName = lo.name

    On Error Resume Next
    lo.QueryTable.Delete
    On Error GoTo 0

    Dim lo2 As ListObject
    On Error Resume Next
    lo.Delete
    On Error GoTo 0

    Set lo2 = sh.ListObjects.Add(xlSrcRange, rng, , xlYes)
    lo2.name = loName

    Set FreezeListObject = lo2
    Exit Function

fin:
    Set FreezeListObject = lo
End Function

Private Sub WaitListObjectReady(ByVal lo As ListObject, ByVal conn As WorkbookConnection, ByVal stageLabel As String, ByVal timeoutSec As Double, ByVal tStage0 As Double)
    Dim tTimeout0 As Double
    tTimeout0 = Timer

    Do
        DoEvents
        StatusStage stageLabel, tStage0

        Dim qt As QueryTable
        Set qt = Nothing
        On Error Resume Next
        Set qt = lo.QueryTable
        On Error GoTo 0

        Dim isQtRef As Boolean, isConnRef As Boolean
        Dim hasHold As Boolean
        isQtRef = False
        isConnRef = False
        hasHold = False

        If Not qt Is Nothing Then
            isQtRef = True
            On Error Resume Next
            hasHold = qt.refreshing
            On Error GoTo 0
        End If

        If (Not isQtRef) And (Not conn Is Nothing) Then
            isConnRef = True
            On Error Resume Next
            hasHold = conn.refreshing
            On Error GoTo 0
        End If

        Dim didKick As Boolean
        didKick = False

        If (Not isQtRef) And (Not isConnRef) And hasHold Then
            If (Not didKick) Then
                didKick = True
                On Error Resume Next
                If Not qt Is Nothing Then
                    qt.BackgroundQuery = False
                    qt.Refresh BackgroundQuery:=False
                ElseIf Not conn Is Nothing Then
                    If conn.Type = xlConnectionTypeOLEDB Then conn.OLEDBConnection.BackgroundQuery = False
                    conn.Refresh
                End If
                On Error GoTo 0
            End If
        End If

        If (Not isQtRef) And (Not isConnRef) And (Not hasHold) Then Exit Do

        If ElapsedSec(tTimeout0) > timeoutSec Then
            Err.Raise vbObjectError + 513, "WaitListObjectReady", "Timeout al cargar " & stageLabel & "."
        End If
    Loop
End Sub

Private Function EnsureStage(ByVal sh As Worksheet, ByVal loName As String, ByVal conn As WorkbookConnection, ByVal stageLabel As String, ByVal showProgress As Boolean) As ListObject
    Dim tStage0 As Double
    tStage0 = Timer

    StatusStage stageLabel, tStage0
    DoEvents

    Dim lo As ListObject
    Set lo = EnsureTableForConnection(sh, conn, loName)

    RefreshListObject lo, conn
    WaitListObjectReady lo, conn, stageLabel, 900, tStage0

    ForceTextIdentityColumns lo

    Dim attempt As Long
    For attempt = 1 To 2
        If Not HasImportPlaceholder(lo) Then Exit For
        RefreshListObject lo, conn
        WaitListObjectReady lo, conn, stageLabel, 900, tStage0
        ForceTextIdentityColumns lo
    Next attempt

    If HasImportPlaceholder(lo) Then
        DeleteAllTablesByName sh.parent, loName
        Set lo = EnsureTableForConnection(sh, conn, loName)
        RefreshListObject lo, conn
        WaitListObjectReady lo, conn, stageLabel, 900, tStage0
        ForceTextIdentityColumns lo
    End If

    If HasImportPlaceholder(lo) Then
        Err.Raise vbObjectError + 514, "EnsureStage", "La carga de " & stageLabel & " no terminó (placeholder 'Importando datos')."
    End If

    Set lo = FreezeListObject(lo)
    ForceTextIdentityColumns lo

    Dim secStage As Double
    secStage = ElapsedSec(tStage0)

    Dim msg As String
    msg = "Ya cargó " & stageLabel & "." & vbCrLf & _
          "Tiempo: " & FormatElapsed(secStage) & " (" & Format(secStage, "0.0") & " s)" & vbCrLf & _
          "Total: " & FormatElapsed(ElapsedSec(mT0Total))

    Application.StatusBar = stageLabel & " listo. Etapa " & Format(secStage, "0.0") & " s (" & FormatElapsed(secStage) & ") | Total " & FormatElapsed(ElapsedSec(mT0Total))

    AppendStageLog stageLabel, secStage

    If showProgress Then
        Debug.Print msg
        If StrComp(stageLabel, "ALERTAS", vbTextCompare) <> 0 Then
            MsgBox msg, vbInformation, "Fondos"
        End If
    End If

    Set EnsureStage = lo
End Function

Private Sub DeleteQueryAndConnection(ByVal qName As String)
    On Error Resume Next
    ThisWorkbook.Queries.Item(qName).Delete
    ThisWorkbook.Connections("PQ_" & qName).Delete
    On Error GoTo 0
End Sub

'======================
' Principal
'======================
Public Sub CrearQueryFondos(ByVal rutaArchivo As String, ByVal arg2 As Variant, ByVal arg3 As Variant, _
                            Optional ByVal arg4 As Variant, _
                            Optional ByVal arg5 As Variant, _
                            Optional ByVal arg6 As Variant)
    On Error GoTo EH

    mT0Total = Timer
    mStageLog = vbNullString

    DbgReset
    DebugWorkbookStatus ThisWorkbook
    DbgAdd "RutaArchivo: " & rutaArchivo

    Dim esRescate As Boolean
    Dim mesesSel As Long

    If VarType(arg2) = vbBoolean Then
        esRescate = CBool(arg2)
        mesesSel = CoerceLong(arg3, 6)
    ElseIf VarType(arg3) = vbBoolean Then
        mesesSel = CoerceLong(arg2, 6)
        esRescate = CBool(arg3)
    Else
        mesesSel = CoerceLong(arg2, 6)
        esRescate = CoerceBool(arg3, False)
    End If

    mesesSel = 6
    DbgAdd "mesesSel (forzado): " & CStr(mesesSel)
    DbgAdd "esRescate: " & BoolTxt(esRescate)

    Dim activar As Boolean
    activar = True

    Dim entidadPrefix As String
    entidadPrefix = "FONDOS"

    Dim showProg As Boolean
    showProg = True

    If Not IsMissing(arg4) Then
        If IsBoolLike(arg4) Then
            activar = CoerceBool(arg4, True)
            If Not IsMissing(arg5) Then
                If Len(CoerceText(arg5)) > 0 Then entidadPrefix = UCase$(CoerceText(arg5))
            End If
            If Not IsMissing(arg6) Then
                showProg = CoerceBool(arg6, False)
            ElseIf Not IsMissing(arg5) Then
                showProg = CoerceBool(arg5, False)
            End If
        Else
            If Len(CoerceText(arg4)) > 0 Then entidadPrefix = UCase$(CoerceText(arg4))
            If Not IsMissing(arg5) Then showProg = CoerceBool(arg5, False)
        End If
    End If

    DbgAdd "entidadPrefix: " & entidadPrefix
    DbgAdd "showProg: " & BoolTxt(showProg)
    DbgAdd "activar: " & BoolTxt(activar)

    SafeApp True

    Dim opCode As String
    If esRescate Then opCode = "RES" Else opCode = "SUS"
    DbgAdd "opCode: " & opCode

    On Error Resume Next
    ThisWorkbook.Queries.Item("RAW_SUS").Delete
    ThisWorkbook.Queries.Item("SUS").Delete
    ThisWorkbook.Queries.Item("SUS_ALERTAS").Delete
    ThisWorkbook.Queries.Item("RAW_RES").Delete
    ThisWorkbook.Queries.Item("RES").Delete
    ThisWorkbook.Queries.Item("RES_ALERTAS").Delete
    On Error GoTo EH

    If esRescate Then
        ThisWorkbook.Queries.Add name:="RAW_RES", Formula:=M_RAW_RES(rutaArchivo)
        ThisWorkbook.Queries.Add name:="RES", Formula:=M_RES(mesesSel)
        ThisWorkbook.Queries.Add name:="RES_ALERTAS", Formula:=M_ALERTAS("RES")
    Else
        ThisWorkbook.Queries.Add name:="RAW_SUS", Formula:=M_RAW_SUS(rutaArchivo)
        ThisWorkbook.Queries.Add name:="SUS", Formula:=M_SUS(mesesSel)
        ThisWorkbook.Queries.Add name:="SUS_ALERTAS", Formula:=M_ALERTAS("SUS")
    End If

    Dim shRaw As Worksheet, shMain As Worksheet, shAL As Worksheet
    Set shRaw = EnsureSheet("RAW_WORK"): ClearSheetButKeepName shRaw
    Set shMain = EnsureSheet("MAIN_WORK"): ClearSheetButKeepName shMain
    Set shAL = EnsureSheet("ALERTAS_WORK"): ClearSheetButKeepName shAL

    ' Limpieza heredada (vía antigua de gráficos)
    DeleteLegacyGraficoSheets

    Dim connRaw As WorkbookConnection, connMain As WorkbookConnection, connAL As WorkbookConnection
    If esRescate Then
        Set connRaw = EnsurePQConnection("RAW_RES")
        Set connMain = EnsurePQConnection("RES")
        Set connAL = EnsurePQConnection("RES_ALERTAS")
    Else
        Set connRaw = EnsurePQConnection("RAW_SUS")
        Set connMain = EnsurePQConnection("SUS")
        Set connAL = EnsurePQConnection("SUS_ALERTAS")
    End If

    Dim loRaw As ListObject, loMain As ListObject, loAL As ListObject
    Set loRaw = EnsureStage(shRaw, "RAW_WORK", connRaw, "RAW", showProg)
    Set loMain = EnsureStage(shMain, "MAIN_WORK", connMain, opCode, showProg)
    Set loAL = EnsureStage(shAL, "ALERTAS_WORK", connAL, "ALERTAS", showProg)

    Dim minD As Date, maxD As Date
    Dim gotDates As Boolean

    gotDates = GetMinMaxDateFromLO(loMain, "FECHA PROCESO", minD, maxD)
    DbgAdd "GetMinMaxDateFromLO loMain(FECHA PROCESO): " & BoolTxt(gotDates)

    If Not gotDates Then
        gotDates = GetMinMaxDateFromLO(loRaw, "FECHA PROCESO", minD, maxD)
        DbgAdd "GetMinMaxDateFromLO loRaw(FECHA PROCESO): " & BoolTxt(gotDates)
    End If

    Dim fin As Date, ini As Date
    If gotDates Then
        ini = FirstDayOfMonth(minD)
        fin = LastDayOfMonth(maxD)
        DbgAdd "minD (data): " & Format$(minD, "yyyy-mm-dd")
        DbgAdd "maxD (data): " & Format$(maxD, "yyyy-mm-dd")
        DbgAdd "ini (mes de minD): " & Format$(ini, "yyyy-mm-dd")
        DbgAdd "fin (mes de maxD): " & Format$(fin, "yyyy-mm-dd")
    Else
        fin = DateSerial(Year(Date), Month(Date), 0)
        ini = DateSerial(Year(fin), Month(fin) - (mesesSel - 1), 1)
        DbgAdd "Fallback a Date()"
        DbgAdd "ini: " & Format$(ini, "yyyy-mm-dd")
        DbgAdd "fin: " & Format$(fin, "yyyy-mm-dd")
    End If

    Dim suf As String
    suf = MesAbrevES(ini) & "_" & MesAbrevES(fin) & "_" & Year(fin)
    DbgAdd "suf: " & suf

    Dim nmRaw As String, nmMain As String, nmAL As String
    nmRaw = "RAW_" & entidadPrefix & "_" & opCode & "_" & suf
    nmMain = entidadPrefix & "_" & opCode & "_" & suf
    nmAL = entidadPrefix & "_" & opCode & "_ALERTAS_" & suf

    Dim shNmRaw As String, shNmMain As String, shNmAL As String
    shNmRaw = SanitizeSheetName(nmRaw)
    shNmMain = SanitizeSheetName(nmMain)
    shNmAL = SanitizeSheetName(nmAL)

    If DEBUG_RENAME Then
        Call TryDeleteSheetVerbose(ThisWorkbook, shNmRaw)
        Call TryDeleteSheetVerbose(ThisWorkbook, shNmMain)
        Call TryDeleteSheetVerbose(ThisWorkbook, shNmAL)
    Else
        DeleteSheetIfExists ThisWorkbook, shNmRaw
        DeleteSheetIfExists ThisWorkbook, shNmMain
        DeleteSheetIfExists ThisWorkbook, shNmAL
    End If

    FreeSheetName ThisWorkbook, shNmRaw, shRaw
    FreeSheetName ThisWorkbook, shNmMain, shMain
    FreeSheetName ThisWorkbook, shNmAL, shAL

    DeleteAllTablesByName ThisWorkbook, nmRaw
    DeleteAllTablesByName ThisWorkbook, nmMain
    DeleteAllTablesByName ThisWorkbook, nmAL

    SetTableNameSafe ThisWorkbook, loRaw, nmRaw
    SetTableNameSafe ThisWorkbook, loMain, nmMain
    SetTableNameSafe ThisWorkbook, loAL, nmAL

    RenameSheetExact shRaw, nmRaw
    RenameSheetExact shMain, nmMain
    RenameSheetExact shAL, nmAL

    If BUILD_GRAFICOS Then
        modFondosGraficos.BuildGraficosAlertasEnHoja loAL, entidadPrefix & " " & opCode
    End If

    If Not KEEP_PQ_QUERIES Then
        If esRescate Then
            DeleteQueryAndConnection "RAW_RES"
            DeleteQueryAndConnection "RES"
            DeleteQueryAndConnection "RES_ALERTAS"
        Else
            DeleteQueryAndConnection "RAW_SUS"
            DeleteQueryAndConnection "SUS"
            DeleteQueryAndConnection "SUS_ALERTAS"
        End If
    End If

    SafeApp False

    Dim totalMsg As String
    totalMsg = "Proceso terminado." & vbCrLf & vbCrLf & _
               mStageLog & vbCrLf & vbCrLf & _
               "Total: " & FormatElapsed(ElapsedSec(mT0Total))

    Application.StatusBar = "Listo. Total " & FormatElapsed(ElapsedSec(mT0Total))
    Debug.Print totalMsg

    If showProg Then
        MsgBox totalMsg, vbInformation, "Fondos"
    End If

    If activar Then
        shMain.Activate
        shMain.Range("A1").Select
    End If

    Exit Sub

EH:
    Dim errDesc As String
    errDesc = Err.Description
    If Len(Trim$(errDesc)) = 0 Then errDesc = "(sin descripción)"

    If DEBUG_RENAME Then
        DbgAdd ""
        DbgAdd "ERROR FINAL: " & CStr(Err.Number) & " | " & errDesc
        DebugWorkbookStatus ThisWorkbook
        DebugListHojas 80
        DbgShow "DEBUG Fondos (falló)"
    End If

    Dim msg As String
    msg = "CrearQueryFondos falló." & vbCrLf & _
          "Error " & Err.Number & vbCrLf & _
          errDesc

    SafeApp False
    Err.Raise Err.Number, "CrearQueryFondos", msg
End Sub

'======================
' Coerciones seguras (privadas)
'======================
Private Function UnwrapValue(ByVal v As Variant) As Variant
    On Error GoTo fin
    If IsObject(v) Then
        If TypeName(v) = "Range" Then
            If v.Cells.CountLarge > 0 Then
                UnwrapValue = v.Cells(1, 1).Value2
                Exit Function
            End If
        End If
    End If
fin:
    UnwrapValue = v
End Function

Private Function CoerceText(ByVal v As Variant) As String
    v = UnwrapValue(v)
    If IsError(v) Then
        CoerceText = vbNullString
    Else
        CoerceText = Trim$(CStr(v))
    End If
End Function

Private Function IsBoolLike(ByVal v As Variant) As Boolean
    v = UnwrapValue(v)

    If VarType(v) = vbBoolean Then
        IsBoolLike = True
        Exit Function
    End If

    If IsNumeric(v) Then
        IsBoolLike = True
        Exit Function
    End If

    Dim s As String
    s = UCase$(Trim$(CStr(v)))
    IsBoolLike = (s = "TRUE" Or s = "FALSE" Or s = "VERDADERO" Or s = "FALSO" Or s = "SI" Or s = "NO" Or s = "1" Or s = "0")
End Function

Private Function CoerceBool(ByVal v As Variant, Optional ByVal def As Boolean = False) As Boolean
    v = UnwrapValue(v)

    If VarType(v) = vbBoolean Then
        CoerceBool = CBool(v)
        Exit Function
    End If

    If IsNumeric(v) Then
        CoerceBool = (CDbl(v) <> 0)
        Exit Function
    End If

    Dim s As String
    s = UCase$(Trim$(CStr(v)))

    Select Case s
        Case "TRUE", "VERDADERO", "SI", "1"
            CoerceBool = True
        Case "FALSE", "FALSO", "NO", "0", ""
            CoerceBool = False
        Case Else
            CoerceBool = def
    End Select
End Function

Private Function CoerceLong(ByVal v As Variant, Optional ByVal def As Long = 0) As Long
    v = UnwrapValue(v)

    If IsError(v) Or IsEmpty(v) Or Len(Trim$(CStr(v))) = 0 Then
        CoerceLong = def
        Exit Function
    End If

    If IsNumeric(v) Then
        CoerceLong = CLng(CDbl(v))
        Exit Function
    End If

    On Error GoTo fin
    CoerceLong = CLng(CDbl(v))
    Exit Function
fin:
    CoerceLong = def
End Function

