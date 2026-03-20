'==========================
' modPQ_Fondos (produccion)
'==========================
Option Explicit

Private Const KEEP_PQ_QUERIES     As Boolean = True
Private Const DEBUG_RENAME        As Boolean = False
Private Const DEBUG_RENAME_MSGBOX As Boolean = False
Private Const DEBUG_LOAD          As Boolean = False
Private Const BUILD_GRAFICOS      As Boolean = True
Private Const USE_NUMERO_DOC      As Boolean = True   ' True = alertas por NUMERO DE DOCUMENTO si hay tabla de clientes

Private mT0Total   As Double
Private mStageLog  As String
Private mDbg       As String
Private mLoadLog   As String

'======================
' Estado Application
'======================
Private mAppFrozen          As Boolean
Private mPrevScreenUpdating As Boolean
Private mPrevEnableEvents   As Boolean
Private mPrevDisplayAlerts  As Boolean
Private mPrevCalculation    As XlCalculation
Private mPrevStatusBar      As Variant

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
    On Error GoTo 0
End Sub

Private Sub UnfreezeForRefresh()
    On Error Resume Next
    Application.ScreenUpdating = True
    Application.EnableEvents = True
    Application.DisplayAlerts = True
    On Error GoTo 0
End Sub

Private Sub RefreezeAfterRefresh()
    On Error Resume Next
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    On Error GoTo 0
End Sub

'======================
' Load debug log
'======================
Private Sub LdReset()
    mLoadLog = vbNullString
End Sub

Private Sub LdAdd(ByVal s As String)
    If Not DEBUG_LOAD Then Exit Sub
    If Len(mLoadLog) = 0 Then
        mLoadLog = s
    Else
        mLoadLog = mLoadLog & vbCrLf & s
    End If
End Sub

Private Sub LdShow(Optional ByVal titulo As String = "DEBUG Carga PQ")
    If Not DEBUG_LOAD Then Exit Sub
    If Len(mLoadLog) = 0 Then
        MsgBox "(sin entradas)", vbInformation, titulo
    Else
        MsgBox mLoadLog, vbInformation, titulo
    End If
End Sub

'======================
' DEBUG helpers (renombrado)
'======================
Private Sub DbgReset()
    mDbg = vbNullString
End Sub

Private Sub DbgAdd(ByVal s As String)
    If Not DEBUG_RENAME Then Exit Sub
    If Len(mDbg) = 0 Then mDbg = s Else mDbg = mDbg & vbCrLf & s
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
    If wb.MultiUserEditing Then TryGetMultiUserEditing = "SI" Else TryGetMultiUserEditing = "NO"
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

Private Sub DebugListHojas(Optional ByVal maxItems As Long = 60)
    Dim ws As Worksheet
    Dim k As Long
    DbgAdd "Hojas (hasta " & CStr(maxItems) & "):"
    k = 0
    For Each ws In ThisWorkbook.Worksheets
        k = k + 1
        If k > maxItems Then DbgAdd "  ... (mas hojas omitidas)": Exit For
        DbgAdd "  - " & ws.name & " | Visible=" & CStr(ws.Visible) & " | Len=" & Len(ws.name)
    Next ws
End Sub

'======================
' Tiempo
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
        Application.StatusBar = "Cargando " & stageLabel & "... " & FormatElapsed(ElapsedSec(tStage0)) & _
                                " | Total " & FormatElapsed(ElapsedSec(mT0Total))
    End If
End Sub

Private Sub AppendStageLog(ByVal stageLabel As String, ByVal secStage As Double)
    Dim line As String
    line = stageLabel & ": " & FormatElapsed(secStage) & " (" & Format(secStage, "0.0") & " s)"
    If Len(mStageLog) = 0 Then mStageLog = line Else mStageLog = mStageLog & vbCrLf & line
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
        Set sh = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        sh.name = nm
    End If
    Set EnsureSheet = sh
End Function

Private Sub ClearSheetButKeepName(ByVal sh As Worksheet)
    Dim lo As ListObject, qt As QueryTable, pt As PivotTable, co As ChartObject
    On Error Resume Next
    For Each pt In sh.PivotTables:  pt.TableRange2.Clear: Next pt
    For Each co In sh.ChartObjects: co.Delete:            Next co
    For Each lo In sh.ListObjects:  lo.Delete:            Next lo
    For Each qt In sh.QueryTables:  qt.Delete:            Next qt
    sh.Cells.Clear
    On Error GoTo 0
End Sub

Private Sub DeleteAllTablesByName(ByVal wb As Workbook, ByVal tableName As String)
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.name, tableName, vbTextCompare) = 0 Then
                On Error Resume Next: lo.Delete: On Error GoTo 0
            End If
        Next lo
    Next ws
End Sub

Private Function TableNameExists(ByVal wb As Workbook, ByVal tableName As String) As Boolean
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.name, tableName, vbTextCompare) = 0 Then
                TableNameExists = True: Exit Function
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
    If Err.Number = 0 Then On Error GoTo 0: Exit Sub
    Err.Clear: On Error GoTo 0
    For k = 2 To 50
        nm = desiredName & "_" & CStr(k)
        If Not TableNameExists(wb, nm) Then
            On Error Resume Next: lo.name = nm: On Error GoTo 0
            Exit Sub
        End If
    Next k
End Sub

Private Sub DeleteSheetIfExists(ByVal wb As Workbook, ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next: Set ws = wb.Worksheets(sheetName): On Error GoTo 0
    If Not ws Is Nothing Then
        On Error Resume Next: ws.Visible = xlSheetVisible: ws.Delete: On Error GoTo 0
    End If
End Sub

Private Sub DeleteLegacyGraficoSheets()
    Dim i As Long, ws As Worksheet, nm As String
    For i = ThisWorkbook.Worksheets.count To 1 Step -1
        Set ws = ThisWorkbook.Worksheets(i)
        nm = UCase$(ws.name)
        If nm = "AUX_WORK" Or nm = "CHARTS_WORK" Then
            On Error Resume Next: ws.Visible = xlSheetVisible: ws.Delete: On Error GoTo 0
        ElseIf Left$(nm, 4) = "AUX_" Then
            If InStr(1, nm, "_SUS_", vbTextCompare) > 0 Or InStr(1, nm, "_RES_", vbTextCompare) > 0 Then
                On Error Resume Next: ws.Visible = xlSheetVisible: ws.Delete: On Error GoTo 0
            End If
        ElseIf InStr(1, nm, "_GRAFICOS_", vbTextCompare) > 0 Then
            If InStr(1, nm, "_SUS_", vbTextCompare) > 0 Or InStr(1, nm, "_RES_", vbTextCompare) > 0 Then
                On Error Resume Next: ws.Visible = xlSheetVisible: ws.Delete: On Error GoTo 0
            End If
        End If
    Next i
End Sub

Private Function TryDeleteSheetVerbose(ByVal wb As Workbook, ByVal sheetName As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next: Set ws = wb.Worksheets(sheetName): On Error GoTo 0
    If ws Is Nothing Then
        DbgAdd "DeleteSheetIfExists: '" & sheetName & "' no existe."
        TryDeleteSheetVerbose = True: Exit Function
    End If
    DbgAdd "DeleteSheetIfExists: intentando borrar '" & ws.name & "' (Visible=" & CStr(ws.Visible) & ")"
    On Error GoTo EH
    ws.Visible = xlSheetVisible: ws.Delete
    DbgAdd "DeleteSheetIfExists: borrada OK '" & sheetName & "'."
    TryDeleteSheetVerbose = True: Exit Function
EH:
    DbgAdd "DeleteSheetIfExists: ERROR al borrar '" & sheetName & "' | " & CStr(Err.Number) & " | " & Err.Description
    TryDeleteSheetVerbose = False
End Function

'======================
' Nombre seguro de hoja
'======================
Private Function SanitizeSheetName(ByVal desired As String) As String
    Dim nm As String
    nm = desired
    nm = Replace(nm, "[", "("): nm = Replace(nm, "]", ")")
    nm = Replace(nm, ":", " - "): nm = Replace(nm, "\", " - ")
    nm = Replace(nm, "/", " - "): nm = Replace(nm, "?", " - ")
    nm = Replace(nm, "*", " - ")
    nm = Trim$(nm)
    If Len(nm) = 0 Then nm = "Hoja"
    If Len(nm) > 31 Then nm = Left$(nm, 31)
    SanitizeSheetName = nm
End Function

Private Sub FreeSheetName(ByVal wb As Workbook, ByVal safeName As String, Optional ByVal exceptSheet As Worksheet)
    Dim ws As Worksheet, base As String, tmp As String, k As Long
    On Error Resume Next: Set ws = wb.Worksheets(safeName): On Error GoTo 0
    If ws Is Nothing Then DbgAdd "FreeSheetName: '" & safeName & "' no existe.": Exit Sub
    If Not exceptSheet Is Nothing Then
        If ws Is exceptSheet Then DbgAdd "FreeSheetName: '" & safeName & "' es la misma hoja destino.": Exit Sub
    End If
    base = Left$(safeName, 20)
    If Len(base) = 0 Then base = "OLD"
    For k = 1 To 50
        tmp = SanitizeSheetName(base & "_OLD_" & Format$(k, "00"))
        On Error Resume Next
        ws.name = tmp
        If Err.Number = 0 Then On Error GoTo 0: Exit Sub
        Err.Clear: On Error GoTo 0
    Next k
End Sub

Private Sub RenameSheetExact(ByVal sh As Worksheet, ByVal desired As String)
    Dim nm As String
    nm = SanitizeSheetName(desired)
    DbgAdd "RenameSheetExact: destino '" & nm & "' | actual '" & sh.name & "'"
    FreeSheetName sh.parent, nm, sh
    On Error GoTo fallback
    sh.name = nm
    DbgAdd "RenameSheetExact: OK '" & sh.name & "'"
    Exit Sub
fallback:
    DbgAdd "RenameSheetExact: ERROR | " & CStr(Err.Number) & " | " & Err.Description
    Err.Clear: RenameSheetSafe sh, nm
End Sub

Private Sub RenameSheetSafe(ByVal sh As Worksheet, ByVal desired As String)
    Dim nm As String, base As String, k As Long, sufX As String, maxBase As Long
    nm = SanitizeSheetName(desired): base = nm: k = 0
    On Error GoTo exists
    sh.name = nm: Exit Sub
exists:
    Err.Clear
    Do
        k = k + 1: sufX = "_" & CStr(k): maxBase = 31 - Len(sufX)
        If maxBase < 1 Then maxBase = 1
        nm = Left$(base, maxBase) & sufX
        On Error GoTo exists
        sh.name = nm: Exit Sub
    Loop
End Sub

'======================
' Columnas texto identidad
'======================
Private Function StripDiacriticsUpper(ByVal s As String) As String
    Dim t As String
    t = s
    t = Replace(t, Chr(193), "A"): t = Replace(t, Chr(192), "A"): t = Replace(t, Chr(194), "A"): t = Replace(t, Chr(196), "A")
    t = Replace(t, Chr(201), "E"): t = Replace(t, Chr(200), "E"): t = Replace(t, Chr(202), "E"): t = Replace(t, Chr(203), "E")
    t = Replace(t, Chr(205), "I"): t = Replace(t, Chr(204), "I"): t = Replace(t, Chr(206), "I"): t = Replace(t, Chr(207), "I")
    t = Replace(t, Chr(211), "O"): t = Replace(t, Chr(210), "O"): t = Replace(t, Chr(212), "O"): t = Replace(t, Chr(214), "O")
    t = Replace(t, Chr(218), "U"): t = Replace(t, Chr(217), "U"): t = Replace(t, Chr(219), "U"): t = Replace(t, Chr(220), "U")
    t = Replace(t, Chr(209), "N")
    StripDiacriticsUpper = t
End Function

Private Function CanonColName(ByVal s As String) As String
    Dim t As String
    t = UCase$(Trim$(s))
    t = Replace(t, Chr$(160), " ")
    t = StripDiacriticsUpper(t)
    t = Replace(t, Chr(176), ""): t = Replace(t, Chr(186), "")
    t = Replace(t, " ", "")
    CanonColName = t
End Function

Private Function FindListColumnByName(ByVal lo As ListObject, ByVal colName As String) As ListColumn
    Dim lc As ListColumn, want As String
    want = CanonColName(colName)
    For Each lc In lo.ListColumns
        If CanonColName(lc.name) = want Then Set FindListColumnByName = lc: Exit Function
    Next lc
    Set FindListColumnByName = Nothing
End Function

Private Sub ForceTextColumnByName(ByVal lo As ListObject, ByVal colName As String)
    Dim lc As ListColumn
    On Error GoTo fin
    If lo Is Nothing Then Exit Sub
    Set lc = FindListColumnByName(lo, colName)
    If lc Is Nothing Then Exit Sub
    lc.Range.NumberFormat = "@"
fin:
End Sub

Private Sub IgnoreNumberAsTextByName(ByVal lo As ListObject, ByVal colName As String)
    Dim lc As ListColumn
    On Error GoTo fin
    If lo Is Nothing Then Exit Sub
    Set lc = FindListColumnByName(lo, colName)
    If lc Is Nothing Then Exit Sub
    On Error Resume Next
    If Not lc.DataBodyRange Is Nothing Then lc.DataBodyRange.Errors(xlNumberAsText).Ignore = True
    lc.Range.Errors(xlNumberAsText).Ignore = True
    On Error GoTo 0
fin:
End Sub

Private Sub ForceTextIdentityColumns(ByVal lo As ListObject)
    ForceTextColumnByName lo, "CUC":                 IgnoreNumberAsTextByName lo, "CUC"
    ForceTextColumnByName lo, "N OP":                IgnoreNumberAsTextByName lo, "N OP"
    ForceTextColumnByName lo, "N CERTIFICADO":       IgnoreNumberAsTextByName lo, "N CERTIFICADO"
    ForceTextColumnByName lo, "NRO OPERACION BANCO": IgnoreNumberAsTextByName lo, "NRO OPERACION BANCO"
    ForceTextColumnByName lo, "NUMERO DE DOCUMENTO": IgnoreNumberAsTextByName lo, "NUMERO DE DOCUMENTO"
End Sub

'======================
' Estilo de tabla
'======================
Private Sub ApplyTableStyle(ByVal lo As ListObject, ByVal styleIndex As Long)
    On Error Resume Next
    If Not lo Is Nothing Then
        lo.TableStyle = "TableStyleLight" & CStr(styleIndex)
    End If
    On Error GoTo 0
End Sub

'======================
' Buscar tabla de clientes en el workbook
' Criterio: ListObject que tenga columnas CUC y NUMERO DE DOCUMENTO
'======================
Private Function FindClientesLO() As ListObject
    Dim ws As Worksheet, lo As ListObject
    Dim tieneCUC As Boolean, tieneDoc As Boolean
    Set FindClientesLO = Nothing
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            tieneCUC = Not FindListColumnByName(lo, "CUC") Is Nothing
            tieneDoc = Not FindListColumnByName(lo, "NUMERO DE DOCUMENTO") Is Nothing
            If tieneCUC And tieneDoc Then
                Set FindClientesLO = lo
                Exit Function
            End If
        Next lo
    Next ws
End Function

'======================
' Fechas para sufijo
'======================
Private Function LastDayOfMonth(ByVal d As Date) As Date
    LastDayOfMonth = DateSerial(Year(d), Month(d) + 1, 0)
End Function

Private Function FirstDayOfMonth(ByVal d As Date) As Date
    FirstDayOfMonth = DateSerial(Year(d), Month(d), 1)
End Function

Private Function TryCoerceExcelDate(ByVal v As Variant, ByRef outD As Date) As Boolean
    Dim n As Double
    On Error GoTo fin
    If IsError(v) Or IsEmpty(v) Then GoTo fin
    If IsDate(v) Then outD = CDate(v): TryCoerceExcelDate = True: Exit Function
    If IsNumeric(v) Then
        n = CDbl(v)
        If n > 0# And n < 60000# Then
            outD = DateSerial(1899, 12, 30) + n: TryCoerceExcelDate = True: Exit Function
        End If
    End If
fin:
    TryCoerceExcelDate = False
End Function

Private Function GetMinMaxDateFromLO(ByVal lo As ListObject, ByVal colName As String, ByRef outMin As Date, ByRef outMax As Date) As Boolean
    Dim lc As ListColumn, c As Range, d As Date, gotAny As Boolean
    On Error GoTo fin
    GetMinMaxDateFromLO = False
    If lo Is Nothing Then Exit Function
    Set lc = FindListColumnByName(lo, colName)
    If lc Is Nothing Then Exit Function
    If lc.DataBodyRange Is Nothing Then Exit Function
    gotAny = False
    For Each c In lc.DataBodyRange.Cells
        If TryCoerceExcelDate(c.Value2, d) Then
            If Not gotAny Then
                outMin = d: outMax = d: gotAny = True
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

Private Function MesAbrevES(ByVal d As Date) As String
    Select Case Month(d)
        Case 1:  MesAbrevES = "ENE": Case 2:  MesAbrevES = "FEB": Case 3:  MesAbrevES = "MAR"
        Case 4:  MesAbrevES = "ABR": Case 5:  MesAbrevES = "MAY": Case 6:  MesAbrevES = "JUN"
        Case 7:  MesAbrevES = "JUL": Case 8:  MesAbrevES = "AGO": Case 9:  MesAbrevES = "SEP"
        Case 10: MesAbrevES = "OCT": Case 11: MesAbrevES = "NOV": Case 12: MesAbrevES = "DIC"
        Case Else: MesAbrevES = "MES"
    End Select
End Function

Private Sub MLine(ByRef m As String, ByVal line As String)
    If Len(m) = 0 Then m = line Else m = m & vbCrLf & line
End Sub

'======================
' Power Query M
'======================
Private Function M_RAW_SUS(ByVal rutaArchivo As String) As String
    Dim m As String
    Dim pathEsc As String
    pathEsc = Replace(rutaArchivo, """", """""""""")
    MLine m, "let"
    MLine m, "  Ruta = """ & pathEsc & ""","
    MLine m, "  Libro = Excel.Workbook(File.Contents(Ruta), null, true),"
    MLine m, "  Keys = {""SUS"",""SUSC"",""SUBSCR"",""SUSCR""},"
    MLine m, "  ConNombre = if Table.HasColumns(Libro, ""Name"") then Libro else Table.AddColumn(Libro, ""Name"", each null, type text),"
    MLine m, "  Preferidas = Table.SelectRows(ConNombre, (r)=> let n = Text.Upper(Text.From(try r[Name] otherwise """")) in List.AnyTrue(List.Transform(Keys, (k)=> Text.Contains(n, k)))),"
    MLine m, "  Base = if Table.RowCount(Preferidas) > 0 then Preferidas{0}[Data] else ConNombre{0}[Data],"
    MLine m, "  PromoTest = Table.ColumnNames(Table.PromoteHeaders(Table.FirstN(Base, 1), [PromoteAllScalars=true])),"
    MLine m, "  IsA1 = List.AnyTrue(List.Transform(PromoTest, each Text.Contains(Text.Upper(Text.From(_)), ""OP""))),"
    MLine m, "  DataReady = if IsA1"
    MLine m, "              then Base"
    MLine m, "              else let"
    MLine m, "                     Skipped = Table.Skip(Base, 10),"
    MLine m, "                     CN      = Table.ColumnNames(Skipped),"
    MLine m, "                     Removed = if List.Count(CN) > 0 then Table.RemoveColumns(Skipped, {List.First(CN)}) else Skipped"
    MLine m, "                   in Removed,"
    MLine m, "  Promoted = Table.PromoteHeaders(DataReady, [PromoteAllScalars=true]),"
    MLine m, "  TargetNames = {"
    MLine m, "    ""N OP"", ""TIPO OPERACION"", ""N CERTIFICADO"", ""CUC"", ""TIPO PERSONA"","
    MLine m, "    ""AGENCIA ORIGEN"", ""ESTADO"", ""FECHA PROCESO"", ""FECHA ABONO DISPONIBLE"","
    MLine m, "    ""FECHA OPERACION"", ""FONDO"", ""MONTO"", ""CUOTAS"", ""VALOR CUOTA"","
    MLine m, "    ""VIA PAGO"", ""FORMA PAGO"", ""PROMOTOR"", ""PROPOSITO"","
    MLine m, "    ""NOMBRE/RAZON SOCIAL PARTICIPE"", ""NRO DE TRASPASO"", ""FONDO ORIGEN"","
    MLine m, "    ""ORIGEN"", ""NRO OPERACION BANCO"", ""MONTO SOLES"""
    MLine m, "  },"
    MLine m, "  ActualCols = Table.ColumnNames(Promoted),"
    MLine m, "  N = List.Min({List.Count(TargetNames), List.Count(ActualCols)}),"
    MLine m, "  Pairs = List.Select("
    MLine m, "    List.Zip({List.FirstN(ActualCols, N), List.FirstN(TargetNames, N)}),"
    MLine m, "    each _{0} <> _{1}"
    MLine m, "  ),"
    MLine m, "  Renamed = if List.Count(Pairs) > 0 then Table.RenameColumns(Promoted, Pairs) else Promoted,"
    MLine m, "  Selected = Table.SelectColumns(Renamed, List.FirstN(TargetNames, N), MissingField.Ignore),"
    MLine m, "  WithID = Table.TransformColumns(Selected, {"
    MLine m, "    {""N OP"", each if _ = null then null else Text.Trim(Text.From(_)), type text},"
    MLine m, "    {""CUC"",  each if _ = null then null else Text.Trim(Text.From(_)), type text}"
    MLine m, "  }, null, MissingField.Ignore),"
    MLine m, "  Typed = Table.TransformColumnTypes(WithID, {"
    MLine m, "    {""FECHA PROCESO"",                  type date},"
    MLine m, "    {""FECHA ABONO DISPONIBLE"",         type date},"
    MLine m, "    {""FECHA OPERACION"",                type date},"
    MLine m, "    {""MONTO"",                          type number},"
    MLine m, "    {""CUOTAS"",                         type number},"
    MLine m, "    {""VALOR CUOTA"",                    type number},"
    MLine m, "    {""MONTO SOLES"",                    type number},"
    MLine m, "    {""N OP"",                           type text},"
    MLine m, "    {""TIPO OPERACION"",                 type text},"
    MLine m, "    {""N CERTIFICADO"",                  type text},"
    MLine m, "    {""CUC"",                            type text},"
    MLine m, "    {""TIPO PERSONA"",                   type text},"
    MLine m, "    {""AGENCIA ORIGEN"",                 type text},"
    MLine m, "    {""ESTADO"",                         type text},"
    MLine m, "    {""FONDO"",                          type text},"
    MLine m, "    {""VIA PAGO"",                       type text},"
    MLine m, "    {""FORMA PAGO"",                     type text},"
    MLine m, "    {""PROMOTOR"",                       type text},"
    MLine m, "    {""PROPOSITO"",                      type text},"
    MLine m, "    {""NOMBRE/RAZON SOCIAL PARTICIPE"",  type text},"
    MLine m, "    {""NRO DE TRASPASO"",                type text},"
    MLine m, "    {""FONDO ORIGEN"",                   type text},"
    MLine m, "    {""ORIGEN"",                         type text},"
    MLine m, "    {""NRO OPERACION BANCO"",            type text}"
    MLine m, "  }),"
    MLine m, "  RAW_SUS = Typed"
    MLine m, "in"
    MLine m, "  RAW_SUS"
    M_RAW_SUS = m
End Function

Private Function M_RAW_RES(ByVal rutaArchivo As String) As String
    Dim m As String
    Dim pathEsc As String
    pathEsc = Replace(rutaArchivo, """", """""""""")
    MLine m, "let"
    MLine m, "  Ruta = """ & pathEsc & ""","
    MLine m, "  Libro = Excel.Workbook(File.Contents(Ruta), null, true),"
    MLine m, "  Keys = {""RES"",""RESC"",""RESCAT""},"
    MLine m, "  ConNombre = if Table.HasColumns(Libro, ""Name"") then Libro else Table.AddColumn(Libro, ""Name"", each null, type text),"
    MLine m, "  Preferidas = Table.SelectRows(ConNombre, (r)=> let n = Text.Upper(Text.From(try r[Name] otherwise """")) in List.AnyTrue(List.Transform(Keys, (k)=> Text.Contains(n, k)))),"
    MLine m, "  Base = if Table.RowCount(Preferidas) > 0 then Preferidas{0}[Data] else ConNombre{0}[Data],"
    MLine m, "  Skip10 = Table.Skip(Base, 10),"
    MLine m, "  CNames = Table.ColumnNames(Skip10),"
    MLine m, "  Pre = if List.Count(CNames) > 0 then Table.RemoveColumns(Skip10, { List.First(CNames) }) else Skip10,"
    MLine m, "  Promoted = Table.PromoteHeaders(Pre, [PromoteAllScalars=true]),"
    MLine m, "  Target = { ""N OP"", ""CUC"", ""TIPOPERSONA"", ""ESTADO"", ""FECHA PROCESO"", ""FECHA OPERACION"", ""FECHA DE PAGO"", ""FONDO"", ""MONTO"", ""MONTO NETO"", ""CUOTAS"", ""VIA SOLICITUD"", ""VIA PAGO"", ""FORMA PAGO"", ""PROMOTOR"", ""PROPOSITO"", ""NOMBRE/RAZON SOCIAL PARTICIPE"", ""USUARIO CREACION"", ""FECHA CREACION"", ""TRASPASO"", ""CUENTA"", ""__UNNAMED1"", ""__UNNAMED2"", ""OPERACION"", ""TIPO FIRMA"", ""RESCATE EN AGENTE COLOCADOR"", ""AGENTE COLOCADOR"", ""MONTO SOLES"", ""FECHA SOLICITUD"", ""VALOR CUOTA"", ""MONTO RETENCION"", ""CUOTAS RETENCION"" },"
    MLine m, "  CN = Table.ColumnNames(Promoted),"
    MLine m, "  N = if List.Count(Target) < List.Count(CN) then List.Count(Target) else List.Count(CN),"
    MLine m, "  Dups = List.Intersect({Target, CN}),"
    MLine m, "  Freed = if List.Count(Dups) > 0 then Table.RenameColumns(Promoted, List.Transform(Dups, each {_, _ & ""__OLD""}), MissingField.Ignore) else Promoted,"
    MLine m, "  CN2 = Table.ColumnNames(Freed),"
    MLine m, "  FirstN = List.FirstN(CN2, N),"
    MLine m, "  Pairs = List.Zip({ FirstN, List.FirstN(Target, N) }),"
    MLine m, "  Renamed = if N > 0 then Table.RenameColumns(Freed, Pairs, MissingField.Ignore) else Freed,"
    MLine m, "  OldCols = List.Select(Table.ColumnNames(Renamed), each Text.EndsWith(_, ""__OLD"")),"
    MLine m, "  Pruned = if List.Count(OldCols) > 0 then Table.RemoveColumns(Renamed, OldCols) else Renamed,"
    MLine m, "  WithID = Table.TransformColumns(Pruned, {{""N OP"", each if _ = null then null else Text.Trim(Text.From(_)), type text}, {""CUC"", each if _ = null then null else Text.Trim(Text.From(_)), type text}}, null, MissingField.Ignore),"
    MLine m, "  Filtered = if List.Contains(Table.ColumnNames(WithID), ""N OP"") then Table.SelectRows(WithID, each let v = Record.FieldOrDefault(_, ""N OP"", null) in v <> null and Text.Trim(Text.From(v)) <> """") else WithID,"
    MLine m, "  Keep = List.Select(Target, each List.Contains(Table.ColumnNames(Filtered), _)),"
    MLine m, "  RAW_RES = if List.Count(Keep) > 0 then Table.SelectColumns(Filtered, Keep, MissingField.Ignore) else Filtered"
    MLine m, "in"
    MLine m, "  RAW_RES"
    M_RAW_RES = m
End Function

Private Function M_SUS(ByVal mesesSel As Long) As String
    Dim m As String
    MLine m, "let"
    MLine m, "  Source = RAW_SUS,"
    MLine m, "  MesesSel = " & CStr(IIf(mesesSel <= 0, 6, mesesSel)) & ","
    MLine m, "  WithTmp = Table.AddColumn(Source, ""__FechaTmp"", each try Date.From([FECHA PROCESO]) otherwise try Date.FromText(Text.From([FECHA PROCESO]), ""es-PE"") otherwise try Date.FromText(Text.From([FECHA PROCESO]), ""en-US"") otherwise null, type date),"
    MLine m, "  DateList = List.RemoveNulls(Table.Column(WithTmp, ""__FechaTmp"")),"
    MLine m, "  FinMes = if List.Count(DateList) > 0 then Date.EndOfMonth(List.Max(DateList)) else Date.EndOfMonth(DateTime.Date(DateTime.LocalNow())),"
    MLine m, "  IniMes = Date.StartOfMonth(Date.AddMonths(FinMes, -(MesesSel - 1))),"
    MLine m, "  F1 = Table.SelectRows(WithTmp, each [__FechaTmp] <> null and [__FechaTmp] >= IniMes and [__FechaTmp] <= FinMes),"
    MLine m, "  F2 = if List.Contains(Table.ColumnNames(F1), ""ESTADO"") then Table.SelectRows(F1, each Text.Upper(Text.Trim(Text.From([ESTADO]))) = ""PRE"") else F1,"
    MLine m, "  F3 = if List.Contains(Table.ColumnNames(F2), ""TIPO OPERACION"") then Table.AddColumn(F2, ""__TO__"", each Text.Upper(Text.Start(Text.Trim(Text.From(Record.Field(_, ""TIPO OPERACION""))), 3))) else Table.AddColumn(F2, ""__TO__"", each ""SUS""),"
    MLine m, "  F4 = Table.SelectRows(F3, each [__TO__] = ""SUS""),"
    MLine m, "  F5 = if List.Contains(Table.ColumnNames(F4), ""NRO DE TRASPASO"") then Table.SelectRows(F4, each Record.Field(_, ""NRO DE TRASPASO"") = null or Text.Trim(Text.From(Record.Field(_, ""NRO DE TRASPASO""))) = """") else F4,"
    MLine m, "  F6 = if List.Contains(Table.ColumnNames(F5), ""TIPO PERSONA"") then Table.TransformColumns(F5, {{""TIPO PERSONA"", each if Text.Upper(Text.Trim(Text.From(_))) = ""MAN"" then ""NAT"" else Text.From(_), type text}}) else F5,"
    MLine m, "  F7 = if List.Contains(Table.ColumnNames(F6), ""TIPO OPERACION"") then Table.RemoveColumns(F6, {""TIPO OPERACION""}) else F6,"
    MLine m, "  TargetFull = { ""N OP"", ""CUC"", ""TIPO PERSONA"", ""ESTADO"", ""FECHA PROCESO"", ""FECHA OPERACION"", ""FECHA ABONO DISPONIBLE"", ""FONDO"", ""MONTO"", ""MONTO SOLES"", ""CUOTAS"", ""VIA PAGO"", ""FORMA PAGO"", ""PROMOTOR"", ""PROPOSITO"", ""NOMBRE/RAZON SOCIAL PARTICIPE"", ""NRO DE TRASPASO"", ""FONDO ORIGEN"", ""ORIGEN"", ""NRO OPERACION BANCO"", ""VALOR CUOTA"" },"
    MLine m, "  Keep = List.Select(TargetFull, each List.Contains(Table.ColumnNames(F7), _)),"
    MLine m, "  Ordered = if List.Count(Keep) > 0 then Table.SelectColumns(F7, Keep, MissingField.Ignore) else F7,"
    MLine m, "  Typed = Table.TransformColumnTypes(Ordered, {{""MONTO"", type number}, {""CUOTAS"", type number}, {""MONTO SOLES"", type number}, {""VALOR CUOTA"", type number}}, ""es-PE""),"
    MLine m, "  ToDrop = List.Intersect({{""__FechaTmp"", ""__TO__""}, Table.ColumnNames(Typed)}),"
    MLine m, "  SUS = if List.Count(ToDrop) > 0 then Table.RemoveColumns(Typed, ToDrop) else Typed"
    MLine m, "in"
    MLine m, "  SUS"
    M_SUS = m
End Function

Private Function M_RES(ByVal mesesSel As Long) As String
    Dim m As String
    MLine m, "let"
    MLine m, "  Source = RAW_RES,"
    MLine m, "  MesesSel = " & CStr(IIf(mesesSel <= 0, 6, mesesSel)) & ","
    MLine m, "  WithTmp = Table.AddColumn(Source, ""__FechaTmp"", each try Date.From([FECHA PROCESO]) otherwise try Date.FromText(Text.From([FECHA PROCESO]), ""es-PE"") otherwise try Date.FromText(Text.From([FECHA PROCESO]), ""en-US"") otherwise null, type date),"
    MLine m, "  DateList = List.RemoveNulls(Table.Column(WithTmp, ""__FechaTmp"")),"
    MLine m, "  FinMes = if List.Count(DateList) > 0 then Date.EndOfMonth(List.Max(DateList)) else Date.EndOfMonth(DateTime.Date(DateTime.LocalNow())),"
    MLine m, "  IniMes = Date.StartOfMonth(Date.AddMonths(FinMes, -(MesesSel - 1))),"
    MLine m, "  F1 = Table.SelectRows(WithTmp, each [__FechaTmp] <> null and [__FechaTmp] >= IniMes and [__FechaTmp] <= FinMes),"
    MLine m, "  F2 = if List.Contains(Table.ColumnNames(F1), ""ESTADO"") then Table.SelectRows(F1, each Text.Upper(Text.Trim(Text.From([ESTADO]))) = ""PRE"") else F1,"
    MLine m, "  F3 = if List.Contains(Table.ColumnNames(F2), ""TRASPASO"") then Table.SelectRows(F2, each Text.Upper(Text.Trim(Text.From([TRASPASO]))) = ""NO"") else F2,"
    MLine m, "  F4 = if List.Contains(Table.ColumnNames(F3), ""TIPOPERSONA"") then Table.RenameColumns(F3, {{""TIPOPERSONA"", ""TIPO PERSONA""}}) else F3,"
    MLine m, "  F5 = if List.Contains(Table.ColumnNames(F4), ""TIPO PERSONA"") then Table.TransformColumns(F4, {{""TIPO PERSONA"", each if Text.Upper(Text.Trim(Text.From(_))) = ""MAN"" then ""NAT"" else Text.From(_), type text}}) else F4,"
    MLine m, "  F6 = Table.TransformColumnTypes(F5, {{""MONTO"", type number}, {""MONTO NETO"", type number}, {""CUOTAS"", type number}, {""MONTO SOLES"", type number}}, ""es-PE""),"
    MLine m, "  Unnamed = List.Select(Table.ColumnNames(F6), each Text.StartsWith(_, ""__UNNAMED"") or Text.StartsWith(_, ""Column"") or Text.StartsWith(_, ""Columna"")),"
    MLine m, "  BaseDrop = {""__FechaTmp"", ""TRASPASO"", ""USUARIO CREACION"", ""FECHA CREACION"", ""CUENTA"", ""OPERACION"", ""TIPO FIRMA"", ""RESCATE EN AGENTE COLOCADOR"", ""AGENTE COLOCADOR""},"
    MLine m, "  ToDrop = List.Intersect({List.Union({BaseDrop, Unnamed}), Table.ColumnNames(F6)}),"
    MLine m, "  Clean = if List.Count(ToDrop) > 0 then Table.RemoveColumns(F6, ToDrop) else F6,"
    MLine m, "  TargetRes = { ""N OP"", ""CUC"", ""TIPO PERSONA"", ""ESTADO"", ""FECHA PROCESO"", ""FECHA OPERACION"", ""FECHA DE PAGO"", ""FONDO"", ""MONTO"", ""MONTO NETO"", ""CUOTAS"", ""VIA SOLICITUD"", ""VIA PAGO"", ""FORMA PAGO"", ""PROMOTOR"", ""PROPOSITO"", ""NOMBRE/RAZON SOCIAL PARTICIPE"", ""MONTO SOLES"", ""FECHA SOLICITUD"", ""VALOR CUOTA"", ""MONTO RETENCION"", ""CUOTAS RETENCION"" },"
    MLine m, "  Keep = List.Select(TargetRes, each List.Contains(Table.ColumnNames(Clean), _)),"
    MLine m, "  RES = if List.Count(Keep) > 0 then Table.SelectColumns(Clean, Keep, MissingField.Ignore) else Clean"
    MLine m, "in"
    MLine m, "  RES"
    M_RES = m
End Function

'======================
' M_ALERTAS
' Si clientesTableName tiene valor, los CUC del mismo NUMERO DE DOCUMENTO
' se consolidan antes de calcular las metricas de riesgo.
' El calculo es:
'   SUMA_MONTOS      = sum de sumas diarias del grupo
'   NUM_OPERACIONES  = count de dias del grupo
'   PROMEDIO_MONTOS  = SUMA_MONTOS / NUM_OPERACIONES  (promedio correcto sobre dias)
'   ULTIMA_OPERACION = monto del dia mas reciente del grupo
'   DESVIACION       = (ULTIMA - PROMEDIO) / PROMEDIO * 100
'   NIVEL_RIESGO     = recomputado sobre DESVIACION del grupo
'======================
Private Function M_ALERTAS(ByVal srcQueryName As String, _
                            Optional ByVal clientesTableName As String = "") As String
    Dim m As String
    Dim useDoc As Boolean
    useDoc = (USE_NUMERO_DOC And Len(clientesTableName) > 0)

    MLine m, "let"
    MLine m, "  Origen = " & srcQueryName & ","
    MLine m, "  EnsureEdad = if List.Contains(Table.ColumnNames(Origen), ""EDAD"") then Origen else Table.AddColumn(Origen, ""EDAD"", each null, Int64.Type),"
    MLine m, "  Typed = Table.TransformColumnTypes(EnsureEdad, {{""MONTO SOLES"", type number}}, ""es-PE""),"
    MLine m, "  WithDate = Table.AddColumn(Typed, ""__Fecha"", each let v = [FECHA PROCESO] in"
    MLine m, "                try Date.From(v) otherwise"
    MLine m, "                try Date.FromText(Text.From(v), ""es-PE"") otherwise"
    MLine m, "                try Date.FromText(Text.From(v), ""en-US"") otherwise null, type date),"

    ' Filtro extendido: excluye filas sin fecha Y sin CUC valido
    MLine m, "  F = Table.SelectRows(WithDate, each [__Fecha] <> null and [CUC] <> null and Text.Length(Text.Trim(Text.From([CUC]))) > 0),"

    If useDoc Then
        MLine m, "  CliRaw = Excel.CurrentWorkbook(){[Name=""" & clientesTableName & """]}[Content],"
        MLine m, "  CliTyped = Table.TransformColumnTypes(CliRaw, {{""CUC"", type text}, {""NUMERO DE DOCUMENTO"", type text}}, ""es-PE""),"
        MLine m, "  CliSel = Table.Distinct(Table.SelectColumns(CliTyped, {""CUC"", ""NUMERO DE DOCUMENTO""}, MissingField.Ignore)),"
        MLine m, "  FConDoc = Table.NestedJoin(F, {""CUC""}, CliSel, {""CUC""}, ""__cli"", JoinKind.LeftOuter),"
        MLine m, "  FExpDoc = Table.ExpandTableColumn(FConDoc, ""__cli"", {""NUMERO DE DOCUMENTO""}, {""NUMERO DE DOCUMENTO""}),"
        MLine m, "  FDoc = Table.TransformColumns(FExpDoc, {{""NUMERO DE DOCUMENTO"", each if _ = null then ""SIN_DOC"" else Text.Trim(Text.From(_)), type text}}),"

        ' Operaciones individuales antes de la agregacion diaria
        MLine m, "  OpsTot = Table.Group(FDoc, {""NUMERO DE DOCUMENTO""}, {{""NUM_OPERACIONES_TOTALES"", each Table.RowCount(_), Int64.Type}}),"

        ' Agregacion diaria por NUMERO DE DOCUMENTO
        MLine m, "  Daily = Table.Group(FDoc, {""NUMERO DE DOCUMENTO"", ""__Fecha""}, {{""MontoDia"", each List.Sum(List.RemoveNulls([MONTO SOLES])), type number}}),"

        ' Metricas por NUMERO DE DOCUMENTO
        MLine m, "  Agg0 = Table.Group(Daily, {""NUMERO DE DOCUMENTO""}, {"
        MLine m, "          {""SUMA_MONTOS"",        each List.Sum(List.RemoveNulls([MontoDia])), type number},"
        MLine m, "          {""NUM_DIAS_OPERACION"", each Table.RowCount(_), Int64.Type},"
        MLine m, "          {""PROMEDIO_MONTOS"",    each try Number.Round(List.Average(List.RemoveNulls([MontoDia])), 2) otherwise null, type number},"
        MLine m, "          {""ULTIMA_OPERACION"",   each let t = Table.Sort(_, {{""__Fecha"", Order.Ascending}}) in try List.Last(t[MontoDia]) otherwise null, type number}"
        MLine m, "        }),"

        ' Join para agregar NUM_OPERACIONES_TOTALES
        MLine m, "  Agg1 = Table.NestedJoin(Agg0, {""NUMERO DE DOCUMENTO""}, OpsTot, {""NUMERO DE DOCUMENTO""}, ""__ops"", JoinKind.LeftOuter),"
        MLine m, "  Agg  = Table.ExpandTableColumn(Agg1, ""__ops"", {""NUM_OPERACIONES_TOTALES""}, {""NUM_OPERACIONES_TOTALES""}),"

        ' TIPO_PERSONA desde el origen
        MLine m, "  MetaDoc = Table.Group(FDoc, {""NUMERO DE DOCUMENTO""}, {"
        MLine m, "              {""TIPO_PERSONA"", each try List.First(List.RemoveNulls([#""TIPO PERSONA""])) otherwise null, type text}"
        MLine m, "            }),"
        MLine m, "  JoinMeta = Table.NestedJoin(Agg, {""NUMERO DE DOCUMENTO""}, MetaDoc, {""NUMERO DE DOCUMENTO""}, ""__meta"", JoinKind.LeftOuter),"
        MLine m, "  Expanded = Table.ExpandTableColumn(JoinMeta, ""__meta"", {""TIPO_PERSONA""}, {""TIPO_PERSONA""}),"

        MLine m, "  WithDesv = Table.AddColumn(Expanded, ""DESVIACION_MEDIA_%"","
        MLine m, "    each let p = [PROMEDIO_MONTOS], u = [ULTIMA_OPERACION] in"
        MLine m, "         if p = null or p = 0 or u = null then null else ((u - p) / p) * 100.0,"
        MLine m, "    type number),"
        MLine m, "  WithNivel = Table.AddColumn(WithDesv, ""NIVEL_RIESGO"","
        MLine m, "    each let d = Record.Field(_, ""DESVIACION_MEDIA_%"") in"
        MLine m, "         if d = null then null else if d < 50 then 1 else if d <= 100 then 2 else 3,"
        MLine m, "    Int64.Type),"
        MLine m, "  Selected = Table.SelectColumns(WithNivel,"
        MLine m, "    {""NUMERO DE DOCUMENTO"", ""TIPO_PERSONA"", ""SUMA_MONTOS"", ""NUM_OPERACIONES_TOTALES"", ""NUM_DIAS_OPERACION"","
        MLine m, "     ""PROMEDIO_MONTOS"", ""ULTIMA_OPERACION"", ""DESVIACION_MEDIA_%"", ""NIVEL_RIESGO""},"
        MLine m, "    MissingField.Ignore),"
    Else
        ' Operaciones individuales antes de la agregacion diaria
        MLine m, "  OpsTot = Table.Group(F, {""CUC""}, {{""NUM_OPERACIONES_TOTALES"", each Table.RowCount(_), Int64.Type}}),"

        ' Agregacion diaria por CUC
        MLine m, "  Daily = Table.Group(F, {""CUC"", ""__Fecha""}, {{""MontoDia"", each List.Sum(List.RemoveNulls([MONTO SOLES])), type number}}),"

        MLine m, "  Agg0 = Table.Group(Daily, {""CUC""}, {"
        MLine m, "          {""SUMA_MONTOS"",        each List.Sum(List.RemoveNulls([MontoDia])), type number},"
        MLine m, "          {""NUM_DIAS_OPERACION"", each Table.RowCount(_), Int64.Type},"
        MLine m, "          {""PROMEDIO_MONTOS"",    each try Number.Round(List.Average(List.RemoveNulls([MontoDia])), 2) otherwise null, type number},"
        MLine m, "          {""ULTIMA_OPERACION"",   each let t = Table.Sort(_, {{""__Fecha"", Order.Ascending}}) in try List.Last(t[MontoDia]) otherwise null, type number}"
        MLine m, "        }),"

        ' Join para agregar NUM_OPERACIONES_TOTALES
        MLine m, "  Agg1 = Table.NestedJoin(Agg0, {""CUC""}, OpsTot, {""CUC""}, ""__ops"", JoinKind.LeftOuter),"
        MLine m, "  Agg  = Table.ExpandTableColumn(Agg1, ""__ops"", {""NUM_OPERACIONES_TOTALES""}, {""NUM_OPERACIONES_TOTALES""}),"

        ' TIPO_PERSONA y EDAD desde el origen
        MLine m, "  Meta = Table.Group(F, {""CUC""}, {"
        MLine m, "          {""EDAD"",        each try List.First(List.RemoveNulls([EDAD])) otherwise null, type any},"
        MLine m, "          {""TIPO_PERSONA"", each try List.First(List.RemoveNulls([#""TIPO PERSONA""])) otherwise null, type text}"
        MLine m, "        }),"
        MLine m, "  JoinMeta = Table.NestedJoin(Agg, {""CUC""}, Meta, {""CUC""}, ""__meta"", JoinKind.LeftOuter),"
        MLine m, "  Expanded = Table.ExpandTableColumn(JoinMeta, ""__meta"", {""EDAD"", ""TIPO_PERSONA""}, {""EDAD"", ""TIPO_PERSONA""}),"

        MLine m, "  WithDesv = Table.AddColumn(Expanded, ""DESVIACION_MEDIA_%"","
        MLine m, "    each let p = [PROMEDIO_MONTOS], u = [ULTIMA_OPERACION] in"
        MLine m, "         if p = null or p = 0 or u = null then null else ((u - p) / p) * 100.0,"
        MLine m, "    type number),"
        MLine m, "  WithNivel = Table.AddColumn(WithDesv, ""NIVEL_RIESGO"","
        MLine m, "    each let d = Record.Field(_, ""DESVIACION_MEDIA_%"") in"
        MLine m, "         if d = null then null else if d < 50 then 1 else if d <= 100 then 2 else 3,"
        MLine m, "    Int64.Type),"
        MLine m, "  Selected = Table.SelectColumns(WithNivel,"
        MLine m, "    {""CUC"", ""EDAD"", ""TIPO_PERSONA"", ""SUMA_MONTOS"", ""NUM_OPERACIONES_TOTALES"", ""NUM_DIAS_OPERACION"","
        MLine m, "     ""PROMEDIO_MONTOS"", ""ULTIMA_OPERACION"", ""DESVIACION_MEDIA_%"", ""NIVEL_RIESGO""},"
        MLine m, "    MissingField.Ignore),"
    End If

    MLine m, "  Sorted = Table.Sort(Selected, {{""DESVIACION_MEDIA_%"", Order.Descending}})"
    MLine m, "in"
    MLine m, "  Sorted"
    M_ALERTAS = m
End Function
'======================
' Conexiones PQ
'======================
Private Function EnsurePQConnection(ByVal queryName As String) As WorkbookConnection
    Dim wb As Workbook, conn As WorkbookConnection, cs As String
    Set wb = ThisWorkbook
    On Error Resume Next
    Set conn = wb.Connections("Consulta - " & queryName)
    If conn Is Nothing Then Set conn = wb.Connections("Query - " & queryName)
    If conn Is Nothing Then Set conn = wb.Connections(queryName)
    On Error GoTo 0
    LdAdd "[EnsurePQConn] '" & queryName & "' conn existente: " & BoolTxt(Not conn Is Nothing)
    If Not conn Is Nothing Then
        LdAdd "  -> Name: " & conn.name & " | Type: " & CStr(conn.Type)
        Set EnsurePQConnection = conn: Exit Function
    End If
    cs = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & queryName & ";Extended Properties="""""
    LdAdd "  -> Creando nueva conn. CS=" & cs
    Set conn = wb.Connections.Add2( _
        name:="Consulta - " & queryName, _
        Description:="", _
        ConnectionString:=cs, _
        CommandText:=queryName, _
        lCmdtype:=xlCmdDefault, _
        CreateModelConnection:=False, _
        ImportRelationships:=False)
    LdAdd "  -> Conn creada: " & BoolTxt(Not conn Is Nothing)
    Set EnsurePQConnection = conn
End Function

Private Function QueryExists(ByVal qName As String) As Boolean
    Dim q As Object
    On Error Resume Next: Set q = ThisWorkbook.Queries.Item(qName): On Error GoTo 0
    QueryExists = Not q Is Nothing
End Function

'======================
' Carga sincronica bloqueante
'======================
Private Function EnsureTableForConnection(ByVal sh As Worksheet, _
                                          ByVal conn As WorkbookConnection, _
                                          ByVal loName As String) As ListObject
    Dim lo As ListObject, qt As QueryTable
    Dim qName As String, cs As String
    Dim nRows As Long, errN As Long, errD As String

    LdAdd "[EnsureTableForConn] loName='" & loName & "' hoja='" & sh.name & "'"
    On Error Resume Next: Set lo = sh.ListObjects(loName): On Error GoTo 0
    If Not lo Is Nothing Then
        LdAdd "  -> ListObject ya existe. Filas=" & IIf(lo.DataBodyRange Is Nothing, "0", CStr(lo.DataBodyRange.Rows.count))
        Set EnsureTableForConnection = lo: Exit Function
    End If

    qName = conn.OLEDBConnection.CommandText
    cs = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & qName & ";Extended Properties="""""
    LdAdd "  -> qName='" & qName & "' | QueryExists=" & BoolTxt(QueryExists(qName))

    UnfreezeForRefresh
    LdAdd "  -> Application desbloqueado para Refresh"

    On Error GoTo EH_Add
    Set qt = sh.ListObjects.Add(SourceType:=xlSrcExternal, Source:=cs, Destination:=sh.Range("A1")).QueryTable
    On Error GoTo 0
    LdAdd "  -> QueryTable creado: " & BoolTxt(Not qt Is Nothing)

    On Error GoTo EH_Cfg
    With qt
        .CommandType = xlCmdDefault
        .CommandText = Array(qName)
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = loName
    End With
    On Error GoTo 0
    LdAdd "  -> Propiedades configuradas."

    LdAdd "  -> Llamando Refresh..."
    On Error GoTo EH_Ref
    qt.Refresh BackgroundQuery:=False
    On Error GoTo 0
    LdAdd "  -> Refresh completado."

    RefreezeAfterRefresh

    Set lo = qt.ListObject
    If lo Is Nothing Then GoTo EH_NoData
    nRows = 0
    If Not lo.DataBodyRange Is Nothing Then nRows = lo.DataBodyRange.Rows.count
    LdAdd "  -> Filas cargadas: " & CStr(nRows)
    On Error Resume Next: lo.name = loName: On Error GoTo 0
    Set EnsureTableForConnection = lo
    Exit Function

EH_Add:
    errN = Err.Number: errD = Err.Description: RefreezeAfterRefresh
    LdAdd "  -> ERROR Add: " & CStr(errN) & " | " & errD
    Err.Raise errN, "EnsureTableForConnection.Add", errD
EH_Cfg:
    errN = Err.Number: errD = Err.Description: RefreezeAfterRefresh
    LdAdd "  -> ERROR Cfg: " & CStr(errN) & " | " & errD
    Err.Raise errN, "EnsureTableForConnection.Cfg", errD
EH_Ref:
    errN = Err.Number: errD = Err.Description: RefreezeAfterRefresh
    LdAdd "  -> ERROR Refresh: " & CStr(errN) & " | " & errD
    Err.Raise errN, "EnsureTableForConnection.Ref", errD
EH_NoData:
    RefreezeAfterRefresh
    Err.Raise vbObjectError + 515, "EnsureTableForConnection", "La consulta '" & loName & "' no produjo datos."
End Function

Private Sub RefreshTableBlocking(ByVal lo As ListObject, ByVal conn As WorkbookConnection)
    LdAdd "[RefreshTableBlocking] lo='" & IIf(lo Is Nothing, "Nothing", lo.name) & "'"
    UnfreezeForRefresh
    On Error Resume Next
    If Not lo Is Nothing Then
        If Not lo.QueryTable Is Nothing Then
            lo.QueryTable.BackgroundQuery = False
            lo.QueryTable.Refresh BackgroundQuery:=False
            RefreezeAfterRefresh
            LdAdd "  -> Refresh via QueryTable. Err=" & CStr(Err.Number)
            On Error GoTo 0: Exit Sub
        End If
    End If
    If Not conn Is Nothing Then
        If conn.Type = xlConnectionTypeOLEDB Then conn.OLEDBConnection.BackgroundQuery = False
        conn.Refresh
        LdAdd "  -> Refresh via conn. Err=" & CStr(Err.Number)
    End If
    On Error GoTo 0
    RefreezeAfterRefresh
End Sub

Private Function HasImportPlaceholder(ByVal lo As ListObject) As Boolean
    Dim r As Range, s As String
    On Error GoTo fin
    HasImportPlaceholder = False
    If lo Is Nothing Then Exit Function
    If lo.DataBodyRange Is Nothing Then Exit Function
    Set r = lo.DataBodyRange.Cells(1, 1)
    s = CStr(r.Value2)
    If InStr(1, s, "Importando", vbTextCompare) > 0 Then
        HasImportPlaceholder = True
        LdAdd "  -> Placeholder: '" & s & "'"
    End If
    Exit Function
fin:
    HasImportPlaceholder = False
End Function

Private Function FreezeListObject(ByVal lo As ListObject) As ListObject
    LdAdd "[FreezeListObject] lo='" & IIf(lo Is Nothing, "Nothing", lo.name) & "'"
    On Error Resume Next
    If Not lo Is Nothing Then
        If Not lo.QueryTable Is Nothing Then
            lo.QueryTable.Delete
            LdAdd "  -> QueryTable desconectada. Err=" & CStr(Err.Number)
        End If
    End If
    On Error GoTo 0
    LdAdd "  -> Filas=" & IIf(lo Is Nothing, "Nothing", _
          IIf(lo.DataBodyRange Is Nothing, "0", CStr(lo.DataBodyRange.Rows.count)))
    Set FreezeListObject = lo
End Function

'======================
' Carga de etapa con orden garantizado
'======================
Private Function EnsureStage(ByVal sh As Worksheet, _
                              ByVal loName As String, _
                              ByVal conn As WorkbookConnection, _
                              ByVal stageLabel As String, _
                              ByVal showProgress As Boolean) As ListObject
    Dim tStage0 As Double, lo As ListObject
    Dim attempt As Long, secStage As Double, msg As String, nRows As Long

    tStage0 = Timer
    LdAdd "": LdAdd "====== STAGE: " & stageLabel & " ======"
    LdAdd "Hoja='" & sh.name & "' loName='" & loName & "'"
    StatusStage stageLabel, tStage0
    DoEvents

    Set lo = EnsureTableForConnection(sh, conn, loName)
    LdAdd "Tras EnsureTableForConnection: " & IIf(lo Is Nothing, "Nothing", "OK")

    If HasImportPlaceholder(lo) Then
        LdAdd "Placeholder detectado, reintentos..."
        For attempt = 1 To 3
            DoEvents: RefreshTableBlocking lo, conn
            If Not HasImportPlaceholder(lo) Then LdAdd "  Resuelto en reintento " & CStr(attempt): Exit For
        Next attempt
    End If

    If HasImportPlaceholder(lo) Then
        Err.Raise vbObjectError + 514, "EnsureStage", "La carga de " & stageLabel & " no termino."
    End If

    ForceTextIdentityColumns lo
    Set lo = FreezeListObject(lo)
    ForceTextIdentityColumns lo

    nRows = 0
    If Not lo Is Nothing Then
        If Not lo.DataBodyRange Is Nothing Then nRows = lo.DataBodyRange.Rows.count
    End If
    LdAdd "Filas finales: " & CStr(nRows)

    secStage = ElapsedSec(tStage0)
    msg = "Carga de " & stageLabel & " completada." & vbCrLf & _
          "Filas: " & CStr(nRows) & vbCrLf & _
          "Tiempo: " & FormatElapsed(secStage) & " (" & Format(secStage, "0.0") & " s)" & vbCrLf & _
          "Total acumulado: " & FormatElapsed(ElapsedSec(mT0Total))

    Application.StatusBar = stageLabel & " listo. " & Format(secStage, "0.0") & " s | Total " & FormatElapsed(ElapsedSec(mT0Total))
    AppendStageLog stageLabel, secStage
    Debug.Print msg
    If showProgress Then MsgBox msg, vbInformation, "Fondos"

    Set EnsureStage = lo
End Function

Private Sub DeleteQueryAndConnection(ByVal qName As String)
    On Error Resume Next
    ThisWorkbook.Queries.Item(qName).Delete
    ThisWorkbook.Connections("Consulta - " & qName).Delete
    On Error GoTo 0
End Sub

'======================
' Principal
'======================
Public Sub CrearQueryFondos(ByVal rutaArchivo As String, ByVal arg2 As Variant, ByVal arg3 As Variant, _
                            Optional ByVal arg4 As Variant, _
                            Optional ByVal arg5 As Variant, _
                            Optional ByVal arg6 As Variant)
    Dim esRescate As Boolean, mesesSel As Long, activar As Boolean
    Dim entidadPrefix As String, showProg As Boolean, opCode As String
    Dim shRaw As Worksheet, shMain As Worksheet, shAL As Worksheet
    Dim connRaw As WorkbookConnection, connMain As WorkbookConnection, connAL As WorkbookConnection
    Dim loRaw As ListObject, loMAIN As ListObject, loAL As ListObject
    Dim loClientes As ListObject, clientesTableName As String
    Dim minD As Date, maxD As Date, gotDates As Boolean
    Dim finD As Date, iniD As Date, suf As String
    Dim nmRaw As String, nmMain As String, nmAL As String
    Dim shNmRaw As String, shNmMain As String, shNmAL As String
    Dim totalMsg As String, errDesc As String, msgErr As String

    On Error GoTo EH

    mT0Total = Timer: mStageLog = vbNullString
DbgReset:     LdReset
    LdAdd "=== CrearQueryFondos INICIO ==="
    LdAdd "Ruta: " & rutaArchivo
    DebugWorkbookStatus ThisWorkbook

    If VarType(arg2) = vbBoolean Then
        esRescate = CBool(arg2): mesesSel = CoerceLong(arg3, 6)
    ElseIf VarType(arg3) = vbBoolean Then
        mesesSel = CoerceLong(arg2, 6): esRescate = CBool(arg3)
    Else
        mesesSel = CoerceLong(arg2, 6): esRescate = CoerceBool(arg3, False)
    End If
    If mesesSel < 1 Then mesesSel = 6

    activar = True: entidadPrefix = "FONDOS": showProg = True
    If Not IsMissing(arg4) Then
        If IsBoolLike(arg4) Then
            activar = CoerceBool(arg4, True)
            If Not IsMissing(arg5) Then
                If Len(CoerceText(arg5)) > 0 Then entidadPrefix = UCase$(CoerceText(arg5))
            End If
            If Not IsMissing(arg6) Then showProg = CoerceBool(arg6, True)
        Else
            If Len(CoerceText(arg4)) > 0 Then entidadPrefix = UCase$(CoerceText(arg4))
            If Not IsMissing(arg5) Then showProg = CoerceBool(arg5, True)
        End If
    End If

    LdAdd "esRescate=" & BoolTxt(esRescate) & " meses=" & CStr(mesesSel) & " prefix=" & entidadPrefix

    ' Buscar tabla de clientes antes de congelar Application
    clientesTableName = vbNullString
    If USE_NUMERO_DOC Then
        Set loClientes = FindClientesLO()
        If Not loClientes Is Nothing Then
            clientesTableName = loClientes.name
            LdAdd "Tabla clientes encontrada: '" & clientesTableName & "'"
        Else
            LdAdd "Tabla clientes NO encontrada. Alertas por CUC."
        End If
    End If

    SafeApp True

    If esRescate Then opCode = "RES" Else opCode = "SUS"

    On Error Resume Next
    ThisWorkbook.Queries.Item("RAW_SUS").Delete
    ThisWorkbook.Queries.Item("SUS").Delete
    ThisWorkbook.Queries.Item("SUS_ALERTAS").Delete
    ThisWorkbook.Queries.Item("RAW_RES").Delete
    ThisWorkbook.Queries.Item("RES").Delete
    ThisWorkbook.Queries.Item("RES_ALERTAS").Delete
    On Error GoTo EH

    LdAdd "Creando queries M..."
    If esRescate Then
        ThisWorkbook.Queries.Add name:="RAW_RES", Formula:=M_RAW_RES(rutaArchivo)
        ThisWorkbook.Queries.Add name:="RES", Formula:=M_RES(mesesSel)
        ThisWorkbook.Queries.Add name:="RES_ALERTAS", Formula:=M_ALERTAS("RES", clientesTableName)
    Else
        ThisWorkbook.Queries.Add name:="RAW_SUS", Formula:=M_RAW_SUS(rutaArchivo)
        ThisWorkbook.Queries.Add name:="SUS", Formula:=M_SUS(mesesSel)
        ThisWorkbook.Queries.Add name:="SUS_ALERTAS", Formula:=M_ALERTAS("SUS", clientesTableName)
    End If
    LdAdd "Queries creadas."

    Set shRaw = EnsureSheet("RAW_WORK"):      ClearSheetButKeepName shRaw
    Set shMain = EnsureSheet("MAIN_WORK"):    ClearSheetButKeepName shMain
    Set shAL = EnsureSheet("ALERTAS_WORK"): ClearSheetButKeepName shAL
    DeleteLegacyGraficoSheets

    If esRescate Then
        Set connRaw = EnsurePQConnection("RAW_RES")
        Set connMain = EnsurePQConnection("RES")
        Set connAL = EnsurePQConnection("RES_ALERTAS")
    Else
        Set connRaw = EnsurePQConnection("RAW_SUS")
        Set connMain = EnsurePQConnection("SUS")
        Set connAL = EnsurePQConnection("SUS_ALERTAS")
    End If

    LdAdd "Iniciando carga de etapas..."
    Set loRaw = EnsureStage(shRaw, "RAW_WORK", connRaw, "RAW", showProg)
    Set loMAIN = EnsureStage(shMain, "MAIN_WORK", connMain, opCode, showProg)
    Set loAL = EnsureStage(shAL, "ALERTAS_WORK", connAL, "ALERTAS", showProg)

    ' Aplicar estilos de tabla
    ApplyTableStyle loRaw, 8
    ApplyTableStyle loMAIN, 9
    ApplyTableStyle loAL, 10

    LdAdd "=== CARGA COMPLETADA ==="
    LdShow "DEBUG Carga PQ - Resultado"

    gotDates = GetMinMaxDateFromLO(loMAIN, "FECHA PROCESO", minD, maxD)
    If Not gotDates Then gotDates = GetMinMaxDateFromLO(loRaw, "FECHA PROCESO", minD, maxD)
    If gotDates Then
        iniD = FirstDayOfMonth(minD): finD = LastDayOfMonth(maxD)
    Else
        finD = DateSerial(Year(Date), Month(Date), 0)
        iniD = DateSerial(Year(finD), Month(finD) - (mesesSel - 1), 1)
    End If

    suf = MesAbrevES(iniD) & "_" & MesAbrevES(finD) & "_" & Year(finD)
    nmRaw = "RAW_" & entidadPrefix & "_" & opCode & "_" & suf
    nmMain = entidadPrefix & "_" & opCode & "_" & suf
    nmAL = entidadPrefix & "_" & opCode & "_ALERTAS_" & suf
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
    SetTableNameSafe ThisWorkbook, loMAIN, nmMain
    SetTableNameSafe ThisWorkbook, loAL, nmAL

    RenameSheetExact shRaw, nmRaw
    RenameSheetExact shMain, nmMain
    RenameSheetExact shAL, nmAL

    If BUILD_GRAFICOS Then
        modFondosGraficos.BuildGraficosAlertasEnHoja loAL, loMAIN, loClientes, opCode
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

    totalMsg = "Proceso terminado." & vbCrLf & vbCrLf & _
               mStageLog & vbCrLf & vbCrLf & _
               "Total: " & FormatElapsed(ElapsedSec(mT0Total))
    Application.StatusBar = "Listo. Total " & FormatElapsed(ElapsedSec(mT0Total))
    Debug.Print totalMsg
    MsgBox totalMsg, vbInformation, "Fondos"

    If activar Then shMain.Activate: shMain.Range("A1").Select

    Exit Sub

EH:
    errDesc = Err.Description
    If Len(Trim$(errDesc)) = 0 Then errDesc = "(sin descripcion)"
    LdAdd "": LdAdd "!!! ERROR FINAL: " & CStr(Err.Number) & " | " & errDesc
    LdShow "DEBUG Carga PQ - ERROR"
    SafeApp False
    msgErr = "CrearQueryFondos fallo." & vbCrLf & "Error " & Err.Number & vbCrLf & errDesc
    Err.Raise Err.Number, "CrearQueryFondos", msgErr
End Sub

'======================
' Coerciones seguras
'======================
Private Function UnwrapValue(ByVal v As Variant) As Variant
    On Error GoTo fin
    If IsObject(v) Then
        If TypeName(v) = "Range" Then
            If v.Cells.CountLarge > 0 Then UnwrapValue = v.Cells(1, 1).Value2: Exit Function
        End If
    End If
fin:
    UnwrapValue = v
End Function

Private Function CoerceText(ByVal v As Variant) As String
    v = UnwrapValue(v)
    If IsError(v) Then CoerceText = vbNullString Else CoerceText = Trim$(CStr(v))
End Function

Private Function IsBoolLike(ByVal v As Variant) As Boolean
    Dim s As String
    v = UnwrapValue(v)
    If VarType(v) = vbBoolean Then IsBoolLike = True: Exit Function
    If IsNumeric(v) Then IsBoolLike = True: Exit Function
    s = UCase$(Trim$(CStr(v)))
    IsBoolLike = (s = "TRUE" Or s = "FALSE" Or s = "VERDADERO" Or s = "FALSO" Or s = "SI" Or s = "NO" Or s = "1" Or s = "0")
End Function

Private Function CoerceBool(ByVal v As Variant, Optional ByVal def As Boolean = False) As Boolean
    Dim s As String
    v = UnwrapValue(v)
    If VarType(v) = vbBoolean Then CoerceBool = CBool(v): Exit Function
    If IsNumeric(v) Then CoerceBool = (CDbl(v) <> 0): Exit Function
    s = UCase$(Trim$(CStr(v)))
    Select Case s
        Case "TRUE", "VERDADERO", "SI", "1": CoerceBool = True
        Case Else:                            CoerceBool = def
    End Select
End Function

Private Function CoerceLong(ByVal v As Variant, Optional ByVal def As Long = 0) As Long
    v = UnwrapValue(v)
    If IsError(v) Or IsEmpty(v) Or Len(Trim$(CStr(v))) = 0 Then CoerceLong = def: Exit Function
    If IsNumeric(v) Then CoerceLong = CLng(CDbl(v)): Exit Function
    On Error GoTo fin
    CoerceLong = CLng(CDbl(v)): Exit Function
fin:
    CoerceLong = def
End Function

