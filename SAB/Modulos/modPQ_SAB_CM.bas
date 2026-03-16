Option Explicit

'==========================
' modPQ_SAB_CM
' Carga y procesa transacciones SAB - Cambio de Moneda.
' Genera hojas MAIN, ALERTAS_COM y ALERTAS_VEN con sufijo de periodo.
' Punto de entrada: CrearQuerySAB_CM(rutaArchivo, mesesSel, showProgress)
'
' Estructura de datos CM:
'   RAW  : archivo fuente, hoja unica, saltar filas de encabezado institucional
'   MAIN : detalle con columnas Documento, Fecha, Tipo Persona,
'          Moneda Ori, Total Neto / Monto Ori / Monto Des / Gan/Per PEN
'   ALERTAS_COM : agrupado por Documento + TIPO_PERSONA, operaciones en USD
'   ALERTAS_VEN : agrupado por Documento + TIPO_PERSONA, operaciones en PEN
'==========================

Private Const KEEP_PQ_QUERIES  As Boolean = True
Private Const BUILD_GRAFICOS   As Boolean = True
Private Const TABLE_STYLE      As String  = "TableStyleLight9"

Private mAppFrozen             As Boolean
Private mPrevScreenUpdating    As Boolean
Private mPrevEnableEvents      As Boolean
Private mPrevDisplayAlerts     As Boolean
Private mPrevCalculation       As XlCalculation
Private mPrevStatusBar         As Variant
Private mT0Total               As Double
Private mStageLog              As String

'======================
' Estado Application
'======================
Private Sub SafeApp(ByVal freeze As Boolean)
    On Error Resume Next
    With Application
        If freeze Then
            If Not mAppFrozen Then
                mPrevScreenUpdating = .ScreenUpdating
                mPrevEnableEvents   = .EnableEvents
                mPrevDisplayAlerts  = .DisplayAlerts
                mPrevCalculation    = .Calculation
                mPrevStatusBar      = .StatusBar
                mAppFrozen          = True
            End If
            .ScreenUpdating = False
            .EnableEvents   = False
            .DisplayAlerts  = False
            .Calculation    = xlCalculationManual
        Else
            If mAppFrozen Then
                .ScreenUpdating = mPrevScreenUpdating
                .EnableEvents   = mPrevEnableEvents
                .DisplayAlerts  = mPrevDisplayAlerts
                .Calculation    = mPrevCalculation
                .StatusBar      = mPrevStatusBar
                mAppFrozen      = False
            Else
                .StatusBar = False
            End If
        End If
    End With
    On Error GoTo 0
End Sub

'======================
' Tiempo
'======================
Private Function ElapsedSec(ByVal t0 As Double) As Double
    Dim t As Double: t = Timer
    If t < t0 Then t = t + 86400#
    ElapsedSec = t - t0
End Function

Private Function FormatElapsed(ByVal secs As Double) As String
    Dim s As Long: If secs < 0 Then secs = 0: s = CLng(secs)
    Dim hh As Long: hh = s \ 3600
    Dim mm As Long: mm = (s \ 60) Mod 60
    Dim ss As Long: ss = s Mod 60
    If hh > 0 Then
        FormatElapsed = Format$(hh, "00") & ":" & Format$(mm, "00") & ":" & Format$(ss, "00")
    Else
        FormatElapsed = Format$(mm, "00") & ":" & Format$(ss, "00")
    End If
End Function

Private Sub StatusStage(ByVal label As String, ByVal t0 As Double)
    Application.StatusBar = "Cargando " & label & "... " & FormatElapsed(ElapsedSec(t0)) & _
                            IIf(mT0Total > 0, " | Total " & FormatElapsed(ElapsedSec(mT0Total)), "")
End Sub

Private Sub AppendStageLog(ByVal label As String, ByVal sec As Double)
    Dim line As String: line = label & ": " & FormatElapsed(sec) & " (" & Format(sec, "0.0") & " s)"
    If Len(mStageLog) = 0 Then mStageLog = line Else mStageLog = mStageLog & vbCrLf & line
End Sub

'======================
' Hojas
'======================
Private Function EnsureSheet(ByVal nm As String) As Worksheet
    Dim sh As Worksheet
    On Error Resume Next: Set sh = ThisWorkbook.Worksheets(nm): On Error GoTo 0
    If sh Is Nothing Then
        Set sh = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.Count))
        sh.Name = nm
    End If
    Set EnsureSheet = sh
End Function

Private Sub ClearSheetButKeepName(ByVal sh As Worksheet)
    Dim lo As ListObject, qt As QueryTable, co As ChartObject
    On Error Resume Next
    For Each co In sh.ChartObjects: co.Delete: Next co
    For Each lo In sh.ListObjects:  lo.Delete:  Next lo
    For Each qt In sh.QueryTables:  qt.Delete:  Next qt
    sh.Cells.Clear
    On Error GoTo 0
End Sub

Private Function SanitizeSheetName(ByVal desired As String) As String
    Dim nm As String
    nm = Replace(desired, "[", "("): nm = Replace(nm, "]", ")")
    nm = Replace(nm, ":", " - "):    nm = Replace(nm, "\", " - ")
    nm = Replace(nm, "/", " - "):    nm = Replace(nm, "?", " - ")
    nm = Replace(nm, "*", " - "):    nm = Trim$(nm)
    If Len(nm) = 0 Then nm = "Hoja"
    If Len(nm) > 31 Then nm = Left$(nm, 31)
    SanitizeSheetName = nm
End Function

Private Sub FreeSheetName(ByVal wb As Workbook, ByVal safeName As String, _
                           Optional ByVal exceptSheet As Worksheet = Nothing)
    Dim ws As Worksheet
    On Error Resume Next: Set ws = wb.Worksheets(safeName): On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    If Not exceptSheet Is Nothing Then If ws Is exceptSheet Then Exit Sub
    Dim base As String: base = Left$(safeName, 20): If Len(base) = 0 Then base = "OLD"
    Dim k As Long, tmp As String
    For k = 1 To 50
        tmp = SanitizeSheetName(base & "_OLD_" & Format$(k, "00"))
        On Error Resume Next: ws.Name = tmp
        If Err.Number = 0 Then On Error GoTo 0: Exit Sub
        Err.Clear: On Error GoTo 0
    Next k
End Sub

Private Sub RenameSheetExact(ByVal sh As Worksheet, ByVal desired As String)
    Dim nm As String: nm = SanitizeSheetName(desired)
    FreeSheetName sh.Parent, nm, sh
    On Error Resume Next: sh.Name = nm: On Error GoTo 0
End Sub

Private Sub DeleteSheetIfExists(ByVal wb As Workbook, ByVal sheetName As String)
    Dim ws As Worksheet
    On Error Resume Next: Set ws = wb.Worksheets(sheetName): On Error GoTo 0
    If Not ws Is Nothing Then On Error Resume Next: ws.Delete: On Error GoTo 0
End Sub

Private Sub DeleteAllTablesByName(ByVal wb As Workbook, ByVal tableName As String)
    Dim ws As Worksheet, lo As ListObject
    For Each ws In wb.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, tableName, vbTextCompare) = 0 Then
                On Error Resume Next: lo.Delete: On Error GoTo 0
            End If
        Next lo
    Next ws
End Sub

Private Sub SetTableNameSafe(ByVal wb As Workbook, ByVal lo As ListObject, _
                              ByVal desiredName As String)
    If Len(Trim$(desiredName)) = 0 Then Exit Sub
    On Error Resume Next: lo.Name = desiredName
    If Err.Number = 0 Then On Error GoTo 0: Exit Sub
    Err.Clear: On Error GoTo 0
    Dim k As Long, nm As String
    For k = 2 To 50
        nm = desiredName & "_" & CStr(k)
        On Error Resume Next: lo.Name = nm
        If Err.Number = 0 Then On Error GoTo 0: Exit Sub
        Err.Clear: On Error GoTo 0
    Next k
End Sub

'======================
' Fechas y sufijo
'======================
Private Function FirstDayOfMonth(ByVal d As Date) As Date
    FirstDayOfMonth = DateSerial(Year(d), Month(d), 1)
End Function

Private Function LastDayOfMonth(ByVal d As Date) As Date
    LastDayOfMonth = DateSerial(Year(d), Month(d) + 1, 0)
End Function

Private Function MesAbrevES(ByVal d As Date) As String
    Select Case Month(d)
        Case 1:  MesAbrevES = "ENE": Case 2:  MesAbrevES = "FEB"
        Case 3:  MesAbrevES = "MAR": Case 4:  MesAbrevES = "ABR"
        Case 5:  MesAbrevES = "MAY": Case 6:  MesAbrevES = "JUN"
        Case 7:  MesAbrevES = "JUL": Case 8:  MesAbrevES = "AGO"
        Case 9:  MesAbrevES = "SEP": Case 10: MesAbrevES = "OCT"
        Case 11: MesAbrevES = "NOV": Case 12: MesAbrevES = "DIC"
        Case Else: MesAbrevES = "MES"
    End Select
End Function

Private Function TryCoerceExcelDate(ByVal v As Variant, ByRef outD As Date) As Boolean
    On Error GoTo fin
    If IsError(v) Or IsEmpty(v) Then GoTo fin
    If IsDate(v) Then outD = CDate(v): TryCoerceExcelDate = True: Exit Function
    If IsNumeric(v) Then
        Dim n As Double: n = CDbl(v)
        If n > 0 And n < 60000 Then
            outD = DateSerial(1899, 12, 30) + n
            TryCoerceExcelDate = True: Exit Function
        End If
    End If
fin:
    TryCoerceExcelDate = False
End Function

Private Function GetMinMaxDateFromLO(ByVal lo As ListObject, ByVal colName As String, _
                                     ByRef outMin As Date, ByRef outMax As Date) As Boolean
    GetMinMaxDateFromLO = False
    If lo Is Nothing Then Exit Function
    Dim lc As ListColumn
    On Error Resume Next: Set lc = lo.ListColumns(colName): On Error GoTo 0
    If lc Is Nothing Then Exit Function
    If lc.DataBodyRange Is Nothing Then Exit Function
    Dim c As Range, d As Date, gotAny As Boolean
    For Each c In lc.DataBodyRange.Cells
        If TryCoerceExcelDate(c.Value2, d) Then
            If Not gotAny Then outMin = d: outMax = d: gotAny = True
            Else: If d < outMin Then outMin = d: If d > outMax Then outMax = d
            End If
        End If
    Next c
    GetMinMaxDateFromLO = gotAny
End Function

'======================
' Power Query helpers
'======================
Private Sub MLine(ByRef buf As String, ByVal s As String)
    If buf = "" Then buf = s Else buf = buf & vbCrLf & s
End Sub

Private Sub UpsertWorkbookQuery(ByVal qName As String, ByVal mFormula As String)
    Dim q As WorkbookQuery
    On Error Resume Next: Set q = ThisWorkbook.Queries.Item(qName): On Error GoTo 0
    If q Is Nothing Then
        ThisWorkbook.Queries.Add Name:=qName, Formula:=mFormula
    Else
        q.Formula = mFormula
    End If
End Sub

Private Function EnsurePQConnection(ByVal queryName As String) As WorkbookConnection
    Dim wb As Workbook: Set wb = ThisWorkbook
    Dim conn As WorkbookConnection
    On Error Resume Next
    Set conn = wb.Connections("Consulta - " & queryName)
    If conn Is Nothing Then Set conn = wb.Connections("Query - " & queryName)
    If conn Is Nothing Then Set conn = wb.Connections("PQ_" & queryName)
    If conn Is Nothing Then Set conn = wb.Connections(queryName)
    On Error GoTo 0
    If conn Is Nothing Then
        Dim cs As String
        cs = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;" & _
             "Location=" & queryName & ";Extended Properties="""""""
        On Error Resume Next
        Set conn = wb.Connections.Add2( _
            Name:="Consulta - " & queryName, Description:="", _
            ConnectionString:=cs, CommandText:=queryName, _
            lCmdtype:=xlCmdSql, CreateModelConnection:=False, _
            ImportRelationships:=False)
        On Error GoTo 0
    End If
    Set EnsurePQConnection = conn
End Function

Private Sub RefreshConnectionSync(ByVal conn As WorkbookConnection)
    If conn Is Nothing Then Exit Sub
    On Error Resume Next
    If conn.Type = xlConnectionTypeOLEDB Then
        conn.OLEDBConnection.BackgroundQuery = False
        conn.OLEDBConnection.Refresh
    Else
        conn.Refresh
    End If
    Application.CalculateUntilAsyncQueriesDone
    On Error GoTo 0
End Sub

Private Function EnsureTableForConnection(ByVal sh As Worksheet, _
                                           ByVal conn As WorkbookConnection, _
                                           ByVal loName As String) As ListObject
    Dim lo As ListObject
    On Error Resume Next: Set lo = sh.ListObjects(loName): On Error GoTo 0
    If lo Is Nothing Then
        Set lo = sh.ListObjects.Add(SourceType:=xlSrcExternal, Source:=conn, _
                                    LinkSource:=True, XlListObjectHasHeaders:=xlYes, _
                                    Destination:=sh.Range("A1"))
        On Error Resume Next: lo.Name = loName: On Error GoTo 0
    End If
    On Error Resume Next
    If Not lo.QueryTable Is Nothing Then
        With lo.QueryTable
            .BackgroundQuery   = False: .RefreshStyle     = xlOverwriteCells
            .AdjustColumnWidth = True:  .PreserveColumnInfo = True
            .Refresh BackgroundQuery:=False
        End With
    End If
    Application.CalculateUntilAsyncQueriesDone
    On Error GoTo 0
    On Error Resume Next: lo.TableStyle = TABLE_STYLE: On Error GoTo 0
    Set EnsureTableForConnection = lo
End Function

Private Function EnsureStage(ByVal sh As Worksheet, ByVal loName As String, _
                              ByVal conn As WorkbookConnection, _
                              ByVal stageLabel As String, _
                              ByVal showProgress As Boolean) As ListObject
    Dim t0 As Double: t0 = Timer
    StatusStage stageLabel, t0: DoEvents
    Dim lo As ListObject
    Set lo = EnsureTableForConnection(sh, conn, loName)
    RefreshConnectionSync conn
    Dim sec As Double: sec = ElapsedSec(t0)
    Application.StatusBar = stageLabel & " listo. " & Format(sec, "0.0") & " s"
    AppendStageLog stageLabel, sec
    Set EnsureStage = lo
End Function

Private Sub DeleteQueryAndConnection(ByVal qName As String)
    On Error Resume Next
    ThisWorkbook.Queries.Item(qName).Delete
    ThisWorkbook.Connections("Consulta - " & qName).Delete
    ThisWorkbook.Connections("PQ_" & qName).Delete
    On Error GoTo 0
End Sub

'======================
' M queries (CM = Cambio de Moneda)
' Estructura del archivo fuente SAB CM:
'   - Excel con una hoja de datos
'   - Encabezados institucionales en las primeras filas (numero variable)
'   - Columnas: Documento, Fecha, Tipo Persona, Moneda Ori,
'               Total Neto / Monto Ori / Monto Des / Gan/Per PEN / Gan/Per
'======================
Private Function M_CM_RAW(ByVal rutaArchivo As String) As String
    Dim m As String
    Dim pathEsc As String: pathEsc = Replace(rutaArchivo, """", """""""""")

    MLine m, "let"
    MLine m, "  Ruta   = """ & pathEsc & ""","
    MLine m, "  Libro  = Excel.Workbook(File.Contents(Ruta), null, true),"
    MLine m, "  Sheet0 = Libro{0}[Data],"
    MLine m, ""
    MLine m, "  // Detectar fila de encabezados buscando 'Documento'"
    MLine m, "  ColCount = Table.ColumnCount(Sheet0),"
    MLine m, "  RowCount  = Table.RowCount(Sheet0),"
    MLine m, "  HdrIdx = List.First(List.Select(List.Numbers(0, Number.Min(RowCount, 30)),"
    MLine m, "    (i) => List.AnyTrue(List.Transform(Record.ToList(Sheet0{i}),"
    MLine m, "      (v) => try Text.Upper(Text.Trim(Text.From(v))) = ""DOCUMENTO"" otherwise false))), null),"
    MLine m, "  HdrRow   = if HdrIdx = null then 0 else HdrIdx,"
    MLine m, "  Skipped  = Table.Skip(Sheet0, HdrRow),"
    MLine m, "  Promoted = Table.PromoteHeaders(Skipped, [PromoteAllScalars=true]),"
    MLine m, "  TrimCols = Table.TransformColumnNames(Promoted, each Text.Trim(_)),"
    MLine m, "  CN       = Table.ColumnNames(TrimCols),"
    MLine m, "  NonEmpty = Table.SelectColumns(TrimCols,"
    MLine m, "               List.Select(CN, (c) => List.NonNullCount(Table.Column(TrimCols, c)) > 0),"
    MLine m, "               MissingField.Ignore),"
    MLine m, ""
    MLine m, "  // Parseo de montos: eliminar simbolos de moneda, convertir separadores"
    MLine m, "  ToNum = (v as any) as nullable number =>"
    MLine m, "    let"
    MLine m, "      t0 = try Text.From(v) otherwise null,"
    MLine m, "      t1 = if t0 = null then null else Text.Trim(t0),"
    MLine m, "      t2 = if t1 = null then null else"
    MLine m, "             Text.Replace(Text.Replace(Text.Replace(Text.Replace(t1,"
    MLine m, "               ""S/"", """"), ""$"", """"), ""USD"", """"), "" "", """"),"
    MLine m, "      hasDot = if t2 = null then false else Text.Contains(t2, "".""),"
    MLine m, "      hasCom = if t2 = null then false else Text.Contains(t2, "",""),"
    MLine m, "      t3 = if t2 = null then null"
    MLine m, "           else if hasDot and hasCom then"
    MLine m, "                  Text.Replace(Text.Replace(t2, ""."", """"), "","", ""."")"
    MLine m, "           else if hasCom and not hasDot then Text.Replace(t2, "","", ""."")"
    MLine m, "           else t2,"
    MLine m, "      n = try Number.FromText(t3, ""en-US"") otherwise try Number.From(t3) otherwise null"
    MLine m, "    in n,"
    MLine m, ""
    MLine m, "  NumCols = {""Total Neto"", ""Monto Ori"", ""Monto Des"","
    MLine m, "             ""Gan/Per PEN"", ""Gan/Per"", ""Gan Per PEN"", ""Gan Per""},"
    MLine m, "  Numd = List.Accumulate(NumCols, NonEmpty,"
    MLine m, "           (st, c) => if List.Contains(Table.ColumnNames(st), c)"
    MLine m, "                      then Table.TransformColumns(st, {{c, each ToNum(_), type number}})"
    MLine m, "                      else st),"
    MLine m, ""
    MLine m, "  SAB_CM_RAW = Numd"
    MLine m, "in"
    MLine m, "  SAB_CM_RAW"

    M_CM_RAW = m
End Function

Private Function M_CM_MAIN(ByVal mesesSel As Long) As String
    Dim m As String
    If mesesSel <= 0 Then mesesSel = 6

    MLine m, "let"
    MLine m, "  Source = SAB_CM_RAW,"
    MLine m, "  CN0    = Table.ColumnNames(Source),"
    MLine m, "  Pick   = (alts as list) as nullable text =>"
    MLine m, "    List.First(List.Select(alts, each List.Contains(CN0, _)), null),"
    MLine m, ""
    MLine m, "  C_Doc  = Pick({""Documento""}),"
    MLine m, "  C_Fec  = Pick({""Fecha"", ""FECHA"", ""Fecha Oper"", ""Fecha Operacion""}),"
    MLine m, "  C_TP   = Pick({""Tipo Persona"", ""TIPO PERSONA"", ""TipoPersona""}),"
    MLine m, "  C_Mon  = Pick({""Moneda Ori"", ""Moneda Origen"", ""Moneda""}),"
    MLine m, "  C_Tot  = Pick({""Total Neto"", ""TotalNeto""}),"
    MLine m, "  C_MOri = Pick({""Monto Ori"", ""Monto Origen""}),"
    MLine m, "  C_MDes = Pick({""Monto Des"", ""Monto Destino""}),"
    MLine m, "  C_GPEN = Pick({""Gan/Per PEN"", ""Gan Per PEN"", ""Ganancia/Perdida PEN""}),"
    MLine m, "  C_G    = Pick({""Gan/Per"", ""Gan Per""}),"
    MLine m, ""
    MLine m, "  RenPairs = List.RemoveNulls({"
    MLine m, "    if C_Doc  <> null then {C_Doc,  ""Documento""}    else null,"
    MLine m, "    if C_Fec  <> null then {C_Fec,  ""Fecha""}        else null,"
    MLine m, "    if C_TP   <> null then {C_TP,   ""Tipo Persona""} else null,"
    MLine m, "    if C_Mon  <> null then {C_Mon,  ""Moneda Ori""}   else null,"
    MLine m, "    if C_Tot  <> null then {C_Tot,  ""Total Neto""}   else null,"
    MLine m, "    if C_MOri <> null then {C_MOri, ""Monto Ori""}    else null,"
    MLine m, "    if C_MDes <> null then {C_MDes, ""Monto Des""}    else null,"
    MLine m, "    if C_GPEN <> null then {C_GPEN, ""Gan/Per PEN""}  else null,"
    MLine m, "    if C_G    <> null then {C_G,    ""Gan/Per""}      else null"
    MLine m, "  }),"
    MLine m, "  Ren = if List.Count(RenPairs) > 0"
    MLine m, "        then Table.RenameColumns(Source, RenPairs, MissingField.Ignore)"
    MLine m, "        else Source,"
    MLine m, ""
    MLine m, "  // Parsear fechas con soporte ES (dd-MMM-yy) y fallback ISO"
    MLine m, "  MonthMap = #table(type table [ES=text, EN=text], {"
    MLine m, "    {""ENE"",""JAN""},{""FEB"",""FEB""},{""MAR"",""MAR""},{""ABR"",""APR""},"
    MLine m, "    {""MAY"",""MAY""},{""JUN"",""JUN""},{""JUL"",""JUL""},{""AGO"",""AUG""},"
    MLine m, "    {""SET"",""SEP""},{""SEP"",""SEP""},{""OCT"",""OCT""},{""NOV"",""NOV""},{""DIC"",""DEC""}}),"
    MLine m, "  ToDateES = (x as any) as nullable date =>"
    MLine m, "    let"
    MLine m, "      d0  = try Date.From(x) otherwise null,"
    MLine m, "      t0  = if d0 <> null then null"
    MLine m, "            else try Text.Upper(Text.Trim(Text.From(x))) otherwise null,"
    MLine m, "      t1  = if t0 = null then null"
    MLine m, "            else if Text.Contains(t0, ""-"") then t0"
    MLine m, "            else if Text.Length(t0) >= 9"
    MLine m, "                 then Text.Start(t0,2) & ""-"" & Text.Range(t0,2,3) & ""-"" & Text.Range(t0,5,4)"
    MLine m, "                 else t0,"
    MLine m, "      monES = if t1 = null then null else Text.Range(t1, 3, 3),"
    MLine m, "      row   = if monES = null then #table(type table [ES=text, EN=text], {})"
    MLine m, "              else Table.SelectRows(MonthMap, each [ES] = monES),"
    MLine m, "      monEN = if Table.RowCount(row) = 1 then row{0}[EN] else monES,"
    MLine m, "      tEN   = if t1 = null then null"
    MLine m, "              else Text.Start(t1,2) & ""-"" & monEN & ""-"" & Text.End(t1,4),"
    MLine m, "      d1    = if d0 <> null then d0"
    MLine m, "              else try Date.FromText(tEN, ""en-US"") otherwise null"
    MLine m, "    in d1,"
    MLine m, ""
    MLine m, "  WithFecha = if List.Contains(Table.ColumnNames(Ren), ""Fecha"")"
    MLine m, "              then Table.TransformColumns(Ren, {{""Fecha"", each ToDateES(_), type date}})"
    MLine m, "              else Ren,"
    MLine m, ""
    MLine m, "  // Filtrar por rango de meses desde el maximo de la data"
    MLine m, "  Dates  = if List.Contains(Table.ColumnNames(WithFecha), ""Fecha"")"
    MLine m, "           then List.RemoveNulls(Table.Column(WithFecha, ""Fecha""))"
    MLine m, "           else {},"
    MLine m, "  FinMes = if List.Count(Dates) > 0"
    MLine m, "           then Date.EndOfMonth(List.Max(Dates))"
    MLine m, "           else Date.EndOfMonth(DateTime.Date(DateTime.LocalNow())),"
    MLine m, "  IniMes = Date.StartOfMonth(Date.AddMonths(FinMes, -" & CStr(mesesSel - 1) & ")),"
    MLine m, "  Fil    = if List.Contains(Table.ColumnNames(WithFecha), ""Fecha"")"
    MLine m, "           then Table.SelectRows(WithFecha,"
    MLine m, "                  each [Fecha] <> null and [Fecha] >= IniMes and [Fecha] <= FinMes)"
    MLine m, "           else WithFecha,"
    MLine m, ""
    MLine m, "  SAB_CM_MAIN = Table.Sort(Fil, {{""Fecha"", Order.Ascending}})"
    MLine m, "in"
    MLine m, "  SAB_CM_MAIN"

    M_CM_MAIN = m
End Function

Private Function M_CM_ALERTAS(ByVal which As String) As String
    Dim m As String
    Dim op As String: op = UCase$(Trim$(which))
    If op <> "COM" And op <> "VEN" Then op = "COM"

    ' Monto: usar la primera columna disponible en orden de prioridad
    Dim montoExpr As String
    montoExpr = "let cols = Table.ColumnNames(_)," & vbCrLf & _
                "        c = List.First(List.Select({""Total Neto"",""Monto Des"",""Monto Ori"",""Gan/Per PEN"",""Gan/Per""}, each List.Contains(cols, _)), null)" & vbCrLf & _
                "    in if c = null then null else List.Sum(List.RemoveNulls(Table.Column(_, c)))"

    Dim monFiltro As String
    If op = "COM" Then
        monFiltro = "let m = Text.Upper(Text.Trim(Text.From([Moneda Ori]))) in" & vbCrLf & _
                    "        Text.Contains(m, ""USD"") or Text.Contains(m, ""DOLAR"") or m = ""$"""
    Else
        monFiltro = "let m = Text.Upper(Text.Trim(Text.From([Moneda Ori]))) in" & vbCrLf & _
                    "        Text.Contains(m, ""PEN"") or Text.Contains(m, ""SOL"") or Text.Contains(m, ""S/"")"
    End If

    MLine m, "let"
    MLine m, "    Origen0  = SAB_CM_MAIN,"
    MLine m, ""
    MLine m, "    // Filtrar por moneda"
        ' placeholder - replaced below
    End If
    MLine m, "    OrigenMon = if List.Contains(Table.ColumnNames(Origen0), ""Moneda Ori"")"
    MLine m, "                then Table.SelectRows(Origen0, each " & monFiltro & ")"
    MLine m, "                else Origen0,"
    MLine m, ""
    MLine m, "    // Normalizar tipo persona"
    MLine m, "    NormTP = (s as nullable text) as text =>"
    MLine m, "      let t = Text.Upper(Text.Trim(if s = null then """" else s))"
    MLine m, "      in if Text.Contains(t, ""NAT"") or t = ""PN"" then ""NATURAL"""
    MLine m, "         else if Text.Contains(t, ""JUR"") or t = ""PJ"" then ""JURIDICA"""
    MLine m, "         else t,"
    MLine m, ""
    MLine m, "    WithTP = if List.Contains(Table.ColumnNames(OrigenMon), ""Tipo Persona"")"
    MLine m, "             then Table.TransformColumns(OrigenMon, {{""Tipo Persona"", NormTP, type text}})"
    MLine m, "             else OrigenMon,"
    MLine m, ""
    MLine m, "    // Parsear fecha"
    MLine m, "    WithDate = Table.AddColumn(WithTP, ""__Fecha"","
    MLine m, "      each try Date.From([Fecha]) otherwise"
    MLine m, "           try Date.FromText(Text.From([Fecha]), ""es-PE"") otherwise"
    MLine m, "           try Date.FromText(Text.From([Fecha]), ""en-US"") otherwise null,"
    MLine m, "      type date),"
    MLine m, ""
    MLine m, "    F = Table.SelectRows(WithDate,"
    MLine m, "          each [__Fecha] <> null"
    MLine m, "              and [Documento] <> null"
    MLine m, "              and Text.Trim(Text.From([Documento])) <> """"),"
    MLine m, ""
    MLine m, "    // Columna de monto por prioridad"
    MLine m, "    CN_F   = Table.ColumnNames(F),"
    MLine m, "    MtoCol = List.First(List.Select("
    MLine m, "               {""Total Neto"",""Monto Des"",""Monto Ori"",""Gan/Per PEN"",""Gan/Per""},"
    MLine m, "               each List.Contains(CN_F, _)), null),"
    MLine m, "    WithMto = if MtoCol = null then F"
    MLine m, "              else Table.AddColumn(F, ""__Monto"", each"
    MLine m, "                     try Record.Field(_, MtoCol) otherwise null, type number),"
    MLine m, ""
    MLine m, "    // Agregar diario por Documento + Tipo Persona"
    MLine m, "    Daily = Table.Group(WithMto, {""Documento"", ""Tipo Persona"", ""__Fecha""},"
    MLine m, "              {{""MontoDia"", each List.Sum(List.RemoveNulls([__Monto])), type number}}),"
    MLine m, ""
    MLine m, "    Agg = Table.Group(Daily, {""Documento"", ""Tipo Persona""},"
    MLine m, "            {"
    MLine m, "              {""SUMA_MONTOS"",      each List.Sum(List.RemoveNulls([MontoDia])), type number},"
    MLine m, "              {""NUM_OPERACIONES"",  each Table.RowCount(_), Int64.Type},"
    MLine m, "              {""PROMEDIO_MONTOS"",  each try Number.Round(List.Average(List.RemoveNulls([MontoDia])), 2) otherwise null, type number},"
    MLine m, "              {""ULTIMA_OPERACION"", each let t = Table.Sort(_, {{""__Fecha"", Order.Ascending}}) in try List.Last(t[MontoDia]) otherwise null, type number}"
    MLine m, "            }),"
    MLine m, ""
    MLine m, "    WithDesv = Table.AddColumn(Agg, ""DESVIACION_MEDIA_%"","
    MLine m, "      each let p = [PROMEDIO_MONTOS], u = [ULTIMA_OPERACION]"
    MLine m, "           in if p = null or p = 0 or u = null then null else ((u - p) / p) * 100.0,"
    MLine m, "      type number),"
    MLine m, ""
    MLine m, "    WithNivel = Table.AddColumn(WithDesv, ""NIVEL_RIESGO"","
    MLine m, "      each let d = [#""DESVIACION_MEDIA_%""]"
    MLine m, "           in if d = null then null else if d < 50 then 1 else if d <= 100 then 2 else 3,"
    MLine m, "      Int64.Type),"
    MLine m, ""
    MLine m, "    // Renombrar Tipo Persona a TIPO_PERSONA para consistencia"
    MLine m, "    Renamed = Table.RenameColumns(WithNivel, {{""Tipo Persona"", ""TIPO_PERSONA""}}, MissingField.Ignore),"
    MLine m, ""
    MLine m, "    Selected = Table.SelectColumns(Renamed,"
    MLine m, "      {""Documento"", ""TIPO_PERSONA"", ""SUMA_MONTOS"", ""NUM_OPERACIONES"","
    MLine m, "       ""PROMEDIO_MONTOS"", ""ULTIMA_OPERACION"", ""DESVIACION_MEDIA_%"", ""NIVEL_RIESGO""},"
    MLine m, "      MissingField.Ignore),"
    MLine m, ""
    MLine m, "    Sorted = Table.Sort(Selected, {{""DESVIACION_MEDIA_%"", Order.Descending}})"
    MLine m, "in"
    MLine m, "    Sorted"

    M_CM_ALERTAS = m
End Function

' Placeholder para resolver referencia en tiempo de compilacion
End Function

'======================
' Punto de entrada publico
'======================
Public Sub CrearQuerySAB_CM(ByVal rutaArchivo As String, _
                             ByVal mesesSel As Long, _
                             Optional ByVal showProgress As Boolean = False)
    On Error GoTo EH

    mT0Total  = Timer
    mStageLog = vbNullString

    If mesesSel <= 0 Then mesesSel = 6

    SafeApp True

    UpsertWorkbookQuery "SAB_CM_RAW",         M_CM_RAW(rutaArchivo)
    UpsertWorkbookQuery "SAB_CM_MAIN",        M_CM_MAIN(mesesSel)
    UpsertWorkbookQuery "SAB_CM_ALERTAS_COM", M_CM_ALERTAS("COM")
    UpsertWorkbookQuery "SAB_CM_ALERTAS_VEN", M_CM_ALERTAS("VEN")

    Dim shRaw   As Worksheet: Set shRaw   = EnsureSheet("SAB_CM_RAW_WORK")
    Dim shMain  As Worksheet: Set shMain  = EnsureSheet("SAB_CM_MAIN_WORK")
    Dim shAlCom As Worksheet: Set shAlCom = EnsureSheet("SAB_CM_AL_COM_WORK")
    Dim shAlVen As Worksheet: Set shAlVen = EnsureSheet("SAB_CM_AL_VEN_WORK")

    ClearSheetButKeepName shRaw
    ClearSheetButKeepName shMain
    ClearSheetButKeepName shAlCom
    ClearSheetButKeepName shAlVen

    Dim connRaw   As WorkbookConnection: Set connRaw   = EnsurePQConnection("SAB_CM_RAW")
    Dim connMain  As WorkbookConnection: Set connMain  = EnsurePQConnection("SAB_CM_MAIN")
    Dim connAlCom As WorkbookConnection: Set connAlCom = EnsurePQConnection("SAB_CM_ALERTAS_COM")
    Dim connAlVen As WorkbookConnection: Set connAlVen = EnsurePQConnection("SAB_CM_ALERTAS_VEN")

    Dim loRaw   As ListObject: Set loRaw   = EnsureStage(shRaw,   "SAB_CM_RAW",         connRaw,   "RAW",    showProgress)
    Dim loMain  As ListObject: Set loMain  = EnsureStage(shMain,  "SAB_CM_MAIN",        connMain,  "MAIN",   showProgress)
    Dim loAlCom As ListObject: Set loAlCom = EnsureStage(shAlCom, "SAB_CM_ALERTAS_COM", connAlCom, "AL_COM", showProgress)
    Dim loAlVen As ListObject: Set loAlVen = EnsureStage(shAlVen, "SAB_CM_ALERTAS_VEN", connAlVen, "AL_VEN", showProgress)

    Dim minD As Date, maxD As Date, gotDates As Boolean
    gotDates = GetMinMaxDateFromLO(loMain, "Fecha", minD, maxD)
    If Not gotDates Then gotDates = GetMinMaxDateFromLO(loRaw, "Fecha", minD, maxD)

    Dim ini As Date, fin As Date, suf As String
    If gotDates Then
        ini = FirstDayOfMonth(minD): fin = LastDayOfMonth(maxD)
    Else
        fin = DateSerial(Year(Date), Month(Date), 0)
        ini = DateSerial(Year(fin), Month(fin) - (mesesSel - 1), 1)
    End If
    suf = MesAbrevES(ini) & "_" & MesAbrevES(fin) & "_" & Year(fin)

    Dim nmRaw   As String: nmRaw   = SanitizeSheetName("SAB_CM_RAW_"     & suf)
    Dim nmMain  As String: nmMain  = SanitizeSheetName("SAB_CM_"         & suf)
    Dim nmAlCom As String: nmAlCom = SanitizeSheetName("SAB_CM_AL_COM_"  & suf)
    Dim nmAlVen As String: nmAlVen = SanitizeSheetName("SAB_CM_AL_VEN_"  & suf)

    DeleteSheetIfExists ThisWorkbook, nmRaw
    DeleteSheetIfExists ThisWorkbook, nmMain
    DeleteSheetIfExists ThisWorkbook, nmAlCom
    DeleteSheetIfExists ThisWorkbook, nmAlVen

    FreeSheetName ThisWorkbook, nmRaw,   shRaw
    FreeSheetName ThisWorkbook, nmMain,  shMain
    FreeSheetName ThisWorkbook, nmAlCom, shAlCom
    FreeSheetName ThisWorkbook, nmAlVen, shAlVen

    DeleteAllTablesByName ThisWorkbook, nmRaw
    DeleteAllTablesByName ThisWorkbook, nmMain
    DeleteAllTablesByName ThisWorkbook, nmAlCom
    DeleteAllTablesByName ThisWorkbook, nmAlVen

    SetTableNameSafe ThisWorkbook, loRaw,   nmRaw
    SetTableNameSafe ThisWorkbook, loMain,  nmMain
    SetTableNameSafe ThisWorkbook, loAlCom, nmAlCom
    SetTableNameSafe ThisWorkbook, loAlVen, nmAlVen

    RenameSheetExact shRaw,   nmRaw
    RenameSheetExact shMain,  nmMain
    RenameSheetExact shAlCom, nmAlCom
    RenameSheetExact shAlVen, nmAlVen

    If BUILD_GRAFICOS Then
        modSABGraficos.BuildGraficosCMEnHoja loAlCom, loMain, "COM"
        modSABGraficos.BuildGraficosCMEnHoja loAlVen, loMain, "VEN"
    End If

    If Not KEEP_PQ_QUERIES Then
        DeleteQueryAndConnection "SAB_CM_RAW"
        DeleteQueryAndConnection "SAB_CM_MAIN"
        DeleteQueryAndConnection "SAB_CM_ALERTAS_COM"
        DeleteQueryAndConnection "SAB_CM_ALERTAS_VEN"
    End If

    SafeApp False

    Dim totalMsg As String
    totalMsg = "SAB - Cambio de Moneda cargado." & vbCrLf & vbCrLf & _
               mStageLog & vbCrLf & vbCrLf & _
               "Total: " & FormatElapsed(ElapsedSec(mT0Total))

    Application.StatusBar = "SAB CM listo. Total " & FormatElapsed(ElapsedSec(mT0Total))
    Debug.Print totalMsg

    If showProgress Then MsgBox totalMsg, vbInformation, "SAB CM"

    shMain.Activate
    shMain.Range("A1").Select
    Exit Sub

EH:
    SafeApp False
    MsgBox "Error en CrearQuerySAB_CM: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

