Option Explicit

'==========================
' modPQ_SAB_CM
' Carga y procesa transacciones SAB - Cambio de Moneda.
' Genera hojas MAIN, ALERTAS_COM y ALERTAS_VEN con sufijo de periodo.
'
' Estructura de queries:
'   SAB_CM_RAW          : lee archivo, detecta hoja CambioMon_*, salta 3 filas
'   SAB_CM              : parsea fechas DDMMMYYYY, nombre canónico de columnas
'   SAB_CM_ALERTAS_COM  : agrupa por Documento, filtra Moneda Ori = USD
'   SAB_CM_ALERTAS_VEN  : agrupa por Documento, filtra Moneda Ori = PEN
'
' Punto de entrada: CrearQuerySAB_CM(rutaArchivo, mesesSel, opMode, showProgress)
'   opMode: "AMBOS" | "SOLO_COM" | "SOLO_VEN"
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
    Dim s As Long: If secs < 0 Then secs = 0
    s = CLng(secs)
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
            .BackgroundQuery   = False: .RefreshStyle      = xlOverwriteCells
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
' M queries - Cambio de Moneda
'======================
Private Function M_CM_RAW(ByVal rutaArchivo As String) As String
    Dim m As String
    Dim p As String: p = Replace(rutaArchivo, """", """""""""")

    MLine m, "let"
    MLine m, "  Ruta   = """ & p & ""","
    MLine m, "  Origen = Excel.Workbook(File.Contents(Ruta), null, true),"
    MLine m, "  Candidatas = Table.SelectRows(Origen, each Value.Is([Data], type table)),"
    MLine m, "  Preferidas = Table.SelectRows(Candidatas, each Text.StartsWith(Text.From([Name]), ""CambioMon_"")),"
    MLine m, "  EncabezadosEsperados = {"
    MLine m, "    ""Transac"",""Fecha"",""Cuenta"",""Documento"",""Tipo Persona"",""Tipo Doc"",""OfCta"","
    MLine m, "    ""Como se Entero"",""Ref"",""Moneda Ori"",""Monto Ori"",""Moneda Des"",""Monto Des"","
    MLine m, "    ""TC"",""TCBanco"",""MonExp"",""Total Neto"",""Mon G/P"",""Gan/Per"",""Gan/Per PEN"","
    MLine m, "    ""Cbte"",""Canal"",""Flujo en GAM"",""Confirmacion Correo""},"
    MLine m, "  FnCoincideEstructura = (t as table) as logical =>"
    MLine m, "    let"
    MLine m, "      Saltar3  = Table.Skip(t, 3),"
    MLine m, "      Promover = try Table.PromoteHeaders(Saltar3, [PromoteAllScalars=true]) otherwise null,"
    MLine m, "      TrimCols = if Promover = null then null else Table.TransformColumnNames(Promover, each Text.Trim(_)),"
    MLine m, "      Cols     = if TrimCols = null then {} else Table.ColumnNames(TrimCols),"
    MLine m, "      EsSubset = List.Count(List.Difference(EncabezadosEsperados, Cols)) = 0"
    MLine m, "    in EsSubset,"
    MLine m, "  HojaElegida ="
    MLine m, "    if Table.RowCount(Preferidas) > 0 then"
    MLine m, "      Preferidas{0}"
    MLine m, "    else"
    MLine m, "      let"
    MLine m, "        Marcadas  = Table.AddColumn(Candidatas, ""Match"", each FnCoincideEstructura([Data]), type logical),"
    MLine m, "        Coinciden = Table.SelectRows(Marcadas, each [Match] = true)"
    MLine m, "      in"
    MLine m, "        if Table.RowCount(Coinciden) > 0"
    MLine m, "        then Coinciden{0}"
    MLine m, "        else error ""No se encontro hoja CambioMon_* ni hoja con estructura esperada"","
    MLine m, "  DatosHoja    = HojaElegida[Data],"
    MLine m, "  Saltar3Filas = Table.Skip(DatosHoja, 3),"
    MLine m, "  PromoverEnc  = Table.PromoteHeaders(Saltar3Filas, [PromoteAllScalars=true]),"
    MLine m, "  SAB_CM_RAW   = PromoverEnc"
    MLine m, "in"
    MLine m, "  SAB_CM_RAW"

    M_CM_RAW = m
End Function

Private Function M_CM_MAIN() As String
    Dim m As String

    MLine m, "let"
    MLine m, "  Source   = SAB_CM_RAW,"
    MLine m, "  TrimCols = Table.TransformColumnNames(Source, each Text.Trim(_)),"
    MLine m, "  MonthMap = #table("
    MLine m, "    type table [ES=text, EN=text],"
    MLine m, "    {{""ENE"",""JAN""},{""FEB"",""FEB""},{""MAR"",""MAR""},{""ABR"",""APR""},"
    MLine m, "     {""MAY"",""MAY""},{""JUN"",""JUN""},{""JUL"",""JUL""},{""AGO"",""AUG""},"
    MLine m, "     {""SET"",""SEP""},{""SEP"",""SEP""},{""OCT"",""OCT""},{""NOV"",""NOV""},{""DIC"",""DEC""}}),"
    MLine m, "  ToDate_DDMMMYYYY = (x as any) as nullable date =>"
    MLine m, "    let"
    MLine m, "      t0    = Text.Upper(Text.Trim(Text.From(x))),"
    MLine m, "      t1    = Text.Replace(Text.Replace(Text.Replace(Text.Replace(Text.Replace(t0,""-"",""""),""/"",""""),"" "",""""),""."",""""),""_"",""""),"
    MLine m, "      day   = if Text.Length(t1) >= 9 then Text.Start(t1, 2) else null,"
    MLine m, "      monES = if Text.Length(t1) >= 9 then Text.Range(t1, 2, 3) else null,"
    MLine m, "      year  = if Text.Length(t1) >= 9 then Text.End(t1, 4) else null,"
    MLine m, "      row   = if monES=null then MonthMap else Table.SelectRows(MonthMap, each [ES]=monES),"
    MLine m, "      monEN = if Table.RowCount(row)=1 then row{0}[EN] else monES,"
    MLine m, "      ds    = if day<>null and monEN<>null and year<>null then day & ""-"" & monEN & ""-"" & year else null,"
    MLine m, "      d     = try Date.FromText(ds, ""en-US"") otherwise try Date.From(x) otherwise null"
    MLine m, "    in d,"
    MLine m, "  WithFecha ="
    MLine m, "    if List.Contains(Table.ColumnNames(TrimCols), ""Fecha"")"
    MLine m, "    then Table.TransformColumns(TrimCols, {{""Fecha"", each ToDate_DDMMMYYYY(_), type date}})"
    MLine m, "    else TrimCols,"
    MLine m, "  SAB_CM = WithFecha"
    MLine m, "in"
    MLine m, "  SAB_CM"

    M_CM_MAIN = m
End Function

Private Function M_CM_ALERTAS(ByVal moneda As String) As String
    Dim m As String
    Dim monFilter As String
    Dim monLabel  As String

    If UCase$(moneda) = "COM" Then
        monFilter = "List.Contains({""USD"",""US$"",""$""}, s)"
        monLabel  = "USD"
    Else
        monFilter = "List.Contains({""PEN"",""S/"",""S/.""}, s)"
        monLabel  = "PEN"
    End If

    MLine m, "let"
    MLine m, "  Source    = SAB_CM,"
    MLine m, "  ColMonOri = ""Moneda Ori"","
    MLine m, "  ColTot    = ""Total Neto"","
    MLine m, "  ColFecha  = ""Fecha"","
    MLine m, "  ColDoc    = ""Documento"","
    MLine m, "  TrimCols  = Table.TransformColumnNames(Source, each Text.Trim(_)),"
    MLine m, "  WithMon   = Table.AddColumn("
    MLine m, "    TrimCols, ""__MonOri"","
    MLine m, "    each let"
    MLine m, "      v0 = try Record.Field(_, ColMonOri) otherwise null,"
    MLine m, "      s  = if v0=null then null else Text.Upper(Text.Trim(Text.From(v0))),"
    MLine m, "      mm = if s=null then null"
    MLine m, "           else if List.Contains({""USD"",""US$"",""$""}, s) then ""USD"""
    MLine m, "           else if List.Contains({""PEN"",""S/"",""S/.""}, s) then ""PEN"""
    MLine m, "           else s"
    MLine m, "    in mm,"
    MLine m, "    type text"
    MLine m, "  ),"
    MLine m, "  Filtro    = Table.SelectRows(WithMon, each [__MonOri] = """ & monLabel & """),"
    MLine m, "  EnsureCols ="
    MLine m, "    let cols = Table.ColumnNames(Filtro)"
    MLine m, "    in if List.ContainsAll(cols, {ColTot, ColFecha, ColDoc})"
    MLine m, "       then Filtro"
    MLine m, "       else error ""Faltan columnas requeridas"","
    MLine m, "  Typed     = Table.TransformColumnTypes("
    MLine m, "                EnsureCols,"
    MLine m, "                {{ColTot, type number}, {ColFecha, type date}, {ColDoc, type text}},"
    MLine m, "                ""es-PE""),"
    MLine m, "  F         = Table.SelectRows("
    MLine m, "                Typed,"
    MLine m, "                each Record.Field(_, ColFecha) <> null"
    MLine m, "                 and Record.Field(_, ColDoc)   <> null"
    MLine m, "                 and Record.Field(_, ColTot)   <> null),"
    MLine m, "  Daily     = Table.Group("
    MLine m, "                F, {ColDoc, ColFecha},"
    MLine m, "                {{""MontoDia"", each List.Sum(List.RemoveNulls(Table.Column(_, ColTot))), type number}}),"
    MLine m, "  Agg       = Table.Group("
    MLine m, "                Daily, {ColDoc},"
    MLine m, "                {"
    MLine m, "                  {""SUMA_MONTOS"",      each List.Sum(Table.Column(_, ""MontoDia"")), type number},"
    MLine m, "                  {""NUM_OPERACIONES"",  each Table.RowCount(_), Int64.Type},"
    MLine m, "                  {""PROMEDIO_MONTOS"",  each try Number.Round(List.Average(Table.Column(_, ""MontoDia"")), 2) otherwise null, type number},"
    MLine m, "                  {""ULTIMA_OPERACION"","
    MLine m, "                    each let t = Table.Sort(_, {{ColFecha, Order.Ascending}})"
    MLine m, "                         in  try List.Last(Table.Column(t, ""MontoDia"")) otherwise null,"
    MLine m, "                    type number}"
    MLine m, "                }),"
    MLine m, "  Meta      = Table.Group("
    MLine m, "                F, {ColDoc},"
    MLine m, "                {"
    MLine m, "                  {""EDAD"",         each try List.First(List.RemoveNulls(Table.Column(_, ""EDAD""))) otherwise null, type any},"
    MLine m, "                  {""TIPO_PERSONA"", each try List.First(List.RemoveNulls(Table.Column(_, ""Tipo Persona""))) otherwise null, type text}"
    MLine m, "                }),"
    MLine m, "  JoinAgg   = Table.NestedJoin(Agg, {ColDoc}, Meta, {ColDoc}, ""meta"", JoinKind.LeftOuter),"
    MLine m, "  Expanded  = Table.ExpandTableColumn(JoinAgg, ""meta"", {""EDAD"",""TIPO_PERSONA""}, {""EDAD"",""TIPO_PERSONA""}),"
    MLine m, "  WithDesv  = Table.AddColumn("
    MLine m, "                Expanded, ""DESVIACION_MEDIA_%"","
    MLine m, "                each let p=[PROMEDIO_MONTOS], u=[ULTIMA_OPERACION]"
    MLine m, "                     in if p=null or p=0 then null else ((u - p)/p) * 100,"
    MLine m, "                type number),"
    MLine m, "  WithNivel = Table.AddColumn("
    MLine m, "                WithDesv, ""NIVEL_RIESGO"","
    MLine m, "                each let d = try Record.Field(_, ""DESVIACION_MEDIA_%"") otherwise null"
    MLine m, "                     in if d=null then null else if d < 50 then 1 else if d <= 100 then 2 else 3,"
    MLine m, "                Int64.Type),"
    MLine m, "  Selected  = Table.SelectColumns("
    MLine m, "                WithNivel,"
    MLine m, "                {ColDoc,""EDAD"",""TIPO_PERSONA"",""SUMA_MONTOS"",""NUM_OPERACIONES"","
    MLine m, "                 ""PROMEDIO_MONTOS"",""ULTIMA_OPERACION"",""DESVIACION_MEDIA_%"",""NIVEL_RIESGO""},"
    MLine m, "                MissingField.Ignore),"
    MLine m, "  Sorted    = Table.Sort(Selected, {{""DESVIACION_MEDIA_%"", Order.Descending}})"
    MLine m, "in"
    MLine m, "  Sorted"

    M_CM_ALERTAS = m
End Function

'======================
' Punto de entrada publico
' opMode: "AMBOS" | "SOLO_COM" | "SOLO_VEN"
'======================
Public Sub CrearQuerySAB_CM(ByVal rutaArchivo As String, _
                             ByVal mesesSel As Long, _
                             Optional ByVal opMode As String = "AMBOS", _
                             Optional ByVal showProgress As Boolean = False)
    On Error GoTo EH

    mT0Total  = Timer
    mStageLog = vbNullString

    If mesesSel <= 0 Then mesesSel = 6
    If Len(Trim$(opMode)) = 0 Then opMode = "AMBOS"

    Dim makeCOM As Boolean: makeCOM = (UCase$(opMode) = "AMBOS" Or UCase$(opMode) = "SOLO_COM")
    Dim makeVEN As Boolean: makeVEN = (UCase$(opMode) = "AMBOS" Or UCase$(opMode) = "SOLO_VEN")

    SafeApp True

    ' Upsert queries
    UpsertWorkbookQuery "SAB_CM_RAW", M_CM_RAW(rutaArchivo)
    UpsertWorkbookQuery "SAB_CM",     M_CM_MAIN()
    If makeCOM Then UpsertWorkbookQuery "SAB_CM_ALERTAS_COM", M_CM_ALERTAS("COM")
    If makeVEN Then UpsertWorkbookQuery "SAB_CM_ALERTAS_VEN", M_CM_ALERTAS("VEN")

    ' Hojas de trabajo
    Dim shRaw  As Worksheet: Set shRaw  = EnsureSheet("SAB_CM_RAW_WORK")
    Dim shMain As Worksheet: Set shMain = EnsureSheet("SAB_CM_WORK")

    ClearSheetButKeepName shRaw
    ClearSheetButKeepName shMain

    ' Conexiones y carga de RAW y MAIN
    Dim connRaw  As WorkbookConnection: Set connRaw  = EnsurePQConnection("SAB_CM_RAW")
    Dim connMain As WorkbookConnection: Set connMain = EnsurePQConnection("SAB_CM")

    Dim loRaw  As ListObject: Set loRaw  = EnsureStage(shRaw,  "SAB_CM_RAW", connRaw,  "RAW CM",  showProgress)
    Dim loMain As ListObject: Set loMain = EnsureStage(shMain, "SAB_CM",     connMain, "MAIN CM", showProgress)

    ' Alertas segun opMode
    Dim shAlCom As Worksheet, shAlVen As Worksheet
    Dim loAlCom As ListObject, loAlVen As ListObject

    If makeCOM Then
        Set shAlCom = EnsureSheet("SAB_CM_AL_COM_WORK")
        ClearSheetButKeepName shAlCom
        Dim connAlCom As WorkbookConnection: Set connAlCom = EnsurePQConnection("SAB_CM_ALERTAS_COM")
        Set loAlCom = EnsureStage(shAlCom, "SAB_CM_ALERTAS_COM", connAlCom, "AL_COM", showProgress)
    End If

    If makeVEN Then
        Set shAlVen = EnsureSheet("SAB_CM_AL_VEN_WORK")
        ClearSheetButKeepName shAlVen
        Dim connAlVen As WorkbookConnection: Set connAlVen = EnsurePQConnection("SAB_CM_ALERTAS_VEN")
        Set loAlVen = EnsureStage(shAlVen, "SAB_CM_ALERTAS_VEN", connAlVen, "AL_VEN", showProgress)
    End If

    ' Sufijo de periodo desde loMain
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

    ' Nombres finales
    Dim nmRaw  As String: nmRaw  = SanitizeSheetName("SAB_CM_RAW_" & suf)
    Dim nmMain As String: nmMain = SanitizeSheetName("SAB_CM_"     & suf)

    DeleteSheetIfExists ThisWorkbook, nmRaw
    DeleteSheetIfExists ThisWorkbook, nmMain
    FreeSheetName ThisWorkbook, nmRaw,  shRaw
    FreeSheetName ThisWorkbook, nmMain, shMain
    DeleteAllTablesByName ThisWorkbook, nmRaw
    DeleteAllTablesByName ThisWorkbook, nmMain
    SetTableNameSafe ThisWorkbook, loRaw,  nmRaw
    SetTableNameSafe ThisWorkbook, loMain, nmMain
    RenameSheetExact shRaw,  nmRaw
    RenameSheetExact shMain, nmMain

    If makeCOM Then
        Dim nmAlCom As String: nmAlCom = SanitizeSheetName("SAB_CM_AL_COM_" & suf)
        DeleteSheetIfExists ThisWorkbook, nmAlCom
        FreeSheetName ThisWorkbook, nmAlCom, shAlCom
        DeleteAllTablesByName ThisWorkbook, nmAlCom
        SetTableNameSafe ThisWorkbook, loAlCom, nmAlCom
        RenameSheetExact shAlCom, nmAlCom
    End If

    If makeVEN Then
        Dim nmAlVen As String: nmAlVen = SanitizeSheetName("SAB_CM_AL_VEN_" & suf)
        DeleteSheetIfExists ThisWorkbook, nmAlVen
        FreeSheetName ThisWorkbook, nmAlVen, shAlVen
        DeleteAllTablesByName ThisWorkbook, nmAlVen
        SetTableNameSafe ThisWorkbook, loAlVen, nmAlVen
        RenameSheetExact shAlVen, nmAlVen
    End If

    ' Graficos
    If BUILD_GRAFICOS Then
        If makeCOM And Not loAlCom Is Nothing Then
            modSABGraficos.BuildGraficosCMEnHoja loAlCom, loMain, "COM"
        End If
        If makeVEN And Not loAlVen Is Nothing Then
            modSABGraficos.BuildGraficosCMEnHoja loAlVen, loMain, "VEN"
        End If
    End If

    ' Limpiar queries si corresponde
    If Not KEEP_PQ_QUERIES Then
        DeleteQueryAndConnection "SAB_CM_RAW"
        DeleteQueryAndConnection "SAB_CM"
        If makeCOM Then DeleteQueryAndConnection "SAB_CM_ALERTAS_COM"
        If makeVEN Then DeleteQueryAndConnection "SAB_CM_ALERTAS_VEN"
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
