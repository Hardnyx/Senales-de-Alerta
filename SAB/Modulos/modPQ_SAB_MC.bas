Option Explicit

'==========================
' modPQ_SAB_MC
' Carga y procesa transacciones SAB - Movimiento de Caja.
' Genera hojas RAW, MAIN, ALERTAS_DEP y ALERTAS_RET con sufijo de periodo.
' Punto de entrada: CrearQuerySAB_MC(rutaArchivo, mesesSel, showProgress)
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
    Dim t As Double
    t = Timer
    If t < t0 Then t = t + 86400#
    ElapsedSec = t - t0
End Function

Private Function FormatElapsed(ByVal secs As Double) As String
    Dim s As Long
    If secs < 0 Then secs = 0
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
    If mT0Total <= 0 Then
        Application.StatusBar = "Cargando " & label & "... " & FormatElapsed(ElapsedSec(t0))
    Else
        Application.StatusBar = "Cargando " & label & "... " & FormatElapsed(ElapsedSec(t0)) & _
                                " | Total " & FormatElapsed(ElapsedSec(mT0Total))
    End If
End Sub

Private Sub AppendStageLog(ByVal label As String, ByVal sec As Double)
    Dim line As String
    line = label & ": " & FormatElapsed(sec) & " (" & Format(sec, "0.0") & " s)"
    If Len(mStageLog) = 0 Then
        mStageLog = line
    Else
        mStageLog = mStageLog & vbCrLf & line
    End If
End Sub

'======================
' Hojas
'======================
Private Function EnsureSheet(ByVal nm As String) As Worksheet
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
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
    nm = Replace(desired, "[", "(")
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

Private Sub FreeSheetName(ByVal wb As Workbook, ByVal safeName As String, _
                           Optional ByVal exceptSheet As Worksheet = Nothing)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(safeName)
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    If Not exceptSheet Is Nothing Then
        If ws Is exceptSheet Then Exit Sub
    End If
    Dim base As String: base = Left$(safeName, 20)
    If Len(base) = 0 Then base = "OLD"
    Dim k As Long, tmp As String
    For k = 1 To 50
        tmp = SanitizeSheetName(base & "_OLD_" & Format$(k, "00"))
        On Error Resume Next
        ws.Name = tmp
        If Err.Number = 0 Then On Error GoTo 0: Exit Sub
        Err.Clear
        On Error GoTo 0
    Next k
End Sub

Private Sub RenameSheetExact(ByVal sh As Worksheet, ByVal desired As String)
    Dim nm As String: nm = SanitizeSheetName(desired)
    FreeSheetName sh.Parent, nm, sh
    On Error Resume Next
    sh.Name = nm
    On Error GoTo 0
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
    On Error Resume Next
    lo.Name = desiredName
    If Err.Number = 0 Then On Error GoTo 0: Exit Sub
    Err.Clear: On Error GoTo 0
    Dim k As Long, nm As String
    For k = 2 To 50
        nm = desiredName & "_" & CStr(k)
        On Error Resume Next
        lo.Name = nm
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
        Case 1:  MesAbrevES = "ENE"
        Case 2:  MesAbrevES = "FEB"
        Case 3:  MesAbrevES = "MAR"
        Case 4:  MesAbrevES = "ABR"
        Case 5:  MesAbrevES = "MAY"
        Case 6:  MesAbrevES = "JUN"
        Case 7:  MesAbrevES = "JUL"
        Case 8:  MesAbrevES = "AGO"
        Case 9:  MesAbrevES = "SEP"
        Case 10: MesAbrevES = "OCT"
        Case 11: MesAbrevES = "NOV"
        Case 12: MesAbrevES = "DIC"
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
            TryCoerceExcelDate = True
            Exit Function
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
    On Error Resume Next
    Set lc = lo.ListColumns(colName)
    On Error GoTo 0
    If lc Is Nothing Then Exit Function
    If lc.DataBodyRange Is Nothing Then Exit Function
    Dim c As Range, d As Date, gotAny As Boolean
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
End Function

'======================
' Power Query helpers
'======================
Private Sub MLine(ByRef buf As String, ByVal s As String)
    If buf = "" Then buf = s Else buf = buf & vbCrLf & s
End Sub

Private Sub UpsertWorkbookQuery(ByVal qName As String, ByVal mFormula As String)
    Dim q As WorkbookQuery
    On Error Resume Next
    Set q = ThisWorkbook.Queries.Item(qName)
    On Error GoTo 0
    If q Is Nothing Then
        ThisWorkbook.Queries.Add Name:=qName, Formula:=mFormula
    Else
        q.Formula = mFormula
    End If
End Sub

Private Function EnsurePQConnection(ByVal queryName As String) As WorkbookConnection
    Dim conn As WorkbookConnection
    Dim connName As String: connName = "PQ_" & queryName
    On Error Resume Next
    Set conn = ThisWorkbook.Connections(connName)
    On Error GoTo 0
    If conn Is Nothing Then
        Dim cs  As String: cs  = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & queryName & ";Extended Properties=" & Chr$(34) & Chr$(34)
        Dim cmd As String: cmd = "SELECT * FROM [" & queryName & "]"
        On Error Resume Next
        Set conn = ThisWorkbook.Connections.Add2(connName, "", cs, cmd, xlCmdSql)
        If conn Is Nothing Then Set conn = ThisWorkbook.Connections.Add(connName, "", cs, cmd, xlCmdSql)
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
                                           ByVal loName As String, _
                                           ByVal conn As WorkbookConnection) As ListObject
    Dim lo As ListObject
    ' Borrar tabla previa si existe (igual que en modPQ_SAB_TipoCambio)
    On Error Resume Next
    Set lo = sh.ListObjects(loName)
    On Error GoTo 0
    If Not lo Is Nothing Then
        On Error Resume Next: lo.Delete: On Error GoTo 0
        Set lo = Nothing
    End If
    ' Crear tabla vinculada a la conexion.
    ' NO se hace Refresh aqui: RefreshConnectionSync ya cargo la data antes de llamar
    ' a esta funcion. Un segundo Refresh duplicaria el tiempo de carga.
    Set lo = sh.ListObjects.Add(SourceType:=xlSrcExternal, Source:=conn, _
                                LinkSource:=True, XlListObjectHasHeaders:=xlYes, _
                                Destination:=sh.Range("A1"))
    On Error Resume Next
    lo.Name = loName
    If Not lo.QueryTable Is Nothing Then
        lo.QueryTable.BackgroundQuery  = False
        lo.QueryTable.RefreshStyle     = xlOverwriteCells
        lo.QueryTable.AdjustColumnWidth = True
        lo.QueryTable.PreserveColumnInfo = True
    End If
    lo.TableStyle = TABLE_STYLE
    On Error GoTo 0
    Set EnsureTableForConnection = lo
End Function

Private Function EnsureStage(ByVal sh As Worksheet, ByVal loName As String, _
                              ByVal conn As WorkbookConnection, _
                              ByVal stageLabel As String, _
                              ByVal showProgress As Boolean) As ListObject
    Dim t0 As Double: t0 = Timer
    StatusStage stageLabel, t0
    DoEvents
    ' Refrescar conexion ANTES de crear la tabla (orden del original)
    RefreshConnectionSync conn
    Dim lo As ListObject
    Set lo = EnsureTableForConnection(sh, loName, conn)
    Dim secStage As Double: secStage = ElapsedSec(t0)
    Application.StatusBar = stageLabel & " listo. " & Format(secStage, "0.0") & " s"
    AppendStageLog stageLabel, secStage
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
' M queries (MC = Movimiento de Caja)
'======================
Private Function M_MC_RAW(ByVal rutaArchivo As String) As String
    Dim m As String
    Dim pathEsc As String
    pathEsc = Replace(rutaArchivo, """", """""""""")

    MLine m, "let"
    MLine m, "  Ruta = """ & pathEsc & ""","
    MLine m, "  Libro = Excel.Workbook(File.Contents(Ruta), null, true),"
    MLine m, "  Base0 = Libro{0}[Data],"
    MLine m, "  Skip = Table.Skip(Base0, 10),"
    MLine m, "  Promoted = Table.PromoteHeaders(Skip, [PromoteAllScalars=true]),"
    MLine m, "  TrimCols = Table.TransformColumnNames(Promoted, each Text.Trim(_)),"
    MLine m, "  NonEmptyCols = Table.SelectColumns(TrimCols, List.Select(Table.ColumnNames(TrimCols), (c)=> List.NonNullCount(Table.Column(TrimCols, c))>0 ), MissingField.Ignore),"
    MLine m, "  CN = Table.ColumnNames(NonEmptyCols),"
    MLine m, "  ColFecha = if List.Contains(CN, ""Fecha"") then ""Fecha"" else if List.Contains(CN, ""FECHA"") then ""FECHA"" else if List.Contains(CN, ""Fec"") then ""Fec"" else CN{0},"
    MLine m, "  WithFechaTxt = Table.AddColumn(NonEmptyCols, ""__FechaTxt"", each Text.Upper(Text.Trim(Text.From(Record.Field(_, ColFecha)))), type text),"
    MLine m, "  MonTmp = Table.AddColumn(WithFechaTxt, ""Moneda"", each let s=[__FechaTxt], pL=Text.PositionOf(s, ""("", Occurrence.Last), pR=Text.PositionOf(s, "")"", Occurrence.Last) in if pL>=0 and pR>pL then Text.Middle(s, pL+1, pR-pL-1) else null, type text),"
    MLine m, "  MonUp = Table.TransformColumns(MonTmp, {{""Moneda"", each if _=null then null else Text.Upper(Text.Trim(_)), type text}}),"
    MLine m, "  MonFill = Table.FillDown(MonUp, {""Moneda""}),"
    MLine m, "  Filtrado = Table.SelectRows(MonFill, each let s=[__FechaTxt] in not Text.StartsWith(s, ""TOTAL"") and not (Text.Contains(s, ""("") and Text.Contains(s, "")""))),"
    MLine m, "  ToNum = (v as any) as nullable number =>"
    MLine m, "    let"
    MLine m, "      t0 = try Text.From(v) otherwise null,"
    MLine m, "      t1 = if t0=null then null else Text.Trim(t0),"
    MLine m, "      t2 = if t1=null then null else Text.Replace(Text.Replace(Text.Replace(t1, ""S/"", """"), ""$"", """"), "" "", """"),"
    MLine m, "      hasDot = if t2=null then false else Text.Contains(t2, "".""),"
    MLine m, "      hasCom = if t2=null then false else Text.Contains(t2, "",""),"
    MLine m, "      t3 = if t2=null then null else if hasDot and hasCom then Text.Replace(Text.Replace(t2, ""."", """"), "","", ""."") else if hasCom and not hasDot then Text.Replace(t2, "","", ""."") else t2,"
    MLine m, "      n  = try Number.FromText(t3, ""en-US"") otherwise try Number.From(t3) otherwise null"
    MLine m, "    in n,"
    MLine m, "  MaybeNumCols = {""Dep" & Chr(243) & "sito"", ""Deposito"", ""Retiro"", ""Saldo"", ""Monto"", ""Abono"", ""Cargo""},"
    MLine m, "  Numd = List.Accumulate(MaybeNumCols, Filtrado, (st,c)=> if List.Contains(Table.ColumnNames(st), c) then Table.TransformColumns(st, {{c, each ToNum(_), type number}}) else st),"
    MLine m, "  SAB_MC_RAW = Table.RemoveColumns(Numd, {""__FechaTxt""})"
    MLine m, "in"
    MLine m, "  SAB_MC_RAW"

    M_MC_RAW = m
End Function

Private Function M_MC_MAIN(ByVal mesesSel As Long) As String
    Dim m As String
    If mesesSel <= 0 Then mesesSel = 6

    MLine m, "let"
    MLine m, "  Source = SAB_MC_RAW,"
    MLine m, "  CN0 = Table.ColumnNames(Source),"
    MLine m, "  Pick = (alts as list) as nullable text => let hit = List.First(List.Select(alts, each List.Contains(CN0, _)), null) in hit,"

    MLine m, "  C_Fecha  = Pick({""Fecha"",""FECHA"",""Fec"",""FECHA MOV"",""Fecha Mov""}),"
    MLine m, "  C_Trans  = Pick({""Transac"",""TRANSAC"",""Transacci" & Chr(243) & "n"",""Transaccion""}),"
    MLine m, "  C_Cuenta = Pick({""Cuenta"",""CUENTA"",""Cta"",""Nro Cuenta"",""Nro. Cuenta"",""N" & Chr(186) & " Cuenta""}),"
    MLine m, "  C_Nombre = Pick({""Nombre"",""A La Orden"",""ALaOrden"",""A_la_Orden""}),"
    MLine m, "  C_Ope    = Pick({""Ope"",""OPE""}),"
    MLine m, "  C_Tipo   = Pick({""Tipo"",""TIPO""}),"
    MLine m, "  C_FPag   = Pick({""FPag"",""F. Pag."",""F. Pago"",""Fecha Pago""}),"
    MLine m, "  C_Clase  = Pick({""Clase"",""CLASE""}),"
    MLine m, "  C_ALaOr  = Pick({""ALaOrden"",""A La Orden"",""Nombre""}),"
    MLine m, "  C_Dep    = Pick({""Dep" & Chr(243) & "sito"",""Deposito"",""Abono""}),"
    MLine m, "  C_Ret    = Pick({""Retiro"",""Cargo""}),"
    MLine m, "  C_CtaLiq = Pick({""CtaLiq"",""Cta Liq"",""Cta Liquidez"",""Cuenta Liquidaci" & Chr(243) & "n"",""Cuenta Liquidacion""}),"
    MLine m, "  C_Est    = Pick({""Estado"",""ESTADO""}),"
    MLine m, "  C_Obs    = Pick({""Observaciones"",""Obs""}),"
    MLine m, "  C_Mon    = Pick({""Moneda""}),"

    MLine m, "  RenPairs = List.RemoveNulls({"
    MLine m, "    if C_Fecha<>null  then {C_Fecha,  ""Fecha""}         else null,"
    MLine m, "    if C_Trans<>null  then {C_Trans,  ""Transac""}       else null,"
    MLine m, "    if C_Cuenta<>null then {C_Cuenta, ""Cuenta""}        else null,"
    MLine m, "    if C_Nombre<>null then {C_Nombre, ""Nombre""}        else null,"
    MLine m, "    if C_Ope<>null    then {C_Ope,    ""Ope""}           else null,"
    MLine m, "    if C_Tipo<>null   then {C_Tipo,   ""Tipo""}          else null,"
    MLine m, "    if C_FPag<>null   then {C_FPag,   ""FPag""}          else null,"
    MLine m, "    if C_Clase<>null  then {C_Clase,  ""Clase""}         else null,"
    MLine m, "    if C_ALaOr<>null  then {C_ALaOr,  ""ALaOrden""}      else null,"
    MLine m, "    if C_Dep<>null    then {C_Dep,    ""Dep" & Chr(243) & "sito""}     else null,"
    MLine m, "    if C_Ret<>null    then {C_Ret,    ""Retiro""}        else null,"
    MLine m, "    if C_CtaLiq<>null then {C_CtaLiq, ""CtaLiq""}        else null,"
    MLine m, "    if C_Est<>null    then {C_Est,    ""Estado""}        else null,"
    MLine m, "    if C_Obs<>null    then {C_Obs,    ""Observaciones""} else null,"
    MLine m, "    if C_Mon<>null    then {C_Mon,    ""Moneda""}        else null"
    MLine m, "  }),"
    MLine m, "  Ren = if List.Count(RenPairs)>0 then Table.RenameColumns(Source, RenPairs, MissingField.Ignore) else Source,"

    MLine m, "  MonthMap = #table(type table [ES=text, EN=text], {"
    MLine m, "     {""ENE"",""JAN""},{""FEB"",""FEB""},{""MAR"",""MAR""},{""ABR"",""APR""},{""MAY"",""MAY""},{""JUN"",""JUN""},"
    MLine m, "     {""JUL"",""JUL""},{""AGO"",""AUG""},{""SET"",""SEP""},{""SEP"",""SEP""},{""OCT"",""OCT""},{""NOV"",""NOV""},{""DIC"",""DEC""}}),"
    MLine m, "  ToDateES = (x as any) as nullable date =>"
    MLine m, "    let"
    MLine m, "      d0 = try Date.From(x) otherwise null,"
    MLine m, "      t0 = if d0<>null then null else try Text.Upper(Text.Trim(Text.From(x))) otherwise null,"
    MLine m, "      t1 = if t0=null then null else if Text.Contains(t0, ""-"") then t0 else if Text.Length(t0)>=9 then Text.Start(t0,2) & ""-"" & Text.Range(t0,2,3) & ""-"" & Text.Range(t0,5,4) else t0,"
    MLine m, "      monES = if t1=null then null else Text.Range(t1,3,3),"
    MLine m, "      row = if monES=null then #table(type table [ES=text, EN=text], {}) else Table.SelectRows(MonthMap, each [ES]=monES),"
    MLine m, "      monEN = if Table.RowCount(row)=1 then row{0}[EN] else monES,"
    MLine m, "      tEN = if t1=null then null else Text.Start(t1,2) & ""-"" & monEN & ""-"" & Text.End(t1,4),"
    MLine m, "      d1 = if d0<>null then d0 else (try Date.FromText(tEN, ""en-US"") otherwise null)"
    MLine m, "    in d1,"

    MLine m, "  WithFecha = if List.Contains(Table.ColumnNames(Ren), ""Fecha"") then Table.TransformColumns(Ren, {{""Fecha"", each ToDateES(_), type date}}) else Ren,"
    MLine m, "  ClaseUp = if List.Contains(Table.ColumnNames(WithFecha), ""Clase"") then Table.TransformColumns(WithFecha, {{""Clase"", each if _=null then null else Text.Upper(Text.Trim(Text.From(_))), type text}}) else WithFecha,"
    MLine m, "  FilClase = if List.Contains(Table.ColumnNames(ClaseUp), ""Clase"") then Table.SelectRows(ClaseUp, each [Clase] = ""DPE"" or [Clase] = ""DFS"" or [Clase] = ""RAF"" or [Clase] = ""RFS"") else ClaseUp,"

    MLine m, "  CN1 = Table.ColumnNames(FilClase),"
    MLine m, "  FixDR = if List.Contains(CN1, ""Dep" & Chr(243) & "sito"") and List.Contains(CN1, ""Retiro"") then"
    MLine m, "    let"
    MLine m, "      A1 = Table.AddColumn(FilClase, ""__Dep"", each if [Dep" & Chr(243) & "sito] <> null and [Dep" & Chr(243) & "sito] <> 0 then [Dep" & Chr(243) & "sito] else null, type number),"
    MLine m, "      A2 = Table.AddColumn(A1, ""__Ret"", each if [Retiro] <> null and [Retiro] <> 0 and ([Dep" & Chr(243) & "sito] = null or [Dep" & Chr(243) & "sito] = 0) then [Retiro] else null, type number),"
    MLine m, "      Rm = Table.RemoveColumns(A2, {""Dep" & Chr(243) & "sito"",""Retiro""}),"
    MLine m, "      Rn = Table.RenameColumns(Rm, {{""__Dep"",""Dep" & Chr(243) & "sito""},{""__Ret"",""Retiro""}})"
    MLine m, "    in Rn"
    MLine m, "  else FilClase,"

    MLine m, "  Target = {""Fecha"",""Transac"",""Cuenta"",""Nombre"",""Ope"",""Tipo"",""FPag"",""Clase"",""ALaOrden"",""Dep" & Chr(243) & "sito"",""Retiro"",""CtaLiq"",""Estado"",""Observaciones"",""Moneda""},"
    MLine m, "  Present = List.Intersect({Target, Table.ColumnNames(FixDR)}),"
    MLine m, "  Sel = Table.SelectColumns(FixDR, Present, MissingField.Ignore),"
    MLine m, "  AddMissing = List.Accumulate(List.Difference(Target, Present), Sel, (st,c)=> Table.AddColumn(st, c, each null)),"

    MLine m, "  Dates = if List.Contains(Table.ColumnNames(AddMissing), ""Fecha"") then List.RemoveNulls(Table.Column(AddMissing, ""Fecha"")) else {},"
    MLine m, "  FinMes = if List.Count(Dates)>0 then Date.EndOfMonth(List.Max(Dates)) else Date.EndOfMonth(DateTime.Date(DateTime.LocalNow())),"
    MLine m, "  IniMes = Date.StartOfMonth(Date.AddMonths(FinMes, -" & CStr(mesesSel - 1) & ")),"
    MLine m, "  F = if List.Contains(Table.ColumnNames(AddMissing), ""Fecha"") then Table.SelectRows(AddMissing, each [Fecha] <> null and [Fecha] >= IniMes and [Fecha] <= FinMes) else AddMissing,"

    MLine m, "  SAB_MC_MAIN = Table.Sort(F, {{""Fecha"", Order.Ascending}})"
    MLine m, "in"
    MLine m, "  SAB_MC_MAIN"

    M_MC_MAIN = m
End Function

Private Function M_MC_ALERTAS(ByVal which As String) As String
    Dim m As String
    Dim op As String: op = UCase$(Trim$(which))
    If op <> "DEP" And op <> "RET" Then op = "DEP"

    Dim montoCol As String
    Dim montoColRef As String
    If op = "DEP" Then
        montoCol    = "Dep" & Chr(243) & "sito"
        montoColRef = "[#""Dep" & Chr(243) & "sito""]"
    Else
        montoCol    = "Retiro"
        montoColRef = "[Retiro]"
    End If

    MLine m, "let"
    MLine m, "    Origen0 = SAB_MC_MAIN,"
    MLine m, "    Origen = Table.SelectRows(Origen0, each " & montoColRef & " <> null and " & montoColRef & " <> 0),"
    MLine m, ""
    MLine m, "    Typed = Table.TransformColumnTypes(Origen, {{""" & montoCol & """, type number}}, ""es-PE""),"
    MLine m, ""
    MLine m, "    WithDate = Table.AddColumn("
    MLine m, "        Typed,"
    MLine m, "        ""__Fecha"","
    MLine m, "        each"
    MLine m, "            let v = [Fecha] in"
    MLine m, "                try Date.From(v) otherwise"
    MLine m, "                try Date.FromText(Text.From(v), ""es-PE"") otherwise"
    MLine m, "                try Date.FromText(Text.From(v), ""en-US"") otherwise"
    MLine m, "                null,"
    MLine m, "        type date"
    MLine m, "    ),"
    MLine m, ""
    MLine m, "    F = Table.SelectRows("
    MLine m, "        WithDate,"
    MLine m, "        each [__Fecha] <> null"
    MLine m, "            and [Cuenta] <> null"
    MLine m, "            and Text.Trim(Text.From([Cuenta])) <> """""
    MLine m, "    ),"
    MLine m, ""
    MLine m, "    Daily = Table.Group("
    MLine m, "        F,"
    MLine m, "        {""Cuenta"", ""__Fecha""},"
    MLine m, "        {{""MontoDia"", each List.Sum(List.RemoveNulls([" & IIf(op = "DEP", "#""Dep" & Chr(243) & "sito""", "Retiro") & "])), type number}}"
    MLine m, "    ),"
    MLine m, ""
    MLine m, "    Agg = Table.Group("
    MLine m, "        Daily,"
    MLine m, "        {""Cuenta""},"
    MLine m, "        {"
    MLine m, "            {""SUMA_MONTOS"",      each List.Sum(List.RemoveNulls([MontoDia])), type number},"
    MLine m, "            {""NUM_OPERACIONES"",  each Table.RowCount(_), Int64.Type},"
    MLine m, "            {""PROMEDIO_MONTOS"",  each try Number.Round(List.Average(List.RemoveNulls([MontoDia])), 2) otherwise null, type number},"
    MLine m, "            {""ULTIMA_OPERACION"", each let t = Table.Sort(_, {{""__Fecha"", Order.Ascending}}) in try List.Last(t[MontoDia]) otherwise null, type number}"
    MLine m, "        }"
    MLine m, "    ),"
    MLine m, ""
    MLine m, "    Meta = Table.Group("
    MLine m, "        F,"
    MLine m, "        {""Cuenta""},"
    MLine m, "        {"
    MLine m, "            {""CLASE"",  each try Text.From(List.First(List.RemoveNulls([Clase])))  otherwise null, type text},"
    MLine m, "            {""MONEDA"", each try Text.From(List.First(List.RemoveNulls([Moneda]))) otherwise null, type text}"
    MLine m, "        }"
    MLine m, "    ),"
    MLine m, ""
    MLine m, "    JoinAgg  = Table.NestedJoin(Agg, {""Cuenta""}, Meta, {""Cuenta""}, ""meta"", JoinKind.LeftOuter),"
    MLine m, "    Expanded = Table.ExpandTableColumn(JoinAgg, ""meta"", {""CLASE"", ""MONEDA""}, {""CLASE"", ""MONEDA""}),"
    MLine m, ""
    MLine m, "    WithDesv = Table.AddColumn("
    MLine m, "        Expanded,"
    MLine m, "        ""DESVIACION_MEDIA_%"","
    MLine m, "        each"
    MLine m, "            let p = [PROMEDIO_MONTOS], u = [ULTIMA_OPERACION]"
    MLine m, "            in if p = null or p = 0 or u = null then null else ((u - p) / p) * 100.0,"
    MLine m, "        type number"
    MLine m, "    ),"
    MLine m, ""
    MLine m, "    WithNivel = Table.AddColumn("
    MLine m, "        WithDesv,"
    MLine m, "        ""NIVEL_RIESGO"","
    MLine m, "        each"
    MLine m, "            let d = [#""DESVIACION_MEDIA_%""]"
    MLine m, "            in if d = null then null else if d < 50 then 1 else if d <= 100 then 2 else 3,"
    MLine m, "        Int64.Type"
    MLine m, "    ),"
    MLine m, ""
    MLine m, "    Selected = Table.SelectColumns("
    MLine m, "        WithNivel,"
    MLine m, "        {""Cuenta"", ""CLASE"", ""MONEDA"", ""SUMA_MONTOS"", ""NUM_OPERACIONES"", ""PROMEDIO_MONTOS"", ""ULTIMA_OPERACION"", ""DESVIACION_MEDIA_%"", ""NIVEL_RIESGO""},"
    MLine m, "        MissingField.Ignore"
    MLine m, "    ),"
    MLine m, ""
    MLine m, "    Sorted = Table.Sort(Selected, {{""DESVIACION_MEDIA_%"", Order.Descending}})"
    MLine m, "in"
    MLine m, "    Sorted"

    M_MC_ALERTAS = m
End Function

'======================
' Punto de entrada publico
'======================
Public Sub CrearQuerySAB_MC(ByVal rutaArchivo As String, _
                             ByVal mesesSel As Long, _
                             Optional ByVal opMode As String = "AMBOS", _
                             Optional ByVal showProgress As Boolean = False)
    On Error GoTo EH

    mT0Total  = Timer
    mStageLog = vbNullString

    If mesesSel <= 0 Then mesesSel = 6
    If Len(Trim$(opMode)) = 0 Then opMode = "AMBOS"

    Dim makeDep As Boolean: makeDep = (UCase$(opMode) = "AMBOS" Or UCase$(opMode) = "SOLO_DEPOSITO")
    Dim makeRet As Boolean: makeRet = (UCase$(opMode) = "AMBOS" Or UCase$(opMode) = "SOLO_RETIRO")

    SafeApp True

    ' Upsert de las queries necesarias
    UpsertWorkbookQuery "SAB_MC_RAW",  M_MC_RAW(rutaArchivo)
    UpsertWorkbookQuery "SAB_MC_MAIN", M_MC_MAIN(mesesSel)
    If makeDep Then UpsertWorkbookQuery "SAB_MC_ALERTAS_DEP", M_MC_ALERTAS("DEP")
    If makeRet Then UpsertWorkbookQuery "SAB_MC_ALERTAS_RET", M_MC_ALERTAS("RET")

    ' Hojas de trabajo
    Dim shRaw   As Worksheet: Set shRaw   = EnsureSheet("SAB_MC_RAW_WORK")
    Dim shMain  As Worksheet: Set shMain  = EnsureSheet("SAB_MC_MAIN_WORK")
    Dim shAlDep As Worksheet
    Dim shAlRet As Worksheet
    ClearSheetButKeepName shRaw
    ClearSheetButKeepName shMain

    ' Conexiones
    Dim connRaw  As WorkbookConnection: Set connRaw  = EnsurePQConnection("SAB_MC_RAW")
    Dim connMain As WorkbookConnection: Set connMain = EnsurePQConnection("SAB_MC_MAIN")
    Dim connAlDep As WorkbookConnection
    Dim connAlRet As WorkbookConnection

    If makeDep Then
        Set shAlDep = EnsureSheet("SAB_MC_AL_DEP_WORK")
        ClearSheetButKeepName shAlDep
        Set connAlDep = EnsurePQConnection("SAB_MC_ALERTAS_DEP")
    End If
    If makeRet Then
        Set shAlRet = EnsureSheet("SAB_MC_AL_RET_WORK")
        ClearSheetButKeepName shAlRet
        Set connAlRet = EnsurePQConnection("SAB_MC_ALERTAS_RET")
    End If

    ' Paso 1: refrescar TODAS las conexiones en orden (como en el original)
    Dim tStage As Double
    tStage = Timer
    Application.StatusBar = "Cargando RAW..."
    RefreshConnectionSync connRaw
    AppendStageLog "RAW", ElapsedSec(tStage)

    tStage = Timer
    Application.StatusBar = "Cargando MAIN..."
    RefreshConnectionSync connMain
    AppendStageLog "MAIN", ElapsedSec(tStage)

    If makeDep Then
        tStage = Timer
        Application.StatusBar = "Cargando ALERTAS DEP..."
        RefreshConnectionSync connAlDep
        AppendStageLog "AL_DEP", ElapsedSec(tStage)
    End If
    If makeRet Then
        tStage = Timer
        Application.StatusBar = "Cargando ALERTAS RET..."
        RefreshConnectionSync connAlRet
        AppendStageLog "AL_RET", ElapsedSec(tStage)
    End If

    ' Paso 2: crear TODAS las tablas vinculadas (datos ya cargados en PQ)
    Application.StatusBar = "Creando tablas..."
    Dim loRaw   As ListObject: Set loRaw   = EnsureTableForConnection(shRaw,  "SAB_MC_RAW",  connRaw)
    Dim loMain  As ListObject: Set loMain  = EnsureTableForConnection(shMain, "SAB_MC_MAIN", connMain)
    Dim loAlDep As ListObject
    Dim loAlRet As ListObject
    If makeDep Then Set loAlDep = EnsureTableForConnection(shAlDep, "SAB_MC_ALERTAS_DEP", connAlDep)
    If makeRet Then Set loAlRet = EnsureTableForConnection(shAlRet, "SAB_MC_ALERTAS_RET", connAlRet)

    ' Sufijo de periodo desde loMain
    Dim minD As Date, maxD As Date, gotDates As Boolean
    gotDates = GetMinMaxDateFromLO(loMain, "Fecha", minD, maxD)
    If Not gotDates Then gotDates = GetMinMaxDateFromLO(loRaw, "Fecha", minD, maxD)

    Dim ini As Date, fin As Date, suf As String
    If gotDates Then
        ini = FirstDayOfMonth(minD)
        fin = LastDayOfMonth(maxD)
    Else
        fin = DateSerial(Year(Date), Month(Date), 0)
        ini = DateSerial(Year(fin), Month(fin) - (mesesSel - 1), 1)
    End If
    suf = MesAbrevES(ini) & "_" & MesAbrevES(fin) & "_" & Year(fin)

    ' Renombrar RAW y MAIN
    Dim nmRaw  As String: nmRaw  = SanitizeSheetName("SAB_MC_RAW_" & suf)
    Dim nmMain As String: nmMain = SanitizeSheetName("SAB_MC_"     & suf)

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

    ' Renombrar alertas segun opMode
    If makeDep Then
        Dim nmAlDep As String: nmAlDep = SanitizeSheetName("SAB_MC_AL_DEP_" & suf)
        DeleteSheetIfExists ThisWorkbook, nmAlDep
        FreeSheetName ThisWorkbook, nmAlDep, shAlDep
        DeleteAllTablesByName ThisWorkbook, nmAlDep
        SetTableNameSafe ThisWorkbook, loAlDep, nmAlDep
        RenameSheetExact shAlDep, nmAlDep
    End If

    If makeRet Then
        Dim nmAlRet As String: nmAlRet = SanitizeSheetName("SAB_MC_AL_RET_" & suf)
        DeleteSheetIfExists ThisWorkbook, nmAlRet
        FreeSheetName ThisWorkbook, nmAlRet, shAlRet
        DeleteAllTablesByName ThisWorkbook, nmAlRet
        SetTableNameSafe ThisWorkbook, loAlRet, nmAlRet
        RenameSheetExact shAlRet, nmAlRet
    End If

    ' Graficos de alertas
    If BUILD_GRAFICOS Then
        If makeDep And Not loAlDep Is Nothing Then
            modSABGraficos.BuildGraficosAlertasEnHoja loAlDep, loMain, "DEP", suf
        End If
        If makeRet And Not loAlRet Is Nothing Then
            modSABGraficos.BuildGraficosAlertasEnHoja loAlRet, loMain, "RET", suf
        End If
    End If

    ' Limpiar queries si corresponde
    If Not KEEP_PQ_QUERIES Then
        DeleteQueryAndConnection "SAB_MC_RAW"
        DeleteQueryAndConnection "SAB_MC_MAIN"
        If makeDep Then DeleteQueryAndConnection "SAB_MC_ALERTAS_DEP"
        If makeRet Then DeleteQueryAndConnection "SAB_MC_ALERTAS_RET"
    End If

    SafeApp False

    Dim totalMsg As String
    totalMsg = "SAB - Movimiento de Caja cargado." & vbCrLf & vbCrLf & _
               mStageLog & vbCrLf & vbCrLf & _
               "Total: " & FormatElapsed(ElapsedSec(mT0Total))

    Application.StatusBar = "SAB MC listo. Total " & FormatElapsed(ElapsedSec(mT0Total))
    Debug.Print totalMsg

    If showProgress Then MsgBox totalMsg, vbInformation, "SAB MC"

    shMain.Activate
    shMain.Range("A1").Select
    Exit Sub

EH:
    SafeApp False
    MsgBox "Error en CrearQuerySAB_MC: " & Err.Number & " - " & Err.Description, vbCritical
End Sub
