Option Explicit

'==========================
' modPQ_SAB_MC
' Carga SAB Movimiento de Caja.
' RAW y MAIN via PQ. Alertas DEP y RET calculadas en VBA desde loMain
' para evitar re-evaluacion de la cadena RAW->MAIN por cada query de alerta.
' Punto de entrada: CrearQuerySAB_MC(rutaArchivo, mesesSel, opMode, showProgress)
'==========================

Private Const BUILD_GRAFICOS As Boolean = True
Private Const TABLE_STYLE    As String  = "TableStyleLight9"

Private mAppFrozen          As Boolean
Private mPrevScreenUpdating As Boolean
Private mPrevEnableEvents   As Boolean
Private mPrevDisplayAlerts  As Boolean
Private mPrevCalculation    As XlCalculation
Private mPrevStatusBar      As Variant
Private mT0Total            As Double
Private mStageLog           As String

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

Private Sub AppendStageLog(ByVal label As String, ByVal sec As Double)
    Dim line As String
    line = label & ": " & FormatElapsed(sec) & " (" & Format(sec, "0.0") & " s)"
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
    nm = Replace(desired, "[", "("):  nm = Replace(nm, "]", ")")
    nm = Replace(nm, ":", " - "):     nm = Replace(nm, "\", " - ")
    nm = Replace(nm, "/", " - "):     nm = Replace(nm, "?", " - ")
    nm = Replace(nm, "*", " - "):     nm = Trim$(nm)
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
    Dim base As String: base = Left$(safeName, 20)
    If Len(base) = 0 Then base = "OLD"
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
    Dim conn As WorkbookConnection
    Dim connName As String: connName = "PQ_" & queryName
    On Error Resume Next: Set conn = ThisWorkbook.Connections(connName): On Error GoTo 0
    If conn Is Nothing Then
        Dim cs  As String
        Dim cmd As String
        cs  = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & _
              queryName & ";Extended Properties=" & Chr$(34) & Chr$(34)
        cmd = "SELECT * FROM [" & queryName & "]"
        On Error Resume Next
        Set conn = ThisWorkbook.Connections.Add2(connName, "", cs, cmd, xlCmdSql)
        If conn Is Nothing Then Set conn = ThisWorkbook.Connections.Add(connName, "", cs, cmd, xlCmdSql)
        On Error GoTo 0
    End If
    Set EnsurePQConnection = conn
End Function

Private Function EnsureTableForConnection(ByVal sh As Worksheet, _
                                           ByVal loName As String, _
                                           ByVal conn As WorkbookConnection) As ListObject
    Dim lo As ListObject
    On Error Resume Next: Set lo = sh.ListObjects(loName): On Error GoTo 0
    If Not lo Is Nothing Then On Error Resume Next: lo.Delete: On Error GoTo 0: Set lo = Nothing
    Set lo = sh.ListObjects.Add(SourceType:=xlSrcExternal, Source:=conn, _
                                LinkSource:=True, XlListObjectHasHeaders:=xlYes, _
                                Destination:=sh.Range("A1"))
    On Error Resume Next: lo.Name = loName: On Error GoTo 0
    On Error Resume Next
    If Not lo.QueryTable Is Nothing Then
        With lo.QueryTable
            .BackgroundQuery    = False
            .RefreshStyle       = xlOverwriteCells
            .AdjustColumnWidth  = True
            .PreserveColumnInfo = True
            .Refresh BackgroundQuery:=False
        End With
    End If
    Application.CalculateUntilAsyncQueriesDone
    On Error GoTo 0
    On Error Resume Next: lo.TableStyle = TABLE_STYLE: On Error GoTo 0
    Set EnsureTableForConnection = lo
End Function

'======================
' M queries
'======================
Private Function M_MC_RAW(ByVal rutaArchivo As String) As String
    Dim m As String
    Dim p As String: p = Replace(rutaArchivo, """", """""""""")

    MLine m, "let"
    MLine m, "  Ruta = """ & p & ""","
    MLine m, "  Libro = Excel.Workbook(File.Contents(Ruta), null, true),"
    MLine m, "  Base0 = Libro{0}[Data],"
    MLine m, "  Skip = Table.Skip(Base0, 10),"
    MLine m, "  Promoted = Table.PromoteHeaders(Skip, [PromoteAllScalars=true]),"
    MLine m, "  TrimCols = Table.TransformColumnNames(Promoted, each Text.Trim(_)),"
    MLine m, "  NonEmptyCols = Table.SelectColumns(TrimCols,"
    MLine m, "    List.Select(Table.ColumnNames(TrimCols), (c) => List.NonNullCount(Table.Column(TrimCols, c)) > 0),"
    MLine m, "    MissingField.Ignore),"
    MLine m, "  CN = Table.ColumnNames(NonEmptyCols),"
    MLine m, "  ColFecha = if List.Contains(CN, ""Fecha"") then ""Fecha"""
    MLine m, "             else if List.Contains(CN, ""FECHA"") then ""FECHA"""
    MLine m, "             else if List.Contains(CN, ""Fec"")   then ""Fec"" else CN{0},"
    MLine m, "  WithFechaTxt = Table.AddColumn(NonEmptyCols, ""__FechaTxt"","
    MLine m, "    each Text.Upper(Text.Trim(Text.From(Record.Field(_, ColFecha)))), type text),"
    MLine m, "  MonTmp = Table.AddColumn(WithFechaTxt, ""Moneda"","
    MLine m, "    each let s=[__FechaTxt],"
    MLine m, "             pL=Text.PositionOf(s, ""("", Occurrence.Last),"
    MLine m, "             pR=Text.PositionOf(s, "")"", Occurrence.Last)"
    MLine m, "         in if pL>=0 and pR>pL then Text.Middle(s, pL+1, pR-pL-1) else null, type text),"
    MLine m, "  MonFill = Table.FillDown("
    MLine m, "    Table.TransformColumns(MonTmp,"
    MLine m, "      {{""Moneda"", each if _=null then null else Text.Upper(Text.Trim(_)), type text}}),"
    MLine m, "    {""Moneda""}),"
    MLine m, "  Filtrado = Table.SelectRows(MonFill,"
    MLine m, "    each let s=[__FechaTxt]"
    MLine m, "         in not Text.StartsWith(s, ""TOTAL"")"
    MLine m, "            and not (Text.Contains(s, ""("") and Text.Contains(s, "")""))),"
    MLine m, "  ToNum = (v as any) as nullable number =>"
    MLine m, "    let t0 = try Text.From(v) otherwise null,"
    MLine m, "        t1 = if t0=null then null else Text.Trim(t0),"
    MLine m, "        t2 = if t1=null then null"
    MLine m, "             else Text.Replace(Text.Replace(Text.Replace(t1,""S/"",""""),""$"",""""),"" "",""""),"
    MLine m, "        hD = if t2=null then false else Text.Contains(t2, "".""),"
    MLine m, "        hC = if t2=null then false else Text.Contains(t2, "",""),"
    MLine m, "        t3 = if t2=null then null"
    MLine m, "             else if hD and hC then Text.Replace(Text.Replace(t2,""."",""""),"","",""."") "
    MLine m, "             else if hC and not hD then Text.Replace(t2,"","",""."")"
    MLine m, "             else t2,"
    MLine m, "        n  = try Number.FromText(t3, ""en-US"") otherwise try Number.From(t3) otherwise null"
    MLine m, "    in n,"
    MLine m, "  MaybeNumCols = {""Dep" & Chr(243) & "sito"",""Deposito"",""Retiro"",""Saldo"",""Monto"",""Abono"",""Cargo""},"
    MLine m, "  Numd = List.Accumulate(MaybeNumCols, Filtrado,"
    MLine m, "    (st,c) => if List.Contains(Table.ColumnNames(st), c)"
    MLine m, "              then Table.TransformColumns(st, {{c, each ToNum(_), type number}})"
    MLine m, "              else st),"
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
    MLine m, "  CN0    = Table.ColumnNames(Source),"
    MLine m, "  Pick   = (alts as list) as nullable text =>"
    MLine m, "    List.First(List.Select(alts, each List.Contains(CN0, _)), null),"

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
    MLine m, "    if C_Fecha <>null then {C_Fecha,  ""Fecha""}         else null,"
    MLine m, "    if C_Trans <>null then {C_Trans,  ""Transac""}       else null,"
    MLine m, "    if C_Cuenta<>null then {C_Cuenta, ""Cuenta""}        else null,"
    MLine m, "    if C_Nombre<>null then {C_Nombre, ""Nombre""}        else null,"
    MLine m, "    if C_Ope   <>null then {C_Ope,    ""Ope""}           else null,"
    MLine m, "    if C_Tipo  <>null then {C_Tipo,   ""Tipo""}          else null,"
    MLine m, "    if C_FPag  <>null then {C_FPag,   ""FPag""}          else null,"
    MLine m, "    if C_Clase <>null then {C_Clase,  ""Clase""}         else null,"
    MLine m, "    if C_ALaOr <>null then {C_ALaOr,  ""ALaOrden""}      else null,"
    MLine m, "    if C_Dep   <>null then {C_Dep,    ""Dep" & Chr(243) & "sito""}     else null,"
    MLine m, "    if C_Ret   <>null then {C_Ret,    ""Retiro""}        else null,"
    MLine m, "    if C_CtaLiq<>null then {C_CtaLiq, ""CtaLiq""}        else null,"
    MLine m, "    if C_Est   <>null then {C_Est,    ""Estado""}        else null,"
    MLine m, "    if C_Obs   <>null then {C_Obs,    ""Observaciones""} else null,"
    MLine m, "    if C_Mon   <>null then {C_Mon,    ""Moneda""}        else null"
    MLine m, "  }),"
    MLine m, "  Ren = if List.Count(RenPairs)>0"
    MLine m, "        then Table.RenameColumns(Source, RenPairs, MissingField.Ignore)"
    MLine m, "        else Source,"

    ' Parseo de fechas DDMMMYYYY: un solo TransformColumns con if/else inline.
    ' Sin pasos lazy anidados. PQ evalua un unico step con 12 ramas if/else de costo
    ' fijo por fila, sin overhead de wrappers de tabla adicionales.
    MLine m, "  WithFecha ="
    MLine m, "    if not List.Contains(Table.ColumnNames(Ren), ""Fecha"") then Ren"
    MLine m, "    else Table.TransformColumns(Ren, {{""Fecha"","
    MLine m, "      each let t  = if _ = null then null"
    MLine m, "                    else Text.Upper(Text.Trim(Text.From(_))),"
    MLine m, "               ok = t <> null and Text.Length(t) >= 9,"
    MLine m, "               dd = if ok then Text.Start(t, 2)     else null,"
    MLine m, "               ms = if ok then Text.Middle(t, 2, 3) else null,"
    MLine m, "               yy = if ok then Text.End(t, 4)       else null,"
    MLine m, "               mn = if ms = null  then null"
    MLine m, "                    else if ms = ""ENE"" then ""01"" else if ms = ""FEB"" then ""02"""
    MLine m, "                    else if ms = ""MAR"" then ""03"" else if ms = ""ABR"" then ""04"""
    MLine m, "                    else if ms = ""MAY"" then ""05"" else if ms = ""JUN"" then ""06"""
    MLine m, "                    else if ms = ""JUL"" then ""07"" else if ms = ""AGO"" then ""08"""
    MLine m, "                    else if ms = ""SET"" then ""09"" else if ms = ""SEP"" then ""09"""
    MLine m, "                    else if ms = ""OCT"" then ""10"" else if ms = ""NOV"" then ""11"""
    MLine m, "                    else if ms = ""DIC"" then ""12"" else ms,"
    MLine m, "               iso = if dd=null or mn=null or yy=null then null"
    MLine m, "                     else yy & ""-"" & mn & ""-"" & dd"
    MLine m, "           in try Date.FromText(iso, ""en-US"") otherwise null,"
    MLine m, "      type date}}),"

    MLine m, "  ClaseUp  = if List.Contains(Table.ColumnNames(WithFecha), ""Clase"")"
    MLine m, "             then Table.TransformColumns(WithFecha, {{""Clase"","
    MLine m, "               each if _=null then null else Text.Upper(Text.Trim(Text.From(_))), type text}})"
    MLine m, "             else WithFecha,"
    MLine m, "  FilClase = if List.Contains(Table.ColumnNames(ClaseUp), ""Clase"")"
    MLine m, "             then Table.SelectRows(ClaseUp,"
    MLine m, "               each [Clase]=""DPE"" or [Clase]=""DFS"" or [Clase]=""RAF"" or [Clase]=""RFS"")"
    MLine m, "             else ClaseUp,"

    MLine m, "  CN1   = Table.ColumnNames(FilClase),"
    MLine m, "  FixDR = if List.Contains(CN1, ""Dep" & Chr(243) & "sito"") and List.Contains(CN1, ""Retiro"") then"
    MLine m, "    let"
    MLine m, "      A1 = Table.AddColumn(FilClase, ""__Dep"","
    MLine m, "             each if [Dep" & Chr(243) & "sito]<>null and [Dep" & Chr(243) & "sito]<>0"
    MLine m, "                  then [Dep" & Chr(243) & "sito] else null, type number),"
    MLine m, "      A2 = Table.AddColumn(A1, ""__Ret"","
    MLine m, "             each if [Retiro]<>null and [Retiro]<>0"
    MLine m, "                      and ([Dep" & Chr(243) & "sito]=null or [Dep" & Chr(243) & "sito]=0)"
    MLine m, "                  then [Retiro] else null, type number),"
    MLine m, "      Rm = Table.RemoveColumns(A2, {""Dep" & Chr(243) & "sito"",""Retiro""}),"
    MLine m, "      Rn = Table.RenameColumns(Rm, {{""__Dep"",""Dep" & Chr(243) & "sito""},{""__Ret"",""Retiro""}})"
    MLine m, "    in Rn"
    MLine m, "  else FilClase,"

    MLine m, "  Target = {""Fecha"",""Transac"",""Cuenta"",""Nombre"",""Ope"",""Tipo"",""FPag"","
    MLine m, "            ""Clase"",""ALaOrden"",""Dep" & Chr(243) & "sito"",""Retiro"","
    MLine m, "            ""CtaLiq"",""Estado"",""Observaciones"",""Moneda""},"
    MLine m, "  Present    = List.Intersect({Target, Table.ColumnNames(FixDR)}),"
    MLine m, "  Sel        = Table.SelectColumns(FixDR, Present, MissingField.Ignore),"
    MLine m, "  AddMissing = List.Accumulate(List.Difference(Target, Present), Sel,"
    MLine m, "                 (st,c) => Table.AddColumn(st, c, each null)),"

    MLine m, "  Dates  = if List.Contains(Table.ColumnNames(AddMissing), ""Fecha"")"
    MLine m, "           then List.RemoveNulls(Table.Column(AddMissing, ""Fecha"")) else {},"
    MLine m, "  FinMes = if List.Count(Dates)>0 then Date.EndOfMonth(List.Max(Dates))"
    MLine m, "           else Date.EndOfMonth(DateTime.Date(DateTime.LocalNow())),"
    MLine m, "  IniMes = Date.StartOfMonth(Date.AddMonths(FinMes, -" & CStr(mesesSel - 1) & ")),"
    MLine m, "  F      = if List.Contains(Table.ColumnNames(AddMissing), ""Fecha"")"
    MLine m, "           then Table.SelectRows(AddMissing,"
    MLine m, "             each [Fecha]<>null and [Fecha]>=IniMes and [Fecha]<=FinMes)"
    MLine m, "           else AddMissing,"
    MLine m, "  SAB_MC_MAIN = Table.Sort(F, {{""Fecha"", Order.Ascending}})"
    MLine m, "in"
    MLine m, "  SAB_MC_MAIN"

    M_MC_MAIN = m
End Function

'======================
' Alertas en VBA desde loMain
'
' Calcula la tabla de alertas directamente desde el array en memoria de loMain.
' Equivale a la logica M original (group by Cuenta/dia, agregacion, desviacion)
' pero sin re-evaluar la cadena RAW->MAIN en PQ.
'
' which: "DEP" -> columna Deposito  |  "RET" -> columna Retiro
'======================
Private Function BuildAlertasVBA(ByVal loMain As ListObject, _
                                   ByVal which As String, _
                                   ByVal shAl As Worksheet, _
                                   ByVal loAlName As String) As ListObject
    Set BuildAlertasVBA = Nothing
    If loMain Is Nothing Then Exit Function
    If loMain.DataBodyRange Is Nothing Then Exit Function

    Dim op As String: op = UCase$(Trim$(which))
    If op <> "DEP" And op <> "RET" Then op = "DEP"

    Dim depName As String: depName = "Dep" & Chr(243) & "sito"

    ' Localizar indices de columna
    Dim colFecha  As Long: colFecha  = 0
    Dim colCuenta As Long: colCuenta = 0
    Dim colMonto  As Long: colMonto  = 0
    Dim colClase  As Long: colClase  = 0
    Dim colMoneda As Long: colMoneda = 0
    Dim i As Long

    For i = 1 To loMain.ListColumns.Count
        Select Case loMain.ListColumns(i).Name
            Case "Fecha":                      colFecha  = i
            Case "Cuenta":                     colCuenta = i
            Case depName, "Deposito", "Abono": If op = "DEP" Then colMonto = i
            Case "Retiro", "Cargo":            If op = "RET" Then colMonto = i
            Case "Clase":                      colClase  = i
            Case "Moneda":                     colMoneda = i
        End Select
    Next i

    If colCuenta = 0 Or colMonto = 0 Then Exit Function

    Dim nRows As Long: nRows = loMain.DataBodyRange.Rows.Count
    Dim data  As Variant: data = loMain.DataBodyRange.Value2

    ' Diccionarios (late binding, sin referencia a Scripting Runtime)
    Dim dDay  As Object: Set dDay  = CreateObject("Scripting.Dictionary") ' "Cuenta|DateLng" -> suma monto dia
    Dim dMeta As Object: Set dMeta = CreateObject("Scripting.Dictionary") ' "Cuenta" -> "Clase|Moneda"
    Dim dSum  As Object: Set dSum  = CreateObject("Scripting.Dictionary") ' "Cuenta" -> suma total
    Dim dNOp  As Object: Set dNOp  = CreateObject("Scripting.Dictionary") ' "Cuenta" -> num dias distintos
    Dim dMaxD As Object: Set dMaxD = CreateObject("Scripting.Dictionary") ' "Cuenta" -> max date (Long)

    Dim vM As Variant, dM As Double
    Dim vC As Variant, sC As String
    Dim vF As Variant, dF As Date, lF As Long
    Dim dayKey As String

    For i = 1 To nRows
        ' Monto
        vM = data(i, colMonto)
        If IsEmpty(vM) Or IsNull(vM) Or IsError(vM) Then GoTo SkipRow
        On Error Resume Next: dM = CDbl(vM): On Error GoTo 0
        If dM = 0 Then GoTo SkipRow

        ' Cuenta
        vC = data(i, colCuenta)
        If IsEmpty(vC) Or IsNull(vC) Or IsError(vC) Then GoTo SkipRow
        sC = Trim$(CStr(vC))
        If Len(sC) = 0 Then GoTo SkipRow

        ' Fecha
        If colFecha = 0 Then GoTo SkipRow
        vF = data(i, colFecha)
        If Not TryCoerceExcelDate(vF, dF) Then GoTo SkipRow
        lF = CLng(CDbl(CDate(dF)))

        ' Acumular monto diario
        dayKey = sC & "|" & CStr(lF)
        If dDay.Exists(dayKey) Then
            dDay(dayKey) = CDbl(dDay(dayKey)) + dM
        Else
            dDay.Add dayKey, dM
        End If

        ' Meta: primera Clase y Moneda no nulas por Cuenta
        If Not dMeta.Exists(sC) Then
            Dim sCl As String: sCl = ""
            Dim sMn As String: sMn = ""
            If colClase > 0 Then
                If Not IsEmpty(data(i, colClase)) And Not IsNull(data(i, colClase)) Then
                    sCl = Trim$(CStr(data(i, colClase)))
                End If
            End If
            If colMoneda > 0 Then
                If Not IsEmpty(data(i, colMoneda)) And Not IsNull(data(i, colMoneda)) Then
                    sMn = Trim$(CStr(data(i, colMoneda)))
                End If
            End If
            dMeta.Add sC, sCl & "|" & sMn
        End If

SkipRow:
    Next i

    ' Segunda pasada sobre dDay: construir dSum, dNOp, dMaxD por Cuenta
    Dim kk As Variant, pts() As String, sCta As String, lDate As Long, monDia As Double
    For Each kk In dDay.Keys
        pts    = Split(CStr(kk), "|")
        sCta   = pts(0)
        lDate  = CLng(pts(1))
        monDia = CDbl(dDay(kk))

        If Not dSum.Exists(sCta) Then
            dSum.Add sCta, 0#: dNOp.Add sCta, 0&: dMaxD.Add sCta, 0&
        End If
        dSum(sCta) = CDbl(dSum(sCta)) + monDia
        dNOp(sCta) = CLng(dNOp(sCta)) + 1
        If lDate > CLng(dMaxD(sCta)) Then dMaxD(sCta) = lDate
    Next kk

    Dim nCtas As Long: nCtas = dSum.Count
    If nCtas = 0 Then Exit Function

    ' Construir array de salida: 9 columnas
    ReDim outArr(1 To nCtas, 1 To 9) As Variant

    Dim r As Long: r = 0
    Dim sCta2 As Variant, suma As Double, nOp As Long
    Dim prom As Double, ultima As Double
    Dim desv As Variant, nivel As Variant
    Dim metaStr As String, mParts() As String

    For Each sCta2 In dSum.Keys
        r      = r + 1
        suma   = CDbl(dSum(sCta2))
        nOp    = CLng(dNOp(sCta2))
        prom   = IIf(nOp > 0, suma / nOp, 0)
        ultima = CDbl(dDay(CStr(sCta2) & "|" & CStr(CLng(dMaxD(sCta2)))))

        If prom <> 0 Then
            desv = ((ultima - prom) / prom) * 100#
        Else
            desv = Null
        End If

        If IsNull(desv) Then
            nivel = Null
        ElseIf CDbl(desv) < 50 Then
            nivel = 1
        ElseIf CDbl(desv) <= 100 Then
            nivel = 2
        Else
            nivel = 3
        End If

        metaStr = ""
        If dMeta.Exists(CStr(sCta2)) Then metaStr = CStr(dMeta(CStr(sCta2)))
        mParts = Split(metaStr, "|")

        outArr(r, 1) = CStr(sCta2)
        outArr(r, 2) = mParts(0)
        outArr(r, 3) = IIf(UBound(mParts) >= 1, mParts(1), "")
        outArr(r, 4) = Round(suma,   2)
        outArr(r, 5) = nOp
        outArr(r, 6) = Round(prom,   2)
        outArr(r, 7) = Round(ultima, 2)
        outArr(r, 8) = desv
        outArr(r, 9) = nivel
    Next sCta2

    ' Escribir encabezados y datos
    ClearSheetButKeepName shAl

    Dim hdrs As Variant
    hdrs = Array("Cuenta", "CLASE", "MONEDA", "SUMA_MONTOS", "NUM_OPERACIONES", _
                 "PROMEDIO_MONTOS", "ULTIMA_OPERACION", "DESVIACION_MEDIA_%", "NIVEL_RIESGO")
    Dim j As Long
    For j = 0 To 8: shAl.Cells(1, j + 1).Value = hdrs(j): Next j
    shAl.Range(shAl.Cells(2, 1), shAl.Cells(nCtas + 1, 9)).Value = outArr

    ' Crear ListObject estatico
    Dim loAl As ListObject
    Set loAl = shAl.ListObjects.Add(xlSrcRange, _
                   shAl.Range(shAl.Cells(1, 1), shAl.Cells(nCtas + 1, 9)), , xlYes)
    On Error Resume Next: loAl.Name = loAlName: On Error GoTo 0
    On Error Resume Next: loAl.TableStyle = TABLE_STYLE: On Error GoTo 0

    ' Ordenar por DESVIACION_MEDIA_% descendente
    On Error Resume Next
    loAl.Sort.SortFields.Clear
    loAl.Sort.SortFields.Add Key:=loAl.ListColumns("DESVIACION_MEDIA_%").DataBodyRange, _
                              SortOn:=xlSortOnValues, Order:=xlDescending, _
                              DataOption:=xlSortNormal
    With loAl.Sort
        .Header      = xlYes
        .MatchCase   = False
        .Orientation = xlTopToBottom
        .Apply
    End With
    On Error GoTo 0

    Set BuildAlertasVBA = loAl
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

    ' Solo RAW y MAIN en PQ; alertas se calculan en VBA desde loMain
    UpsertWorkbookQuery "SAB_MC_RAW",  M_MC_RAW(rutaArchivo)
    UpsertWorkbookQuery "SAB_MC_MAIN", M_MC_MAIN(mesesSel)

    Dim shRaw  As Worksheet: Set shRaw  = EnsureSheet("SAB_MC_RAW_WORK")
    Dim shMain As Worksheet: Set shMain = EnsureSheet("SAB_MC_MAIN_WORK")
    ClearSheetButKeepName shRaw
    ClearSheetButKeepName shMain

    Dim connRaw  As WorkbookConnection: Set connRaw  = EnsurePQConnection("SAB_MC_RAW")
    Dim connMain As WorkbookConnection: Set connMain = EnsurePQConnection("SAB_MC_MAIN")

    Dim tStage As Double

    tStage = Timer
    Application.StatusBar = "Cargando RAW..."
    Dim loRaw  As ListObject: Set loRaw  = EnsureTableForConnection(shRaw,  "SAB_MC_RAW",  connRaw)
    AppendStageLog "RAW", ElapsedSec(tStage)

    tStage = Timer
    Application.StatusBar = "Cargando MAIN..."
    Dim loMain As ListObject: Set loMain = EnsureTableForConnection(shMain, "SAB_MC_MAIN", connMain)
    AppendStageLog "MAIN", ElapsedSec(tStage)

    ' Alertas calculadas en VBA
    Dim shAlDep As Worksheet, shAlRet As Worksheet
    Dim loAlDep As ListObject, loAlRet As ListObject

    If makeDep Then
        tStage = Timer
        Application.StatusBar = "Calculando alertas DEP..."
        Set shAlDep = EnsureSheet("SAB_MC_AL_DEP_WORK")
        ClearSheetButKeepName shAlDep
        Set loAlDep = BuildAlertasVBA(loMain, "DEP", shAlDep, "SAB_MC_ALERTAS_DEP")
        AppendStageLog "AL_DEP", ElapsedSec(tStage)
    End If

    If makeRet Then
        tStage = Timer
        Application.StatusBar = "Calculando alertas RET..."
        Set shAlRet = EnsureSheet("SAB_MC_AL_RET_WORK")
        ClearSheetButKeepName shAlRet
        Set loAlRet = BuildAlertasVBA(loMain, "RET", shAlRet, "SAB_MC_ALERTAS_RET")
        AppendStageLog "AL_RET", ElapsedSec(tStage)
    End If

    ' Sufijo de periodo
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

    ' Renombrar RAW y MAIN
    Dim nmRaw  As String: nmRaw  = SanitizeSheetName("SAB_MC_RAW_" & suf)
    Dim nmMain As String: nmMain = SanitizeSheetName("SAB_MC_"     & suf)

    DeleteSheetIfExists ThisWorkbook, nmRaw:   FreeSheetName ThisWorkbook, nmRaw,  shRaw
    DeleteSheetIfExists ThisWorkbook, nmMain:  FreeSheetName ThisWorkbook, nmMain, shMain
    DeleteAllTablesByName ThisWorkbook, nmRaw: DeleteAllTablesByName ThisWorkbook, nmMain
    SetTableNameSafe ThisWorkbook, loRaw,  nmRaw
    SetTableNameSafe ThisWorkbook, loMain, nmMain
    RenameSheetExact shRaw,  nmRaw
    RenameSheetExact shMain, nmMain

    ' Renombrar alertas
    If makeDep Then
        Dim nmAlDep As String: nmAlDep = SanitizeSheetName("SAB_MC_AL_DEP_" & suf)
        DeleteSheetIfExists ThisWorkbook, nmAlDep: FreeSheetName ThisWorkbook, nmAlDep, shAlDep
        DeleteAllTablesByName ThisWorkbook, nmAlDep
        SetTableNameSafe ThisWorkbook, loAlDep, nmAlDep
        RenameSheetExact shAlDep, nmAlDep
    End If

    If makeRet Then
        Dim nmAlRet As String: nmAlRet = SanitizeSheetName("SAB_MC_AL_RET_" & suf)
        DeleteSheetIfExists ThisWorkbook, nmAlRet: FreeSheetName ThisWorkbook, nmAlRet, shAlRet
        DeleteAllTablesByName ThisWorkbook, nmAlRet
        SetTableNameSafe ThisWorkbook, loAlRet, nmAlRet
        RenameSheetExact shAlRet, nmAlRet
    End If

    ' Graficos
    If BUILD_GRAFICOS Then
        If makeDep And Not loAlDep Is Nothing Then
            modSABGraficos.BuildGraficosAlertasEnHoja loAlDep, loMain, "DEP", suf
        End If
        If makeRet And Not loAlRet Is Nothing Then
            modSABGraficos.BuildGraficosAlertasEnHoja loAlRet, loMain, "RET", suf
        End If
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
