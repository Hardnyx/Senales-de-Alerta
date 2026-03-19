Option Explicit

'==========================
' modPQ_SAB_MC
' RAW via PQ (lectura del archivo externo).
' MAIN y Alertas en VBA para evitar re-evaluacion lazy de PQ sobre 61k filas.
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

' Convierte serial Excel o Date a Date. Devuelve True si tuvo exito.
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

' Parsea string DDMMMYYYY (ej. "27SEP2024") a Date.
' Devuelve 0 si falla.
Private Function ParseDDMMMYYYY(ByVal s As String) As Date
    ParseDDMMMYYYY = 0
    If Len(s) < 9 Then Exit Function
    s = UCase$(Trim$(s))
    Dim dd As Integer, yy As Integer, mm As Integer
    Dim ms As String
    On Error GoTo fin
    dd = CInt(Left$(s, 2))
    ms = Mid$(s, 3, 3)
    yy = CInt(Right$(s, 4))
    Select Case ms
        Case "ENE": mm = 1:  Case "FEB": mm = 2:  Case "MAR": mm = 3
        Case "ABR": mm = 4:  Case "MAY": mm = 5:  Case "JUN": mm = 6
        Case "JUL": mm = 7:  Case "AGO": mm = 8:  Case "SET": mm = 9
        Case "SEP": mm = 9:  Case "OCT": mm = 10: Case "NOV": mm = 11
        Case "DIC": mm = 12
        Case Else: Exit Function
    End Select
    If dd < 1 Or dd > 31 Or mm < 1 Or mm > 12 Or yy < 1900 Then Exit Function
    ParseDDMMMYYYY = DateSerial(yy, mm, dd)
    Exit Function
fin:
    ParseDDMMMYYYY = 0
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
' Power Query helpers (solo para RAW)
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
' M query: solo RAW
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

'======================
' Helper: buscar indice de columna por lista de alternativas (case-insensitive)
'======================
Private Function PickColIdx(ByVal colNames() As String, ByVal alts() As String) As Long
    Dim i As Long, j As Long
    For j = 0 To UBound(alts)
        For i = 0 To UBound(colNames)
            If StrComp(colNames(i), alts(j), vbTextCompare) = 0 Then
                PickColIdx = i + 1  ' 1-based como ListColumns
                Exit Function
            End If
        Next i
    Next j
    PickColIdx = 0
End Function

'======================
' MAIN en VBA
' Lee loRaw, aplica la misma logica que M_MC_MAIN pero en VBA nativo:
'   - Renombrado de columnas via Pick
'   - Parseo de fechas DDMMMYYYY con Select Case (nanosegundos por fila)
'   - Filtro por Clase (DPE/DFS/RAF/RFS)
'   - Separacion exclusiva DEP/RET
'   - Filtro por rango de meses
'   - Ordenar por Fecha
' Devuelve el ListObject creado en shMain.
'======================
Private Function BuildMainVBA(ByVal loRaw As ListObject, _
                               ByVal mesesSel As Long, _
                               ByVal shMain As Worksheet, _
                               ByVal loMainName As String) As ListObject
    Set BuildMainVBA = Nothing
    If loRaw Is Nothing Then Exit Function
    If loRaw.DataBodyRange Is Nothing Then Exit Function

    Dim depName As String: depName = "Dep" & Chr(243) & "sito"

    ' --- Leer nombres de columna de loRaw ---
    Dim nCols As Long: nCols = loRaw.ListColumns.Count
    ReDim rawColNames(0 To nCols - 1) As String
    Dim i As Long
    For i = 1 To nCols
        rawColNames(i - 1) = loRaw.ListColumns(i).Name
    Next i

    ' --- Pick: mapear columnas fuente a indices canonicos ---
    Dim cFecha  As Long: cFecha  = PickColIdx(rawColNames, Split("Fecha|FECHA|Fec|FECHA MOV|Fecha Mov", "|"))
    Dim cTrans  As Long: cTrans  = PickColIdx(rawColNames, Split("Transac|TRANSAC|Transacci" & Chr(243) & "n|Transaccion", "|"))
    Dim cCuenta As Long: cCuenta = PickColIdx(rawColNames, Split("Cuenta|CUENTA|Cta|Nro Cuenta|Nro. Cuenta|N" & Chr(186) & " Cuenta", "|"))
    Dim cNombre As Long: cNombre = PickColIdx(rawColNames, Split("Nombre|A La Orden|ALaOrden|A_la_Orden", "|"))
    Dim cOpe    As Long: cOpe    = PickColIdx(rawColNames, Split("Ope|OPE", "|"))
    Dim cTipo   As Long: cTipo   = PickColIdx(rawColNames, Split("Tipo|TIPO", "|"))
    Dim cFPag   As Long: cFPag   = PickColIdx(rawColNames, Split("FPag|F. Pag.|F. Pago|Fecha Pago", "|"))
    Dim cClase  As Long: cClase  = PickColIdx(rawColNames, Split("Clase|CLASE", "|"))
    Dim cALaOr  As Long: cALaOr  = PickColIdx(rawColNames, Split("ALaOrden|A La Orden|Nombre", "|"))
    Dim cDep    As Long: cDep    = PickColIdx(rawColNames, Split(depName & "|Deposito|Abono", "|"))
    Dim cRet    As Long: cRet    = PickColIdx(rawColNames, Split("Retiro|Cargo", "|"))
    Dim cCtaLiq As Long: cCtaLiq = PickColIdx(rawColNames, Split("CtaLiq|Cta Liq|Cta Liquidez|Cuenta Liquidaci" & Chr(243) & "n|Cuenta Liquidacion", "|"))
    Dim cEst    As Long: cEst    = PickColIdx(rawColNames, Split("Estado|ESTADO", "|"))
    Dim cObs    As Long: cObs    = PickColIdx(rawColNames, Split("Observaciones|Obs", "|"))
    Dim cMon    As Long: cMon    = PickColIdx(rawColNames, Split("Moneda", "|"))

    If cFecha = 0 Then Exit Function

    ' --- Leer datos en array ---
    Dim nRows As Long: nRows = loRaw.DataBodyRange.Rows.Count
    Dim raw   As Variant: raw = loRaw.DataBodyRange.Value2

    ' --- Target: 15 columnas canonicas de MAIN ---
    ' Fecha, Transac, Cuenta, Nombre, Ope, Tipo, FPag, Clase,
    ' ALaOrden, Deposito, Retiro, CtaLiq, Estado, Observaciones, Moneda
    Dim TARGET_COLS As Long: TARGET_COLS = 15

    ' Indices en outArr (1-based en columna)
    Const O_FECHA  As Long = 1:  Const O_TRANSAC As Long = 2:  Const O_CUENTA As Long = 3
    Const O_NOMBRE As Long = 4:  Const O_OPE     As Long = 5:  Const O_TIPO   As Long = 6
    Const O_FPAG   As Long = 7:  Const O_CLASE   As Long = 8:  Const O_ALAOR  As Long = 9
    Const O_DEP    As Long = 10: Const O_RET     As Long = 11: Const O_CTALIQ As Long = 12
    Const O_EST    As Long = 13: Const O_OBS     As Long = 14: Const O_MON    As Long = 15

    ' Pre-alocar output (worst case = nRows)
    ReDim outArr(1 To nRows, 1 To TARGET_COLS) As Variant

    Dim r    As Long: r = 0
    Dim vF   As Variant, dF As Date
    Dim vCl  As Variant, sCl As String
    Dim vDep As Variant, vRet As Variant, nDep As Double, nRet As Double
    Dim hasDep As Boolean, hasRet As Boolean

    ' Rango de fechas
    Dim today As Date: today = Date
    Dim finMes As Date: finMes = DateSerial(Year(today), Month(today) + 1, 0)
    Dim iniMes As Date: iniMes = DateSerial(Year(finMes), Month(finMes) - (mesesSel - 1), 1)

    For i = 1 To nRows
        ' --- Fecha ---
        vF = raw(i, cFecha)
        ' Intentar como serial/Date primero; si falla, intentar parsear como texto DDMMMYYYY
        If Not TryCoerceExcelDate(vF, dF) Then
            If IsEmpty(vF) Or IsNull(vF) Or IsError(vF) Then GoTo SkipRow
            dF = ParseDDMMMYYYY(CStr(vF))
            If dF = 0 Then GoTo SkipRow
        End If

        ' --- Filtro Clase (DPE / DFS / RAF / RFS) ---
        If cClase > 0 Then
            vCl = raw(i, cClase)
            If IsEmpty(vCl) Or IsNull(vCl) Or IsError(vCl) Then GoTo SkipRow
            sCl = UCase$(Trim$(CStr(vCl)))
            Select Case sCl
                Case "DPE", "DFS", "RAF", "RFS"
                    ' OK
                Case Else
                    GoTo SkipRow
            End Select
        End If

        ' --- Filtro rango de fechas ---
        If dF < iniMes Or dF > finMes Then GoTo SkipRow

        ' --- Separacion exclusiva DEP / RET ---
        hasDep = False: hasDep = False
        nDep = 0: nRet = 0
        If cDep > 0 Then
            vDep = raw(i, cDep)
            If Not (IsEmpty(vDep) Or IsNull(vDep) Or IsError(vDep)) Then
                On Error Resume Next: nDep = CDbl(vDep): On Error GoTo 0
                hasDep = (nDep <> 0)
            End If
        End If
        If cRet > 0 Then
            vRet = raw(i, cRet)
            If Not (IsEmpty(vRet) Or IsNull(vRet) Or IsError(vRet)) Then
                On Error Resume Next: nRet = CDbl(vRet): On Error GoTo 0
                hasRet = (nRet <> 0)
            End If
        End If

        ' Si tiene DEP, DEP prevalece; RET solo si DEP es null/0
        Dim finalDep As Variant: finalDep = Null
        Dim finalRet As Variant: finalRet = Null
        If hasDep Then
            finalDep = nDep
        ElseIf hasRet Then
            finalRet = nRet
        Else
            GoTo SkipRow   ' fila sin movimiento
        End If

        ' --- Escribir fila ---
        r = r + 1
        outArr(r, O_FECHA) = CDbl(CDate(dF))   ' serial numerico para que Excel lo reconozca como fecha

        If cTrans  > 0 Then outArr(r, O_TRANSAC) = raw(i, cTrans)
        If cCuenta > 0 Then outArr(r, O_CUENTA)  = raw(i, cCuenta)
        If cNombre > 0 Then outArr(r, O_NOMBRE)  = raw(i, cNombre)
        If cOpe    > 0 Then outArr(r, O_OPE)     = raw(i, cOpe)
        If cTipo   > 0 Then outArr(r, O_TIPO)    = raw(i, cTipo)
        If cFPag   > 0 Then outArr(r, O_FPAG)    = raw(i, cFPag)
        outArr(r, O_CLASE) = sCl
        If cALaOr  > 0 Then outArr(r, O_ALAOR)   = raw(i, cALaOr)
        outArr(r, O_DEP)   = finalDep
        outArr(r, O_RET)   = finalRet
        If cCtaLiq > 0 Then outArr(r, O_CTALIQ)  = raw(i, cCtaLiq)
        If cEst    > 0 Then outArr(r, O_EST)      = raw(i, cEst)
        If cObs    > 0 Then outArr(r, O_OBS)      = raw(i, cObs)
        If cMon    > 0 Then outArr(r, O_MON)      = raw(i, cMon)

SkipRow:
    Next i

    If r = 0 Then Exit Function

    ' --- Ordenar en memoria por Fecha (burbuja rapida, pocos meses = casi ordenado ya) ---
    ' Para 61k filas usar QuickSort en columna O_FECHA
    QuickSortByCol outArr, 1, r, O_FECHA

    ' --- Escribir en hoja ---
    ClearSheetButKeepName shMain

    Dim hdrs As Variant
    hdrs = Array("Fecha", "Transac", "Cuenta", "Nombre", "Ope", "Tipo", "FPag", _
                 "Clase", "ALaOrden", depName, "Retiro", "CtaLiq", "Estado", "Observaciones", "Moneda")
    Dim j As Long
    For j = 0 To 14: shMain.Cells(1, j + 1).Value = hdrs(j): Next j

    shMain.Range(shMain.Cells(2, 1), shMain.Cells(r + 1, TARGET_COLS)).Value = outArr

    ' Formatear columna Fecha como fecha
    shMain.Columns(O_FECHA).NumberFormat = "dd/mm/yyyy"

    ' --- Crear ListObject estatico ---
    Dim loMain As ListObject
    Set loMain = shMain.ListObjects.Add(xlSrcRange, _
                     shMain.Range(shMain.Cells(1, 1), shMain.Cells(r + 1, TARGET_COLS)), , xlYes)
    On Error Resume Next: loMain.Name = loMainName: On Error GoTo 0
    On Error Resume Next: loMain.TableStyle = TABLE_STYLE: On Error GoTo 0

    Set BuildMainVBA = loMain
End Function

'======================
' QuickSort in-place sobre array 2D por columna sortCol (valores numericos/fecha)
'======================
Private Sub QuickSortByCol(ByRef arr() As Variant, ByVal lo As Long, ByVal hi As Long, ByVal sortCol As Long)
    If lo >= hi Then Exit Sub
    Dim pivot As Double: pivot = CDbl(arr((lo + hi) \ 2, sortCol))
    Dim i As Long: i = lo
    Dim j As Long: j = hi
    Dim tmp As Variant
    Dim c As Long
    Do While i <= j
        Do While CDbl(arr(i, sortCol)) < pivot: i = i + 1: Loop
        Do While CDbl(arr(j, sortCol)) > pivot: j = j - 1: Loop
        If i <= j Then
            For c = 1 To UBound(arr, 2)
                tmp = arr(i, c): arr(i, c) = arr(j, c): arr(j, c) = tmp
            Next c
            i = i + 1: j = j - 1
        End If
    Loop
    If lo < j Then QuickSortByCol arr, lo, j, sortCol
    If i < hi Then QuickSortByCol arr, i, hi, sortCol
End Sub

'======================
' Alertas en VBA desde loMain
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

    Dim dDay  As Object: Set dDay  = CreateObject("Scripting.Dictionary")
    Dim dMeta As Object: Set dMeta = CreateObject("Scripting.Dictionary")

    Dim vM As Variant, dM As Double
    Dim vC As Variant, sC As String
    Dim vF As Variant, dF As Date, lF As Long
    Dim dayKey As String

    For i = 1 To nRows
        vM = data(i, colMonto)
        If IsEmpty(vM) Or IsNull(vM) Or IsError(vM) Then GoTo SkipAl
        On Error Resume Next: dM = CDbl(vM): On Error GoTo 0
        If dM = 0 Then GoTo SkipAl

        vC = data(i, colCuenta)
        If IsEmpty(vC) Or IsNull(vC) Or IsError(vC) Then GoTo SkipAl
        sC = Trim$(CStr(vC))
        If Len(sC) = 0 Then GoTo SkipAl

        If colFecha = 0 Then GoTo SkipAl
        vF = data(i, colFecha)
        If Not TryCoerceExcelDate(vF, dF) Then GoTo SkipAl
        lF = CLng(CDbl(CDate(dF)))

        dayKey = sC & "|" & CStr(lF)
        If dDay.Exists(dayKey) Then
            dDay(dayKey) = CDbl(dDay(dayKey)) + dM
        Else
            dDay.Add dayKey, dM
        End If

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
SkipAl:
    Next i

    Dim dSum  As Object: Set dSum  = CreateObject("Scripting.Dictionary")
    Dim dNOp  As Object: Set dNOp  = CreateObject("Scripting.Dictionary")
    Dim dMaxD As Object: Set dMaxD = CreateObject("Scripting.Dictionary")

    Dim kk As Variant, pts() As String, sCta As String, lDate As Long, monDia As Double
    For Each kk In dDay.Keys
        pts    = Split(CStr(kk), "|")
        sCta   = pts(0)
        lDate  = CLng(pts(1))
        monDia = CDbl(dDay(kk))
        If Not dSum.Exists(sCta) Then
            dSum.Add sCta, 0#: dNOp.Add sCta, 0&: dMaxD.Add sCta, 0&
        End If
        dSum(sCta)  = CDbl(dSum(sCta)) + monDia
        dNOp(sCta)  = CLng(dNOp(sCta)) + 1
        If lDate > CLng(dMaxD(sCta)) Then dMaxD(sCta) = lDate
    Next kk

    Dim nCtas As Long: nCtas = dSum.Count
    If nCtas = 0 Then Exit Function

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

    ClearSheetButKeepName shAl

    Dim hdrs As Variant
    hdrs = Array("Cuenta", "CLASE", "MONEDA", "SUMA_MONTOS", "NUM_OPERACIONES", _
                 "PROMEDIO_MONTOS", "ULTIMA_OPERACION", "DESVIACION_MEDIA_%", "NIVEL_RIESGO")
    Dim j As Long
    For j = 0 To 8: shAl.Cells(1, j + 1).Value = hdrs(j): Next j
    shAl.Range(shAl.Cells(2, 1), shAl.Cells(nCtas + 1, 9)).Value = outArr

    Dim loAl As ListObject
    Set loAl = shAl.ListObjects.Add(xlSrcRange, _
                   shAl.Range(shAl.Cells(1, 1), shAl.Cells(nCtas + 1, 9)), , xlYes)
    On Error Resume Next: loAl.Name = loAlName: On Error GoTo 0
    On Error Resume Next: loAl.TableStyle = TABLE_STYLE: On Error GoTo 0

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

    ' Solo RAW en PQ
    UpsertWorkbookQuery "SAB_MC_RAW", M_MC_RAW(rutaArchivo)

    Dim shRaw  As Worksheet: Set shRaw  = EnsureSheet("SAB_MC_RAW_WORK")
    Dim shMain As Worksheet: Set shMain = EnsureSheet("SAB_MC_MAIN_WORK")
    ClearSheetButKeepName shRaw
    ClearSheetButKeepName shMain

    Dim connRaw As WorkbookConnection: Set connRaw = EnsurePQConnection("SAB_MC_RAW")

    Dim tStage As Double

    tStage = Timer
    Application.StatusBar = "Cargando RAW..."
    Dim loRaw As ListObject: Set loRaw = EnsureTableForConnection(shRaw, "SAB_MC_RAW", connRaw)
    AppendStageLog "RAW", ElapsedSec(tStage)

    tStage = Timer
    Application.StatusBar = "Construyendo MAIN..."
    Dim loMain As ListObject: Set loMain = BuildMainVBA(loRaw, mesesSel, shMain, "SAB_MC_MAIN")
    AppendStageLog "MAIN", ElapsedSec(tStage)

    ' Alertas en VBA
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

    ' Renombrar RAW y MAIN
    Dim nmRaw  As String: nmRaw  = SanitizeSheetName("SAB_MC_RAW_"  & suf)
    Dim nmMain As String: nmMain = SanitizeSheetName("SAB_MC_"      & suf)

    DeleteSheetIfExists ThisWorkbook, nmRaw:  FreeSheetName ThisWorkbook, nmRaw,  shRaw
    DeleteSheetIfExists ThisWorkbook, nmMain: FreeSheetName ThisWorkbook, nmMain, shMain
    DeleteAllTablesByName ThisWorkbook, nmRaw:  DeleteAllTablesByName ThisWorkbook, nmMain
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
