Option Explicit

'==========================
' modPQ_SAB_CM
' Carga y procesa transacciones SAB - Cambio de Moneda desde archivo TSV.
'
' El archivo fuente es un TSV exportado como .XLS con estructura:
'   Linea 1 : titulo "Cambios de Moneda"
'   Linea 2 : rango "del DDMMMYYYY al DDMMMYYYY"
'   Linea 3 : vacia
'   Linea 4 : encabezados (24 columnas separadas por tab)
'   Lineas 5+: datos
'
' Todo el procesamiento es en VBA puro, sin OLEDB ni PQ.
' Punto de entrada: CrearQuerySAB_CM(rutaArchivo, mesesSel, opMode, showProgress)
'   opMode: "AMBOS" | "SOLO_COM" | "SOLO_VEN"
'==========================

Private Const BUILD_GRAFICOS  As Boolean = True
Private Const TABLE_STYLE     As String = "TableStyleLight9"
Private Const SKIP_ROWS       As Long = 3

Private mAppFrozen            As Boolean
Private mPrevScreenUpdating   As Boolean
Private mPrevEnableEvents     As Boolean
Private mPrevDisplayAlerts    As Boolean
Private mPrevCalculation      As XlCalculation
Private mPrevStatusBar        As Variant
Private mT0Total              As Double
Private mStageLog             As String

'======================
' Estado Application
'======================
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
        Set sh = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        sh.Name = nm
    End If
    Set EnsureSheet = sh
End Function

Private Sub ClearSheetButKeepName(ByVal sh As Worksheet)
    Dim lo As ListObject, co As ChartObject
    On Error Resume Next
    For Each co In sh.ChartObjects: co.Delete: Next co
    For Each lo In sh.ListObjects: lo.Delete: Next lo
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
    FreeSheetName sh.parent, nm, sh
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

Private Function ParseDDMMMYYYY(ByVal s As String) As Date
    ParseDDMMMYYYY = 0
    If Len(s) < 9 Then Exit Function
    s = UCase$(Trim$(s))
    Dim dd As Integer, yy As Integer, mm As Integer, ms As String
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
    On Error Resume Next
    Set lc = lo.ListColumns(colName)
    On Error GoTo 0
    
    If lc Is Nothing Then Exit Function
    If lc.DataBodyRange Is Nothing Then Exit Function
    
    Dim c As Range, d As Date, gotAny As Boolean
    
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
End Function

'======================
' Elimina separador de miles y simbolos de moneda
'======================
Private Function CleanNum(ByVal s As String) As Double
    On Error Resume Next
    s = Trim$(s)
    s = Replace(s, ",", "")
    s = Replace(s, "S/", "")
    s = Replace(s, "$", "")
    s = Replace(s, " ", "")
    CleanNum = CDbl(s)
    On Error GoTo 0
End Function

'======================
' Normaliza Tipo Persona a "J" o "N"
'======================
Private Function NormTipoPersonaCM(ByVal s As String) As String
    Dim t As String: t = UCase$(Trim$(s))
    t = Replace(t, Chr(205), "I"): t = Replace(t, Chr(201), "E")
    t = Replace(t, Chr(211), "O"): t = Replace(t, Chr(218), "U")
    t = Replace(t, Chr(193), "A"): t = Replace(t, Chr(209), "N")
    If InStr(t, "JUR") > 0 Or t = "J" Or t = "PJ" Then
        NormTipoPersonaCM = "J"
    Else
        NormTipoPersonaCM = "N"
    End If
End Function

'======================
' Normaliza nombre de columna para busqueda: mayusculas, sin diacriticos
'======================
Private Function NormHdr(ByVal s As String) As String
    Dim t As String: t = UCase$(Trim$(s))
    t = Replace(t, Chr(193), "A"): t = Replace(t, Chr(201), "E")
    t = Replace(t, Chr(205), "I"): t = Replace(t, Chr(211), "O")
    t = Replace(t, Chr(218), "U"): t = Replace(t, Chr(209), "N")
    NormHdr = t
End Function

'======================
' LoadCM_TSV
' Lee el archivo TSV directamente en VBA.
' Salta SKIP_ROWS lineas, toma la siguiente como encabezados,
' y carga el resto como datos en sh.
'======================
Private Function LoadCM_TSV(ByVal rutaArchivo As String, _
                              ByVal sh As Worksheet, _
                              ByVal loName As String) As ListObject
    Set LoadCM_TSV = Nothing

    Dim nF As Integer: nF = FreeFile
    On Error GoTo errLoad
    Open rutaArchivo For Input As #nF

    Dim linea As String
    Dim i As Long

    For i = 1 To SKIP_ROWS
        If Not EOF(nF) Then Line Input #nF, linea
    Next i

    If EOF(nF) Then Close #nF: Exit Function
    Line Input #nF, linea

    ' Eliminar BOM si existe
    Do While Len(linea) > 0 And Asc(Left$(linea, 1)) > 127 And Asc(Left$(linea, 1)) < 32
        linea = Mid$(linea, 2)
    Loop

    Dim hdrs() As String: hdrs = Split(linea, Chr(9))
    Dim nCols As Long: nCols = UBound(hdrs) + 1

    ' Leer datos en buffer
    Dim buf() As String
    ReDim buf(0 To 4999)
    Dim nRows As Long: nRows = 0

    Do While Not EOF(nF)
        Line Input #nF, linea
        If Len(Trim$(linea)) = 0 Then GoTo NextLine
        If nRows > UBound(buf) Then ReDim Preserve buf(0 To nRows + 4999)
        buf(nRows) = linea
        nRows = nRows + 1
NextLine:
    Loop
    Close #nF

    If nRows = 0 Then Exit Function

    ClearSheetButKeepName sh

    Dim j As Long
    For j = 0 To nCols - 1
        sh.Cells(1, j + 1).Value = Trim$(hdrs(j))
    Next j

    Dim outArr() As Variant
    ReDim outArr(1 To nRows, 1 To nCols)
    Dim r As Long, cols() As String

    For r = 1 To nRows
        cols = Split(buf(r - 1), Chr(9))
        For j = 0 To nCols - 1
            If j <= UBound(cols) Then
                outArr(r, j + 1) = Trim$(cols(j))
            End If
        Next j
    Next r

    sh.Range(sh.Cells(2, 1), sh.Cells(nRows + 1, nCols)).NumberFormat = "@"
    sh.Range(sh.Cells(2, 1), sh.Cells(nRows + 1, nCols)).Value = outArr

    Dim lo As ListObject
    Set lo = sh.ListObjects.Add(xlSrcRange, _
        sh.Range(sh.Cells(1, 1), sh.Cells(nRows + 1, nCols)), , xlYes)
    On Error Resume Next: lo.Name = loName: On Error GoTo 0
    On Error Resume Next: lo.TableStyle = TABLE_STYLE: On Error GoTo 0

    Set LoadCM_TSV = lo
    Exit Function

errLoad:
    On Error Resume Next: Close #nF: On Error GoTo 0
End Function

'======================
' BuildMainCM_VBA
' Parsea fechas DDMMMYYYY, filtra por rango de meses,
' limpia numeros con separador de miles.
'======================
Private Function BuildMainCM_VBA(ByVal loRaw As ListObject, _
                                  ByVal mesesSel As Long, _
                                  ByVal shMain As Worksheet, _
                                  ByVal loMainName As String) As ListObject
    Set BuildMainCM_VBA = Nothing
    If loRaw Is Nothing Then Exit Function
    If loRaw.DataBodyRange Is Nothing Then Exit Function

    Dim nCols As Long: nCols = loRaw.ListColumns.count

    Dim cFecha   As Long: cFecha = 0
    Dim cTotNeto As Long: cTotNeto = 0
    Dim cMtoOri  As Long: cMtoOri = 0
    Dim cMtoDes  As Long: cMtoDes = 0

    Dim i As Long
    For i = 1 To nCols
        Select Case NormHdr(loRaw.ListColumns(i).Name)
            Case "FECHA":      cFecha = i
            Case "TOTAL NETO": cTotNeto = i
            Case "MONTO ORI":  cMtoOri = i
            Case "MONTO DES":  cMtoDes = i
        End Select
    Next i

    If cFecha = 0 Then Exit Function

    Dim nRows As Long: nRows = loRaw.DataBodyRange.rows.count
    Dim raw As Variant: raw = loRaw.DataBodyRange.Value2

    ' Primera pasada: fecha maxima para derivar rango
    Dim maxDtRaw As Date: maxDtRaw = 0
    Dim tmpD As Date, vF As Variant
    For i = 1 To nRows
        vF = raw(i, cFecha)
        If Not (IsEmpty(vF) Or IsNull(vF) Or IsError(vF)) Then
            tmpD = ParseDDMMMYYYY(CStr(vF))
            If tmpD > maxDtRaw Then maxDtRaw = tmpD
        End If
    Next i

    Dim finMes As Date, iniMes As Date
    If maxDtRaw > 0 Then
        finMes = DateSerial(Year(maxDtRaw), Month(maxDtRaw) + 1, 0)
    Else
        finMes = DateSerial(Year(Date), Month(Date) + 1, 0)
    End If
    iniMes = DateSerial(Year(finMes), Month(finMes) - (mesesSel - 1), 1)

    ' Segunda pasada: filtrar por rango, parsear fechas, limpiar numeros
    ReDim outArr(1 To nRows, 1 To nCols) As Variant
    Dim r As Long: r = 0
    Dim dF As Date, j As Long, sNum As String

    For i = 1 To nRows
        vF = raw(i, cFecha)
        If IsEmpty(vF) Or IsNull(vF) Or IsError(vF) Then GoTo SkipMainCM
        dF = ParseDDMMMYYYY(CStr(vF))
        If dF = 0 Then GoTo SkipMainCM
        If dF < iniMes Or dF > finMes Then GoTo SkipMainCM

        r = r + 1
        For j = 1 To nCols
            outArr(r, j) = raw(i, j)
        Next j

        outArr(r, cFecha) = CDbl(CDate(dF))

        If cTotNeto > 0 Then
            sNum = CStr(outArr(r, cTotNeto))
            If InStr(sNum, ",") > 0 Then outArr(r, cTotNeto) = CleanNum(sNum)
        End If
        If cMtoOri > 0 Then
            sNum = CStr(outArr(r, cMtoOri))
            If InStr(sNum, ",") > 0 Then outArr(r, cMtoOri) = CleanNum(sNum)
        End If
        If cMtoDes > 0 Then
            sNum = CStr(outArr(r, cMtoDes))
            If InStr(sNum, ",") > 0 Then outArr(r, cMtoDes) = CleanNum(sNum)
        End If
SkipMainCM:
    Next i

    If r = 0 Then Exit Function

    ClearSheetButKeepName shMain

    For j = 1 To nCols
        shMain.Cells(1, j).Value = loRaw.ListColumns(j).Name
    Next j

    ' Después:
    shMain.Range(shMain.Cells(2, 1), shMain.Cells(r + 1, nCols)).NumberFormat = "@"
    shMain.Range(shMain.Cells(2, 1), shMain.Cells(r + 1, nCols)).Value = outArr
    
    ' Re-escribir columnas numericas como numeros reales
    Dim numColsCM(3) As Long
    numColsCM(0) = cFecha: numColsCM(1) = cTotNeto
    numColsCM(2) = cMtoOri: numColsCM(3) = cMtoDes
    Dim ncm As Long, ncmR As Long
    Dim ncmArr() As Variant
    For ncm = 0 To 3
        If numColsCM(ncm) > 0 Then
            ReDim ncmArr(1 To r, 1 To 1)
            For ncmR = 1 To r: ncmArr(ncmR, 1) = outArr(ncmR, numColsCM(ncm)): Next ncmR
            With shMain.Range(shMain.Cells(2, numColsCM(ncm)), shMain.Cells(r + 1, numColsCM(ncm)))
                .NumberFormat = "General"
                .Value = ncmArr
            End With
        End If
    Next ncm
    shMain.Columns(cFecha).NumberFormat = "dd/mm/yyyy"

    Dim loMain As ListObject
    Set loMain = shMain.ListObjects.Add(xlSrcRange, _
        shMain.Range(shMain.Cells(1, 1), shMain.Cells(r + 1, nCols)), , xlYes)
    On Error Resume Next: loMain.Name = loMainName: On Error GoTo 0
    On Error Resume Next: loMain.TableStyle = TABLE_STYLE: On Error GoTo 0

    Set BuildMainCM_VBA = loMain
End Function

'======================
' BuildAlertasCM_VBA
' Agrupa loMain por Documento, filtra por Moneda Ori.
' which: "COM" = USD  |  "VEN" = PEN
' Columnas de salida: Documento, TIPO_PERSONA, SUMA_MONTOS, NUM_OPERACIONES,
'                     PROMEDIO_MONTOS, ULTIMA_OPERACION, DESVIACION_MEDIA_%, NIVEL_RIESGO, MONEDA
'======================
Private Function BuildAlertasCM_VBA(ByVal loMain As ListObject, _
                                     ByVal which As String, _
                                     ByVal shAl As Worksheet, _
                                     ByVal loAlName As String) As ListObject
    Set BuildAlertasCM_VBA = Nothing
    If loMain Is Nothing Then Exit Function
    If loMain.DataBodyRange Is Nothing Then Exit Function

    Dim op As String: op = UCase$(Trim$(which))
    Dim wantMon As String
    If op = "COM" Then wantMon = "USD" Else wantMon = "PEN"

    Dim cFecha   As Long: cFecha = 0
    Dim cDoc     As Long: cDoc = 0
    Dim cTipoP   As Long: cTipoP = 0
    Dim cMonOri  As Long: cMonOri = 0
    Dim cTotNeto As Long: cTotNeto = 0

    Dim i As Long
    For i = 1 To loMain.ListColumns.count
        Select Case NormHdr(loMain.ListColumns(i).Name)
            Case "FECHA":        cFecha = i
            Case "DOCUMENTO":    cDoc = i
            Case "TIPO PERSONA": cTipoP = i
            Case "MONEDA ORI":   cMonOri = i
            Case "TOTAL NETO":   cTotNeto = i
        End Select
    Next i

    If cDoc = 0 Or cFecha = 0 Or cTotNeto = 0 Or cMonOri = 0 Then Exit Function

    Dim nRows As Long: nRows = loMain.DataBodyRange.rows.count
    Dim data As Variant: data = loMain.DataBodyRange.Value2

    Dim dDay  As Object: Set dDay = CreateObject("Scripting.Dictionary")
    Dim dMeta As Object: Set dMeta = CreateObject("Scripting.Dictionary")

    Dim vM As Variant, dM As Double
    Dim vC As Variant, sDoc As String
    Dim vF As Variant, dF As Date, lF As Long
    Dim dayKey As String

    For i = 1 To nRows
        ' Filtrar por moneda
        Dim sMon As String
        Dim sMonRaw As String: sMonRaw = UCase$(Trim$(CStr(data(i, cMonOri))))
        If wantMon = "USD" Then
            sMon = IIf(sMonRaw = "USD" Or sMonRaw = "US$" Or sMonRaw = "$", "USD", sMonRaw)
        Else
            sMon = IIf(sMonRaw = "PEN" Or sMonRaw = "S/" Or sMonRaw = "S/.", "PEN", sMonRaw)
        End If
        If sMon <> wantMon Then GoTo SkipAlCM

        vM = data(i, cTotNeto)
        If IsEmpty(vM) Or IsNull(vM) Or IsError(vM) Then GoTo SkipAlCM
        On Error Resume Next: dM = CDbl(vM): On Error GoTo 0
        If dM = 0 Then GoTo SkipAlCM

        vC = data(i, cDoc)
        If IsEmpty(vC) Or IsNull(vC) Or IsError(vC) Then GoTo SkipAlCM
        sDoc = Trim$(CStr(vC))
        If Len(sDoc) = 0 Then GoTo SkipAlCM

        vF = data(i, cFecha)
        If Not TryCoerceExcelDate(vF, dF) Then GoTo SkipAlCM
        lF = CLng(CDbl(CDate(dF)))

        dayKey = sDoc & "|" & CStr(lF)
        If dDay.exists(dayKey) Then
            dDay(dayKey) = CDbl(dDay(dayKey)) + dM
        Else
            dDay.Add dayKey, dM
        End If

        If Not dMeta.exists(sDoc) Then
            Dim sTipoP As String: sTipoP = ""
            If cTipoP > 0 Then
                Dim vTP As Variant: vTP = data(i, cTipoP)
                If Not (IsEmpty(vTP) Or IsNull(vTP) Or IsError(vTP)) Then
                    sTipoP = Trim$(CStr(vTP))
                End If
            End If
            dMeta.Add sDoc, sTipoP
        End If
SkipAlCM:
    Next i

    Dim dSum  As Object: Set dSum = CreateObject("Scripting.Dictionary")
    Dim dNOp  As Object: Set dNOp = CreateObject("Scripting.Dictionary")
    Dim dMaxD As Object: Set dMaxD = CreateObject("Scripting.Dictionary")

    Dim kk As Variant, pts() As String, sDocK As String, lDate As Long, monDia As Double
    For Each kk In dDay.keys
        pts = Split(CStr(kk), "|")
        sDocK = pts(0)
        lDate = CLng(pts(1))
        monDia = CDbl(dDay(kk))
        If Not dSum.exists(sDocK) Then
            dSum.Add sDocK, 0#: dNOp.Add sDocK, 0&: dMaxD.Add sDocK, 0&
        End If
        dSum(sDocK) = CDbl(dSum(sDocK)) + monDia
        dNOp(sDocK) = CLng(dNOp(sDocK)) + 1
        If lDate > CLng(dMaxD(sDocK)) Then dMaxD(sDocK) = lDate
    Next kk

    Dim nDocs As Long: nDocs = dSum.count
    If nDocs = 0 Then Exit Function

    ReDim outArr(1 To nDocs, 1 To 9) As Variant
    Dim r As Long: r = 0
    Dim sDoc2 As Variant, suma As Double, nOp As Long
    Dim prom As Double, ultima As Double
    Dim desv As Variant, nivel As Variant

    For Each sDoc2 In dSum.keys
        r = r + 1
        suma = CDbl(dSum(sDoc2))
        nOp = CLng(dNOp(sDoc2))
        prom = IIf(nOp > 0, suma / nOp, 0)
        ultima = CDbl(dDay(CStr(sDoc2) & "|" & CStr(CLng(dMaxD(sDoc2)))))

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

        outArr(r, 1) = CStr(sDoc2)
        outArr(r, 2) = IIf(dMeta.exists(CStr(sDoc2)), CStr(dMeta(CStr(sDoc2))), "")
        outArr(r, 3) = Round(suma, 2)
        outArr(r, 4) = nOp
        outArr(r, 5) = Round(prom, 2)
        outArr(r, 6) = Round(ultima, 2)
        outArr(r, 7) = desv
        outArr(r, 8) = nivel
        outArr(r, 9) = wantMon
    Next sDoc2

    ClearSheetButKeepName shAl

    Dim hdrs As Variant
    hdrs = Array("Documento", "TIPO_PERSONA", "SUMA_MONTOS", "NUM_OPERACIONES", _
                 "PROMEDIO_MONTOS", "ULTIMA_OPERACION", "DESVIACION_MEDIA_%", _
                 "NIVEL_RIESGO", "MONEDA")
    Dim j As Long
    For j = 0 To 8: shAl.Cells(1, j + 1).Value = hdrs(j): Next j
    ' Después:
    shAl.Range(shAl.Cells(2, 1), shAl.Cells(nDocs + 1, 9)).NumberFormat = "@"
    shAl.Range(shAl.Cells(2, 1), shAl.Cells(nDocs + 1, 9)).Value = outArr
    
    ' Re-escribir columnas numericas (3=SUMA a 8=NIVEL_RIESGO)
    Dim nalR As Long, nalArr() As Variant, nalC As Long
    For nalC = 3 To 8
        ReDim nalArr(1 To nDocs, 1 To 1)
        For nalR = 1 To nDocs: nalArr(nalR, 1) = outArr(nalR, nalC): Next nalR
        With shAl.Range(shAl.Cells(2, nalC), shAl.Cells(nDocs + 1, nalC))
            .NumberFormat = "General"
            .Value = nalArr
        End With
    Next nalC

    Dim loAL As ListObject
    Set loAL = shAl.ListObjects.Add(xlSrcRange, _
        shAl.Range(shAl.Cells(1, 1), shAl.Cells(nDocs + 1, 9)), , xlYes)
    On Error Resume Next: loAL.Name = loAlName: On Error GoTo 0
    On Error Resume Next: loAL.TableStyle = TABLE_STYLE: On Error GoTo 0

    On Error Resume Next
    loAL.Sort.SortFields.Clear
    loAL.Sort.SortFields.Add key:=loAL.ListColumns("DESVIACION_MEDIA_%").DataBodyRange, _
        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With loAL.Sort
        .header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
    On Error GoTo 0

    Set BuildAlertasCM_VBA = loAL
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

    mT0Total = Timer
    mStageLog = vbNullString

    If mesesSel <= 0 Then mesesSel = 6
    If Len(Trim$(opMode)) = 0 Then opMode = "AMBOS"

    Dim makeCOM As Boolean: makeCOM = (UCase$(opMode) = "AMBOS" Or UCase$(opMode) = "SOLO_COM")
    Dim makeVEN As Boolean: makeVEN = (UCase$(opMode) = "AMBOS" Or UCase$(opMode) = "SOLO_VEN")

    SafeApp True

    ' Hojas de trabajo
    Dim shRaw  As Worksheet: Set shRaw = EnsureSheet("SAB_CM_RAW_WORK")
    Dim shMain As Worksheet: Set shMain = EnsureSheet("SAB_CM_WORK")
    ClearSheetButKeepName shRaw
    ClearSheetButKeepName shMain

    Dim tStage As Double

    ' Cargar RAW desde TSV
    tStage = Timer
    SAB_Progress 0.1, "Cargando RAW CM..."
    Dim loRaw As ListObject: Set loRaw = LoadCM_TSV(rutaArchivo, shRaw, "SAB_CM_RAW")
    If loRaw Is Nothing Then
        Err.Raise 9001, , "No se pudo leer el archivo TSV: " & rutaArchivo
    End If
    AppendStageLog "RAW CM", ElapsedSec(tStage)

    ' Construir MAIN
    tStage = Timer
    SAB_Progress 0.35, "Construyendo MAIN CM..."
    Dim loMain As ListObject: Set loMain = BuildMainCM_VBA(loRaw, mesesSel, shMain, "SAB_CM_MAIN")
    If loMain Is Nothing Then
        Err.Raise 9002, , "MAIN CM vacio: sin datos en el rango de fechas seleccionado."
    End If
    AppendStageLog "MAIN CM", ElapsedSec(tStage)

    ' Alertas
    Dim shAlCom As Worksheet, shAlVen As Worksheet
    Dim loAlCom As ListObject, loAlVen As ListObject

    If makeCOM Then
        tStage = Timer
        SAB_Progress 0.6, "Calculando alertas COM (USD)..."
        Set shAlCom = EnsureSheet("SAB_CM_AL_COM_WORK")
        ClearSheetButKeepName shAlCom
        Set loAlCom = BuildAlertasCM_VBA(loMain, "COM", shAlCom, "SAB_CM_ALERTAS_COM")
        AppendStageLog "AL_COM", ElapsedSec(tStage)
    End If

    If makeVEN Then
        tStage = Timer
        SAB_Progress 0.75, "Calculando alertas VEN (PEN)..."
        Set shAlVen = EnsureSheet("SAB_CM_AL_VEN_WORK")
        ClearSheetButKeepName shAlVen
        Set loAlVen = BuildAlertasCM_VBA(loMain, "VEN", shAlVen, "SAB_CM_ALERTAS_VEN")
        AppendStageLog "AL_VEN", ElapsedSec(tStage)
    End If

    ' Sufijo de periodo desde loMain
    Dim minD As Date, maxD As Date, gotDates As Boolean
    gotDates = GetMinMaxDateFromLO(loMain, "Fecha", minD, maxD)

    Dim ini As Date, fin As Date, suf As String
    If gotDates Then
        ini = FirstDayOfMonth(minD): fin = LastDayOfMonth(maxD)
    Else
        fin = DateSerial(Year(Date), Month(Date), 0)
        ini = DateSerial(Year(fin), Month(fin) - (mesesSel - 1), 1)
    End If
    suf = MesAbrevES(ini) & "_" & MesAbrevES(fin) & "_" & Year(fin)

    ' Renombrar hojas
    Dim nmRaw  As String: nmRaw = SanitizeSheetName("SAB_CM_RAW_" & suf)
    Dim nmMain As String: nmMain = SanitizeSheetName("SAB_CM_" & suf)

    DeleteSheetIfExists ThisWorkbook, nmRaw
    DeleteSheetIfExists ThisWorkbook, nmMain
    FreeSheetName ThisWorkbook, nmRaw, shRaw
    FreeSheetName ThisWorkbook, nmMain, shMain
    DeleteAllTablesByName ThisWorkbook, nmRaw
    DeleteAllTablesByName ThisWorkbook, nmMain
    SetTableNameSafe ThisWorkbook, loRaw, nmRaw
    SetTableNameSafe ThisWorkbook, loMain, nmMain
    RenameSheetExact shRaw, nmRaw
    RenameSheetExact shMain, nmMain

    If makeCOM And Not loAlCom Is Nothing Then
        Dim nmAlCom As String: nmAlCom = SanitizeSheetName("SAB_CM_AL_COM_" & suf)
        DeleteSheetIfExists ThisWorkbook, nmAlCom
        FreeSheetName ThisWorkbook, nmAlCom, shAlCom
        DeleteAllTablesByName ThisWorkbook, nmAlCom
        SetTableNameSafe ThisWorkbook, loAlCom, nmAlCom
        RenameSheetExact shAlCom, nmAlCom
    End If

    If makeVEN And Not loAlVen Is Nothing Then
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

    SafeApp False

    Dim totalMsg As String
    totalMsg = "SAB - Cambio de Moneda cargado." & vbCrLf & vbCrLf & _
               mStageLog & vbCrLf & vbCrLf & _
               "Total: " & FormatElapsed(ElapsedSec(mT0Total))

    SAB_Progress 1#, "SAB CM listo. Total " & FormatElapsed(ElapsedSec(mT0Total))
    Debug.Print totalMsg
    If showProgress Then MsgBox totalMsg, vbInformation, "SAB CM"

    If makeCOM And Not loAlCom Is Nothing Then
        shAlCom.Activate
    ElseIf makeVEN And Not loAlVen Is Nothing Then
        shAlVen.Activate
    Else
        shMain.Activate
    End If
    ActiveSheet.Range("A1").Select
    Exit Sub

EH:
    SafeApp False
    MsgBox "Error en CrearQuerySAB_CM: " & Err.Number & " - " & Err.Description, vbCritical
End Sub
