Option Explicit

'=========================================================
' modFondosGraficos
' Genera 10 graficos de serie temporal (XY Scatter con linea)
' en la hoja de ALERTAS:
'   Columna izquierda : top 5 DESVIACION_MEDIA_% con TIPO PERSONA NAT o MAN
'   Columna derecha   : top 5 DESVIACION_MEDIA_% con TIPO PERSONA JUR
'
' Cada grafico:
'   - Serie solida    : monto diario agregado por fecha (de loMAIN)
'   - Serie punteada  : promedio plano extendido al rango completo del eje
'   - Eje X proporcional a fechas reales, etiquetas "Mmm.AA" (ej: Jul.25)
'   - Ticks en el inicio de cada mes dentro del rango
'   - Titulo con tipo persona, tipo de operacion, clave y desviacion
'
' Requiere hoja auxiliar oculta _GF_HELPER para datos de las series.
'
' Firma publica:
'   BuildGraficosAlertasEnHoja(loAL, loMAIN, loClientes, opCode)
'=========================================================

Private Const HELPER_SH     As String = "_GF_HELPER"
Private Const CHART_W       As Double = 510     ' ~18 cm en puntos
Private Const CHART_H       As Double = 255     ' ~9 cm en puntos
Private Const CHART_GAP_H   As Double = 10      ' separacion vertical entre graficos
Private Const CHART_GAP_C   As Double = 14      ' separacion horizontal entre columnas
Private Const CHART_TOP_MGN As Double = 30      ' margen superior antes del primer grafico
Private Const MAX_PER_COL   As Long = 5
Private Const CLI_BLOCK     As Long = 400       ' filas helper reservadas por cliente

' =========================================================
' Texto
' =========================================================

Private Function StripDiacriticsUpper(ByVal s As String) As String
    Dim t As String
    t = UCase$(Trim$(s))
    t = Replace(t, Chr(193), "A"): t = Replace(t, Chr(192), "A")
    t = Replace(t, Chr(194), "A"): t = Replace(t, Chr(196), "A")
    t = Replace(t, Chr(201), "E"): t = Replace(t, Chr(200), "E")
    t = Replace(t, Chr(202), "E"): t = Replace(t, Chr(203), "E")
    t = Replace(t, Chr(205), "I"): t = Replace(t, Chr(204), "I")
    t = Replace(t, Chr(206), "I"): t = Replace(t, Chr(207), "I")
    t = Replace(t, Chr(211), "O"): t = Replace(t, Chr(210), "O")
    t = Replace(t, Chr(212), "O"): t = Replace(t, Chr(214), "O")
    t = Replace(t, Chr(218), "U"): t = Replace(t, Chr(217), "U")
    t = Replace(t, Chr(219), "U"): t = Replace(t, Chr(220), "U")
    t = Replace(t, Chr(209), "N")
    StripDiacriticsUpper = t
End Function

Private Function CanonColName(ByVal s As String) As String
    Dim t As String
    t = StripDiacriticsUpper(s)
    t = Replace(t, Chr$(160), " ")
    t = Replace(t, Chr(176), "")
    t = Replace(t, Chr(186), "")
    t = Replace(t, " ", "")
    CanonColName = t
End Function

Private Function FindListColumnByName(ByVal lo As ListObject, ByVal colName As String) As ListColumn
    Dim lc As ListColumn
    Dim want As String
    want = CanonColName(colName)
    For Each lc In lo.ListColumns
        If CanonColName(lc.Name) = want Then
            Set FindListColumnByName = lc
            Exit Function
        End If
    Next lc
    Set FindListColumnByName = Nothing
End Function

Private Function LOHasColumn(ByVal lo As ListObject, ByVal colName As String) As Boolean
    LOHasColumn = False
    If lo Is Nothing Then Exit Function
    LOHasColumn = Not (FindListColumnByName(lo, colName) Is Nothing)
End Function

Private Function GetColIdx(ByVal lo As ListObject, ByVal colName As String) As Long
    Dim lc As ListColumn
    Set lc = FindListColumnByName(lo, colName)
    If lc Is Nothing Then GetColIdx = 0 Else GetColIdx = lc.Index
End Function

' =========================================================
' Utilidades
' =========================================================

Private Function SafeDbl(ByVal v As Variant) As Double
    On Error Resume Next
    SafeDbl = CDbl(v)
    On Error GoTo 0
End Function

' NiceFloor: devuelve el mayor numero "bonito" (mantisa 1, 2, 5 x 10^n)
' que sea menor o igual a v. Evita resultados con divisiones entre 3.
' Ejemplos: NiceFloor(6667) = 5000, NiceFloor(3333) = 2000,
'           NiceFloor(4999) = 2000, NiceFloor(5000) = 5000
Private Function NiceFloor(ByVal v As Double) As Double
    If v <= 0 Then NiceFloor = 1: Exit Function
    Dim mag As Double
    mag = 10 ^ Int(log(v) / log(10))
    Dim m As Double
    m = v / mag
    Dim niceM As Double
    If m >= 5 Then
        niceM = 5
    ElseIf m >= 2 Then
        niceM = 2
    Else
        niceM = 1
    End If
    NiceFloor = niceM * mag
End Function

Private Sub DeleteChartsByPrefix(ByVal ws As Worksheet, ByVal pref As String)
    Dim co As ChartObject
    Dim nms() As String
    Dim cnt As Long
    Dim i As Long
    cnt = 0
    On Error Resume Next
    For Each co In ws.ChartObjects
        If StrComp(Left$(co.Name, Len(pref)), pref, vbTextCompare) = 0 Then
            ReDim Preserve nms(cnt)
            nms(cnt) = co.Name
            cnt = cnt + 1
        End If
    Next co
    For i = 0 To cnt - 1
        ws.ChartObjects(nms(i)).Delete
    Next i
    On Error GoTo 0
End Sub

Private Function EnsureHelperSheet(ByVal wb As Workbook) As Worksheet
    Dim wsh As Worksheet
    On Error Resume Next
    Set wsh = wb.Worksheets(HELPER_SH)
    On Error GoTo 0
    If wsh Is Nothing Then
        Set wsh = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        wsh.Name = HELPER_SH
    Else
        wsh.Cells.Clear
    End If
    wsh.Visible = xlSheetVeryHidden
    Set EnsureHelperSheet = wsh
End Function

' Ordena n elementos descendentemente segun dv().
Private Sub SortDesc5(ByVal n As Long, _
    k() As String, dv() As Double, pm() As Double, _
    tp() As String, td() As String)
    Dim i As Long, j As Long
    Dim ts As String, tf As Double
    For i = 0 To n - 2
        For j = 0 To n - i - 2
            If dv(j) < dv(j + 1) Then
                ts = k(j):  k(j) = k(j + 1):   k(j + 1) = ts
                tf = dv(j): dv(j) = dv(j + 1): dv(j + 1) = tf
                tf = pm(j): pm(j) = pm(j + 1): pm(j + 1) = tf
                ts = tp(j): tp(j) = tp(j + 1): tp(j + 1) = ts
                ts = td(j): td(j) = td(j + 1): td(j + 1) = ts
            End If
        Next j
    Next i
End Sub

' =========================================================
' WriteMontoSeries
' Agrega montos diarios para el conjunto de CUCs en wsh
' columnas A-B desde blockStart.
' Devuelve el numero de filas escritas y el rango de fechas.
' La serie de promedio se escribe por separado con WritePromedioSeries.
' =========================================================
Private Function WriteMontoSeries( _
    ByVal wsh As Worksheet, _
    ByVal blockStart As Long, _
    ByVal cucList As String, _
    ByVal iCUC As Long, _
    ByVal iFecha As Long, _
    ByVal iMonto As Long, _
    ByRef arrM As Variant, _
    ByRef minDtOut As Date, _
    ByRef maxDtOut As Date) As Long

    On Error GoTo errExit

    Dim dCUC As Object
    Set dCUC = CreateObject("Scripting.Dictionary")
    Dim parts() As String
    parts = Split(cucList, "|")
    Dim p As Long
    For p = 0 To UBound(parts)
        Dim cc As String
        cc = Trim$(parts(p))
        If cc <> "" And Not dCUC.exists(cc) Then dCUC.Add cc, 1
    Next p

    Dim dDM As Object
    Set dDM = CreateObject("Scripting.Dictionary")

    Dim r As Long
    Dim nRows As Long
    nRows = UBound(arrM, 1)

    For r = 1 To nRows
        Dim sv As String
        sv = Trim$(CStr(arrM(r, iCUC)))
        If Not dCUC.exists(sv) Then GoTo NextR

        Dim rawFecha As Variant
        rawFecha = arrM(r, iFecha)
        If IsEmpty(rawFecha) Or IsNull(rawFecha) Then GoTo NextR
        If Not IsDate(rawFecha) And Not IsNumeric(rawFecha) Then GoTo NextR

        Dim dtVal As Date
        On Error Resume Next
        dtVal = CDate(rawFecha)
        If Err.Number <> 0 Then Err.Clear: On Error GoTo errExit: GoTo NextR
        On Error GoTo errExit

        Dim mVal As Double
        mVal = SafeDbl(arrM(r, iMonto))

        Dim dk As String
        dk = CStr(CLng(CDbl(dtVal)))
        If dDM.exists(dk) Then
            dDM(dk) = dDM(dk) + mVal
        Else
            dDM.Add dk, mVal
        End If
NextR:
    Next r

    If dDM.count = 0 Then
        WriteMontoSeries = 0
        Exit Function
    End If

    ' Ordenar fechas por serial
    Dim sers() As Long
    ReDim sers(dDM.count - 1)
    Dim kk As Long
    kk = 0
    Dim vk As Variant
    For Each vk In dDM.keys
        sers(kk) = CLng(vk)
        kk = kk + 1
    Next vk

    Dim ii As Long, jj As Long, tmp As Long
    For ii = 1 To UBound(sers)
        tmp = sers(ii)
        jj = ii - 1
        Do While jj >= 0 And sers(jj) > tmp
            sers(jj + 1) = sers(jj)
            jj = jj - 1
        Loop
        sers(jj + 1) = tmp
    Next ii

    minDtOut = CDate(sers(0))
    maxDtOut = CDate(sers(UBound(sers)))

    Dim wr As Long
    wr = blockStart
    For ii = 0 To UBound(sers)
        wsh.Cells(wr, 1).Value = CDate(sers(ii))
        wsh.Cells(wr, 2).Value = dDM(CStr(sers(ii)))
        wr = wr + 1
    Next ii
    wsh.Range(wsh.Cells(blockStart, 1), wsh.Cells(wr - 1, 1)).NumberFormat = "dd/mm/yyyy"

    WriteMontoSeries = UBound(sers) + 1
    Exit Function

errExit:
    WriteMontoSeries = 0
End Function

' =========================================================
' WritePromedioSeries
' Escribe exactamente 2 puntos en col D-E del bloque,
' usando los limites del eje (axMinDate, axMaxDate) para que
' la linea de promedio ocupe todo el horizonte visible.
' =========================================================
Private Sub WritePromedioSeries( _
    ByVal wsh As Worksheet, _
    ByVal blockStart As Long, _
    ByVal axMinDate As Date, _
    ByVal axMaxDate As Date, _
    ByVal promedio As Double)

    wsh.Cells(blockStart, 4).Value = axMinDate
    wsh.Cells(blockStart, 5).Value = promedio
    wsh.Cells(blockStart + 1, 4).Value = axMaxDate
    wsh.Cells(blockStart + 1, 5).Value = promedio
    wsh.Cells(blockStart, 4).NumberFormat = "dd/mm/yyyy"
    wsh.Cells(blockStart + 1, 4).NumberFormat = "dd/mm/yyyy"
End Sub

' =========================================================
' CalcAxisBounds
' Calcula los limites del eje X y el MajorUnit para que los
' ticks coincidan con el inicio de cada mes.
' axMin = dia 1 del mes de minDt
' axMax = dia 1 del mes siguiente a maxDt
' majorUnit = (axMax - axMin) / nMonths  (un tick por mes)
' =========================================================
Private Sub CalcAxisBounds( _
    ByVal minDt As Date, _
    ByVal maxDt As Date, _
    ByRef axMin As Double, _
    ByRef axMax As Double, _
    ByRef nMonths As Long, _
    ByRef majorUnit As Double)

    Dim minM As Integer, minY As Integer
    Dim maxM As Integer, maxY As Integer

    minM = Month(minDt): minY = Year(minDt)
    maxM = Month(maxDt): maxY = Year(maxDt)

    maxM = maxM + 1
    If maxM > 12 Then maxM = 1: maxY = maxY + 1

    axMin = CDbl(DateSerial(minY, minM, 1))
    axMax = CDbl(DateSerial(maxY, maxM, 1))

    nMonths = (maxY - minY) * 12 + (maxM - minM)
    If nMonths < 1 Then nMonths = 1

    majorUnit = (axMax - axMin) / CDbl(nMonths)
End Sub

' =========================================================
' CreateScatterChart
' Crea el grafico XY Scatter con linea para un cliente.
'
' Nota sobre el eje X en XY Scatter:
'   En este tipo de grafico el eje de categorias es un eje de valores
'   numerico. NumberFormatLinked = False desvincula el formato de la
'   celda origen y permite aplicar "mmm"".""yy" sobre TickLabels,
'   produciendo etiquetas tipo "Jul.25".
'
' Nota sobre lineas de referencia Y:
'   Excel calcula un MajorUnit automatico con mantisa 1, 2 o 5.
'   Para aumentar la densidad se aplica NiceFloor(autoMU / 2), que
'   siempre devuelve un numero bonito (sin divisiones entre 3).
'   Ejemplos: 10000 -> 5000, 20000 -> 10000, 5000 -> 2000.
' =========================================================
Private Sub CreateScatterChart( _
    ByVal ws As Worksheet, _
    ByVal wsh As Worksheet, _
    ByVal bStart As Long, _
    ByVal bRows As Long, _
    ByVal promedio As Double, _
    ByVal axMin As Double, _
    ByVal axMax As Double, _
    ByVal majorUnit As Double, _
    ByVal cLeft As Double, _
    ByVal cTop As Double, _
    ByVal cName As String, _
    ByVal titleText As String)

    On Error GoTo errExit

    Dim co As ChartObject
    Set co = ws.ChartObjects.Add(cLeft, cTop, CHART_W, CHART_H)
    co.Name = cName

    Dim bEnd As Long
    bEnd = bStart + bRows - 1

    With co.Chart
        .ChartType = xlXYScatterLines

        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop

        ' --- Serie 1: Monto (colores por defecto de Excel) ---
        Dim s1 As series
        Set s1 = .SeriesCollection.NewSeries
        s1.Name = "Monto"
        s1.XValues = wsh.Range(wsh.Cells(bStart, 1), wsh.Cells(bEnd, 1))
        s1.Values = wsh.Range(wsh.Cells(bStart, 2), wsh.Cells(bEnd, 2))
        s1.MarkerStyle = xlMarkerStyleDiamond
        s1.MarkerSize = 5

        ' --- Serie 2: Promedio plano extendido (linea punteada naranja) ---
        Dim s2 As series
        Set s2 = .SeriesCollection.NewSeries
        s2.Name = "Promedio: " & Format(promedio, "#,##0.00")
        s2.XValues = wsh.Range(wsh.Cells(bStart, 4), wsh.Cells(bStart + 1, 4))
        s2.Values = wsh.Range(wsh.Cells(bStart, 5), wsh.Cells(bStart + 1, 5))
        s2.MarkerStyle = xlMarkerStyleNone
        With s2.Format.line
            .DashStyle = msoLineDash
            .ForeColor.RGB = RGB(237, 125, 49)
            .Weight = 1.5
        End With

        ' --- Titulo en negrita ---
        .HasTitle = True
        .ChartTitle.Text = titleText
        With .ChartTitle.Font
            .Size = 18
            .Bold = True
        End With

        ' --- Eje X: ticks exactamente en inicio de cada mes ---
        ' NumberFormatLinked = False desvincula el formato de la celda origen.
        ' El codigo "mmm"".""yy" produce etiquetas tipo "Jul.25".
        Dim axX As Axis
        Set axX = .Axes(xlCategory)
        With axX
            .MinimumScaleIsAuto = False
            .MaximumScaleIsAuto = False
            .MinimumScale = axMin
            .MaximumScale = axMax
            .MajorUnitIsAuto = False
            .majorUnit = majorUnit
            .MajorTickMark = xlOutside
            .MinorTickMark = xlNone
            .TickLabels.NumberFormatLinked = False
            .TickLabels.NumberFormat = "[$-409]mmm"". ""yy"
            .TickLabels.Font.Size = 10
        End With

        ' --- Eje Y: separador de millares, escala automatica inicial ---
        Dim axY As Axis
        Set axY = .Axes(xlValue)
        With axY
            .MajorUnitIsAuto = True
            .TickLabels.NumberFormat = "#,##0"
            .TickLabels.Font.Size = 10
        End With

        ' --- Leyenda ---
        .HasLegend = True
        With .Legend
            .Font.Size = 10
            .Position = xlLegendPositionRight
        End With

        ' --- Borde ---
        .ChartArea.Border.LineStyle = xlContinuous
        .ChartArea.Border.Color = RGB(190, 190, 190)
        .ChartArea.Border.Weight = xlHairline

        On Error Resume Next
        .PlotArea.Left = 58
        On Error GoTo 0
    End With

    Exit Sub
errExit:
End Sub

' =========================================================
' PUBLIC: BuildGraficosAlertasEnHoja
'
' loAL      : ListObject de ALERTAS
' loMAIN    : ListObject de transacciones (CUC, FECHA PROCESO, MONTO)
' loClientes: ListObject de clientes (puede ser Nothing)
' opCode    : "SUS" o "RES"
' =========================================================
Public Sub BuildGraficosAlertasEnHoja( _
    ByVal loAL As ListObject, _
    ByVal loMAIN As ListObject, _
    ByVal loClientes As ListObject, _
    ByVal opCode As String)

    On Error GoTo fin

    If loAL Is Nothing Then Exit Sub
    If loAL.DataBodyRange Is Nothing Then Exit Sub
    If loMAIN Is Nothing Then Exit Sub
    If loMAIN.DataBodyRange Is Nothing Then Exit Sub

    Dim keyCol As String
    If LOHasColumn(loAL, "NUMERO DE DOCUMENTO") Then
        keyCol = "NUMERO DE DOCUMENTO"
    ElseIf LOHasColumn(loAL, "CUC") Then
        keyCol = "CUC"
    Else
        Exit Sub
    End If

    If Not LOHasColumn(loAL, "DESVIACION_MEDIA_%") Then Exit Sub
    If Not LOHasColumn(loAL, "PROMEDIO_MONTOS") Then Exit Sub

    Dim ws As Worksheet: Set ws = loAL.parent
    Dim wb As Workbook:  Set wb = ws.parent

    Dim opLabel As String
    Select Case UCase$(Trim$(opCode))
        Case "SUS": opLabel = "Suscripciones"
        Case "RES": opLabel = "Rescates"
        Case Else:  opLabel = opCode
    End Select

    ' =========================================================
    ' 1. Diccionarios desde loClientes
    '    dTP  : CUC    -> TIPO PERSONA
    '    dTD  : CUC    -> TIPO DE DOCUMENTO
    '    dNTP : NumDoc -> TIPO PERSONA
    '    dNTD : NumDoc -> TIPO DE DOCUMENTO
    '    dNC  : NumDoc -> "CUC1|CUC2|..." (para modo NumDoc)
    ' =========================================================
    Dim dTP  As Object: Set dTP = CreateObject("Scripting.Dictionary")
    Dim dTD  As Object: Set dTD = CreateObject("Scripting.Dictionary")
    Dim dNTP As Object: Set dNTP = CreateObject("Scripting.Dictionary")
    Dim dNTD As Object: Set dNTD = CreateObject("Scripting.Dictionary")
    Dim dNC  As Object: Set dNC = CreateObject("Scripting.Dictionary")

    If Not loClientes Is Nothing Then
        If Not loClientes.DataBodyRange Is Nothing Then
            Dim iC_cuc As Long, iC_tp As Long, iC_td As Long, iC_nd As Long
            iC_cuc = GetColIdx(loClientes, "CUC")
            iC_tp = GetColIdx(loClientes, "TIPO PERSONA")
            iC_td = GetColIdx(loClientes, "TIPO DE DOCUMENTO")
            iC_nd = GetColIdx(loClientes, "NUMERO DE DOCUMENTO")

            Dim arrCL As Variant
            arrCL = loClientes.DataBodyRange.Value

            Dim ri As Long
            For ri = 1 To UBound(arrCL, 1)
                Dim sCUC_c As String, sTP_c As String
                Dim sTD_c  As String, sND_c  As String
                sCUC_c = IIf(iC_cuc > 0, Trim$(CStr(arrCL(ri, iC_cuc))), "")
                sTP_c = IIf(iC_tp > 0, Trim$(UCase$(CStr(arrCL(ri, iC_tp)))), "")
                sTD_c = IIf(iC_td > 0, Trim$(UCase$(CStr(arrCL(ri, iC_td)))), "")
                sND_c = IIf(iC_nd > 0, Trim$(CStr(arrCL(ri, iC_nd))), "")

                If sCUC_c <> "" And Not dTP.exists(sCUC_c) Then
                    dTP.Add sCUC_c, sTP_c
                    dTD.Add sCUC_c, sTD_c
                End If
                If sND_c <> "" Then
                    If Not dNTP.exists(sND_c) Then
                        dNTP.Add sND_c, sTP_c
                        dNTD.Add sND_c, sTD_c
                    End If
                    If dNC.exists(sND_c) Then
                        If sCUC_c <> "" Then dNC(sND_c) = dNC(sND_c) & "|" & sCUC_c
                    Else
                        dNC.Add sND_c, sCUC_c
                    End If
                End If
            Next ri
        End If
    End If

    ' =========================================================
    ' 2. Extraer candidatos NAT/MAN y JUR desde loAL
    ' =========================================================
    Dim iKey As Long, iDv As Long, iPm As Long
    iKey = GetColIdx(loAL, keyCol)
    iDv = GetColIdx(loAL, "DESVIACION_MEDIA_%")
    iPm = GetColIdx(loAL, "PROMEDIO_MONTOS")

    Const BUF As Long = 256
    Dim nK(BUF) As String, nDv(BUF) As Double, nPm(BUF) As Double
    Dim nTP(BUF) As String, nTD(BUF) As String
    Dim jK(BUF) As String, jDv(BUF) As Double, jPm(BUF) As Double
    Dim jTP(BUF) As String, jTD(BUF) As String
    Dim nCnt As Long, jCnt As Long
    nCnt = 0: jCnt = 0

    Dim arrAL As Variant
    arrAL = loAL.DataBodyRange.Value

    Dim ai As Long
    For ai = 1 To UBound(arrAL, 1)
        Dim sKeyV As String, dDvV As Double, dPmV As Double
        sKeyV = Trim$(CStr(arrAL(ai, iKey)))
        dDvV = SafeDbl(arrAL(ai, iDv))
        dPmV = SafeDbl(arrAL(ai, iPm))

        Dim sTPv As String, sTDv As String
        sTPv = "": sTDv = ""
        If keyCol = "CUC" Then
            If dTP.exists(sKeyV) Then sTPv = dTP(sKeyV)
            If dTD.exists(sKeyV) Then sTDv = dTD(sKeyV)
        Else
            If dNTP.exists(sKeyV) Then sTPv = dNTP(sKeyV)
            If dNTD.exists(sKeyV) Then sTDv = dNTD(sKeyV)
        End If

        If sTPv = "JUR" Then
            If jCnt < BUF Then
                jK(jCnt) = sKeyV: jDv(jCnt) = dDvV: jPm(jCnt) = dPmV
                jTP(jCnt) = sTPv: jTD(jCnt) = sTDv
                jCnt = jCnt + 1
            End If
        Else
            If nCnt < BUF Then
                nK(nCnt) = sKeyV: nDv(nCnt) = dDvV: nPm(nCnt) = dPmV
                nTP(nCnt) = sTPv: nTD(nCnt) = sTDv
                nCnt = nCnt + 1
            End If
        End If
    Next ai

    SortDesc5 nCnt, nK, nDv, nPm, nTP, nTD
    SortDesc5 jCnt, jK, jDv, jPm, jTP, jTD

    If nCnt > MAX_PER_COL Then nCnt = MAX_PER_COL
    If jCnt > MAX_PER_COL Then jCnt = MAX_PER_COL

    ' =========================================================
    ' 3. Array de loMAIN para iteracion rapida
    ' =========================================================
    Dim iM_cuc As Long, iM_fch As Long, iM_mto As Long
    iM_cuc = GetColIdx(loMAIN, "CUC")
    iM_fch = GetColIdx(loMAIN, "FECHA PROCESO")
    iM_mto = GetColIdx(loMAIN, "MONTO SOLES")

    If iM_cuc = 0 Or iM_fch = 0 Or iM_mto = 0 Then GoTo fin

    Dim arrM As Variant
    arrM = loMAIN.DataBodyRange.Value

    ' =========================================================
    ' 4. Hoja helper y limpieza de graficos anteriores
    ' =========================================================
    Dim wsh As Worksheet
    Set wsh = EnsureHelperSheet(wb)

    DeleteChartsByPrefix ws, "GF_AL_"

    ' =========================================================
    ' 5. Posicion inicial de graficos a la derecha de loAL
    ' =========================================================
    Dim lastALCol As Long
    lastALCol = loAL.Range.Column + loAL.Range.Columns.count

    Dim chartLeft1   As Double
    Dim chartLeft2   As Double
    Dim chartTopBase As Double
    chartLeft1 = ws.Cells(1, lastALCol + 1).Left
    chartLeft2 = chartLeft1 + CHART_W + CHART_GAP_C
    chartTopBase = ws.Cells(loAL.Range.Row, 1).top + CHART_TOP_MGN

    Dim cliIdx As Long
    cliIdx = 0

    ' =========================================================
    ' 6. Graficos NAT / MAN (columna izquierda)
    ' =========================================================
    Dim ci As Long
    For ci = 0 To nCnt - 1

        Dim cucListN As String
        cucListN = ""
        If keyCol = "CUC" Then
            cucListN = nK(ci)
        Else
            If dNC.exists(nK(ci)) Then cucListN = dNC(nK(ci))
        End If
        If cucListN = "" Then GoTo NextNAT

        Dim bStartN As Long
        bStartN = 1 + cliIdx * CLI_BLOCK

        Dim minDtN As Date, maxDtN As Date
        Dim rowsN As Long
        rowsN = WriteMontoSeries(wsh, bStartN, cucListN, _
                                 iM_cuc, iM_fch, iM_mto, arrM, _
                                 minDtN, maxDtN)
        If rowsN = 0 Then GoTo NextNAT

        Dim axMinN As Double, axMaxN As Double
        Dim nMthN As Long, mjUN As Double
        CalcAxisBounds minDtN, maxDtN, axMinN, axMaxN, nMthN, mjUN

        WritePromedioSeries wsh, bStartN, CDate(axMinN), CDate(axMaxN), nPm(ci)

        Dim tpLblN As String, tdLblN As String, titleN As String
        tpLblN = IIf(nTP(ci) <> "", nTP(ci), "NAT")
        If keyCol = "CUC" Then
            tdLblN = "CUC"
        ElseIf nTD(ci) <> "" Then
            tdLblN = nTD(ci)
        Else
            tdLblN = "DOC"
        End If
        titleN = "[" & tpLblN & "] " & opLabel & " " & tdLblN & ": " & nK(ci) & _
                 " | Desviacion: " & Format(nDv(ci), "0.00") & "%"

        Dim cTopN As Double
        cTopN = chartTopBase + CDbl(ci) * (CHART_H + CHART_GAP_H)

        CreateScatterChart ws, wsh, bStartN, rowsN, nPm(ci), _
                           axMinN, axMaxN, mjUN, _
                           chartLeft1, cTopN, "GF_AL_N" & Format(ci + 1, "00"), titleN

        cliIdx = cliIdx + 1
NextNAT:
    Next ci

    ' =========================================================
    ' 7. Graficos JUR (columna derecha)
    ' =========================================================
    Dim cj As Long
    For cj = 0 To jCnt - 1

        Dim cucListJ As String
        cucListJ = ""
        If keyCol = "CUC" Then
            cucListJ = jK(cj)
        Else
            If dNC.exists(jK(cj)) Then cucListJ = dNC(jK(cj))
        End If
        If cucListJ = "" Then GoTo NextJUR

        Dim bStartJ As Long
        bStartJ = 1 + cliIdx * CLI_BLOCK

        Dim minDtJ As Date, maxDtJ As Date
        Dim rowsJ As Long
        rowsJ = WriteMontoSeries(wsh, bStartJ, cucListJ, _
                                 iM_cuc, iM_fch, iM_mto, arrM, _
                                 minDtJ, maxDtJ)
        If rowsJ = 0 Then GoTo NextJUR

        Dim axMinJ As Double, axMaxJ As Double
        Dim nMthJ As Long, mjUJ As Double
        CalcAxisBounds minDtJ, maxDtJ, axMinJ, axMaxJ, nMthJ, mjUJ

        WritePromedioSeries wsh, bStartJ, CDate(axMinJ), CDate(axMaxJ), jPm(cj)

        Dim tpLblJ As String, tdLblJ As String, titleJ As String
        tpLblJ = IIf(jTP(cj) <> "", jTP(cj), "JUR")
        If keyCol = "CUC" Then
            tdLblJ = "CUC"
        ElseIf jTD(cj) <> "" Then
            tdLblJ = jTD(cj)
        Else
            tdLblJ = "DOC"
        End If
        titleJ = "[" & tpLblJ & "] " & opLabel & " " & tdLblJ & ": " & jK(cj) & _
                 " | Desviacion: " & Format(jDv(cj), "0.00") & "%"

        Dim cTopJ As Double
        cTopJ = chartTopBase + CDbl(cj) * (CHART_H + CHART_GAP_H)

        CreateScatterChart ws, wsh, bStartJ, rowsJ, jPm(cj), _
                           axMinJ, axMaxJ, mjUJ, _
                           chartLeft2, cTopJ, "GF_AL_J" & Format(cj + 1, "00"), titleJ

        cliIdx = cliIdx + 1
NextJUR:
    Next cj

fin:
End Sub


