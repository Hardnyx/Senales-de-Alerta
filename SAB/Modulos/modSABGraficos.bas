Option Explicit

'=========================================================
' modSABGraficos
' Genera hasta 10 graficos de serie temporal (XY Scatter con linea)
' en la hoja de ALERTAS SAB - Movimiento de Caja:
'   Columna unica: top 10 DESVIACION_MEDIA_% de la tabla de alertas
'
' Cada grafico:
'   - Serie solida   : monto diario agregado por Cuenta y fecha (de loMAIN)
'   - Serie punteada : promedio plano extendido al rango completo del eje
'   - Eje X proporcional a fechas reales, etiquetas "Mmm.AA" (ej: Jul.25)
'   - Ticks en el inicio de cada mes dentro del rango
'   - Titulo con operacion, cuenta, clase, moneda y desviacion
'
' Requiere hoja auxiliar oculta _GF_SAB_HELPER para datos de las series.
'
' Firma publica:
'   BuildGraficosAlertasEnHoja(loAL, loMAIN, which, suf)
'   which = "DEP" o "RET"
'   suf   = sufijo de periodo (ej: "ENE_JUN_2025")
'=========================================================

Private Const HELPER_SH     As String = "_GF_SAB_HELPER"
Private Const CHART_W       As Double = 510
Private Const CHART_H       As Double = 255
Private Const CHART_GAP_H   As Double = 10
Private Const CHART_TOP_MGN As Double = 30
Private Const MAX_CHARTS    As Long   = 10
Private Const CLI_BLOCK     As Long   = 400

'=========================================================
' Texto / columnas
'=========================================================
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

'=========================================================
' Utilidades
'=========================================================
Private Function SafeDbl(ByVal v As Variant) As Double
    On Error Resume Next
    SafeDbl = CDbl(v)
    On Error GoTo 0
End Function

Private Function NiceFloor(ByVal v As Double) As Double
    If v <= 0 Then NiceFloor = 1: Exit Function
    Dim mag As Double: mag = 10 ^ Int(Log(v) / Log(10))
    Dim m As Double:   m   = v / mag
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
    Dim cnt As Long, i As Long
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
        Set wsh = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsh.Name = HELPER_SH
    Else
        wsh.Cells.Clear
    End If
    wsh.Visible = xlSheetVeryHidden
    Set EnsureHelperSheet = wsh
End Function

' Ordena n elementos descendentemente segun dv().
Private Sub SortDescN(ByVal n As Long, _
    k() As String, dv() As Double, pm() As Double, _
    cl() As String, mo() As String)
    Dim i As Long, j As Long
    Dim ts As String, tf As Double
    For i = 0 To n - 2
        For j = 0 To n - i - 2
            If dv(j) < dv(j + 1) Then
                ts = k(j):  k(j)  = k(j + 1):  k(j + 1)  = ts
                tf = dv(j): dv(j) = dv(j + 1): dv(j + 1) = tf
                tf = pm(j): pm(j) = pm(j + 1): pm(j + 1) = tf
                ts = cl(j): cl(j) = cl(j + 1): cl(j + 1) = ts
                ts = mo(j): mo(j) = mo(j + 1): mo(j + 1) = ts
            End If
        Next j
    Next i
End Sub

'=========================================================
' WriteMontoSeries
' Agrega montos diarios para una Cuenta en wsh cols A-B desde blockStart.
'=========================================================
Private Function WriteMontoSeries( _
    ByVal wsh As Worksheet, _
    ByVal blockStart As Long, _
    ByVal cuenta As String, _
    ByVal iCuenta As Long, _
    ByVal iFecha As Long, _
    ByVal iMonto As Long, _
    ByRef arrM As Variant, _
    ByRef minDtOut As Date, _
    ByRef maxDtOut As Date) As Long

    On Error GoTo errExit

    Dim dDM As Object
    Set dDM = CreateObject("Scripting.Dictionary")

    Dim r As Long
    Dim nRows As Long: nRows = UBound(arrM, 1)

    For r = 1 To nRows
        If Trim$(CStr(arrM(r, iCuenta))) <> cuenta Then GoTo NextR

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
        If dDM.Exists(dk) Then
            dDM(dk) = dDM(dk) + mVal
        Else
            dDM.Add dk, mVal
        End If
NextR:
    Next r

    If dDM.Count = 0 Then WriteMontoSeries = 0: Exit Function

    Dim sers() As Long
    ReDim sers(dDM.Count - 1)
    Dim kk As Long: kk = 0
    Dim vk As Variant
    For Each vk In dDM.Keys
        sers(kk) = CLng(vk): kk = kk + 1
    Next vk

    Dim ii As Long, jj As Long, tmp As Long
    For ii = 1 To UBound(sers)
        tmp = sers(ii): jj = ii - 1
        Do While jj >= 0 And sers(jj) > tmp
            sers(jj + 1) = sers(jj): jj = jj - 1
        Loop
        sers(jj + 1) = tmp
    Next ii

    minDtOut = CDate(sers(0))
    maxDtOut = CDate(sers(UBound(sers)))

    Dim wr As Long: wr = blockStart
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

'=========================================================
' WritePromedioSeries
'=========================================================
Private Sub WritePromedioSeries( _
    ByVal wsh As Worksheet, _
    ByVal blockStart As Long, _
    ByVal axMinDate As Date, _
    ByVal axMaxDate As Date, _
    ByVal promedio As Double)

    wsh.Cells(blockStart,     4).Value        = axMinDate
    wsh.Cells(blockStart + 1, 4).Value        = axMaxDate
    wsh.Cells(blockStart,     5).Value        = promedio
    wsh.Cells(blockStart + 1, 5).Value        = promedio
    wsh.Cells(blockStart,     4).NumberFormat = "dd/mm/yyyy"
    wsh.Cells(blockStart + 1, 4).NumberFormat = "dd/mm/yyyy"
End Sub

'=========================================================
' CalcAxisBounds
'=========================================================
Private Sub CalcAxisBounds( _
    ByVal minDt As Date, _
    ByVal maxDt As Date, _
    ByRef axMin As Double, _
    ByRef axMax As Double, _
    ByRef nMonths As Long, _
    ByRef majorUnit As Double)

    Dim minM As Integer: minM = Month(minDt)
    Dim minY As Integer: minY = Year(minDt)
    Dim maxM As Integer: maxM = Month(maxDt)
    Dim maxY As Integer: maxY = Year(maxDt)

    maxM = maxM + 1
    If maxM > 12 Then maxM = 1: maxY = maxY + 1

    axMin = CDbl(DateSerial(minY, minM, 1))
    axMax = CDbl(DateSerial(maxY, maxM, 1))

    nMonths = (maxY - minY) * 12 + (maxM - minM)
    If nMonths < 1 Then nMonths = 1

    majorUnit = (axMax - axMin) / CDbl(nMonths)
End Sub

'=========================================================
' CreateScatterChart
'=========================================================
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

    Dim bEnd As Long: bEnd = bStart + bRows - 1

    With co.Chart
        .ChartType = xlXYScatterLines

        Do While .SeriesCollection.Count > 0
            .SeriesCollection(1).Delete
        Loop

        Dim s1 As Series
        Set s1        = .SeriesCollection.NewSeries
        s1.Name       = "Monto"
        s1.XValues    = wsh.Range(wsh.Cells(bStart, 1), wsh.Cells(bEnd, 1))
        s1.Values     = wsh.Range(wsh.Cells(bStart, 2), wsh.Cells(bEnd, 2))
        s1.MarkerStyle = xlMarkerStyleDiamond
        s1.MarkerSize  = 5

        Dim s2 As Series
        Set s2     = .SeriesCollection.NewSeries
        s2.Name    = "Promedio: " & Format(promedio, "#,##0.00")
        s2.XValues = wsh.Range(wsh.Cells(bStart, 4), wsh.Cells(bStart + 1, 4))
        s2.Values  = wsh.Range(wsh.Cells(bStart, 5), wsh.Cells(bStart + 1, 5))
        s2.MarkerStyle = xlMarkerStyleNone
        With s2.Format.Line
            .DashStyle     = msoLineDash
            .ForeColor.RGB = RGB(237, 125, 49)
            .Weight        = 1.5
        End With

        .HasTitle = True
        .ChartTitle.Text = titleText
        With .ChartTitle.Font
            .Size = 18
            .Bold = True
        End With

        Dim axX As Axis
        Set axX = .Axes(xlCategory)
        With axX
            .MinimumScaleIsAuto            = False
            .MaximumScaleIsAuto            = False
            .MinimumScale                  = axMin
            .MaximumScale                  = axMax
            .MajorUnitIsAuto               = False
            .MajorUnit                     = majorUnit
            .MajorTickMark                 = xlOutside
            .MinorTickMark                 = xlNone
            .TickLabels.NumberFormatLinked = False
            .TickLabels.NumberFormat       = "[$-409]mmm"". ""yy"
            .TickLabels.Font.Size          = 10
        End With

        Dim axY As Axis
        Set axY = .Axes(xlValue)
        With axY
            .MajorUnitIsAuto         = True
            .TickLabels.NumberFormat = "#,##0"
            .TickLabels.Font.Size    = 10
        End With

        .HasLegend = True
        With .Legend
            .Font.Size = 10
            .Position  = xlLegendPositionRight
        End With

        .ChartArea.Border.LineStyle = xlContinuous
        .ChartArea.Border.Color     = RGB(190, 190, 190)
        .ChartArea.Border.Weight    = xlHairline

        On Error Resume Next
        .PlotArea.Left = 58
        On Error GoTo 0
    End With

    Exit Sub
errExit:
End Sub

'=========================================================
' PUBLIC: BuildGraficosAlertasEnHoja
'
' loAL  : ListObject de ALERTAS (SAB_MC_ALERTAS_DEP o _RET)
' loMAIN: ListObject de transacciones (SAB_MC_MAIN)
' which : "DEP" o "RET"
' suf   : sufijo de periodo para nombres de graficos
'=========================================================
Public Sub BuildGraficosAlertasEnHoja( _
    ByVal loAL   As ListObject, _
    ByVal loMAIN As ListObject, _
    ByVal which  As String, _
    ByVal suf    As String)

    On Error GoTo fin

    If loAL Is Nothing Then Exit Sub
    If loAL.DataBodyRange Is Nothing Then Exit Sub
    If loMAIN Is Nothing Then Exit Sub
    If loMAIN.DataBodyRange Is Nothing Then Exit Sub

    Dim op As String: op = UCase$(Trim$(which))
    If op <> "DEP" And op <> "RET" Then op = "DEP"

    If Not LOHasColumn(loAL, "DESVIACION_MEDIA_%") Then Exit Sub
    If Not LOHasColumn(loAL, "PROMEDIO_MONTOS") Then Exit Sub
    If Not LOHasColumn(loAL, "Cuenta") Then Exit Sub

    Dim ws As Worksheet: Set ws = loAL.Parent
    Dim wb As Workbook:  Set wb = ws.Parent

    Dim opLabel   As String: opLabel   = IIf(op = "DEP", "Deposito", "Retiro")
    Dim chartPref As String: chartPref = "GF_SAB_" & op & "_"

    ' =========================================================
    ' 1. Extraer top MAX_CHARTS desde loAL
    ' =========================================================
    Dim iKey As Long, iDv As Long, iPm As Long, iCl As Long, iMo As Long
    iKey = GetColIdx(loAL, "Cuenta")
    iDv  = GetColIdx(loAL, "DESVIACION_MEDIA_%")
    iPm  = GetColIdx(loAL, "PROMEDIO_MONTOS")
    iCl  = GetColIdx(loAL, "CLASE")
    iMo  = GetColIdx(loAL, "MONEDA")

    Const BUF As Long = 256
    Dim aK(BUF) As String, aDv(BUF) As Double, aPm(BUF) As Double
    Dim aCl(BUF) As String, aMo(BUF) As String
    Dim nCnt As Long: nCnt = 0

    Dim arrAL As Variant
    arrAL = loAL.DataBodyRange.Value

    Dim ai As Long
    For ai = 1 To UBound(arrAL, 1)
        If nCnt >= BUF Then Exit For
        aK(nCnt)  = Trim$(CStr(arrAL(ai, iKey)))
        aDv(nCnt) = SafeDbl(arrAL(ai, iDv))
        aPm(nCnt) = SafeDbl(arrAL(ai, iPm))
        aCl(nCnt) = IIf(iCl > 0, Trim$(CStr(arrAL(ai, iCl))), "")
        aMo(nCnt) = IIf(iMo > 0, Trim$(CStr(arrAL(ai, iMo))), "")
        nCnt = nCnt + 1
    Next ai

    SortDescN nCnt, aK, aDv, aPm, aCl, aMo
    If nCnt > MAX_CHARTS Then nCnt = MAX_CHARTS

    ' =========================================================
    ' 2. Columnas en loMAIN
    ' =========================================================
    Dim iM_cta As Long, iM_fch As Long, iM_mto As Long
    iM_cta = GetColIdx(loMAIN, "Cuenta")
    iM_fch = GetColIdx(loMAIN, "Fecha")
    iM_mto = GetColIdx(loMAIN, IIf(op = "DEP", "Dep" & Chr(243) & "sito", "Retiro"))

    If iM_cta = 0 Or iM_fch = 0 Or iM_mto = 0 Then GoTo fin

    Dim arrM As Variant
    arrM = loMAIN.DataBodyRange.Value

    ' =========================================================
    ' 3. Hoja helper y limpieza de graficos anteriores
    ' =========================================================
    Dim wsh As Worksheet
    Set wsh = EnsureHelperSheet(wb)

    DeleteChartsByPrefix ws, chartPref

    ' =========================================================
    ' 4. Posicion a la derecha de loAL
    ' =========================================================
    Dim lastALCol As Long
    lastALCol = loAL.Range.Column + loAL.Range.Columns.Count

    Dim chartLeft    As Double
    Dim chartTopBase As Double
    chartLeft    = ws.Cells(1, lastALCol + 1).Left
    chartTopBase = ws.Cells(loAL.Range.Row, 1).Top + CHART_TOP_MGN

    ' =========================================================
    ' 5. Generar graficos
    ' =========================================================
    Dim ci As Long
    For ci = 0 To nCnt - 1

        If aK(ci) = "" Then GoTo NextChart

        Dim bStart As Long
        bStart = 1 + ci * CLI_BLOCK

        Dim minDt As Date, maxDt As Date
        Dim nRows As Long
        nRows = WriteMontoSeries(wsh, bStart, aK(ci), _
                                 iM_cta, iM_fch, iM_mto, arrM, _
                                 minDt, maxDt)
        If nRows = 0 Then GoTo NextChart

        Dim axMin As Double, axMax As Double
        Dim nMth As Long, mjU As Double
        CalcAxisBounds minDt, maxDt, axMin, axMax, nMth, mjU

        WritePromedioSeries wsh, bStart, CDate(axMin), CDate(axMax), aPm(ci)

        Dim titleStr As String
        titleStr = "[" & opLabel & "] Cuenta: " & aK(ci)
        If aCl(ci) <> "" Then titleStr = titleStr & " | Clase: " & aCl(ci)
        If aMo(ci) <> "" Then titleStr = titleStr & " | " & aMo(ci)
        titleStr = titleStr & " | Desviacion: " & Format(aDv(ci), "0.00") & "%"

        Dim cTop  As Double: cTop  = chartTopBase + CDbl(ci) * (CHART_H + CHART_GAP_H)
        Dim cName As String: cName = chartPref & Format(ci + 1, "00")

        CreateScatterChart ws, wsh, bStart, nRows, aPm(ci), _
                           axMin, axMax, mjU, _
                           chartLeft, cTop, cName, titleStr

NextChart:
    Next ci

fin:
End Sub

'=========================================================
' PUBLIC: BuildGraficosCMEnHoja
'
' loAL  : ListObject de ALERTAS CM (SAB_CM_ALERTAS_COM o _VEN)
' loMAIN: ListObject de detalle (SAB_CM_MAIN)
' which : "COM" o "VEN"
'
' Estructura CM:
'   loAL  columnas: Documento, TIPO_PERSONA, DESVIACION_MEDIA_%, PROMEDIO_MONTOS, NIVEL_RIESGO
'   loMAIN columnas: Documento, Fecha, Moneda Ori, Total Neto / Monto Ori / Monto Des / Gan/Per PEN, Tipo Persona
'
' Biparticion NAT (izquierda) / JUR (derecha), top 5 cada una,
' igual que modFondosGraficos.
'=========================================================
Public Sub BuildGraficosCMEnHoja( _
    ByVal loAL   As ListObject, _
    ByVal loMAIN As ListObject, _
    ByVal which  As String)

    On Error GoTo fin

    If loAL Is Nothing Then Exit Sub
    If loAL.DataBodyRange Is Nothing Then Exit Sub
    If loMAIN Is Nothing Then Exit Sub
    If loMAIN.DataBodyRange Is Nothing Then Exit Sub

    Dim op As String: op = UCase$(Trim$(which))
    If op <> "COM" And op <> "VEN" Then op = "COM"

    If Not LOHasColumn(loAL, "DESVIACION_MEDIA_%") Then Exit Sub
    If Not LOHasColumn(loAL, "PROMEDIO_MONTOS") Then Exit Sub
    If Not LOHasColumn(loAL, "Documento") Then Exit Sub

    Dim ws As Worksheet: Set ws = loAL.Parent
    Dim wb As Workbook:  Set wb = ws.Parent

    Dim opLabel   As String: opLabel   = IIf(op = "COM", "Compra", "Venta")
    Dim chartPref As String: chartPref = "GF_CM_" & op & "_"

    ' =========================================================
    ' 1. Extraer candidatos NAT y JUR desde loAL
    ' =========================================================
    Dim iKey As Long, iDv As Long, iPm As Long, iTP As Long
    iKey = GetColIdx(loAL, "Documento")
    iDv  = GetColIdx(loAL, "DESVIACION_MEDIA_%")
    iPm  = GetColIdx(loAL, "PROMEDIO_MONTOS")
    iTP  = GetColIdx(loAL, "TIPO_PERSONA")

    Const BUF As Long = 256
    Dim nK(BUF) As String, nDv(BUF) As Double, nPm(BUF) As Double
    Dim nTP_a(BUF) As String, nPH(BUF) As String
    Dim jK(BUF) As String, jDv(BUF) As Double, jPm(BUF) As Double
    Dim jTP_a(BUF) As String, jPH(BUF) As String
    Dim nCnt As Long: nCnt = 0
    Dim jCnt As Long: jCnt = 0

    Dim arrAL As Variant
    arrAL = loAL.DataBodyRange.Value

    Dim ai As Long
    For ai = 1 To UBound(arrAL, 1)
        Dim sKey As String: sKey = Trim$(CStr(arrAL(ai, iKey)))
        Dim sDv  As Double: sDv  = SafeDbl(arrAL(ai, iDv))
        Dim sPm  As Double: sPm  = SafeDbl(arrAL(ai, iPm))
        Dim sTP  As String
        sTP = IIf(iTP > 0, UCase$(Trim$(CStr(arrAL(ai, iTP)))), "")

        If InStr(sTP, "JUR") > 0 Or sTP = "PJ" Then
            If jCnt < BUF Then
                jK(jCnt)    = sKey: jDv(jCnt)   = sDv: jPm(jCnt)  = sPm
                jTP_a(jCnt) = sTP:  jPH(jCnt)   = ""
                jCnt = jCnt + 1
            End If
        Else
            If nCnt < BUF Then
                nK(nCnt)    = sKey: nDv(nCnt)   = sDv: nPm(nCnt)  = sPm
                nTP_a(nCnt) = sTP:  nPH(nCnt)   = ""
                nCnt = nCnt + 1
            End If
        End If
    Next ai

    SortDescN nCnt, nK, nDv, nPm, nTP_a, nPH
    SortDescN jCnt, jK, jDv, jPm, jTP_a, jPH
    If nCnt > MAX_CHARTS \ 2 Then nCnt = MAX_CHARTS \ 2
    If jCnt > MAX_CHARTS \ 2 Then jCnt = MAX_CHARTS \ 2

    ' =========================================================
    ' 2. Columnas en loMAIN (deteccion por nombre canonico)
    ' =========================================================
    Dim iM_doc As Long, iM_fch As Long
    Dim iM_tp  As Long, iM_mto As Long
    iM_doc = GetColIdx(loMAIN, "Documento")
    iM_fch = GetColIdx(loMAIN, "Fecha")
    iM_tp  = GetColIdx(loMAIN, "Tipo Persona")

    ' Monto: prioridad Total Neto > Monto Des > Monto Ori > Gan/Per PEN > Gan/Per
    Dim mntCandidatos(4) As String
    mntCandidatos(0) = "Total Neto"
    mntCandidatos(1) = "Monto Des"
    mntCandidatos(2) = "Monto Ori"
    mntCandidatos(3) = "Gan/Per PEN"
    mntCandidatos(4) = "Gan/Per"
    Dim mc As Long
    For mc = 0 To 4
        iM_mto = GetColIdx(loMAIN, mntCandidatos(mc))
        If iM_mto > 0 Then Exit For
    Next mc

    ' Moneda (para filtrar COM=USD, VEN=PEN)
    Dim iM_mon As Long
    iM_mon = GetColIdx(loMAIN, "Moneda Ori")

    If iM_doc = 0 Or iM_fch = 0 Or iM_mto = 0 Then GoTo fin

    Dim arrM As Variant
    arrM = loMAIN.DataBodyRange.Value

    ' =========================================================
    ' 3. Hoja helper y limpieza de graficos anteriores
    ' =========================================================
    Dim wsh As Worksheet
    Set wsh = EnsureHelperSheet(wb)

    DeleteChartsByPrefix ws, chartPref

    ' =========================================================
    ' 4. Posicion a la derecha de loAL, dos columnas NAT/JUR
    ' =========================================================
    Dim lastALCol As Long
    lastALCol = loAL.Range.Column + loAL.Range.Columns.Count

    Dim chartLeft1   As Double: chartLeft1   = ws.Cells(1, lastALCol + 1).Left
    Dim chartLeft2   As Double: chartLeft2   = chartLeft1 + CHART_W + 14
    Dim chartTopBase As Double: chartTopBase = ws.Cells(loAL.Range.Row, 1).Top + CHART_TOP_MGN

    Dim cliIdx As Long: cliIdx = 0

    ' =========================================================
    ' 5. Graficos NAT (columna izquierda)
    ' =========================================================
    Dim ci As Long
    For ci = 0 To nCnt - 1
        If nK(ci) = "" Then GoTo NextNAT_CM

        Dim bStartN As Long: bStartN = 1 + cliIdx * CLI_BLOCK

        Dim minDtN As Date, maxDtN As Date, rowsN As Long
        rowsN = WriteMontoSeriesCM(wsh, bStartN, nK(ci), "NATURAL", op, _
                                   iM_doc, iM_fch, iM_mto, iM_tp, iM_mon, arrM, _
                                   minDtN, maxDtN)
        If rowsN = 0 Then GoTo NextNAT_CM

        Dim axMinN As Double, axMaxN As Double, nMthN As Long, mjUN As Double
        CalcAxisBounds minDtN, maxDtN, axMinN, axMaxN, nMthN, mjUN
        WritePromedioSeries wsh, bStartN, CDate(axMinN), CDate(axMaxN), nPm(ci)

        Dim docLblN As String: docLblN = IIf(InStr(nTP_a(ci), "JUR") > 0 Or nTP_a(ci) = "PJ", "RUC", "DNI")
        Dim titleN As String
        titleN = "[NAT] " & opLabel & " " & docLblN & ": " & nK(ci) & _
                 " | Desviacion: " & Format(nDv(ci), "0.00") & "%"

        Dim cTopN As Double: cTopN = chartTopBase + CDbl(ci) * (CHART_H + CHART_GAP_H)
        CreateScatterChart ws, wsh, bStartN, rowsN, nPm(ci), _
                           axMinN, axMaxN, mjUN, _
                           chartLeft1, cTopN, chartPref & "N" & Format(ci + 1, "00"), titleN
        cliIdx = cliIdx + 1
NextNAT_CM:
    Next ci

    ' =========================================================
    ' 6. Graficos JUR (columna derecha)
    ' =========================================================
    Dim cj As Long
    For cj = 0 To jCnt - 1
        If jK(cj) = "" Then GoTo NextJUR_CM

        Dim bStartJ As Long: bStartJ = 1 + cliIdx * CLI_BLOCK

        Dim minDtJ As Date, maxDtJ As Date, rowsJ As Long
        rowsJ = WriteMontoSeriesCM(wsh, bStartJ, jK(cj), "JURIDICA", op, _
                                   iM_doc, iM_fch, iM_mto, iM_tp, iM_mon, arrM, _
                                   minDtJ, maxDtJ)
        If rowsJ = 0 Then GoTo NextJUR_CM

        Dim axMinJ As Double, axMaxJ As Double, nMthJ As Long, mjUJ As Double
        CalcAxisBounds minDtJ, maxDtJ, axMinJ, axMaxJ, nMthJ, mjUJ
        WritePromedioSeries wsh, bStartJ, CDate(axMinJ), CDate(axMaxJ), jPm(cj)

        Dim titleJ As String
        titleJ = "[JUR] " & opLabel & " RUC: " & jK(cj) & _
                 " | Desviacion: " & Format(jDv(cj), "0.00") & "%"

        Dim cTopJ As Double: cTopJ = chartTopBase + CDbl(cj) * (CHART_H + CHART_GAP_H)
        CreateScatterChart ws, wsh, bStartJ, rowsJ, jPm(cj), _
                           axMinJ, axMaxJ, mjUJ, _
                           chartLeft2, cTopJ, chartPref & "J" & Format(cj + 1, "00"), titleJ
        cliIdx = cliIdx + 1
NextJUR_CM:
    Next cj

fin:
End Sub

'=========================================================
' WriteMontoSeriesCM
' Filtra por Documento + TIPO_PERSONA + moneda (COM=USD, VEN=PEN)
' y agrega montos diarios en wsh cols A-B desde blockStart.
'=========================================================
Private Function WriteMontoSeriesCM( _
    ByVal wsh As Worksheet, _
    ByVal blockStart As Long, _
    ByVal docKey As String, _
    ByVal persona As String, _
    ByVal op As String, _
    ByVal iDoc As Long, _
    ByVal iFecha As Long, _
    ByVal iMonto As Long, _
    ByVal iTP As Long, _
    ByVal iMon As Long, _
    ByRef arrM As Variant, _
    ByRef minDtOut As Date, _
    ByRef maxDtOut As Date) As Long

    On Error GoTo errExit

    Dim wantUSD As Boolean: wantUSD = (op = "COM")
    Dim dDM As Object: Set dDM = CreateObject("Scripting.Dictionary")

    Dim r As Long, nRows As Long: nRows = UBound(arrM, 1)

    ' Primer pase: con filtro de moneda
    For r = 1 To nRows
        If NormDoc(CStr(arrM(r, iDoc))) <> docKey Then GoTo NextR_CM

        If iTP > 0 Then
            If NormPersona(CStr(arrM(r, iTP))) <> UCase$(persona) Then GoTo NextR_CM
        End If

        If iMon > 0 Then
            Dim monStr As String: monStr = Trim$(CStr(arrM(r, iMon)))
            Dim okMon As Boolean
            If wantUSD Then
                okMon = (InStr(UCase$(monStr), "USD") > 0 Or InStr(monStr, "$") > 0 Or _
                         InStr(UCase$(monStr), "DOLAR") > 0)
            Else
                okMon = (InStr(UCase$(monStr), "PEN") > 0 Or InStr(monStr, "S/") > 0 Or _
                         InStr(UCase$(monStr), "SOL") > 0)
            End If
            If Not okMon Then GoTo NextR_CM
        End If

        Dim rawFecha As Variant: rawFecha = arrM(r, iFecha)
        If IsEmpty(rawFecha) Or IsNull(rawFecha) Then GoTo NextR_CM
        If Not IsDate(rawFecha) And Not IsNumeric(rawFecha) Then GoTo NextR_CM

        Dim dtVal As Date
        On Error Resume Next: dtVal = CDate(rawFecha)
        If Err.Number <> 0 Then Err.Clear: On Error GoTo errExit: GoTo NextR_CM
        On Error GoTo errExit

        Dim mVal As Double: mVal = SafeDbl(arrM(r, iMonto))
        Dim dk As String:   dk   = CStr(CLng(CDbl(dtVal)))
        If dDM.Exists(dk) Then dDM(dk) = dDM(dk) + mVal Else dDM.Add dk, mVal
NextR_CM:
    Next r

    ' Si no hubo puntos, segundo pase sin filtro de moneda
    If dDM.Count = 0 And iMon > 0 Then
        For r = 1 To nRows
            If NormDoc(CStr(arrM(r, iDoc))) <> docKey Then GoTo NextR_CM2
            If iTP > 0 Then
                If NormPersona(CStr(arrM(r, iTP))) <> UCase$(persona) Then GoTo NextR_CM2
            End If
            Dim rawF2 As Variant: rawF2 = arrM(r, iFecha)
            If IsEmpty(rawF2) Or IsNull(rawF2) Then GoTo NextR_CM2
            Dim dtV2 As Date
            On Error Resume Next: dtV2 = CDate(rawF2)
            If Err.Number <> 0 Then Err.Clear: On Error GoTo errExit: GoTo NextR_CM2
            On Error GoTo errExit
            Dim mv2 As Double: mv2 = SafeDbl(arrM(r, iMonto))
            Dim dk2 As String: dk2 = CStr(CLng(CDbl(dtV2)))
            If dDM.Exists(dk2) Then dDM(dk2) = dDM(dk2) + mv2 Else dDM.Add dk2, mv2
NextR_CM2:
        Next r
    End If

    If dDM.Count = 0 Then WriteMontoSeriesCM = 0: Exit Function

    Dim sers() As Long
    ReDim sers(dDM.Count - 1)
    Dim kk As Long: kk = 0
    Dim vk As Variant
    For Each vk In dDM.Keys: sers(kk) = CLng(vk): kk = kk + 1: Next vk

    Dim ii As Long, jj As Long, tmp As Long
    For ii = 1 To UBound(sers)
        tmp = sers(ii): jj = ii - 1
        Do While jj >= 0 And sers(jj) > tmp: sers(jj + 1) = sers(jj): jj = jj - 1: Loop
        sers(jj + 1) = tmp
    Next ii

    minDtOut = CDate(sers(0))
    maxDtOut = CDate(sers(UBound(sers)))

    Dim wr As Long: wr = blockStart
    For ii = 0 To UBound(sers)
        wsh.Cells(wr, 1).Value = CDate(sers(ii))
        wsh.Cells(wr, 2).Value = dDM(CStr(sers(ii)))
        wr = wr + 1
    Next ii
    wsh.Range(wsh.Cells(blockStart, 1), wsh.Cells(wr - 1, 1)).NumberFormat = "dd/mm/yyyy"

    WriteMontoSeriesCM = UBound(sers) + 1
    Exit Function
errExit:
    WriteMontoSeriesCM = 0
End Function

' Normaliza numero de documento (sin ceros iniciales, sin guiones)
Private Function NormDoc(ByVal s As String) As String
    Dim t As String: t = UCase$(Trim$(s))
    t = Replace(t, "-", ""): t = Replace(t, ".", "")
    Do While Left$(t, 1) = "0" And Len(t) > 1: t = Mid$(t, 2): Loop
    NormDoc = t
End Function

' Normaliza TIPO_PERSONA a "NATURAL" o "JURIDICA"
Private Function NormPersona(ByVal s As String) As String
    Dim t As String: t = UCase$(Trim$(s))
    If InStr(t, "NAT") > 0 Or t = "PN" Then NormPersona = "NATURAL": Exit Function
    If InStr(t, "JUR") > 0 Or t = "PJ" Then NormPersona = "JURIDICA": Exit Function
    NormPersona = t
End Function
