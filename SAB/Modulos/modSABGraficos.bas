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
Private Const CHART_W       As Double = 567
Private Const CHART_H       As Double = 283
Private Const CHART_GAP_H   As Double = 18
Private Const CHART_TOP_MGN As Double = 30
Private Const MAX_CHARTS    As Long = 10
Private Const CLI_BLOCK     As Long = 400

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
    Dim mag As Double: mag = 10 ^ Int(log(v) / log(10))
    Dim m As Double:   m = v / mag
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

Private Function EnsureHelperSheet(ByVal wb As Workbook, ByVal shName As String) As Worksheet
    Dim wsh As Worksheet
    On Error Resume Next
    Set wsh = wb.Worksheets(shName)
    On Error GoTo 0
    If wsh Is Nothing Then
        Set wsh = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.count))
        wsh.Name = shName
    Else
        wsh.Cells.Clear
    End If
    wsh.visible = xlSheetVeryHidden
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
                ts = k(j):  k(j) = k(j + 1):   k(j + 1) = ts
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
        If dDM.exists(dk) Then
            dDM(dk) = dDM(dk) + mVal
        Else
            dDM.Add dk, mVal
        End If
NextR:
    Next r

    If dDM.count = 0 Then WriteMontoSeries = 0: Exit Function

    Dim sers() As Long
    ReDim sers(dDM.count - 1)
    Dim kk As Long: kk = 0
    Dim vk As Variant
    For Each vk In dDM.keys
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

    wsh.Cells(blockStart, 4).Value = axMinDate
    wsh.Cells(blockStart + 1, 4).Value = axMaxDate
    wsh.Cells(blockStart, 5).Value = promedio
    wsh.Cells(blockStart + 1, 5).Value = promedio
    wsh.Cells(blockStart, 4).NumberFormat = "dd/mm/yyyy"
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
    Dim axMaxBase As Double: axMaxBase = CDbl(DateSerial(maxY, maxM, 1))

    nMonths = (maxY - minY) * 12 + (maxM - minM)
    If nMonths < 1 Then nMonths = 1

    majorUnit = Int((axMaxBase - axMin - 1) / CDbl(nMonths)) + 1
    axMax = axMaxBase + 2
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

        Do While .SeriesCollection.count > 0
            .SeriesCollection(1).Delete
        Loop

        Dim s1 As series
        Set s1 = .SeriesCollection.NewSeries
        s1.Name = "Monto"
        s1.XValues = wsh.Range(wsh.Cells(bStart, 1), wsh.Cells(bEnd, 1))
        s1.Values = wsh.Range(wsh.Cells(bStart, 2), wsh.Cells(bEnd, 2))
        s1.MarkerStyle = xlMarkerStyleCircle
        s1.MarkerSize = 6
        s1.Format.line.ForeColor.RGB = RGB(68, 114, 196)
        s1.MarkerForegroundColor = RGB(42, 80, 150)
        s1.MarkerBackgroundColor = RGB(42, 80, 150)

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

        .HasTitle = True
        .ChartTitle.Text = titleText
        With .ChartTitle.Font
            .Size = 18
            .Bold = True
        End With

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

        Dim axY As Axis
        Set axY = .Axes(xlValue)
        With axY
            .MajorUnitIsAuto = True
            .TickLabels.NumberFormat = "#,##0"
            .TickLabels.Font.Size = 10
        End With

        .HasLegend = True
        With .Legend
            .Font.Size = 10
            .Position = xlLegendPositionRight
        End With

        .ChartArea.Border.LineStyle = xlContinuous
        .ChartArea.Border.Color = RGB(190, 190, 190)
        .ChartArea.Border.Weight = xlHairline

        On Error Resume Next
        .PlotArea.Left = 10
        .PlotArea.width = CHART_W - 10 - 120
        On Error GoTo 0

        On Error Resume Next
        .ChartGroups(1).HasDropLines = True
        With .ChartGroups(1).DropLines.Format.line
            .ForeColor.RGB = RGB(0, 0, 0)
            .Weight = 0.75
            .DashStyle = msoLineDashDot
        End With
        On Error GoTo 0
    End With

    Exit Sub
errExit:
End Sub

'=========================================================
' AddAlertTextBox
' Agrega un cuadro de texto de aviso en la hoja ws a la posicion indicada.
' Devuelve la altura del cuadro para poder calcular el offset del siguiente elemento.
'=========================================================
Private Function AddAlertTextBox( _
    ByVal ws As Worksheet, _
    ByVal msg As String, _
    ByVal cLeft As Double, _
    ByVal cTop As Double, _
    ByVal boxWidth As Double, _
    ByVal pref As String, _
    Optional ByVal fixedHeight As Double = 0) As Double

    On Error Resume Next
    Dim shp As Shape
    Dim initH As Double: initH = IIf(fixedHeight > 0, fixedHeight, 60)
    Set shp = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
                                   cLeft, cTop, boxWidth, initH)
    If shp Is Nothing Then AddAlertTextBox = 0: Exit Function

    shp.Name = pref & "_AVISO"
    With shp.TextFrame2
        .WordWrap = msoTrue
        .TextRange.Text = msg
        With .TextRange.Font
            .Size = 11
            .Bold = msoTrue
            .Fill.ForeColor.RGB = RGB(156, 0, 6)
        End With
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
    End With
    With shp.line
        .visible = msoTrue
        .ForeColor.RGB = RGB(255, 199, 206)
        .Weight = 1.5
    End With
    shp.Fill.ForeColor.RGB = RGB(255, 235, 238)
    shp.Fill.visible = msoTrue

    ' Autofit altura solo si no se especifico una altura fija
    If fixedHeight <= 0 Then
        On Error Resume Next
        shp.TextFrame.AutoSize = True
        On Error GoTo 0
    End If

    AddAlertTextBox = shp.height + 10
End Function

'=========================================================
' PUBLIC: BuildGraficosAlertasEnHoja
'
' loAL  : ListObject de ALERTAS MC (SAB_MC_ALERTAS_DEP o _RET)
'         Columna clave: "RUC/NIT" o "Cuenta" (primer columna, detectada dinamicamente)
'         Incluye columna TIPO_PERSONA con valores "N" o "J" (de Clientes_SAB)
' loMAIN: ListObject SAB_MC_MAIN (Cuenta, Fecha, Deposito/Retiro)
' which : "DEP" o "RET"
' suf   : sufijo de periodo para nombres de graficos
'
' Biparticion NAT (N) izquierda / JUR (J) derecha, top 5 cada una.
' Para agregar datos de MAIN por RUC/NIT (que puede cubrir varias Cuentas),
' se carga el diccionario Cuenta->RUC/NIT desde Clientes_SAB via modPQ_SAB_MC.
'=========================================================
Public Sub BuildGraficosAlertasEnHoja( _
    ByVal loAL As ListObject, _
    ByVal loMain As ListObject, _
    ByVal which As String, _
    ByVal suf As String)

    On Error GoTo ErrBuild

    If loAL Is Nothing Then Exit Sub
    If loAL.DataBodyRange Is Nothing Then Exit Sub
    If loMain Is Nothing Then Exit Sub
    If loMain.DataBodyRange Is Nothing Then Exit Sub

    Dim op As String: op = UCase$(Trim$(which))
    If op <> "DEP" And op <> "RET" Then op = "DEP"

    If Not LOHasColumn(loAL, "DESVIACION_MEDIA_%") Then Exit Sub
    If Not LOHasColumn(loAL, "PROMEDIO_MONTOS") And _
       Not LOHasColumn(loAL, "PROMEDIO_MONTOS_SOLES") Then Exit Sub

    Dim ws As Worksheet: Set ws = loAL.parent
    Dim wb As Workbook:  Set wb = ws.parent

    Dim opLabel   As String: opLabel = IIf(op = "DEP", "Deposito", "Retiro")
    Dim chartPref As String: chartPref = "GF_SAB_" & op & "_"

    ' =========================================================
    ' 1. Detectar columna clave (RUC/NIT o Cuenta) y TIPO_PERSONA
    ' =========================================================
    Dim iKey As Long, iDv As Long, iPm As Long, iCl As Long, iMo As Long, iTP As Long
    iKey = GetColIdx(loAL, "RUC/NIT")
    If iKey = 0 Then iKey = GetColIdx(loAL, "Cuenta")
    If iKey = 0 Then iKey = 1  ' fallback a primera columna
    iDv = GetColIdx(loAL, "DESVIACION_MEDIA_%")
    iPm = GetColIdx(loAL, "PROMEDIO_MONTOS_SOLES")
    If iPm = 0 Then iPm = GetColIdx(loAL, "PROMEDIO_MONTOS")
    iCl = GetColIdx(loAL, "CLASE")
    iMo = GetColIdx(loAL, "MONEDA")
    iTP = GetColIdx(loAL, "TIPO_PERSONA")

    ' =========================================================
    ' 2. Separar en NAT (N) y JUR (J) desde loAL
    ' =========================================================
    Const buf As Long = 256
    Dim nK(buf) As String, nDv(buf) As Double, nPm(buf) As Double
    Dim nCl(buf) As String, nMo(buf) As String
    Dim jK(buf) As String, jDv(buf) As Double, jPm(buf) As Double
    Dim jCl(buf) As String, jMo(buf) As String
    Dim nCnt As Long: nCnt = 0
    Dim jCnt As Long: jCnt = 0

    Dim arrAL As Variant
    arrAL = loAL.DataBodyRange.Value

    Dim ai As Long
    For ai = 1 To UBound(arrAL, 1)
        Dim sKeyAL As String: sKeyAL = Trim$(CStr(arrAL(ai, iKey)))
        If sKeyAL = "" Then GoTo NextAL
        Dim sDvAL  As Double: sDvAL = SafeDbl(arrAL(ai, iDv))
        Dim sPmAL  As Double: sPmAL = SafeDbl(arrAL(ai, iPm))
        Dim sClAL As String
        If iCl > 0 Then sClAL = Trim$(CStr(arrAL(ai, iCl))) Else sClAL = ""
        Dim sMoAL As String
        If iMo > 0 Then sMoAL = Trim$(CStr(arrAL(ai, iMo))) Else sMoAL = ""
        Dim sTPAL As String
        If iTP > 0 Then sTPAL = UCase$(Trim$(CStr(arrAL(ai, iTP)))) Else sTPAL = ""

        ' J si el tipo empieza por J o es PJ; todo lo demas es NAT
        Dim isJur As Boolean
        isJur = (sTPAL = "J" Or Left$(sTPAL, 1) = "J" Or sTPAL = "PJ")

        If isJur Then
            If jCnt < buf Then
                jK(jCnt) = sKeyAL: jDv(jCnt) = sDvAL: jPm(jCnt) = sPmAL
                jCl(jCnt) = sClAL: jMo(jCnt) = sMoAL
                jCnt = jCnt + 1
            End If
        Else
            If nCnt < buf Then
                nK(nCnt) = sKeyAL: nDv(nCnt) = sDvAL: nPm(nCnt) = sPmAL
                nCl(nCnt) = sClAL: nMo(nCnt) = sMoAL
                nCnt = nCnt + 1
            End If
        End If
NextAL:
    Next ai

    SortDescN nCnt, nK, nDv, nPm, nCl, nMo
    SortDescN jCnt, jK, jDv, jPm, jCl, jMo

    ' Si no hay clasificacion por tipo (iTP=0 o todo sin clasificar),
    ' usar MAX_CHARTS completo en NAT en lugar de dividirlo a la mitad
    Dim capN As Long, capJ As Long
    If iTP = 0 Or jCnt = 0 Then
        capN = MAX_CHARTS
        capJ = 0
    Else
        capN = MAX_CHARTS \ 2
        capJ = MAX_CHARTS \ 2
    End If
    If nCnt > capN Then nCnt = capN
    If jCnt > capJ Then jCnt = capJ

    ' =========================================================
    ' 3. Cargar diccionario Cuenta->RUC/NIT directamente desde Clientes_SAB
    ' =========================================================
    ' 3. Columna clave
    ' =========================================================
    Dim keyColName As String: keyColName = loAL.ListColumns(iKey).Name

    ' =========================================================
    ' 4. Columnas en loMAIN
    ' =========================================================
    Dim iM_cta As Long, iM_fch As Long, iM_mto As Long, iM_ruc As Long
    iM_cta = GetColIdx(loMain, "Cuenta")
    iM_fch = GetColIdx(loMain, "Fecha")
    iM_mto = GetColIdx(loMain, "Monto en Soles")
    If iM_mto = 0 Then
        iM_mto = GetColIdx(loMain, IIf(op = "DEP", "Dep" & Chr(243) & "sito", "Retiro"))
    End If
    iM_ruc = GetColIdx(loMain, "RUC/NIT")

    If iM_cta = 0 Or iM_fch = 0 Or iM_mto = 0 Then GoTo fin

    ' Si MAIN tiene columna RUC/NIT usarla directamente (sin dict local).
    ' Fallback: dict local si MAIN es de una version anterior sin esa columna.
    Dim dCuentaDoc As Object
    Set dCuentaDoc = CreateObject("Scripting.Dictionary")  ' nunca Nothing, evita Error 5 en VBA And
    Dim useRucCol As Boolean
    useRucCol = (iM_ruc > 0) And (StrComp(keyColName, "RUC/NIT", vbTextCompare) = 0)
    If Not useRucCol And StrComp(keyColName, "RUC/NIT", vbTextCompare) = 0 Then
        Set dCuentaDoc = BuildCuentaDocDictLocal()
    End If

    Dim arrM As Variant
    arrM = loMain.DataBodyRange.Value

    ' =========================================================
    ' 5. Hoja helper y limpieza de graficos anteriores
    ' =========================================================
    Dim wsh As Worksheet
    Set wsh = EnsureHelperSheet(wb, "_GF_MC_" & op & "_HELPER")

    DeleteChartsByPrefix ws, chartPref

    ' =========================================================
    ' 6. Posicion: dos columnas NAT izquierda / JUR derecha
    ' =========================================================
    Dim lastALCol As Long
    lastALCol = loAL.Range.Column + loAL.Range.Columns.count

    Dim chartLeft1   As Double: chartLeft1 = ws.Cells(1, lastALCol + 1).Left
    Dim chartLeft2   As Double: chartLeft2 = chartLeft1 + CHART_W + 14
    Dim chartTopBase As Double: chartTopBase = ws.Cells(loAL.Range.Row, 1).top + CHART_TOP_MGN

    Dim cliIdx As Long: cliIdx = 0

    ' Diagnostico: mostrar estado de columnas clave
    Debug.Print "[SABGraficos] op=" & op & " iM_cta=" & iM_cta & " iM_fch=" & iM_fch & " iM_mto=" & iM_mto & " iM_ruc=" & iM_ruc & " useRucCol=" & useRucCol & " keyColName=" & keyColName

    ' =========================================================
    ' 7. Primer pase: contar generables por tipo separado
    ' =========================================================
    Dim generablesN As Long: generablesN = 0
    Dim generablesJ As Long: generablesJ = 0
    Dim ci As Long, cj As Long
    Dim testRows As Long
    Dim testMin As Date, testMax As Date

    For ci = 0 To nCnt - 1
        If nK(ci) = "" Then GoTo CountNAT
        testRows = WriteMontoSeriesByDoc(wsh, 1, nK(ci), keyColName, _
                                         iM_cta, iM_fch, iM_mto, arrM, dCuentaDoc, _
                                         testMin, testMax, useRucCol, iM_ruc)
        If testRows >= 2 Then generablesN = generablesN + 1
CountNAT:
    Next ci

    For cj = 0 To jCnt - 1
        If jK(cj) = "" Then GoTo CountJUR
        testRows = WriteMontoSeriesByDoc(wsh, 1, jK(cj), keyColName, _
                                         iM_cta, iM_fch, iM_mto, arrM, dCuentaDoc, _
                                         testMin, testMax, useRucCol, iM_ruc)
        If testRows >= 2 Then generablesJ = generablesJ + 1
CountJUR:
    Next cj

    Debug.Print "[SABGraficos] nCnt=" & nCnt & " jCnt=" & jCnt & " generablesN=" & generablesN & " generablesJ=" & generablesJ

    ' Limpiar helper antes del pase real
    wsh.Cells.Clear

    ' =========================================================
    ' 8. Aviso global si no hay ningun grafico posible
    ' =========================================================
    Dim avisoWidth As Double: avisoWidth = CHART_W * 2 + 14

    If generablesN = 0 And generablesJ = 0 Then
        Dim msgCero As String
        msgCero = Chr(9888) & " Sin gr" & Chr(225) & "ficos disponibles para [" & opLabel & "]" & vbCrLf & _
                  "Todos los clientes en el top de alertas tienen una " & _
                  Chr(250) & "nica operaci" & Chr(243) & "n registrada. " & _
                  "No es posible calcular tendencia con un solo punto de datos."
        AddAlertTextBox ws, msgCero, chartLeft1, chartTopBase, avisoWidth, chartPref
        GoTo fin
    End If

    ' =========================================================
    ' 9. Pase real: generar graficos con >= 2 fechas distintas
    ' =========================================================
    Dim nIdxPlot As Long: nIdxPlot = 0   ' indice de posicion vertical NAT
    Dim jIdxPlot As Long: jIdxPlot = 0   ' indice de posicion vertical JUR

    For ci = 0 To nCnt - 1
        If nK(ci) = "" Then GoTo NextNAT

        Dim bStartN As Long: bStartN = 1 + (nIdxPlot + jIdxPlot) * CLI_BLOCK

        Dim minDtN As Date, maxDtN As Date, rowsN As Long
        rowsN = WriteMontoSeriesByDoc(wsh, bStartN, nK(ci), keyColName, _
                                      iM_cta, iM_fch, iM_mto, arrM, dCuentaDoc, _
                                      minDtN, maxDtN, useRucCol, iM_ruc)
        If rowsN < 2 Then GoTo NextNAT

        Dim axMinN As Double, axMaxN As Double, nMthN As Long, mjUN As Double
        CalcAxisBounds minDtN, maxDtN, axMinN, axMaxN, nMthN, mjUN
        WritePromedioSeries wsh, bStartN, CDate(axMinN), CDate(axMaxN), nPm(ci)

        Dim titleN As String
        titleN = "[NAT][" & opLabel & "] " & keyColName & ": " & nK(ci) & _
                 " | Desviacion: " & Format(nDv(ci), "0.00") & "%"
        If nCl(ci) <> "" Then titleN = titleN & " | " & nCl(ci)
        If nMo(ci) <> "" Then titleN = titleN & " | " & nMo(ci)

        Dim cTopN As Double: cTopN = chartTopBase + CDbl(nIdxPlot) * (CHART_H + CHART_GAP_H)
        CreateScatterChart ws, wsh, bStartN, rowsN, nPm(ci), _
                           axMinN, axMaxN, mjUN, _
                           chartLeft1, cTopN, chartPref & "N" & Format(ci + 1, "00"), titleN
        nIdxPlot = nIdxPlot + 1
NextNAT:
    Next ci

    ' Placeholder NAT si no se genero ninguno
    If generablesN = 0 And jCnt > 0 Then
        Dim msgNoNAT As String
        msgNoNAT = "Persona Natural" & vbCrLf & Chr(10) & _
                   "No fue posible generar gr" & Chr(225) & "ficos para clientes " & _
                   "Persona Natural en [" & opLabel & "]." & vbCrLf & _
                   "Los clientes de este tipo no cuentan con suficientes fechas " & _
                   "de operaci" & Chr(243) & "n distintas para mostrar una tendencia."
        AddAlertTextBox ws, msgNoNAT, chartLeft1, chartTopBase, CHART_W, _
                        chartPref & "NAT", CHART_H
    End If

    For cj = 0 To jCnt - 1
        If jK(cj) = "" Then GoTo NextJUR

        Dim bStartJ As Long: bStartJ = 1 + (nIdxPlot + jIdxPlot) * CLI_BLOCK

        Dim minDtJ As Date, maxDtJ As Date, rowsJ As Long
        rowsJ = WriteMontoSeriesByDoc(wsh, bStartJ, jK(cj), keyColName, _
                                      iM_cta, iM_fch, iM_mto, arrM, dCuentaDoc, _
                                      minDtJ, maxDtJ, useRucCol, iM_ruc)
        If rowsJ < 2 Then GoTo NextJUR

        Dim axMinJ As Double, axMaxJ As Double, nMthJ As Long, mjUJ As Double
        CalcAxisBounds minDtJ, maxDtJ, axMinJ, axMaxJ, nMthJ, mjUJ
        WritePromedioSeries wsh, bStartJ, CDate(axMinJ), CDate(axMaxJ), jPm(cj)

        Dim titleJ As String
        titleJ = "[JUR][" & opLabel & "] " & keyColName & ": " & jK(cj) & _
                 " | Desviacion: " & Format(jDv(cj), "0.00") & "%"
        If jCl(cj) <> "" Then titleJ = titleJ & " | " & jCl(cj)
        If jMo(cj) <> "" Then titleJ = titleJ & " | " & jMo(cj)

        Dim cTopJ As Double: cTopJ = chartTopBase + CDbl(jIdxPlot) * (CHART_H + CHART_GAP_H)
        CreateScatterChart ws, wsh, bStartJ, rowsJ, jPm(cj), _
                           axMinJ, axMaxJ, mjUJ, _
                           chartLeft2, cTopJ, chartPref & "J" & Format(cj + 1, "00"), titleJ
        jIdxPlot = jIdxPlot + 1
NextJUR:
    Next cj

    ' Placeholder JUR si no se genero ninguno
    If generablesJ = 0 And nCnt > 0 Then
        Dim msgNoJUR As String
        msgNoJUR = "Persona Jur" & Chr(237) & "dica" & vbCrLf & Chr(10) & _
                   "No fue posible generar gr" & Chr(225) & "ficos para clientes " & _
                   "Persona Jur" & Chr(237) & "dica en [" & opLabel & "]." & vbCrLf & _
                   "Los clientes de este tipo no cuentan con suficientes fechas " & _
                   "de operaci" & Chr(243) & "n distintas para mostrar una tendencia."
        AddAlertTextBox ws, msgNoJUR, chartLeft2, chartTopBase, CHART_W, _
                        chartPref & "JUR", CHART_H
    End If

    Exit Sub

ErrBuild:
    Dim errN As Long: errN = Err.Number
    Dim errD As String: errD = Err.Description
    MsgBox "Error en BuildGraficosAlertasEnHoja [" & op & "]" & vbCrLf & _
           "Numero: " & errN & vbCrLf & _
           "Descripcion: " & errD, vbCritical, "SAB Graficos"

fin:
End Sub

'=========================================================
' BuildCuentaDocDictLocal
' Construye diccionario Cuenta -> "RUC/NIT|Tipo" desde Clientes_SAB.
' Normaliza numeros de cuenta eliminando decimales y Chr(160).
'=========================================================
Private Function BuildCuentaDocDictLocal() As Object
    Dim d As Object: Set d = CreateObject("Scripting.Dictionary")
    Dim ws As Worksheet, lo As ListObject
    Dim colCta As Long, colDoc As Long, colTipo As Long
    Dim i As Long

    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, "Clientes_SAB", vbTextCompare) = 0 Then
                If lo.DataBodyRange Is Nothing Then GoTo NotFound
                colCta = 0: colDoc = 0: colTipo = 0
                For i = 1 To lo.ListColumns.count
                    Dim hdr As String
                    hdr = UCase$(Replace(Replace(NormStr(lo.ListColumns(i).Name), "/", ""), "-", ""))
                    If hdr = "CUENTA" Then colCta = i
                    If hdr = "RUCNIT" Then colDoc = i
                    If hdr = "TIPO" Then colTipo = i
                Next i
                If colCta = 0 Or colDoc = 0 Then GoTo NotFound
                Dim data As Variant: data = lo.DataBodyRange.Value2
                Dim nR As Long: nR = UBound(data, 1)
                Dim vC As Variant, vD As Variant, vT As Variant
                Dim sC As String, sD As String, sT As String
                For i = 1 To nR
                    vC = data(i, colCta): vD = data(i, colDoc)
                    If IsEmpty(vC) Or IsNull(vC) Or IsError(vC) Then GoTo NextCli
                    If IsEmpty(vD) Or IsNull(vD) Or IsError(vD) Then GoTo NextCli
                    sC = NormStr(CStr(vC))
                    sD = NormStr(CStr(vD))
                    sT = ""
                    If colTipo > 0 Then
                        vT = data(i, colTipo)
                        If Not (IsEmpty(vT) Or IsNull(vT) Or IsError(vT)) Then
                            sT = UCase$(NormStr(CStr(vT)))
                        End If
                    End If
                    If Len(sC) > 0 And Len(sD) > 0 Then
                        If Not d.exists(sC) Then d.Add sC, sD & "|" & sT
                    End If
NextCli:
                Next i
                Set BuildCuentaDocDictLocal = d
                Exit Function
            End If
        Next lo
    Next ws
NotFound:
    Set BuildCuentaDocDictLocal = d
End Function

'=========================================================
' NormStr: limpia Chr(160), recorta, elimina decimales si es numero
'=========================================================
Private Function NormStr(ByVal s As String) As String
    Dim i As Integer
    For i = 0 To 31: s = Replace(s, Chr(i), ""): Next i
    s = Replace(s, Chr(160), "")
    s = Trim$(s)
    ' Eliminar decimales si es numero entero (ej. "12345678.0" -> "12345678")
    ' Usar CDec para evitar desbordamiento con RUC/NIT de 11 digitos (>Long max)
    If IsNumeric(s) Then
        On Error Resume Next
        Dim d2 As Double: d2 = CDbl(s)
        If Err.Number = 0 And d2 = Int(d2) Then
            Dim dec2 As Variant: dec2 = CDec(d2)
            If Err.Number = 0 Then s = CStr(dec2)
        End If
        Err.Clear
        On Error GoTo 0
    End If
    NormStr = s
End Function

'=========================================================
' WriteMontoSeriesByDoc
' Escribe serie de montos diarios en wsh A-B desde blockStart.
' Si keyColName = "RUC/NIT" y dCuentaDoc no es Nothing, agrega
' datos de TODAS las Cuentas que mapean al RUC/NIT dado.
' Si keyColName = "Cuenta" o no hay dict, busca directamente por Cuenta.
'=========================================================
Private Function WriteMontoSeriesByDoc( _
    ByVal wsh As Worksheet, _
    ByVal blockStart As Long, _
    ByVal docKey As String, _
    ByVal keyColName As String, _
    ByVal iCuenta As Long, _
    ByVal iFecha As Long, _
    ByVal iMonto As Long, _
    ByRef arrM As Variant, _
    ByVal dCuentaDoc As Object, _
    ByRef minDtOut As Date, _
    ByRef maxDtOut As Date, _
    Optional ByVal useRucCol As Boolean = False, _
    Optional ByVal iRucCol As Long = 0) As Long

    On Error GoTo errExit

    ' Determinar modo de busqueda
    ' useRucCol=True: comparar directamente la columna RUC/NIT de arrM (mas preciso)
    ' byDoc=True:     usar diccionario local (fallback si MAIN no tiene col RUC/NIT)
    Dim byDoc As Boolean
    byDoc = (Not useRucCol) And _
            (StrComp(keyColName, "RUC/NIT", vbTextCompare) = 0) And _
            Not (dCuentaDoc Is Nothing) And (dCuentaDoc.count > 0)
    Dim normDocKey As String: normDocKey = NormStr(docKey)

    Dim dDM As Object: Set dDM = CreateObject("Scripting.Dictionary")

    Dim r As Long
    Dim nRows As Long: nRows = UBound(arrM, 1)

    For r = 1 To nRows
        ' Filtrar por RUC/NIT o Cuenta segun modo
        If useRucCol Then
            ' Comparar columna RUC/NIT de MAIN; si esta vacia, fallback a Cuenta
            ' (ocurre cuando la cuenta no esta en Clientes_SAB y BuildAlertasVBA
            '  uso la Cuenta directamente como clave en la tabla de alertas)
            Dim sRuc As String: sRuc = NormStr(CStr(arrM(r, iRucCol)))
            If Len(sRuc) = 0 Then
                sRuc = NormStr(CStr(arrM(r, iCuenta)))
            End If
            If StrComp(sRuc, normDocKey, vbBinaryCompare) <> 0 Then GoTo NextR2
        ElseIf byDoc Then
            Dim sCta As String: sCta = NormStr(CStr(arrM(r, iCuenta)))
            If Not dCuentaDoc.exists(sCta) Then GoTo NextR2
            Dim rawVal As String: rawVal = CStr(dCuentaDoc(sCta))
            Dim pp As Long: pp = InStr(rawVal, "|")
            Dim ctaDoc As String
            If pp > 0 Then ctaDoc = NormStr(Left$(rawVal, pp - 1)) Else ctaDoc = NormStr(rawVal)
            If StrComp(ctaDoc, normDocKey, vbBinaryCompare) <> 0 Then GoTo NextR2
        Else
            Dim sCtaD As String: sCtaD = NormStr(CStr(arrM(r, iCuenta)))
            If StrComp(sCtaD, normDocKey, vbBinaryCompare) <> 0 Then GoTo NextR2
        End If

        Dim rawFecha As Variant: rawFecha = arrM(r, iFecha)
        If IsEmpty(rawFecha) Or IsNull(rawFecha) Then GoTo NextR2
        If Not IsDate(rawFecha) And Not IsNumeric(rawFecha) Then GoTo NextR2

        Dim dtVal As Date
        On Error Resume Next: dtVal = CDate(rawFecha): If Err.Number <> 0 Then Err.Clear: On Error GoTo errExit: GoTo NextR2
        On Error GoTo errExit

        Dim rawMto As Variant: rawMto = arrM(r, iMonto)
        If IsEmpty(rawMto) Or IsNull(rawMto) Then GoTo NextR2

        Dim mtoVal As Double
        On Error Resume Next: mtoVal = CDbl(rawMto): If Err.Number <> 0 Then Err.Clear: On Error GoTo errExit: GoTo NextR2
        On Error GoTo errExit
        If mtoVal = 0 Then GoTo NextR2

        Dim dtKey As Long: dtKey = CLng(CDbl(dtVal))
        If dDM.exists(dtKey) Then
            dDM(dtKey) = CDbl(dDM(dtKey)) + mtoVal
        Else
            dDM.Add dtKey, mtoVal
        End If
NextR2:
    Next r

    If dDM.count = 0 Then GoTo errExit

    ' Escribir en hoja helper y calcular min/max
    Dim nPts As Long: nPts = 0
    Dim kk As Variant
    Dim gotMin As Boolean: gotMin = False
    For Each kk In dDM.keys
        Dim dtOut As Date: dtOut = CDate(CLng(kk))
        If Not gotMin Then
            minDtOut = dtOut: maxDtOut = dtOut: gotMin = True
        Else
            If dtOut < minDtOut Then minDtOut = dtOut
            If dtOut > maxDtOut Then maxDtOut = dtOut
        End If
        wsh.Cells(blockStart + nPts, 1).Value = dtOut
        wsh.Cells(blockStart + nPts, 1).NumberFormat = "dd/mm/yyyy"
        wsh.Cells(blockStart + nPts, 2).Value = CDbl(dDM(kk))
        nPts = nPts + 1
    Next kk

    WriteMontoSeriesByDoc = nPts
    Exit Function
errExit:
    WriteMontoSeriesByDoc = 0
End Function

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
    ByVal loAL As ListObject, _
    ByVal loMain As ListObject, _
    ByVal which As String)

    On Error GoTo fin

    If loAL Is Nothing Then Exit Sub
    If loAL.DataBodyRange Is Nothing Then Exit Sub
    If loMain Is Nothing Then Exit Sub
    If loMain.DataBodyRange Is Nothing Then Exit Sub

    Dim op As String: op = UCase$(Trim$(which))
    If op <> "COM" And op <> "VEN" Then op = "COM"

    If Not LOHasColumn(loAL, "DESVIACION_MEDIA_%") Then Exit Sub
    If Not LOHasColumn(loAL, "PROMEDIO_MONTOS") Then Exit Sub
    If Not LOHasColumn(loAL, "Documento") Then Exit Sub

    Dim ws As Worksheet: Set ws = loAL.parent
    Dim wb As Workbook:  Set wb = ws.parent

    Dim opLabel   As String: opLabel = IIf(op = "COM", "Compra", "Venta")
    Dim chartPref As String: chartPref = "GF_CM_" & op & "_"

    ' =========================================================
    ' 1. Extraer candidatos NAT y JUR desde loAL (solo nivel 2 o 3)
    ' =========================================================
    Dim iKey As Long, iDv As Long, iPm As Long, iTP As Long, iNR As Long
    iKey = GetColIdx(loAL, "Documento")
    iDv = GetColIdx(loAL, "DESVIACION_MEDIA_%")
    iPm = GetColIdx(loAL, "PROMEDIO_MONTOS")
    iTP = GetColIdx(loAL, "TIPO_PERSONA")
    iNR = GetColIdx(loAL, "NIVEL_RIESGO")

    Const buf As Long = 256
    Dim nK(buf) As String, nDv(buf) As Double, nPm(buf) As Double
    Dim nTP_a(buf) As String, nPH(buf) As String
    Dim jK(buf) As String, jDv(buf) As Double, jPm(buf) As Double
    Dim jTP_a(buf) As String, jPH(buf) As String
    Dim nCnt As Long: nCnt = 0
    Dim jCnt As Long: jCnt = 0

    Dim arrAL As Variant
    arrAL = loAL.DataBodyRange.Value

    Dim ai As Long
    For ai = 1 To UBound(arrAL, 1)
        If iNR > 0 Then
            Dim nrVal As Long: nrVal = CLng(SafeDbl(arrAL(ai, iNR)))
            If nrVal < 2 Then GoTo NextALCM
        End If

        Dim sKey As String: sKey = Trim$(CStr(arrAL(ai, iKey)))
        Dim sDv  As Double: sDv = SafeDbl(arrAL(ai, iDv))
        Dim sPm  As Double: sPm = SafeDbl(arrAL(ai, iPm))
        Dim sTP  As String
        sTP = IIf(iTP > 0, UCase$(Trim$(CStr(arrAL(ai, iTP)))), "")

        If InStr(sTP, "JUR") > 0 Or sTP = "PJ" Then
            If jCnt < buf Then
                jK(jCnt) = sKey: jDv(jCnt) = sDv: jPm(jCnt) = sPm
                jTP_a(jCnt) = sTP: jPH(jCnt) = ""
                jCnt = jCnt + 1
            End If
        Else
            If nCnt < buf Then
                nK(nCnt) = sKey: nDv(nCnt) = sDv: nPm(nCnt) = sPm
                nTP_a(nCnt) = sTP: nPH(nCnt) = ""
                nCnt = nCnt + 1
            End If
        End If
NextALCM:
    Next ai

    SortDescN nCnt, nK, nDv, nPm, nTP_a, nPH
    SortDescN jCnt, jK, jDv, jPm, jTP_a, jPH

    Dim capCharts As Long: capCharts = MAX_CHARTS \ 2
    If nCnt > capCharts Then nCnt = capCharts
    If jCnt > capCharts Then jCnt = capCharts

    ' =========================================================
    ' 2. Columnas en loMAIN
    ' =========================================================
    Dim iM_doc As Long, iM_fch As Long, iM_tp As Long, iM_mto As Long
    iM_doc = GetColIdx(loMain, "Documento")
    iM_fch = GetColIdx(loMain, "Fecha")
    iM_tp = GetColIdx(loMain, "Tipo Persona")

    Dim mntCandidatos(4) As String
    mntCandidatos(0) = "Total Neto":  mntCandidatos(1) = "Monto Des"
    mntCandidatos(2) = "Monto Ori":   mntCandidatos(3) = "Gan/Per PEN"
    mntCandidatos(4) = "Gan/Per"
    Dim mc As Long
    For mc = 0 To 4
        iM_mto = GetColIdx(loMain, mntCandidatos(mc))
        If iM_mto > 0 Then Exit For
    Next mc

    Dim iM_mon As Long: iM_mon = GetColIdx(loMain, "Moneda Ori")

    If iM_doc = 0 Or iM_fch = 0 Or iM_mto = 0 Then GoTo fin

    Dim arrM As Variant
    arrM = loMain.DataBodyRange.Value

    ' =========================================================
    ' 3. Hoja helper y limpieza de graficos anteriores
    ' =========================================================
    Dim wsh As Worksheet
    Set wsh = EnsureHelperSheet(wb, "_GF_CM_" & op & "_HELPER")
    DeleteChartsByPrefix ws, chartPref

    ' Limpiar tambien los textboxes de aviso previos
    Dim shp As Shape
    On Error Resume Next
    For Each shp In ws.Shapes
        If Left$(shp.Name, Len(chartPref)) = chartPref Then shp.Delete
    Next shp
    On Error GoTo fin

    ' =========================================================
    ' 4. Posicion
    ' =========================================================
    Dim lastALCol As Long
    lastALCol = loAL.Range.Column + loAL.Range.Columns.count

    Dim chartLeft1   As Double: chartLeft1 = ws.Cells(1, lastALCol + 1).Left
    Dim chartLeft2   As Double: chartLeft2 = chartLeft1 + CHART_W + 14
    Dim chartTopBase As Double: chartTopBase = ws.Cells(loAL.Range.Row, 1).top + CHART_TOP_MGN

    Dim cliIdx As Long: cliIdx = 0
    Dim nPlot  As Long: nPlot = 0
    Dim jPlot  As Long: jPlot = 0

    ' =========================================================
    ' 5. Graficos NAT (columna izquierda)
    ' =========================================================
    Dim ci As Long
    For ci = 0 To nCnt - 1
        If nK(ci) = "" Then GoTo NextNAT_CM

        Dim bStartN As Long: bStartN = 1 + cliIdx * CLI_BLOCK

        Dim minDtN As Date, maxDtN As Date, rowsN As Long
        rowsN = WriteMontoSeriesCM(wsh, bStartN, NormDoc(nK(ci)), "NATURAL", op, _
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

        Dim cTopN As Double: cTopN = chartTopBase + CDbl(nPlot) * (CHART_H + CHART_GAP_H)
        CreateScatterChart ws, wsh, bStartN, rowsN, nPm(ci), _
                           axMinN, axMaxN, mjUN, _
                           chartLeft1, cTopN, chartPref & "N" & Format(ci + 1, "00"), titleN
        cliIdx = cliIdx + 1
        nPlot = nPlot + 1
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
        rowsJ = WriteMontoSeriesCM(wsh, bStartJ, NormDoc(jK(cj)), "JURIDICA", op, _
                                   iM_doc, iM_fch, iM_mto, iM_tp, iM_mon, arrM, _
                                   minDtJ, maxDtJ)
        If rowsJ = 0 Then GoTo NextJUR_CM

        Dim axMinJ As Double, axMaxJ As Double, nMthJ As Long, mjUJ As Double
        CalcAxisBounds minDtJ, maxDtJ, axMinJ, axMaxJ, nMthJ, mjUJ
        WritePromedioSeries wsh, bStartJ, CDate(axMinJ), CDate(axMaxJ), jPm(cj)

        Dim titleJ As String
        titleJ = "[JUR] " & opLabel & " RUC: " & jK(cj) & _
                 " | Desviacion: " & Format(jDv(cj), "0.00") & "%"

        Dim cTopJ As Double: cTopJ = chartTopBase + CDbl(jPlot) * (CHART_H + CHART_GAP_H)
        CreateScatterChart ws, wsh, bStartJ, rowsJ, jPm(cj), _
                           axMinJ, axMaxJ, mjUJ, _
                           chartLeft2, cTopJ, chartPref & "J" & Format(cj + 1, "00"), titleJ
        cliIdx = cliIdx + 1
        jPlot = jPlot + 1
NextJUR_CM:
    Next cj

    ' =========================================================
    ' 7. Avisos informativos si hay menos de capCharts por tipo
    ' =========================================================
    On Error Resume Next

    If nPlot < capCharts Then
        Dim msgN As String
        msgN = "Solo hay " & nPlot & " cliente(s) Persona Natural en nivel de riesgo no bajo (nivel 2 o 3)."
        Dim shpN As Shape
        Set shpN = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
            chartLeft1, chartTopBase + CDbl(nPlot) * (CHART_H + CHART_GAP_H), CHART_W, 40)
        If Not shpN Is Nothing Then
            shpN.Name = chartPref & "NAT_INFO"
            With shpN.TextFrame2
                .WordWrap = msoTrue
                .TextRange.Text = msgN
                With .TextRange.Font
                    .Size = 10
                    .Bold = msoFalse
                    .Fill.ForeColor.RGB = RGB(80, 80, 80)
                End With
                .TextRange.ParagraphFormat.Alignment = msoAlignLeft
            End With
            shpN.Fill.ForeColor.RGB = RGB(255, 255, 255)
            shpN.Fill.visible = msoTrue
            shpN.line.visible = msoTrue
            shpN.line.ForeColor.RGB = RGB(200, 200, 200)
            shpN.line.Weight = 0.75
        End If
    End If

    If jPlot < capCharts Then
        Dim msgJ As String
        msgJ = "Solo hay " & jPlot & " cliente(s) Persona Juridica en nivel de riesgo no bajo (nivel 2 o 3)."
        Dim shpJ As Shape
        Set shpJ = ws.Shapes.AddTextbox(msoTextOrientationHorizontal, _
            chartLeft2, chartTopBase + CDbl(jPlot) * (CHART_H + CHART_GAP_H), CHART_W, 40)
        If Not shpJ Is Nothing Then
            shpJ.Name = chartPref & "JUR_INFO"
            With shpJ.TextFrame2
                .WordWrap = msoTrue
                .TextRange.Text = msgJ
                With .TextRange.Font
                    .Size = 10
                    .Bold = msoFalse
                    .Fill.ForeColor.RGB = RGB(80, 80, 80)
                End With
                .TextRange.ParagraphFormat.Alignment = msoAlignLeft
            End With
            shpJ.Fill.ForeColor.RGB = RGB(255, 255, 255)
            shpJ.Fill.visible = msoTrue
            shpJ.line.visible = msoTrue
            shpJ.line.ForeColor.RGB = RGB(200, 200, 200)
            shpJ.line.Weight = 0.75
        End If
    End If

    On Error GoTo 0

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
        Dim dk As String:   dk = CStr(CLng(CDbl(dtVal)))
        If dDM.exists(dk) Then dDM(dk) = dDM(dk) + mVal Else dDM.Add dk, mVal
NextR_CM:
    Next r

    ' Si no hubo puntos, segundo pase sin filtro de moneda
    If dDM.count = 0 And iMon > 0 Then
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
            If dDM.exists(dk2) Then dDM(dk2) = dDM(dk2) + mv2 Else dDM.Add dk2, mv2
NextR_CM2:
        Next r
    End If

    If dDM.count = 0 Then WriteMontoSeriesCM = 0: Exit Function

    Dim sers() As Long
    ReDim sers(dDM.count - 1)
    Dim kk As Long: kk = 0
    Dim vk As Variant
    For Each vk In dDM.keys: sers(kk) = CLng(vk): kk = kk + 1: Next vk

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
