Option Explicit

'=========================================================
' modFondosGraficos
' - Genera 3 pivots + 3 gráficos EN LA MISMA HOJA de ALERTAS
' - No crea hojas adicionales
' - Limpia solo los gráficos/pivots creados por este módulo
'=========================================================

'======================
' Normalización de nombres de columnas (robusta)
'======================
Private Function StripDiacriticsUpper(ByVal s As String) As String
    Dim t As String
    t = UCase$(Trim$(s))

    t = Replace(t, "Á", "A"): t = Replace(t, "À", "A"): t = Replace(t, "Â", "A"): t = Replace(t, "Ä", "A")
    t = Replace(t, "É", "E"): t = Replace(t, "È", "E"): t = Replace(t, "Ê", "E"): t = Replace(t, "Ë", "E")
    t = Replace(t, "Í", "I"): t = Replace(t, "Ì", "I"): t = Replace(t, "Î", "I"): t = Replace(t, "Ï", "I")
    t = Replace(t, "Ó", "O"): t = Replace(t, "Ò", "O"): t = Replace(t, "Ô", "O"): t = Replace(t, "Ö", "O")
    t = Replace(t, "Ú", "U"): t = Replace(t, "Ù", "U"): t = Replace(t, "Û", "U"): t = Replace(t, "Ü", "U")
    t = Replace(t, "Ñ", "N")

    StripDiacriticsUpper = t
End Function

Private Function CanonColName(ByVal s As String) As String
    Dim t As String
    t = StripDiacriticsUpper(s)
    t = Replace(t, Chr$(160), " ")
    t = Replace(t, "°", "")
    t = Replace(t, "º", "")
    t = Replace(t, " ", "")
    CanonColName = t
End Function

Private Function FindListColumnByName(ByVal lo As ListObject, ByVal colName As String) As ListColumn
    Dim lc As ListColumn
    Dim want As String
    want = CanonColName(colName)

    For Each lc In lo.ListColumns
        If CanonColName(lc.name) = want Then
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

'======================
' Limpieza selectiva en hoja de alertas
'======================
Private Sub DeleteChartsByPrefix(ByVal ws As Worksheet, ByVal prefix As String)
    Dim co As ChartObject
    On Error Resume Next
    For Each co In ws.ChartObjects
        If StrComp(Left$(co.name, Len(prefix)), prefix, vbTextCompare) = 0 Then
            co.Delete
        End If
    Next co
    On Error GoTo 0
End Sub

Private Sub ClearPivotAreaFromRow(ByVal ws As Worksheet, ByVal startRow As Long)
    Dim pt As PivotTable
    On Error Resume Next
    For Each pt In ws.PivotTables
        If Not pt.TableRange2 Is Nothing Then
            If pt.TableRange2.Row >= startRow Then
                pt.TableRange2.Clear
            End If
        End If
    Next pt
    On Error GoTo 0

    On Error Resume Next
    ws.Range(ws.Cells(startRow, 1), ws.Cells(startRow + 120, 12)).Clear
    On Error GoTo 0
End Sub

Private Function SafePivotNameToken(ByVal wsName As String) As String
    Dim t As String
    t = Replace(wsName, " ", "_")
    t = Replace(t, "-", "_")
    t = Replace(t, ".", "_")
    t = Replace(t, ":", "_")
    t = Replace(t, "/", "_")
    t = Replace(t, "\", "_")
    If Len(t) > 20 Then t = Left$(t, 20)
    SafePivotNameToken = t
End Function

'======================
' Público: construir gráficos en hoja de ALERTAS
'======================
Public Sub BuildGraficosAlertasEnHoja(ByVal loAL As ListObject, Optional ByVal tituloBase As String = "FONDOS ALERTAS")
    On Error GoTo fin

    If loAL Is Nothing Then Exit Sub
    If loAL.DataBodyRange Is Nothing Then Exit Sub

    ' Validar columnas mínimas
    If Not LOHasColumn(loAL, "CUC") Then Exit Sub
    If Not LOHasColumn(loAL, "NIVEL_RIESGO") Then Exit Sub
    If Not LOHasColumn(loAL, "SUMA_MONTOS") Then Exit Sub
    If Not LOHasColumn(loAL, "DESVIACION_MEDIA_%") Then Exit Sub

    Dim ws As Worksheet
    Set ws = loAL.parent

    Dim wb As Workbook
    Set wb = ws.parent

    ' Definir punto de salida: debajo de la tabla de alertas
    Dim startRow As Long
    startRow = loAL.Range.Rows(loAL.Range.Rows.Count).Row + 2

    ' Limpieza: pivots debajo de startRow y charts del módulo
    DeleteChartsByPrefix ws, "GF_AL_"
    ClearPivotAreaFromRow ws, startRow

    Dim token As String
    token = SafePivotNameToken(ws.name)

    Dim ptNivelName As String, ptMontosName As String, ptTopName As String
    ptNivelName = "ptAL_Nivel_" & token
    ptMontosName = "ptAL_Montos_" & token
    ptTopName = "ptAL_Top_" & token

    ' PivotCache
    Dim pc As PivotCache
    Set pc = wb.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=loAL.Range)

    '=========================
    ' Pivot 1: CUCs por nivel
    '=========================
    Dim ptNivel As PivotTable
    Set ptNivel = pc.CreatePivotTable(TableDestination:=ws.Cells(startRow, 1), tableName:=ptNivelName)

    With ptNivel
        .ManualUpdate = True
        With .PivotFields("NIVEL_RIESGO")
            .Orientation = xlRowField
            .Position = 1
        End With
        .AddDataField .PivotFields("CUC"), "CUCs", xlCount
        .RowAxisLayout xlTabularRow
        .ManualUpdate = False
    End With

    '=========================
    ' Pivot 2: Montos por nivel
    '=========================
    Dim ptMontos As PivotTable
    Set ptMontos = pc.CreatePivotTable(TableDestination:=ws.Cells(startRow + 16, 1), tableName:=ptMontosName)

    With ptMontos
        .ManualUpdate = True
        With .PivotFields("NIVEL_RIESGO")
            .Orientation = xlRowField
            .Position = 1
        End With
        .AddDataField .PivotFields("SUMA_MONTOS"), "Suma montos", xlSum
        .RowAxisLayout xlTabularRow
        .ManualUpdate = False
    End With

    '=========================
    ' Pivot 3: Top 10 desviación
    '=========================
    Dim ptTop As PivotTable
    Set ptTop = pc.CreatePivotTable(TableDestination:=ws.Cells(startRow + 32, 1), tableName:=ptTopName)

    Dim df As PivotField
    With ptTop
        .ManualUpdate = True
        With .PivotFields("CUC")
            .Orientation = xlRowField
            .Position = 1
        End With

        Set df = .AddDataField(.PivotFields("DESVIACION_MEDIA_%"), "Desviación", xlMax)

        On Error Resume Next
        .PivotFields("CUC").AutoShow xlAutomatic, xlTop, 10, df.name
        .PivotFields("CUC").AutoSort xlDescending, df.name
        On Error GoTo 0

        .RowAxisLayout xlTabularRow
        .ManualUpdate = False
    End With

    '=========================
    ' Gráficos EN LA MISMA HOJA (a la derecha de los pivots)
    '=========================
    Dim left0 As Double, top0 As Double, w As Double, h As Double
    left0 = ws.Cells(startRow, 7).Left   ' Columna G
    top0 = ws.Cells(startRow, 7).top
    w = 520
    h = 240

    Dim co1 As ChartObject
    Set co1 = ws.ChartObjects.Add(left0, top0, w, h)
    co1.name = "GF_AL_CUCS"
    With co1.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=ptNivel.TableRange1
        .HasTitle = True
        .ChartTitle.Text = tituloBase & " | CUCs por nivel de riesgo"
        On Error Resume Next
        .Legend.Delete
        On Error GoTo 0
    End With

    Dim co2 As ChartObject
    Set co2 = ws.ChartObjects.Add(left0, top0 + h + 18, w, h)
    co2.name = "GF_AL_MONTOS"
    With co2.Chart
        .ChartType = xlColumnClustered
        .SetSourceData Source:=ptMontos.TableRange1
        .HasTitle = True
        .ChartTitle.Text = tituloBase & " | Suma de montos por nivel de riesgo"
        On Error Resume Next
        .Legend.Delete
        On Error GoTo 0
    End With

    Dim co3 As ChartObject
    Set co3 = ws.ChartObjects.Add(left0, top0 + (h + 18) * 2, w, h + 40)
    co3.name = "GF_AL_TOP10"
    With co3.Chart
        .ChartType = xlBarClustered
        .SetSourceData Source:=ptTop.TableRange1
        .HasTitle = True
        .ChartTitle.Text = tituloBase & " | Top 10 desviación media (%)"
        On Error Resume Next
        .Legend.Delete
        On Error GoTo 0
    End With

fin:
End Sub
