Option Explicit

Private Const PQ_NAME As String = "PQ_Clientes_SAB"
Private Const SH_NAME As String = "Clientes_SAB"
Private Const LO_NAME As String = "Clientes_SAB"

' ==========================
' Punto de entrada desde el formulario
' ==========================
Public Sub CrearQueryClientesSAB(ByVal rutaArchivo As String, Optional ByVal showMsgs As Boolean = True)
    On Error GoTo EH

    If Len(Trim$(rutaArchivo)) = 0 Then
        Err.Raise vbObjectError + 1, , "Ruta de archivo vacía."
    End If
    If Dir(rutaArchivo, vbNormal) = "" Then
        Err.Raise vbObjectError + 2, , "El archivo no existe: " & rutaArchivo
    End If

    Dim wb As Workbook
    Set wb = ThisWorkbook

    ' 1) Crear o actualizar la consulta de Power Query
    EnsurePQClientesSAB wb, rutaArchivo

    ' 2) Crear hoja y cargar datos en tabla vinculada a la consulta
    LoadClientesSAB wb, showMsgs

    Exit Sub

EH:
    If showMsgs Then
        MsgBox "Error al cargar Clientes SAB: " & Err.Number & " - " & Err.Description, vbCritical
    End If
End Sub

' ==========================
' Crea o actualiza la consulta PQ_Clientes_SAB
' ==========================
Private Sub EnsurePQClientesSAB(ByVal wb As Workbook, ByVal rutaArchivo As String)
    Dim m As String
    m = BuildMFormulaClientesSAB(rutaArchivo)

    Dim q As WorkbookQuery
    On Error Resume Next
    Set q = wb.Queries(PQ_NAME)
    On Error GoTo 0

    If q Is Nothing Then
        wb.Queries.Add name:=PQ_NAME, Formula:=m
    Else
        q.Formula = m
    End If
End Sub

' ==========================
' M de Power Query para Clientes SAB
' - Lee archivo tabulado
' - Promueve encabezados
' - Filtra filas sin "Cuenta" (elimina la fila "Número de Cuentas: 4920")
' ==========================
Private Function BuildMFormulaClientesSAB(ByVal rutaArchivo As String) As String
    Dim p As String
    p = Replace(rutaArchivo, """", """""")

    Dim m As String
    m = "let" & vbCrLf
    m = m & "    Ruta = """ & p & """," & vbCrLf
    m = m & "    Origen = File.Contents(Ruta)," & vbCrLf
    m = m & "    Csv = Csv.Document(Origen,[Delimiter=""#(tab)"",Encoding=1252,QuoteStyle=QuoteStyle.Csv])," & vbCrLf
    m = m & "    Prom = Table.PromoteHeaders(Csv,[PromoteAllScalars=true])," & vbCrLf
    m = m & "    Fil = Table.SelectRows(Prom, each " & _
                "not ( [Cuenta] = null or Text.Trim(Text.From([Cuenta])) = """" ))" & vbCrLf
    m = m & "in" & vbCrLf
    m = m & "    Fil"

    BuildMFormulaClientesSAB = m
End Function

' ==========================
' Crea hoja, ListObject y refresca desde PQ_Clientes_SAB
' ==========================
Private Sub LoadClientesSAB(ByVal wb As Workbook, ByVal showMsgs As Boolean)
    Dim sh As Worksheet
    Dim lo As ListObject
    Dim qt As QueryTable
    Dim conn As String

    ' 1) Hoja Clientes_SAB
    On Error Resume Next
    Set sh = wb.Worksheets(SH_NAME)
    On Error GoTo 0

    If sh Is Nothing Then
        Set sh = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        sh.name = SH_NAME
    Else
        sh.Cells.Clear
    End If

    ' 2) Limpiar objetos previos
    On Error Resume Next
    For Each lo In sh.ListObjects
        lo.Delete
    Next lo

    For Each qt In sh.QueryTables
        qt.Delete
    Next qt
    On Error GoTo 0

    ' 3) Crear ListObject vinculado a la consulta PQ_Clientes_SAB
    conn = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;" & _
           "Location=" & PQ_NAME & ";Extended Properties="""";"

    Set lo = sh.ListObjects.Add( _
        SourceType:=0, _
        Source:=conn, _
        Destination:=sh.Range("A1") _
    )

    With lo.QueryTable
        .CommandType = xlCmdSql
        .CommandText = "SELECT * FROM [" & PQ_NAME & "]"
        .BackgroundQuery = False
        .RefreshStyle = xlInsertDeleteCells
        .AdjustColumnWidth = True
        .PreserveColumnInfo = True
        .Refresh
    End With

    ' 4) Nombre y estilo de tabla
    lo.name = LO_NAME
    lo.TableStyle = "TableStyleLight14"

    ' 5) Dejar visible la hoja y posicionarse en A1
    sh.Activate
    sh.Range("A1").Select

    If showMsgs Then
        MsgBox "Clientes SAB cargados correctamente en la hoja '" & SH_NAME & "'.", vbInformation
    End If
End Sub


