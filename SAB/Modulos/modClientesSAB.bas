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
        Err.Raise vbObjectError + 1, , "Ruta de archivo vac" & Chr(237) & "a."
    End If
    If Dir(rutaArchivo, vbNormal) = "" Then
        Err.Raise vbObjectError + 2, , "El archivo no existe: " & rutaArchivo
    End If

    Dim wb As Workbook
    Set wb = ThisWorkbook

    ' 1) Limpiar conexion anterior antes de recrear la consulta
    PurgeOldConnection wb

    ' 2) Crear o actualizar la consulta de Power Query
    EnsurePQClientesSAB wb, rutaArchivo

    ' 3) Crear hoja y cargar datos en tabla vinculada a la consulta
    LoadClientesSAB wb, showMsgs

    Exit Sub

EH:
    If showMsgs Then
        MsgBox "Error al cargar Clientes SAB: " & Err.Number & " - " & Err.Description, vbCritical
    End If
End Sub

' ==========================
' Elimina la conexion OLEDB huerfana si existe
' ==========================
Private Sub PurgeOldConnection(ByVal wb As Workbook)
    Dim conn As WorkbookConnection
    On Error Resume Next
    Set conn = wb.Connections("PQ_" & PQ_NAME)
    If Not conn Is Nothing Then conn.Delete
    Set conn = wb.Connections(PQ_NAME)
    If Not conn Is Nothing Then conn.Delete
    On Error GoTo 0
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
        wb.Queries.Add Name:=PQ_NAME, Formula:=m
    Else
        q.Formula = m
    End If
End Sub

' ==========================
' M de Power Query para Clientes SAB
' - Lee archivo tabulado (TSV, encoding 1252)
' - Promueve encabezados
' - Trim a columnas clave: Cuenta, RUC/NIT, Tipo, Nombre
' - Filtra filas sin Cuenta valida
' ==========================
Private Function BuildMFormulaClientesSAB(ByVal rutaArchivo As String) As String
    Dim p As String
    p = Replace(rutaArchivo, """", """""")

    Dim m As String
    m = "let" & vbCrLf
    m = m & "    Ruta   = """ & p & """," & vbCrLf
    m = m & "    Origen = File.Contents(Ruta)," & vbCrLf
    m = m & "    Csv    = Csv.Document(Origen,[Delimiter=""#(tab)"",Encoding=1252,QuoteStyle=QuoteStyle.Csv])," & vbCrLf
    m = m & "    Prom   = Table.PromoteHeaders(Csv,[PromoteAllScalars=true])," & vbCrLf
    m = m & "    Fil    = Table.SelectRows(Prom, each" & vbCrLf
    m = m & "               not ([Cuenta] = null or Text.Trim(Text.From([Cuenta])) = """"))," & vbCrLf
    m = m & "    Cols   = Table.ColumnNames(Fil)," & vbCrLf
    m = m & "    TrimCol = (t as table, col as text) as table =>" & vbCrLf
    m = m & "        if List.Contains(Table.ColumnNames(t), col)" & vbCrLf
    m = m & "        then Table.TransformColumns(t, {{col, each Text.Trim(Text.From(_)), type text}})" & vbCrLf
    m = m & "        else t," & vbCrLf
    m = m & "    T1 = TrimCol(Fil,   ""Cuenta"")," & vbCrLf
    m = m & "    T2 = TrimCol(T1,    ""RUC/NIT"")," & vbCrLf
    m = m & "    T3 = TrimCol(T2,    ""Tipo"")," & vbCrLf
    m = m & "    T4 = TrimCol(T3,    ""Nombre"")," & vbCrLf
    m = m & "    Result = T4" & vbCrLf
    m = m & "in" & vbCrLf
    m = m & "    Result"

    BuildMFormulaClientesSAB = m
End Function

' ==========================
' Crea hoja, ListObject y refresca desde PQ_Clientes_SAB
' ==========================
Private Sub LoadClientesSAB(ByVal wb As Workbook, ByVal showMsgs As Boolean)
    On Error GoTo EH_Load
    Dim sh As Worksheet
    Dim lo As ListObject
    Dim qt As QueryTable

    ' 1) Hoja Clientes_SAB
    On Error Resume Next
    Set sh = wb.Worksheets(SH_NAME)
    On Error GoTo 0

    If sh Is Nothing Then
        Set sh = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        sh.Name = SH_NAME
    Else
        sh.Cells.Clear
    End If

    ' 2) Limpiar ListObjects y QueryTables previos en la hoja
    On Error Resume Next
    For Each lo In sh.ListObjects:  lo.Delete:  Next lo
    For Each qt In sh.QueryTables:  qt.Delete:  Next qt
    On Error GoTo 0

    ' 3) Conexion OLEDB hacia la consulta PQ
    Dim connStr As String
    connStr = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;" & _
              "Location=" & PQ_NAME & ";Extended Properties="""";"

    Dim conn As WorkbookConnection
    On Error Resume Next
    Set conn = wb.Connections.Add2( _
        "PQ_" & PQ_NAME, "", connStr, _
        "SELECT * FROM [" & PQ_NAME & "]", xlCmdSql)
    On Error GoTo 0

    ' 4) ListObject vinculado a la conexion (xlSrcExternal = 1)
    Set lo = sh.ListObjects.Add( _
        SourceType:=xlSrcExternal, _
        Source:=conn, _
        LinkSource:=True, _
        XlListObjectHasHeaders:=xlYes, _
        Destination:=sh.Range("A1"))

    On Error Resume Next
    With lo.QueryTable
        .BackgroundQuery    = False
        .RefreshStyle       = xlOverwriteCells
        .AdjustColumnWidth  = True
        .PreserveColumnInfo = True
        .Refresh BackgroundQuery:=False
    End With
    Application.CalculateUntilAsyncQueriesDone
    On Error GoTo 0

    ' 5) Nombre y estilo
    On Error Resume Next
    lo.Name         = LO_NAME
    lo.TableStyle   = "TableStyleLight14"
    On Error GoTo 0

    ' 6) Posicion
    sh.Activate
    sh.Range("A1").Select

    If showMsgs Then
        MsgBox "Clientes SAB cargados correctamente en la hoja '" & SH_NAME & "'.", vbInformation
    End If
    Exit Sub

EH_Load:
    Dim errMsg As String
    errMsg = "Error al cargar datos desde Power Query." & vbCrLf & vbCrLf & _
             "Numero: " & Err.Number & vbCrLf & _
             "Descripcion: " & Err.Description & vbCrLf & vbCrLf & _
             "Verifique que el archivo exista, que la ruta sea correcta " & _
             "y que Power Query pueda acceder al archivo."
    If showMsgs Then MsgBox errMsg, vbCritical, "Error - Clientes SAB"
End Sub
