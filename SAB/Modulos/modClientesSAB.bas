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

    ' Validar formato antes de continuar
    Dim motivo As String
    If Not ValidarFormatoClientesSAB(rutaArchivo, motivo) Then
        If showMsgs Then
            MsgBox "El archivo no tiene el formato esperado de Clientes SAB." & vbCrLf & vbCrLf & _
                   motivo & vbCrLf & vbCrLf & _
                   "Se requieren al menos las columnas: Cuenta, RUC/NIT.", _
                   vbExclamation, "Formato incorrecto"
        End If
        Exit Sub
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
    p = Replace(rutaArchivo, """", """""""")

    Dim m As String
    m = m & "let" & vbCrLf
    m = m & "    Ruta   = """ & p & """," & vbCrLf
    m = m & "    Origen = File.Contents(Ruta)," & vbCrLf
    m = m & "    Csv    = Csv.Document(Origen,[Delimiter=""#(tab)"",Encoding=1252,QuoteStyle=QuoteStyle.Csv])," & vbCrLf
    m = m & "    Prom   = Table.PromoteHeaders(Csv,[PromoteAllScalars=true])," & vbCrLf
    m = m & "    AllText = Table.TransformColumnTypes(Prom," & vbCrLf
    m = m & "        List.Transform(Table.ColumnNames(Prom), each {_, type text}))," & vbCrLf
    m = m & "    CleanTxt = (s as any) as text =>" & vbCrLf
    m = m & "        let" & vbCrLf
    m = m & "            t0 = if s = null then """" else Text.From(s)," & vbCrLf
    m = m & "            t1 = Text.Replace(t0, Character.FromNumber(160), """")," & vbCrLf
    m = m & "            t2 = Text.Trim(t1)" & vbCrLf
    m = m & "        in t2," & vbCrLf
    m = m & "    TrimCol = (t as table, col as text) as table =>" & vbCrLf
    m = m & "        if List.Contains(Table.ColumnNames(t), col)" & vbCrLf
    m = m & "        then Table.TransformColumns(t, {{col, each CleanTxt(_), type text}})" & vbCrLf
    m = m & "        else t," & vbCrLf
    m = m & "    T1 = TrimCol(AllText, ""Cuenta"")," & vbCrLf
    m = m & "    T2 = TrimCol(T1,      ""RUC/NIT"")," & vbCrLf
    m = m & "    T3 = TrimCol(T2,      ""Tipo"")," & vbCrLf
    m = m & "    T4 = TrimCol(T3,      ""Nombre"")," & vbCrLf
    m = m & "    Fil = Table.SelectRows(T4, each" & vbCrLf
    m = m & "        not ([Cuenta] = null or Text.Trim([Cuenta]) = """"))," & vbCrLf
    m = m & "    MesNum = (mm as text) as number =>" & vbCrLf
    m = m & "        let m2 = Text.Upper(Text.Trim(mm))" & vbCrLf
    m = m & "        in      if m2 = ""ENE"" then 1" & vbCrLf
    m = m & "           else if m2 = ""FEB"" then 2" & vbCrLf
    m = m & "           else if m2 = ""MAR"" then 3" & vbCrLf
    m = m & "           else if m2 = ""ABR"" then 4" & vbCrLf
    m = m & "           else if m2 = ""MAY"" then 5" & vbCrLf
    m = m & "           else if m2 = ""JUN"" then 6" & vbCrLf
    m = m & "           else if m2 = ""JUL"" then 7" & vbCrLf
    m = m & "           else if m2 = ""AGO"" then 8" & vbCrLf
    m = m & "           else if m2 = ""SET"" or m2 = ""SEP"" then 9" & vbCrLf
    m = m & "           else if m2 = ""OCT"" then 10" & vbCrLf
    m = m & "           else if m2 = ""NOV"" then 11" & vbCrLf
    m = m & "           else if m2 = ""DIC"" then 12" & vbCrLf
    m = m & "           else 0," & vbCrLf
    m = m & "    ParseFecha = (s as any) as any =>" & vbCrLf
    m = m & "        let" & vbCrLf
    m = m & "            t  = if s = null then """" else Text.Trim(Text.From(s))," & vbCrLf
    m = m & "            n  = Text.Length(t)," & vbCrLf
    m = m & "            ok = n >= 8," & vbCrLf
    m = m & "            dd  = if ok then try Number.From(Text.Start(t, 2)) otherwise 0 else 0," & vbCrLf
    m = m & "            mmm = if ok then Text.Middle(t, 2, 3) else """"," & vbCrLf
    m = m & "            yrT = if ok then Text.End(t, n - 5) else """"," & vbCrLf
    m = m & "            yrN = try Number.From(yrT) otherwise 0," & vbCrLf
    m = m & "            yr  = if yrN < 100 then yrN + 2000 else yrN," & vbCrLf
    m = m & "            mes = if ok then MesNum(mmm) else 0," & vbCrLf
    m = m & "            res = if ok and dd > 0 and mes > 0 and yr > 1900" & vbCrLf
    m = m & "                  then try #date(yr, mes, dd) otherwise null" & vbCrLf
    m = m & "                  else null" & vbCrLf
    m = m & "        in res," & vbCrLf
    m = m & "    HasCol = (t as table, col as text) as logical =>" & vbCrLf
    m = m & "        List.Contains(Table.ColumnNames(t), col)," & vbCrLf
    m = m & "    F1 = if HasCol(Fil, ""Fecha de Ingreso"")" & vbCrLf
    m = m & "         then Table.TransformColumns(Fil,   {{""Fecha de Ingreso"", each ParseFecha(_), type date}})" & vbCrLf
    m = m & "         else Fil," & vbCrLf
    m = m & "    F2 = if HasCol(F1,  ""Fecha de Bloqueo"")" & vbCrLf
    m = m & "         then Table.TransformColumns(F1,    {{""Fecha de Bloqueo"", each ParseFecha(_), type date}})" & vbCrLf
    m = m & "         else F1," & vbCrLf
    m = m & "    Result = F2" & vbCrLf
    m = m & "in" & vbCrLf
    m = m & "    Result" & vbCrLf

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
        Set sh = wb.Worksheets.Add(After:=wb.Sheets(wb.Sheets.count))
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
            .BackgroundQuery = False
            .RefreshStyle = xlOverwriteCells
            .AdjustColumnWidth = True
            .PreserveColumnInfo = True
            SAB_SetPQRefreshing True
            .Refresh BackgroundQuery:=False
            SAB_SetPQRefreshing False
        End With
    Application.CalculateUntilAsyncQueriesDone
    On Error GoTo 0

    ' 5) Nombre y estilo
    On Error Resume Next
    lo.Name = LO_NAME
    lo.TableStyle = "TableStyleLight14"
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

' ==========================
' Valida que el archivo TSV tenga el formato minimo esperado.
' Devuelve True si es valido; False con motivo en el parametro de salida.
' ==========================
Private Function ValidarFormatoClientesSAB(ByVal rutaArchivo As String, _
                                            ByRef motivo As String) As Boolean
    ValidarFormatoClientesSAB = False
    motivo = ""

    Dim fileNum As Integer
    fileNum = FreeFile
    Dim primeraLinea As String

    On Error GoTo errLectura
    Open rutaArchivo For Input Access Read As #fileNum
    If Not EOF(fileNum) Then Line Input #fileNum, primeraLinea
    Close #fileNum
    On Error GoTo 0

    If Len(Trim$(primeraLinea)) = 0 Then
        motivo = "El archivo est" & Chr(225) & " vac" & Chr(237) & "o o no tiene encabezados."
        Exit Function
    End If

    ' Normalizar: separador puede ser tab
    Dim cols() As String
    cols = Split(primeraLinea, vbTab)

    ' Construir set de columnas presentes (en mayusculas, sin espacios extra)
    Dim dCols As Object: Set dCols = CreateObject("Scripting.Dictionary")
    Dim k As Long
    For k = 0 To UBound(cols)
        Dim colNorm As String
        colNorm = UCase$(Trim$(Replace(cols(k), Chr(160), "")))
        If Len(colNorm) > 0 Then
            If Not dCols.exists(colNorm) Then dCols.Add colNorm, True
        End If
    Next k

    ' Columnas minimas requeridas
    Dim faltantes As String: faltantes = ""
    If Not dCols.exists("CUENTA") Then faltantes = faltantes & "  - Cuenta" & vbCrLf
    If Not dCols.exists("RUC/NIT") And Not dCols.exists("RUCNIT") Then
        faltantes = faltantes & "  - RUC/NIT" & vbCrLf
    End If

    If Len(faltantes) > 0 Then
        motivo = "Columnas faltantes:" & vbCrLf & faltantes & _
                 "Columnas encontradas: " & Join(cols, ", ")
        Exit Function
    End If

    ValidarFormatoClientesSAB = True
    Exit Function

errLectura:
    Close #fileNum
    motivo = "No se pudo leer el archivo: " & Err.Description
End Function