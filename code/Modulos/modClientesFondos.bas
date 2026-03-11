'==========================
' modClientesFondos (producción)
'==========================
Option Explicit

'======================
' Estado Application
'======================
Private mAppFrozen As Boolean
Private mPrevScreenUpdating As Boolean
Private mPrevEnableEvents As Boolean
Private mPrevDisplayAlerts As Boolean
Private mPrevCalculation As XlCalculation
Private mPrevStatusBar As Variant

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
End Sub

'======================
' Helpers
'======================
Private Sub MLine(ByRef BUF As String, ByVal s As String)
    If Len(BUF) = 0 Then
        BUF = s
    Else
        BUF = BUF & vbCrLf & s
    End If
End Sub

Private Function EnsureSheet(ByVal nm As String) As Worksheet
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If sh Is Nothing Then
        Set sh = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        sh.name = nm
    End If
    Set EnsureSheet = sh
End Function

Private Sub ClearSheetHard(ByVal sh As Worksheet)
    Dim lo As ListObject, qt As QueryTable, pt As PivotTable, co As ChartObject, shp As Shape
    On Error Resume Next
    For Each pt In sh.PivotTables: pt.TableRange2.Clear: Next pt
    For Each co In sh.ChartObjects: co.Delete: Next co
    For Each lo In sh.ListObjects: lo.Delete: Next lo
    For Each qt In sh.QueryTables: qt.Delete: Next qt
    For Each shp In sh.Shapes: shp.Delete: Next shp
    sh.Cells.Clear
    On Error GoTo 0
End Sub

Private Function EnsurePQConnection(ByVal qName As String, ByVal connName As String) As WorkbookConnection
    Dim cn As WorkbookConnection
    Dim conStr As String, cmdText As String

    conStr = "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & qName & ";Extended Properties="""""  ' engine M
    cmdText = "SELECT * FROM [" & qName & "]"

    On Error Resume Next
    Set cn = ThisWorkbook.Connections(connName)
    On Error GoTo 0

    If cn Is Nothing Then
        On Error Resume Next
        Set cn = ThisWorkbook.Connections.Add2( _
                    name:=connName, _
                    Description:="Conexión " & qName, _
                    ConnectionString:=conStr, _
                    CommandText:=cmdText, _
                    lCmdtype:=xlCmdSql)
        If cn Is Nothing Then
            Set cn = ThisWorkbook.Connections.Add( _
                        name:=connName, _
                        Description:="Conexión " & qName, _
                        ConnectionString:=conStr, _
                        CommandText:=cmdText, _
                        lCmdtype:=xlCmdSql)
        End If
        On Error GoTo 0
    Else
        On Error Resume Next
        If cn.Type = xlConnectionTypeOLEDB Then
            With cn.OLEDBConnection
                .Connection = conStr
                .CommandText = cmdText
                .CommandType = xlCmdSql
                .BackgroundQuery = False
            End With
        End If
        On Error GoTo 0
    End If

    Set EnsurePQConnection = cn
End Function

Private Sub RefreshListObject(ByVal lo As ListObject)
    On Error Resume Next

    If lo Is Nothing Then Exit Sub

    If Not lo.QueryTable Is Nothing Then
        lo.QueryTable.BackgroundQuery = False
        lo.QueryTable.RefreshStyle = xlOverwriteCells
        lo.QueryTable.AdjustColumnWidth = True
        lo.QueryTable.PreserveColumnInfo = True
        lo.QueryTable.Refresh BackgroundQuery:=False
        If Err.Number <> 0 Then
            Err.Clear
            lo.Refresh
        End If
        Do While lo.QueryTable.refreshing
            DoEvents
        Loop
    Else
        lo.Refresh
    End If

    On Error GoTo 0
End Sub

'======================
' Principal
'======================
Public Sub CrearQueryClientesFondos(ByVal rutaArchivo As String, Optional ByVal activarHoja As Boolean = True)
    On Error GoTo EH
    SafeApp True

    Dim wb As Workbook
    Set wb = ThisWorkbook

    Dim q As WorkbookQuery
    Dim wsOut As Worksheet
    Dim lo As ListObject
    Dim m As String
    Dim rutaEsc As String
    Dim cn As WorkbookConnection
    Dim cnName As String
    Dim qName As String

    qName = "PQ_Clientes_Fondos"
    cnName = "PQ_Clientes_Fondos_Conn"

    rutaEsc = Replace(rutaArchivo, """", """""""") ' escapar comillas para M

    '==========================
    ' Código M (robusto)
    ' - No asume Origen{0}
    ' - Remueve primera columna de forma genérica
    ' - Tipado solo si existen columnas
    ' - Fecha de nacimiento null-safe
    '==========================
    MLine m, "let"
    MLine m, "  Ruta = """ & rutaEsc & ""","
    MLine m, "  Origen = Excel.Workbook(File.Contents(Ruta), null, true),"
    MLine m, "  Candidatos = Table.SelectRows(Origen, each [Data] <> null),"
    MLine m, "  ConTabla = Table.AddColumn(Candidatos, ""__T"", each"
    MLine m, "    let"
    MLine m, "      t0 = [Data],"
    MLine m, "      t1 = Table.Skip(t0, 9),"
    MLine m, "      cn = Table.ColumnNames(t1),"
    MLine m, "      t2 = if List.Count(cn) > 0 then Table.RemoveColumns(t1, {List.First(cn)}, MissingField.Ignore) else t1,"
    MLine m, "      t3 = Table.PromoteHeaders(t2, [PromoteAllScalars=true])"
    MLine m, "    in"
    MLine m, "      t3, type table),"
    MLine m, "  Filtrados = Table.SelectRows(ConTabla, each try Table.HasColumns([__T], {""CUC""}) otherwise false),"
    MLine m, "  Base = if Table.RowCount(Filtrados) > 0 then Filtrados{0}[__T] else ConTabla{0}[__T],"
    MLine m, "  Types0 = {"
    MLine m, "    {""CUC"", type text},"
    MLine m, "    {""NOMBRE"", type text},"
    MLine m, "    {""APELLIDO PATERNO"", type text},"
    MLine m, "    {""APELLIDO MATERNO"", type text},"
    MLine m, "    {""ESTADO"", type text},"
    MLine m, "    {""TIPO DE DOCUMENTO"", type text},"
    MLine m, "    {""NUMERO DE DOCUMENTO"", type text},"
    MLine m, "    {""PAIS"", type text},"
    MLine m, "    {""CATEGORIA"", type text},"
    MLine m, "    {""NO PUBLICIDAD"", type text},"
    MLine m, "    {""TIPO PARTICIPE"", type text},"
    MLine m, "    {""SEGMENTO"", type text},"
    MLine m, "    {""CANAL"", type text},"
    MLine m, "    {""TIPO PERSONA"", type text}"
    MLine m, "  },"
    MLine m, "  Types = List.Select(Types0, each List.Contains(Table.ColumnNames(Base), _{0})),"
    MLine m, "  Tipadas = if List.Count(Types) > 0 then Table.TransformColumnTypes(Base, Types, ""es-PE"") else Base,"
    MLine m, "  FechaOk = if List.Contains(Table.ColumnNames(Tipadas), ""FECHA DE NACIMIENTO"") then"
    MLine m, "              Table.TransformColumns(Tipadas, {{""FECHA DE NACIMIENTO"", each"
    MLine m, "                try Date.From(_) otherwise"
    MLine m, "                try Date.FromText(Text.From(_), ""es-PE"") otherwise"
    MLine m, "                try Date.FromText(Text.From(_), ""en-US"") otherwise null, type date}})"
    MLine m, "            else"
    MLine m, "              Tipadas,"
    MLine m, "  Filas = if List.Contains(Table.ColumnNames(FechaOk), ""CUC"") then"
    MLine m, "            Table.SelectRows(FechaOk, each [CUC] <> null and Text.Trim(Text.From([CUC])) <> """")"
    MLine m, "          else"
    MLine m, "            FechaOk"
    MLine m, "in"
    MLine m, "  Filas"

    '==========================
    ' Crear o actualizar Query
    '==========================
    On Error Resume Next
    Set q = wb.Queries(qName)
    On Error GoTo EH

    If q Is Nothing Then
        wb.Queries.Add name:=qName, Formula:=m
    Else
        q.Formula = m
    End If

    '==========================
    ' Conexión PQ
    '==========================
    Set cn = EnsurePQConnection(qName, cnName)
    If cn Is Nothing Then
        Err.Raise vbObjectError + 5002, "CrearQueryClientesFondos", "No se pudo crear la conexión para " & qName & "."
    End If

    '==========================
    ' Hoja de salida
    '==========================
    Set wsOut = EnsureSheet("Clientes_Fondos")
    ClearSheetHard wsOut

    '==========================
    ' Crear tabla vinculada
    '==========================
    Set lo = wsOut.ListObjects.Add(SourceType:=xlSrcExternal, Source:=cn, Destination:=wsOut.Range("A1"))
    lo.name = "Clientes_Fondos"
    lo.TableStyle = "TableStyleLight14"

    RefreshListObject lo

    If activarHoja Then
        wsOut.Activate
        wsOut.Range("A1").Select
    End If

    SafeApp False
    Exit Sub

EH:
    Application.StatusBar = False
    SafeApp False
    MsgBox "Error: " & Err.Description, vbCritical
End Sub


