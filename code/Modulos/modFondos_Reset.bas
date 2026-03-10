'==========================
' modFondos_Reset
' Elimina hojas y consultas generadas por el proceso de Fondos.
' Conserva: Clientes_Fondos y hojas no generadas por el proceso.
'==========================
Option Explicit

'======================
' Estado Application (local al módulo)
'======================
Private mRstFrozen As Boolean
Private mRstPrevScreenUpdating As Boolean
Private mRstPrevEnableEvents As Boolean
Private mRstPrevDisplayAlerts As Boolean
Private mRstPrevCalculation As XlCalculation
Private mRstPrevStatusBar As Variant

Private Sub RstAppFreeze(ByVal freeze As Boolean)
    On Error Resume Next
    With Application
        If freeze Then
            If Not mRstFrozen Then
                mRstPrevScreenUpdating = .ScreenUpdating
                mRstPrevEnableEvents = .EnableEvents
                mRstPrevDisplayAlerts = .DisplayAlerts
                mRstPrevCalculation = .Calculation
                mRstPrevStatusBar = .StatusBar
                mRstFrozen = True
            End If
            .ScreenUpdating = False
            .EnableEvents = False
            .DisplayAlerts = False
            .Calculation = xlCalculationManual
        Else
            If mRstFrozen Then
                .ScreenUpdating = mRstPrevScreenUpdating
                .EnableEvents = mRstPrevEnableEvents
                .DisplayAlerts = mRstPrevDisplayAlerts
                .Calculation = mRstPrevCalculation
                .StatusBar = mRstPrevStatusBar
                mRstFrozen = False
            Else
                .StatusBar = False
            End If
        End If
    End With
    On Error GoTo 0
End Sub

'======================
' Identificacion de hojas generadas por el proceso
'======================
Private Function EsHojaGenerada(ByVal nm As String) As Boolean
    Dim u As String
    u = UCase$(Trim$(nm))
    EsHojaGenerada = False

    ' Hojas de trabajo temporales
    Select Case u
        Case "RAW_WORK", "MAIN_WORK", "ALERTAS_WORK", "AUX_WORK", "CHARTS_WORK"
            EsHojaGenerada = True
            Exit Function
    End Select

    ' Patron: contiene _SUS_ o _RES_ (hojas renombradas por CrearQueryFondos)
    Dim tieneSus As Boolean, tieneRes As Boolean
    tieneSus = InStr(1, u, "_SUS_", vbBinaryCompare) > 0
    tieneRes = InStr(1, u, "_RES_", vbBinaryCompare) > 0

    If Not (tieneSus Or tieneRes) Then Exit Function

    ' RAW_FONDOS_SUS_... / RAW_FONDOS_RES_...
    If Left$(u, 4) = "RAW_" Then
        EsHojaGenerada = True
        Exit Function
    End If

    ' FONDOS_SUS_... / FONDOS_RES_...
    If Left$(u, 7) = "FONDOS_" Then
        EsHojaGenerada = True
        Exit Function
    End If

    ' *_ALERTAS_* con _SUS_ o _RES_
    If InStr(1, u, "_ALERTAS_", vbBinaryCompare) > 0 Then
        EsHojaGenerada = True
        Exit Function
    End If

    ' AUX_*_SUS_* / AUX_*_RES_*
    If Left$(u, 4) = "AUX_" Then
        EsHojaGenerada = True
        Exit Function
    End If

    ' *_GRAFICOS_*_SUS_* / *_GRAFICOS_*_RES_*
    If InStr(1, u, "_GRAFICOS_", vbBinaryCompare) > 0 Then
        EsHojaGenerada = True
        Exit Function
    End If
End Function

'======================
' Eliminar una conexion por todos sus nombres posibles
'======================
Private Sub EliminarConexion(ByVal wb As Workbook, ByVal queryName As String)
    Dim candidatos As Variant
    Dim i As Long
    candidatos = Array( _
        "Consulta - " & queryName, _
        "Query - " & queryName, _
        "PQ_" & queryName, _
        queryName)

    For i = LBound(candidatos) To UBound(candidatos)
        On Error Resume Next
        wb.Connections(CStr(candidatos(i))).Delete
        On Error GoTo 0
    Next i
End Sub

'======================
' Eliminar consultas y conexiones de Fondos
'======================
Private Sub EliminarConsultasFondos(ByVal wb As Workbook, ByRef log As String)
    Dim nombres As Variant
    Dim i As Long
    nombres = Array("RAW_SUS", "SUS", "SUS_ALERTAS", "RAW_RES", "RES", "RES_ALERTAS")

    For i = LBound(nombres) To UBound(nombres)
        Dim qn As String
        qn = CStr(nombres(i))

        On Error Resume Next
        wb.Queries.Item(qn).Delete
        If Err.Number = 0 Then
            log = log & "  Consulta eliminada: " & qn & vbCrLf
        End If
        Err.Clear
        On Error GoTo 0

        EliminarConexion wb, qn
    Next i
End Sub

'======================
' Construir lista de hojas a eliminar
'======================
Private Function ListarHojasAEliminar(ByVal wb As Workbook) As Collection
    Dim col As New Collection
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If EsHojaGenerada(ws.name) Then
            col.Add ws.name
        End If
    Next ws
    Set ListarHojasAEliminar = col
End Function

'======================
' Construir lista de consultas a eliminar
'======================
Private Function ListarConsultasAEliminar(ByVal wb As Workbook) As Collection
    Dim col As New Collection
    Dim nombres As Variant
    Dim i As Long
    nombres = Array("RAW_SUS", "SUS", "SUS_ALERTAS", "RAW_RES", "RES", "RES_ALERTAS")

    For i = LBound(nombres) To UBound(nombres)
        Dim qn As String
        qn = CStr(nombres(i))
        On Error Resume Next
        Dim dummy As Object
        Set dummy = wb.Queries.Item(qn)
        If Err.Number = 0 Then col.Add qn
        Err.Clear
        On Error GoTo 0
    Next i
    Set ListarConsultasAEliminar = col
End Function

'======================
' Construir texto de confirmacion
'======================
Private Function ArmarTextoConfirmacion(ByVal hojas As Collection, ByVal consultas As Collection) As String
    Dim txt As String
    txt = "Se eliminarán los siguientes elementos:" & vbCrLf & vbCrLf

    If hojas.Count > 0 Then
        txt = txt & "HOJAS (" & hojas.Count & "):" & vbCrLf
        Dim nm As Variant
        For Each nm In hojas
            txt = txt & "  - " & CStr(nm) & vbCrLf
        Next nm
    Else
        txt = txt & "HOJAS: ninguna que eliminar." & vbCrLf
    End If

    txt = txt & vbCrLf

    If consultas.Count > 0 Then
        txt = txt & "CONSULTAS PQ (" & consultas.Count & "):" & vbCrLf
        Dim qn As Variant
        For Each qn In consultas
            txt = txt & "  - " & CStr(qn) & vbCrLf
        Next qn
    Else
        txt = txt & "CONSULTAS PQ: ninguna que eliminar." & vbCrLf
    End If

    txt = txt & vbCrLf & "Se conservará: Clientes_Fondos y demás hojas no generadas por el proceso." & vbCrLf & vbCrLf
    txt = txt & "¿Confirmar eliminación?"

    ArmarTextoConfirmacion = txt
End Function

'======================
' Punto de entrada publico
'======================
Public Sub ResetFondos()
    On Error GoTo EH

    Dim wb As Workbook
    Set wb = ThisWorkbook

    ' Inventario previo a la confirmacion
    Dim hojas As Collection
    Dim consultas As Collection
    Set hojas = ListarHojasAEliminar(wb)
    Set consultas = ListarConsultasAEliminar(wb)

    If hojas.Count = 0 And consultas.Count = 0 Then
        MsgBox "No se encontraron hojas ni consultas generadas por el proceso de Fondos." & vbCrLf & _
               "No hay nada que eliminar.", vbInformation, "Reset Fondos"
        Exit Sub
    End If

    ' Confirmacion
    Dim txtConf As String
    txtConf = ArmarTextoConfirmacion(hojas, consultas)

    If MsgBox(txtConf, vbQuestion + vbYesNo + vbDefaultButton2, "Reset Fondos - Confirmar") = vbNo Then
        MsgBox "Operacion cancelada. No se realizaron cambios.", vbInformation, "Reset Fondos"
        Exit Sub
    End If

    RstAppFreeze True

    Dim log As String
    Dim errores As String
    log = vbNullString
    errores = vbNullString

    ' Eliminar hojas
    Dim i As Long
    For i = wb.Worksheets.Count To 1 Step -1
        Dim ws As Worksheet
        Set ws = wb.Worksheets(i)
        If EsHojaGenerada(ws.name) Then
            Dim nmHoja As String
            nmHoja = ws.name
            On Error Resume Next
            ws.Delete
            If Err.Number = 0 Then
                log = log & "  Hoja eliminada: " & nmHoja & vbCrLf
            Else
                errores = errores & "  No se pudo eliminar hoja '" & nmHoja & "': " & Err.Description & vbCrLf
                Err.Clear
            End If
            On Error GoTo EH
        End If
    Next i

    ' Eliminar consultas y conexiones
    EliminarConsultasFondos wb, log

    RstAppFreeze False

    ' Reporte final
    Dim resumen As String
    If Len(errores) = 0 Then
        resumen = "Reset completado exitosamente." & vbCrLf & vbCrLf
        If Len(log) > 0 Then
            resumen = resumen & "Elementos eliminados:" & vbCrLf & log
        Else
            resumen = resumen & "No se eliminaron elementos (ya estaban limpios)."
        End If
        MsgBox resumen, vbInformation, "Reset Fondos"
    Else
        resumen = "Reset completado con advertencias." & vbCrLf & vbCrLf
        If Len(log) > 0 Then
            resumen = resumen & "Eliminados correctamente:" & vbCrLf & log & vbCrLf
        End If
        resumen = resumen & "Errores encontrados:" & vbCrLf & errores
        MsgBox resumen, vbExclamation, "Reset Fondos"
    End If

    Exit Sub

EH:
    Dim errDesc As String
    errDesc = Err.Description
    If Len(Trim$(errDesc)) = 0 Then errDesc = "(sin descripcion)"

    RstAppFreeze False
    MsgBox "ResetFondos fallo." & vbCrLf & "Error " & Err.Number & vbCrLf & errDesc, _
           vbCritical, "Reset Fondos"
End Sub
