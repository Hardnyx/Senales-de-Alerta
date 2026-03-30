'==========================
' modReset
' Elimina hojas y consultas generadas por el proceso activo.
' Cubre: Fondos (SUS/RES), SAB Movimiento de Caja (MC), SAB Cambio de Moneda (CM).
' Las hojas no generadas por el proceso se conservan.
'==========================
Option Explicit

'======================
' Estado Application
'======================
Private mRstFrozen             As Boolean
Private mRstPrevScreenUpdating As Boolean
Private mRstPrevEnableEvents   As Boolean
Private mRstPrevDisplayAlerts  As Boolean
Private mRstPrevCalculation    As XlCalculation
Private mRstPrevStatusBar      As Variant

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

    ' --- Hojas de trabajo temporales (cualquier proceso) ---
    Select Case u
        Case "RAW_WORK", "MAIN_WORK", "ALERTAS_WORK", "AUX_WORK", "CHARTS_WORK", _
             "SAB_MC_RAW_WORK", "SAB_MC_MAIN_WORK", _
             "SAB_MC_AL_DEP_WORK", "SAB_MC_AL_RET_WORK", _
             "SAB_CM_RAW_WORK", "SAB_CM_MAIN_WORK", _
             "SAB_CM_AL_COM_WORK", "SAB_CM_AL_VEN_WORK"
            EsHojaGenerada = True
            Exit Function
    End Select

    ' --- Fondos: hojas con _SUS_ o _RES_ en el nombre ---
    Dim tieneSus As Boolean: tieneSus = InStr(1, u, "_SUS_", vbBinaryCompare) > 0
    Dim tieneRes As Boolean: tieneRes = InStr(1, u, "_RES_", vbBinaryCompare) > 0

    If tieneSus Or tieneRes Then
        If Left$(u, 4) = "RAW_" Then EsHojaGenerada = True: Exit Function
        If Left$(u, 7) = "FONDOS_" Then EsHojaGenerada = True: Exit Function
        If InStr(1, u, "_ALERTAS_", vbBinaryCompare) > 0 Then EsHojaGenerada = True: Exit Function
        If Left$(u, 4) = "AUX_" Then EsHojaGenerada = True: Exit Function
        If InStr(1, u, "_GRAFICOS_", vbBinaryCompare) > 0 Then EsHojaGenerada = True: Exit Function
        Exit Function
    End If

    ' --- SAB MC: hojas con prefijo SAB_MC_ ---
    If Left$(u, 7) = "SAB_MC_" Then
        EsHojaGenerada = True
        Exit Function
    End If

    ' --- SAB CM: hojas con prefijo SAB_CM_ ---
    If Left$(u, 7) = "SAB_CM_" Then
        EsHojaGenerada = True
        Exit Function
    End If
    ' --- Hojas helper de graficos (ocultas xlSheetVeryHidden) ---
    If Left$(u, 4) = "_GF_" Then
        EsHojaGenerada = True
        Exit Function
    End If

    ' --- Hoja de debug del modulo legacy CM ---
    If u = "CM_DEBUG_TOP5" Then
        EsHojaGenerada = True
        Exit Function
    End If

End Function

'======================
' Eliminar una conexion por todos sus prefijos posibles
'======================
Private Sub EliminarConexion(ByVal wb As Workbook, ByVal queryName As String)
    Dim candidatos(3) As String
    candidatos(0) = "Consulta - " & queryName
    candidatos(1) = "Query - " & queryName
    candidatos(2) = "PQ_" & queryName
    candidatos(3) = queryName
    Dim i As Long
    For i = 0 To 3
        On Error Resume Next: wb.Connections(candidatos(i)).Delete: On Error GoTo 0
    Next i
End Sub

'======================
' Eliminar consultas PQ y sus conexiones
'======================
Private Sub EliminarConsultas(ByVal wb As Workbook, ByRef log As String)
    Dim nombres() As String
    ReDim nombres(13)

    ' Fondos
    nombres(0) = "RAW_SUS"
    nombres(1) = "SUS"
    nombres(2) = "SUS_ALERTAS"
    nombres(3) = "RAW_RES"
    nombres(4) = "RES"
    nombres(5) = "RES_ALERTAS"
    ' SAB MC
    nombres(6) = "SAB_MC_RAW"
    nombres(7) = "SAB_MC_MAIN"
    nombres(8) = "SAB_MC_ALERTAS_DEP"
    nombres(9) = "SAB_MC_ALERTAS_RET"
    ' SAB CM
    nombres(10) = "SAB_CM_RAW"
    nombres(11) = "SAB_CM_MAIN"
    nombres(12) = "SAB_CM_ALERTAS_COM"
    nombres(13) = "SAB_CM_ALERTAS_VEN"

    Dim i As Long, qn As String
    For i = 0 To UBound(nombres)
        qn = nombres(i)
        On Error Resume Next
        wb.Queries.Item(qn).Delete
        If Err.Number = 0 Then log = log & "  Consulta eliminada: " & qn & vbCrLf
        Err.Clear
        On Error GoTo 0
        EliminarConexion wb, qn
    Next i
End Sub

'======================
' Inventario: hojas que se eliminaran
'======================
Private Function ListarHojasAEliminar(ByVal wb As Workbook) As Collection
    Dim col As New Collection
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If EsHojaGenerada(ws.Name) Then col.Add ws.Name
    Next ws
    Set ListarHojasAEliminar = col
End Function

'======================
' Inventario: consultas que se eliminaran
'======================
Private Function ListarConsultasAEliminar(ByVal wb As Workbook) As Collection
    Dim col As New Collection
    Dim nombres() As String
    ReDim nombres(13)

    nombres(0) = "RAW_SUS":             nombres(1) = "SUS"
    nombres(2) = "SUS_ALERTAS":         nombres(3) = "RAW_RES"
    nombres(4) = "RES":                 nombres(5) = "RES_ALERTAS"
    nombres(6) = "SAB_MC_RAW":          nombres(7) = "SAB_MC_MAIN"
    nombres(8) = "SAB_MC_ALERTAS_DEP": nombres(9) = "SAB_MC_ALERTAS_RET"
    nombres(10) = "SAB_CM_RAW":         nombres(11) = "SAB_CM_MAIN"
    nombres(12) = "SAB_CM_ALERTAS_COM": nombres(13) = "SAB_CM_ALERTAS_VEN"

    Dim i As Long, qn As String
    Dim dummy As Object
    For i = 0 To UBound(nombres)
        qn = nombres(i)
        On Error Resume Next
        Set dummy = wb.Queries.Item(qn)
        If Err.Number = 0 Then col.Add qn
        Err.Clear
        On Error GoTo 0
    Next i
    Set ListarConsultasAEliminar = col
End Function

'======================
' Texto de confirmacion
'======================
Private Function ArmarTextoConfirmacion(ByVal hojas As Collection, ByVal consultas As Collection) As String
    Dim txt As String
    Dim nm As Variant, qn As Variant

    txt = "Se eliminaran los siguientes elementos:" & vbCrLf & vbCrLf

    If hojas.count > 0 Then
        txt = txt & "HOJAS (" & hojas.count & "):" & vbCrLf
        For Each nm In hojas
            txt = txt & "  - " & CStr(nm) & vbCrLf
        Next nm
    Else
        txt = txt & "HOJAS: ninguna que eliminar." & vbCrLf
    End If

    txt = txt & vbCrLf

    If consultas.count > 0 Then
        txt = txt & "CONSULTAS PQ (" & consultas.count & "):" & vbCrLf
        For Each qn In consultas
            txt = txt & "  - " & CStr(qn) & vbCrLf
        Next qn
    Else
        txt = txt & "CONSULTAS PQ: ninguna que eliminar." & vbCrLf
    End If

    txt = txt & vbCrLf
    txt = txt & "Las hojas no generadas por el proceso se conservaran." & vbCrLf & vbCrLf
    txt = txt & "Confirmar eliminacion?"

    ArmarTextoConfirmacion = txt
End Function

'======================
' Punto de entrada publico
'======================
Public Sub ResetProceso()
    Dim wb        As Workbook
    Dim hojas     As Collection
    Dim consultas As Collection
    Dim txtConf   As String
    Dim log       As String
    Dim errores   As String
    Dim resumen   As String
    Dim i         As Long
    Dim ws        As Worksheet
    Dim nmHoja    As String
    Dim errNum    As Long
    Dim errDesc   As String

    Set wb = ThisWorkbook
    Set hojas = ListarHojasAEliminar(wb)
    Set consultas = ListarConsultasAEliminar(wb)

    If hojas.count = 0 And consultas.count = 0 Then
        MsgBox "No se encontraron hojas ni consultas generadas por el proceso." & vbCrLf & _
               "No hay nada que eliminar.", vbInformation, "Reset"
        Exit Sub
    End If

    txtConf = ArmarTextoConfirmacion(hojas, consultas)

    If MsgBox(txtConf, vbQuestion + vbYesNo + vbDefaultButton2, "Reset - Confirmar") = vbNo Then
        MsgBox "Operacion cancelada. No se realizaron cambios.", vbInformation, "Reset"
        Exit Sub
    End If

    RstAppFreeze True

    log = vbNullString
    errores = vbNullString

    For i = wb.Worksheets.count To 1 Step -1
        Set ws = wb.Worksheets(i)
        If EsHojaGenerada(ws.Name) Then
            nmHoja = ws.Name
            errNum = 0
            errDesc = vbNullString
            On Error Resume Next
            ws.Visible = xlSheetVisible
            ws.Delete
            errNum = Err.Number
            errDesc = Err.Description
            Err.Clear
            On Error GoTo 0
            If errNum = 0 Then
                log = log & "  Hoja eliminada: " & nmHoja & vbCrLf
            Else
                errores = errores & "  No se pudo eliminar '" & nmHoja & "': " & errDesc & vbCrLf
            End If
        End If
    Next i

    EliminarConsultas wb, log

    RstAppFreeze False

    If Len(errores) = 0 Then
        resumen = "Reset completado exitosamente." & vbCrLf & vbCrLf
        If Len(log) > 0 Then
            resumen = resumen & "Elementos eliminados:" & vbCrLf & log
        Else
            resumen = resumen & "No habia elementos que eliminar."
        End If
        MsgBox resumen, vbInformation, "Reset"
    Else
        resumen = "Reset completado con advertencias." & vbCrLf & vbCrLf
        If Len(log) > 0 Then resumen = resumen & "Eliminados correctamente:" & vbCrLf & log & vbCrLf
        resumen = resumen & "Errores:" & vbCrLf & errores
        MsgBox resumen, vbExclamation, "Reset"
    End If
End Sub


