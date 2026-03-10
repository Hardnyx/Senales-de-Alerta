Option Explicit

' Mantiene una sola instancia del formulario
Private mForm As frmCargaFondos
Private mIsShowing As Boolean

' Devuelve siempre una única instancia del form (reutiliza si ya está cargado)
Private Function GetForm() As frmCargaFondos
    Dim uf As Object

    If Not mForm Is Nothing Then
        For Each uf In VBA.UserForms
            If uf Is mForm Then
                Set GetForm = mForm
                Exit Function
            End If
        Next uf
    End If

    For Each uf In VBA.UserForms
        If TypeName(uf) = "frmCargaFondos" Then
            Set mForm = uf
            Set GetForm = mForm
            Exit Function
        End If
    Next uf

    Set mForm = New frmCargaFondos
    Set GetForm = mForm
End Function

' Acceso público a la instancia actual (por si otros módulos la necesitan)
Public Function CurrentForm() As frmCargaFondos
    On Error Resume Next
    Set CurrentForm = GetForm()
    On Error GoTo 0
End Function

' Compatibilidad con macros existentes
Public Sub AbrirCargaFondos()
    ShowFormModal
End Sub

' Mostrar formulario en MODELESS (por defecto)
Public Sub MostrarFormulario_General()
    ShowFormModeless
End Sub

' Compatibilidad con el nombre anterior
Public Sub MostrarFormulario_CargaFondos()
    MostrarFormulario_General
End Sub

' Muestra la UI en modeless y la trae al frente sin crear instancias duplicadas
Public Sub ShowFormModeless()
    Dim f As frmCargaFondos
    If mIsShowing Then Exit Sub

    mIsShowing = True
    On Error GoTo CleanFail

    Set f = GetForm()

    On Error Resume Next
    f.StartUpPosition = 1 ' CenterOwner
    On Error GoTo 0

    modUF_PollProxy.Attach f

    If f.Visible Then
        f.Hide
        DoEvents
    End If

    f.Show vbModeless

    On Error Resume Next
    modWinChrome.EnableMinimizeBox f
    AppActivate Application.caption
    On Error GoTo 0

CleanExit:
    mIsShowing = False
    Exit Sub

CleanFail:
    Resume CleanExit
End Sub

' Versión modal (útil si algún proceso bloquea la UI)
Public Sub ShowFormModal()
    Dim f As frmCargaFondos
    If mIsShowing Then Exit Sub

    mIsShowing = True
    On Error GoTo CleanFail

    Set f = GetForm()

    On Error Resume Next
    f.StartUpPosition = 1 ' CenterOwner
    On Error GoTo 0

    modUF_PollProxy.Attach f

    ' Nota: EnableMinimizeBox debe llamarse en frmCargaFondos.UserForm_Activate
    f.Show vbModal

CleanExit:
    mIsShowing = False
    Exit Sub

CleanFail:
    Resume CleanExit
End Sub

' Cierra y limpia la referencia del formulario
Public Sub CerrarFormulario_General()
    On Error Resume Next
    If Not mForm Is Nothing Then
        modUF_PollProxy.Detach
        Unload mForm
        Set mForm = Nothing
    End If
    On Error GoTo 0
End Sub

' Llamar opcionalmente desde UserForm_Terminate para limpiar referencia
Public Sub OnFormTerminated()
    On Error Resume Next
    Set mForm = Nothing
    modUF_PollProxy.Detach
    On Error GoTo 0
End Sub

