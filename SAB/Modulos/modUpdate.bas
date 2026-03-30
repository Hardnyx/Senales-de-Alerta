'==========================
' modUpdate
' Actualiza la barra de carga del formulario
'==========================

Public Sub SAB_Progress(ByVal pct As Double, ByVal msg As String)
    Application.StatusBar = msg
    On Error Resume Next
    frmCargaSAB.ProgressToCurrent pct, msg
    On Error GoTo 0
End Sub