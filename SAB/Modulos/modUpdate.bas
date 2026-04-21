'==========================
' modUpdate
' Actualiza la barra de carga del formulario
'==========================
Option Explicit

Public gSABForm As frmCargaSAB
Public gTCDict  As Object        ' Tipo de cambio cargado, disponible para MC
Private mInPQRefresh As Boolean

Public Sub SAB_SetPQRefreshing(ByVal v As Boolean)
    mInPQRefresh = v
End Sub

Public Sub SAB_Progress(ByVal pct As Double, ByVal msg As String)
    Application.StatusBar = msg
    On Error Resume Next
    If Not gSABForm Is Nothing Then gSABForm.ProgressToCurrent pct, msg
    On Error GoTo 0
    If Not mInPQRefresh Then DoEvents
End Sub