Option Explicit

Private mUF As Object
Private mLastEnsure As Double

' ==========================
' API pública (compatibilidad)
' ==========================
Public Sub Attach(ByVal uf As Object)
    On Error Resume Next
    Set mUF = uf
    On Error GoTo 0
End Sub

Public Sub Detach()
    On Error Resume Next
    Set mUF = Nothing
    On Error GoTo 0
End Sub

' Alias usados por tu UserForm nuevo
Public Sub AttachForm(ByVal uf As Object)
    Attach uf
End Sub

Public Sub DetachForm()
    Detach
End Sub

' Alias usados por modUtils (compatibilidad)
Public Sub SetCurrentUF(ByVal uf As Object)
    Attach uf
End Sub

Public Sub StartPoll()
    ' No-op: el progreso se empuja directamente con ProgressToCurrent
End Sub

Public Sub StopPoll()
    ' No-op
End Sub

Public Function IsAttached() As Boolean
    IsAttached = Not (mUF Is Nothing)
End Function

Public Sub EnsureAttached(Optional ByVal force As Boolean = False)
    On Error Resume Next

    If Not (mUF Is Nothing) Then Exit Sub

    Dim nowT As Double
    nowT = Timer
    If nowT < mLastEnsure Then mLastEnsure = 0#

    If Not force Then
        If (nowT - mLastEnsure) < 0.5 Then Exit Sub
    End If
    mLastEnsure = nowT

    Dim uf As Object
    For Each uf In VBA.UserForms
        If LCase$(uf.name) = "frmcargafondos" Then
            Set mUF = uf
            Exit For
        End If
    Next uf

    On Error GoTo 0
End Sub

Private Function Clamp01(ByVal x As Double) As Double
    If x < 0# Then
        Clamp01 = 0#
    ElseIf x > 1# Then
        Clamp01 = 1#
    Else
        Clamp01 = x
    End If
End Function

' ==========================
' Progreso hacia el UserForm
' ==========================
Public Sub ProgressToCurrent(ByVal pct As Double, ByVal msg As String)
    On Error Resume Next

    pct = Clamp01(pct)

    If mUF Is Nothing Then EnsureAttached False
    If mUF Is Nothing Then Exit Sub

    Err.Clear
    CallByName mUF, "ProgressToCurrent", VbMethod, pct, msg
    If Err.Number <> 0 Then
        Err.Clear
        CallByName mUF, "Progress", VbMethod, pct, msg
    End If

    If Err.Number <> 0 Then
        Err.Clear
        CallByName mUF, "SetStatusOnly", VbMethod, msg
    End If

    On Error GoTo 0
End Sub

