Option Explicit

Private WithEvents mBtn As MSForms.CommandButton
Private WithEvents mCbo As MSForms.ComboBox
Private mOwner As Object

Public Sub HookButton(ByVal btn As MSForms.CommandButton, ByVal owner As Object)
    Set mBtn = btn
    Set mOwner = owner
End Sub

Public Sub HookCombo(ByVal cbo As MSForms.ComboBox, ByVal owner As Object)
    Set mCbo = cbo
    Set mOwner = owner
End Sub

Private Sub mBtn_Click()
    On Error Resume Next
    If Not mOwner Is Nothing Then
        CallByName mOwner, "HandleButtonClick", VbMethod, mBtn.Name
    End If
    On Error GoTo 0
End Sub

Private Sub mCbo_Change()
    On Error Resume Next
    If Not mOwner Is Nothing Then
        CallByName mOwner, "HandleComboChange", VbMethod, mCbo.Name
    End If
    On Error GoTo 0
End Sub

