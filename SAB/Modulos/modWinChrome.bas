' ==========================
' modWinChrome
' ==========================
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Function FindWindowA Lib "user32" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowLongPtr Lib "user32" Alias "GetWindowLongPtrA" ( _
        ByVal hWnd As LongPtr, ByVal nIndex As Long) As LongPtr
    Private Declare PtrSafe Function SetWindowLongPtr Lib "user32" Alias "SetWindowLongPtrA" ( _
        ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As LongPtr) As LongPtr
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" ( _
        ByVal hWnd As LongPtr) As Long
    Private Type HWND_T: h As LongPtr: End Type
#Else
    Private Declare Function FindWindowA Lib "user32" ( _
        ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" ( _
        ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function DrawMenuBar Lib "user32" ( _
        ByVal hWnd As Long) As Long
    Private Type HWND_T: h As Long: End Type
#End If

Private Const GWL_STYLE As Long = -16&
Private Const WS_SYSMENU As Long = &H80000
Private Const WS_MINIMIZEBOX As Long = &H20000

Private Function GetUFhWnd(ByVal uf As Object) As HWND_T
    ' Intenta ambas clases de ventana típicas de UserForm.
    Dim classes As Variant, cls As Variant
    classes = Array("ThunderDFrame", "ThunderXFrame")
    Dim h As HWND_T
    For Each cls In classes
#If VBA7 Then
        h.h = FindWindowA(CStr(cls), CStr(uf.caption))
#Else
        h.h = FindWindowA(CStr(cls), CStr(uf.caption))
#End If
        If h.h <> 0 Then Exit For
    Next
    GetUFhWnd = h
End Function

Public Sub EnableMinimizeBox(ByVal uf As Object)
    On Error Resume Next
    Dim w As HWND_T: w = GetUFhWnd(uf)
    If w.h = 0 Then Exit Sub

#If VBA7 Then
    Dim sT As LongPtr
    sT = GetWindowLongPtr(w.h, GWL_STYLE)
    If sT = 0 Then Exit Sub
    sT = sT Or WS_SYSMENU Or WS_MINIMIZEBOX
    SetWindowLongPtr w.h, GWL_STYLE, sT
#Else
    Dim st32 As Long
    st32 = GetWindowLong(w.h, GWL_STYLE)
    If st32 = 0 Then Exit Sub
    st32 = st32 Or WS_SYSMENU Or WS_MINIMIZEBOX
    SetWindowLong w.h, GWL_STYLE, st32
#End If

    DrawMenuBar w.h
    On Error GoTo 0
End Sub