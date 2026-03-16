'==========================
' UserForm: frmCargaSAB
'==========================
Option Explicit

Private gHandlers       As Collection
Private isRunning       As Boolean
Private gStage          As String
Private gSuppressEvents As Boolean
Private gInSync         As Boolean   ' guard de reentrada para SyncControlStates
Private gLastTipo       As String    ' ultimo valor procesado de cbTipoCarga
Private gLastOrg        As String    ' ultimo valor procesado de cbOrigen

' ==========================
' API de progreso (usada por modUF_PollProxy)
' pct: 0 a 1
' ==========================
Public Sub ProgressToCurrent(ByVal pct As Double, ByVal msg As String)
    On Error Resume Next
    If pct < 0 Then pct = 0
    If pct > 1 Then pct = 1
    If Len(Trim$(msg)) > 0 Then
        If StrComp(gStage, msg, vbTextCompare) <> 0 Then
            gStage = msg
            AppendLogLine msg
        End If
    End If
    Application.StatusBar = msg
    Dim fr As MSForms.Frame
    Set fr = GetFrameOrNothing("fraProg")
    If fr Is Nothing Then Exit Sub
    Dim lbBg     As MSForms.Label
    Dim lbFill   As MSForms.Label
    Dim lbPct    As MSForms.Label
    Dim lbStatus As MSForms.Label
    Set lbBg     = GetLabelInFrame(fr, "lblBarBg")
    Set lbFill   = GetLabelInFrame(fr, "lblBar")
    Set lbPct    = GetLabelInFrame(fr, "lblPct")
    Set lbStatus = GetLabelInFrame(fr, "lblStatus")
    If Not lbBg Is Nothing And Not lbFill Is Nothing Then
        Dim wMax As Single
        wMax = lbBg.Width - 2
        If wMax < 0 Then wMax = 0
        lbFill.Width = wMax * pct
        If pct > 0 And lbFill.Width < 1 Then lbFill.Width = 1
    End If
    If Not lbPct    Is Nothing Then lbPct.Caption    = Format$(pct, "0%")
    If Not lbStatus Is Nothing Then lbStatus.Caption = msg
    Me.Repaint
    DoEvents
    On Error GoTo 0
End Sub

Public Sub Progress(ByVal pct As Double, ByVal msg As String)
    ProgressToCurrent pct, msg
End Sub

Public Sub PollTick()
    DoEvents
End Sub

' ==========================
' Ciclo de vida del form
' ==========================
Private Sub UserForm_Initialize()
    Set gHandlers   = New Collection
    gSuppressEvents = True
    BuildOrRefreshUI
    InitCombosDefaults
    gSuppressEvents = False
    On Error Resume Next
    modUF_PollProxy.Attach Me
    On Error GoTo 0
    SetStatusOnly 0, "Listo para iniciar."
    ClearLog
End Sub

Private Sub UserForm_Terminate()
    EndProgressHook
    On Error Resume Next
    modUF_PollProxy.Detach
    On Error GoTo 0
End Sub

' ==========================
' UX: Busy state
' ==========================
Private Sub SetBusy(ByVal running As Boolean, Optional ByVal statusMsg As String = "")
    isRunning = running
    If HasControl("cmdCargar")   Then Me.Controls("cmdCargar").Enabled   = Not running
    If HasControl("cmdExaminar") Then Me.Controls("cmdExaminar").Enabled = Not running
    If HasControl("cbTipoCarga") Then Me.Controls("cbTipoCarga").Enabled = Not running
    If HasControl("cbOrigen")    Then Me.Controls("cbOrigen").Enabled    = Not running
    If HasControl("cbOperacion") Then Me.Controls("cbOperacion").Enabled = Not running
    If HasControl("txtMeses")    Then Me.Controls("txtMeses").Enabled    = Not running
    If HasControl("txtArchivo")  Then Me.Controls("txtArchivo").Enabled  = Not running
    If HasControl("cmdCancelar") Then
        Me.Controls("cmdCancelar").Caption = IIf(running, "Cerrar", "Cancelar")
    End If
    Me.MousePointer = IIf(running, fmMousePointerHourGlass, fmMousePointerDefault)
    If Len(Trim$(statusMsg)) > 0 Then
        SetStatusOnly IIf(running, 0.01, 0), statusMsg
    End If
End Sub

' ==========================
' UI: crear o reutilizar controles
' ==========================
Private Sub BuildOrRefreshUI()
    Me.Caption         = "Cargar Datos"
    Me.StartUpPosition = 1
    Dim x As Single: x = 12
    Dim y As Single: y = 12
    Dim l  As MSForms.Label
    Dim t  As MSForms.TextBox
    Dim cb As MSForms.ComboBox
    Dim b  As MSForms.CommandButton
    Dim fr As MSForms.Frame

    Set l = EnsureLabel(Me, "lblTitulo")
    l.Caption   = "Cargar datos: Transacciones / Clientes"
    l.Left      = x: l.Top = y: l.Width = 420
    l.Font.Bold = True: l.Font.Size = 12
    y = y + 26

    Set l = EnsureLabel(Me, "lblTipoCarga")
    l.Caption = "Tipo de dato:": l.Left = x: l.Top = y: l.Width = 110
    Set cb = EnsureCombo(Me, "cbTipoCarga")
    cb.Left = x + 120: cb.Top = y - 3: cb.Width = 260
    cb.Style = fmStyleDropDownList
    cb.ControlTipText = "Elige si cargar transacciones o clientes."
    AttachCombo cb
    y = y + 28

    Set l = EnsureLabel(Me, "lblOrigen")
    l.Caption = "Origen de datos:": l.Left = x: l.Top = y: l.Width = 110
    Set cb = EnsureCombo(Me, "cbOrigen")
    cb.Left = x + 120: cb.Top = y - 3: cb.Width = 260
    cb.Style = fmStyleDropDownList
    cb.ControlTipText = "Selecciona Fondos o SAB."
    AttachCombo cb
    y = y + 28

    Set l = EnsureLabel(Me, "lblOperacion")
    l.Caption = "Tipo de operacion:": l.Left = x: l.Top = y: l.Width = 110
    Set cb = EnsureCombo(Me, "cbOperacion")
    cb.Left = x + 120: cb.Top = y - 3: cb.Width = 260
    cb.Style = fmStyleDropDownList
    cb.ControlTipText = "Elige el tipo de operacion segun el origen."
    AttachCombo cb
    y = y + 28

    Set l = EnsureLabel(Me, "lblMeses")
    l.Caption = "Ultimos meses:": l.Left = x: l.Top = y: l.Width = 110
    Set t = EnsureTextBox(Me, "txtMeses")
    t.Left = x + 120: t.Top = y - 3: t.Width = 50
    t.ControlTipText = "Cantidad de meses a cargar. Por defecto 6."
    y = y + 28

    Set l = EnsureLabel(Me, "lblArchivo")
    l.Caption = "Archivo origen:": l.Left = x: l.Top = y: l.Width = 110
    Set t = EnsureTextBox(Me, "txtArchivo")
    t.Left = x + 120: t.Top = y - 3: t.Width = 400
    t.ControlTipText = "Ruta del archivo."
    Set b = EnsureButton(Me, "cmdExaminar")
    b.Caption = "Examinar...": b.Left = x + 530: b.Top = y - 5: b.Width = 90
    b.ControlTipText = "Buscar archivo."
    AttachButton b
    y = y + 34

    Set fr = EnsureFrame(Me, "fraProg")
    fr.Caption = " Progreso"
    fr.Left = x: fr.Top = y: fr.Width = 610: fr.Height = 130
    EnsureProgressControls fr
    y = y + fr.Height + 10

    Dim bOK     As MSForms.CommandButton
    Dim bCancel As MSForms.CommandButton
    Set bOK = EnsureButton(Me, "cmdCargar")
    bOK.Caption = "Cargar": bOK.Left = x + 370: bOK.Top = y: bOK.Width = 120
    AttachButton bOK
    Set bCancel = EnsureButton(Me, "cmdCancelar")
    bCancel.Caption = "Cancelar": bCancel.Left = x + 500: bCancel.Top = y: bCancel.Width = 120
    AttachButton bCancel

    Me.Width  = 650
    Me.Height = y + 90
End Sub

Private Sub EnsureProgressControls(ByVal fr As MSForms.Frame)
    Dim txtLog As MSForms.TextBox
    Set txtLog        = EnsureTextBox(fr, "txtProgLog")
    txtLog.Left       = 10: txtLog.Top = 18
    txtLog.Width      = fr.InsideWidth - 20: txtLog.Height = 56
    txtLog.MultiLine  = True: txtLog.Locked = True
    txtLog.ScrollBars = fmScrollBarsVertical
    txtLog.BackColor  = RGB(255, 255, 255)

    Dim lbBg     As MSForms.Label
    Dim lbFill   As MSForms.Label
    Dim lbPct    As MSForms.Label
    Dim lbStatus As MSForms.Label

    Set lbBg = EnsureLabel(fr, "lblBarBg")
    lbBg.Left        = 10: lbBg.Top = txtLog.Top + txtLog.Height + 10
    lbBg.Width       = fr.InsideWidth - 70: lbBg.Height = 12
    lbBg.BackStyle   = fmBackStyleOpaque: lbBg.BackColor = RGB(230, 230, 230)
    lbBg.BorderStyle = fmBorderStyleSingle: lbBg.Caption = ""

    Set lbFill = EnsureLabel(fr, "lblBar")
    lbFill.Left        = lbBg.Left + 1: lbFill.Top = lbBg.Top + 1
    lbFill.Width       = 0: lbFill.Height = lbBg.Height - 2
    lbFill.BackStyle   = fmBackStyleOpaque: lbFill.BackColor = RGB(0, 120, 215)
    lbFill.BorderStyle = fmBorderStyleNone: lbFill.Caption = ""

    Set lbPct = EnsureLabel(fr, "lblPct")
    lbPct.Left      = lbBg.Left + lbBg.Width + 10
    lbPct.Top       = lbBg.Top - 2: lbPct.Width = 40: lbPct.Height = 14
    lbPct.Caption   = "0%": lbPct.TextAlign = fmTextAlignRight

    Set lbStatus = EnsureLabel(fr, "lblStatus")
    lbStatus.Left   = 10: lbStatus.Top = lbBg.Top + lbBg.Height + 8
    lbStatus.Width  = fr.InsideWidth - 20: lbStatus.Height = 14
    lbStatus.Caption = ""
End Sub

' ==========================
' Inicializacion combos
' ==========================
Private Sub InitCombosDefaults()
    gSuppressEvents = True

    If HasControl("cbTipoCarga") Then
        With Me.Controls("cbTipoCarga")
            .Clear
            .AddItem "Seleccionar"
            .AddItem "Transacciones"
            .AddItem "Clientes"
            .ListIndex = 0
        End With
    End If

    If HasControl("cbOrigen") Then
        With Me.Controls("cbOrigen")
            .Clear
            .AddItem "Seleccionar"
            .AddItem "Fondos"
            .AddItem "SAB"
            .ListIndex = 0
        End With
    End If

    SetOperacionOptionsByOrigen

    If HasControl("txtMeses")   Then Me.Controls("txtMeses").Value   = "6"
    If HasControl("txtArchivo") Then Me.Controls("txtArchivo").Value = ""

    gSuppressEvents = False

    ' Sincronizar estado de controles segun valor actual de cbTipoCarga.
    ' Necesario porque los eventos no disparan durante la inicializacion.
    If HasControl("cbTipoCarga") Then gLastTipo = CStr(Me.Controls("cbTipoCarga").Value)
    If HasControl("cbOrigen")    Then gLastOrg  = CStr(Me.Controls("cbOrigen").Value)
    SyncControlStates
End Sub

' Rellena cbOperacion segun origen seleccionado.
Private Sub SetOperacionOptionsByOrigen()
    If Not HasControl("cbOperacion") Then Exit Sub
    Dim cbOp As MSForms.ComboBox
    Set cbOp = Me.Controls("cbOperacion")
    Dim wasSupp As Boolean
    wasSupp         = gSuppressEvents
    gSuppressEvents = True
    cbOp.Clear
    cbOp.AddItem "Seleccionar"
    Dim org As String
    If HasControl("cbOrigen") Then org = UCase$(Trim$(CStr(Me.Controls("cbOrigen").Value)))
    Select Case org
        Case "FONDOS"
            cbOp.AddItem "Suscripcion"
            cbOp.AddItem "Rescate"
        Case "SAB"
            cbOp.AddItem "Movimiento de Caja"
            cbOp.AddItem "Cambio de Moneda"
    End Select
    cbOp.ListIndex  = 0
    gSuppressEvents = wasSupp
End Sub

' Habilita o deshabilita controles segun el tipo de carga y origen actuales.
' Guard gInSync evita reentrada cuando activar/desactivar un control
' dispara eventos de combo que vuelven a llamar aqui.
Private Sub SyncControlStates()
    If gInSync Then Exit Sub
    gInSync = True

    On Error Resume Next

    Dim tipo As String
    Dim org  As String
    tipo = ""
    org  = ""
    If HasControl("cbTipoCarga") Then tipo = UCase$(Trim$(CStr(Me.Controls("cbTipoCarga").Value)))
    If HasControl("cbOrigen")    Then org  = UCase$(Trim$(CStr(Me.Controls("cbOrigen").Value)))

    Select Case tipo
        Case "TRANSACCIONES"
            If HasControl("cbOrigen")    Then Me.Controls("cbOrigen").Enabled    = True
            If HasControl("txtMeses")    Then Me.Controls("txtMeses").Enabled    = True
            If HasControl("cbOperacion") Then
                Me.Controls("cbOperacion").Enabled = (org = "FONDOS" Or org = "SAB")
            End If

        Case "CLIENTES"
            If HasControl("cbOrigen")    Then Me.Controls("cbOrigen").Enabled    = True
            If HasControl("cbOperacion") Then Me.Controls("cbOperacion").Enabled = False
            If HasControl("txtMeses")    Then Me.Controls("txtMeses").Enabled    = False

        Case Else
            If HasControl("cbOrigen")    Then Me.Controls("cbOrigen").Enabled    = False
            If HasControl("cbOperacion") Then Me.Controls("cbOperacion").Enabled = False
            If HasControl("txtMeses")    Then Me.Controls("txtMeses").Enabled    = False
    End Select

    On Error GoTo 0
    gInSync = False
End Sub

Private Function IsPlaceholder(ByVal s As String) As Boolean
    IsPlaceholder = (Len(Trim$(s)) = 0) Or (UCase$(Trim$(s)) = "SELECCIONAR")
End Function

' ==========================
' Progreso: inicio/fin
' ==========================
Private Sub BeginProgressHook()
    On Error Resume Next
    modUF_PollProxy.Attach Me
    ClearLog
    gStage = vbNullString
    SetStatusOnly 0, "Inicializando..."
    On Error GoTo 0
End Sub

Private Sub EndProgressHook()
    On Error Resume Next
    Application.StatusBar = False
    On Error GoTo 0
End Sub

' ==========================
' Acciones (usadas por CCtrlEvents)
' ==========================
Public Sub OnExaminar()
    Dim p As String
    p = PickFileXLS("Selecciona el archivo origen")
    If Len(p) > 0 Then Me.Controls("txtArchivo").Value = p
End Sub

Public Sub OnCargar()
    If isRunning Then Exit Sub

    Dim ruta      As String
    Dim mesesSel  As Long
    Dim origen    As String
    Dim op        As String
    Dim esRescate As Boolean
    Dim tipoCarga As String
    Dim tipoU     As String
    Dim orgU      As String
    Dim opU       As String

    ruta = CStr(Me.Controls("txtArchivo").Value)
    If Len(Trim$(ruta)) = 0 Then
        MsgBox "Selecciona un archivo origen.", vbExclamation: Exit Sub
    End If
    If Dir(ruta, vbNormal) = "" Then
        MsgBox "El archivo no existe en la ruta indicada." & vbCrLf & ruta, vbExclamation
        Exit Sub
    End If

    If HasControl("cbTipoCarga") Then tipoCarga = CStr(Me.Controls("cbTipoCarga").Value)
    tipoU = UCase$(Trim$(tipoCarga))
    If IsPlaceholder(tipoCarga) Then
        MsgBox "Selecciona el tipo de dato a cargar (Transacciones o Clientes).", vbExclamation
        Exit Sub
    End If

    SetBusy True, "Iniciando carga..."
    BeginProgressHook
    On Error GoTo fallo

    ' ==========================
    ' Caso: Clientes
    ' ==========================
    If tipoU = "CLIENTES" Then
        origen = CStr(Me.Controls("cbOrigen").Value)
        If IsPlaceholder(origen) Then
            MsgBox "Selecciona el origen de datos (Fondos o SAB) para los clientes.", vbExclamation
            GoTo salir
        End If
        orgU = UCase$(Trim$(origen))
        Select Case orgU
            Case "SAB"
                ProgressToCurrent 0.05, "Creando consulta de Clientes SAB..."
                Application.Run "CrearQueryClientesSAB", ruta, True
            Case "FONDOS"
                ProgressToCurrent 0.05, "Creando consulta de Clientes Fondos..."
                Application.Run "CrearQueryClientesFondos", ruta, True
            Case Else
                MsgBox "Origen de datos no reconocido para clientes.", vbExclamation
                GoTo salir
        End Select
        ProgressToCurrent 1, "Carga completada."
        EndProgressHook
        SetBusy False, "Listo."
        Unload Me
        Exit Sub
    End If

    ' ==========================
    ' Caso: Transacciones
    ' ==========================
    If tipoU = "TRANSACCIONES" Then
        origen = CStr(Me.Controls("cbOrigen").Value)
        If IsPlaceholder(origen) Then
            MsgBox "Selecciona el origen de datos.", vbExclamation: GoTo salir
        End If
        op = CStr(Me.Controls("cbOperacion").Value)
        If IsPlaceholder(op) Then
            MsgBox "Selecciona el tipo de operacion.", vbExclamation: GoTo salir
        End If
        mesesSel = Val(Me.Controls("txtMeses").Value)
        If mesesSel <= 0 Then mesesSel = 6
        orgU = UCase$(Trim$(origen))
        opU  = UCase$(Trim$(op))

        Select Case orgU
            Case "FONDOS"
                If InStr(1, opU, "SUSCRIP", vbTextCompare) > 0 Then
                    esRescate = False
                ElseIf InStr(1, opU, "RESCATE", vbTextCompare) > 0 Then
                    esRescate = True
                Else
                    MsgBox "Operacion no valida para Fondos.", vbExclamation: GoTo salir
                End If
                ProgressToCurrent 0.05, "Creando consultas de Fondos..."
                Application.Run "CrearQueryFondos", ruta, mesesSel, esRescate, "FONDOS", True

            Case "SAB"
                If InStr(1, opU, "MOVIMIENTO", vbTextCompare) > 0 Then
                    ProgressToCurrent 0.05, "Creando consultas SAB - Movimiento de Caja..."
                    Application.Run "CrearQuerySAB_MC", ruta, mesesSel, True
                ElseIf InStr(1, opU, "CAMBIO", vbTextCompare) > 0 Then
                    ProgressToCurrent 0.05, "Creando consultas SAB - Cambio de Moneda..."
                    Application.Run "CrearQuerySAB_CM", ruta, mesesSel, True
                Else
                    MsgBox "Operacion no valida para SAB.", vbExclamation: GoTo salir
                End If

            Case Else
                MsgBox "Origen de datos no reconocido.", vbExclamation: GoTo salir
        End Select

        ProgressToCurrent 1, "Carga completada."
        EndProgressHook
        SetBusy False, "Listo."
        Unload Me
        Exit Sub
    End If

salir:
    EndProgressHook
    SetBusy False, "Listo."
    Exit Sub
fallo:
    EndProgressHook
    SetBusy False, "Listo."
    SetStatusOnly 0, "Error al cargar."
    MsgBox "Error al cargar: " & Err.Number & " - " & Err.Description, vbCritical
End Sub

Public Sub OnCancelar()
    If isRunning Then
        If MsgBox("Hay una operacion en progreso." & vbCrLf & _
                  "Deseas cerrar de todos modos?", vbQuestion + vbYesNo) = vbNo Then Exit Sub
        EndProgressHook
        Unload Me
        Exit Sub
    End If
    EndProgressHook
    Unload Me
End Sub

Public Sub OnComboChanged(ByVal Name As String)
    If gSuppressEvents Then Exit Sub

    Dim curVal As String

    Select Case Name
        Case "cbTipoCarga"
            If HasControl("cbTipoCarga") Then curVal = CStr(Me.Controls("cbTipoCarga").Value)
            If curVal = gLastTipo Then Exit Sub
            gLastTipo = curVal
            SetOperacionOptionsByOrigen
            SyncControlStates
            SetStatusOnly 0, "Tipo de dato: " & curVal
            ClearLog

        Case "cbOrigen"
            If HasControl("cbOrigen") Then curVal = CStr(Me.Controls("cbOrigen").Value)
            If curVal = gLastOrg Then Exit Sub
            gLastOrg = curVal
            SetOperacionOptionsByOrigen
            SyncControlStates
            SetStatusOnly 0, "Origen cambiado."
            ClearLog
    End Select
End Sub

' ==========================
' Status sin escribir en el log
' ==========================
Private Sub SetStatusOnly(ByVal pct As Double, ByVal msg As String)
    On Error Resume Next
    If pct < 0 Then pct = 0
    If pct > 1 Then pct = 1
    Application.StatusBar = msg
    Dim fr As MSForms.Frame
    Set fr = GetFrameOrNothing("fraProg")
    If fr Is Nothing Then Exit Sub
    Dim lbBg     As MSForms.Label
    Dim lbFill   As MSForms.Label
    Dim lbPct    As MSForms.Label
    Dim lbStatus As MSForms.Label
    Set lbBg     = GetLabelInFrame(fr, "lblBarBg")
    Set lbFill   = GetLabelInFrame(fr, "lblBar")
    Set lbPct    = GetLabelInFrame(fr, "lblPct")
    Set lbStatus = GetLabelInFrame(fr, "lblStatus")
    If Not lbBg Is Nothing And Not lbFill Is Nothing Then
        Dim wMax As Single
        wMax = lbBg.Width - 2
        If wMax < 0 Then wMax = 0
        lbFill.Width = wMax * pct
        If pct > 0 And lbFill.Width < 1 Then lbFill.Width = 1
    End If
    If Not lbPct    Is Nothing Then lbPct.Caption    = Format$(pct, "0%")
    If Not lbStatus Is Nothing Then lbStatus.Caption = msg
    Me.Repaint
    DoEvents
    On Error GoTo 0
End Sub

' ==========================
' Log de progreso (UX)
' ==========================
Private Sub ClearLog()
    Dim fr As MSForms.Frame
    Set fr = GetFrameOrNothing("fraProg")
    If fr Is Nothing Then Exit Sub
    Dim t As MSForms.TextBox
    On Error Resume Next
    Set t = fr.Controls("txtProgLog")
    On Error GoTo 0
    If Not t Is Nothing Then t.Text = ""
End Sub

Private Sub AppendLogLine(ByVal line As String)
    Dim fr As MSForms.Frame
    Set fr = GetFrameOrNothing("fraProg")
    If fr Is Nothing Then Exit Sub
    Dim t As MSForms.TextBox
    On Error Resume Next
    Set t = fr.Controls("txtProgLog")
    On Error GoTo 0
    If t Is Nothing Then Exit Sub
    Dim s As String
    s = t.Text
    If Len(s) > 0 Then s = s & vbCrLf
    s = s & line
    Dim parts()  As String
    Dim i        As Long
    Dim out      As String
    parts = Split(s, vbCrLf)
    If UBound(parts) > 15 Then
        Dim startAt As Long
        startAt = UBound(parts) - 15
        out = ""
        For i = startAt To UBound(parts)
            If Len(out) > 0 Then out = out & vbCrLf
            out = out & parts(i)
        Next i
        t.Text = out
    Else
        t.Text = s
    End If
    t.SelStart = Len(t.Text)
End Sub

' ==========================
' Helpers: Ensure controls
' ==========================
Private Function EnsureLabel(ByVal parent As Object, ByVal nm As String) As MSForms.Label
    Dim lb As MSForms.Label
    On Error Resume Next: Set lb = parent.Controls(nm): On Error GoTo 0
    If lb Is Nothing Then Set lb = parent.Controls.Add("Forms.Label.1", nm, True)
    Set EnsureLabel = lb
End Function

Private Function EnsureTextBox(ByVal parent As Object, ByVal nm As String) As MSForms.TextBox
    Dim tb As MSForms.TextBox
    On Error Resume Next: Set tb = parent.Controls(nm): On Error GoTo 0
    If tb Is Nothing Then Set tb = parent.Controls.Add("Forms.TextBox.1", nm, True)
    Set EnsureTextBox = tb
End Function

Private Function EnsureCombo(ByVal parent As Object, ByVal nm As String) As MSForms.ComboBox
    Dim cb As MSForms.ComboBox
    On Error Resume Next: Set cb = parent.Controls(nm): On Error GoTo 0
    If cb Is Nothing Then Set cb = parent.Controls.Add("Forms.ComboBox.1", nm, True)
    Set EnsureCombo = cb
End Function

Private Function EnsureButton(ByVal parent As Object, ByVal nm As String) As MSForms.CommandButton
    Dim b As MSForms.CommandButton
    On Error Resume Next: Set b = parent.Controls(nm): On Error GoTo 0
    If b Is Nothing Then Set b = parent.Controls.Add("Forms.CommandButton.1", nm, True)
    Set EnsureButton = b
End Function

Private Function EnsureFrame(ByVal parent As Object, ByVal nm As String) As MSForms.Frame
    Dim fr As MSForms.Frame
    On Error Resume Next: Set fr = parent.Controls(nm): On Error GoTo 0
    If fr Is Nothing Then Set fr = parent.Controls.Add("Forms.Frame.1", nm, True)
    Set EnsureFrame = fr
End Function

Private Function GetFrameOrNothing(ByVal nm As String) As MSForms.Frame
    On Error Resume Next
    Set GetFrameOrNothing = Me.Controls(nm)
    On Error GoTo 0
End Function

Private Function GetLabelInFrame(ByVal fr As MSForms.Frame, ByVal nm As String) As MSForms.Label
    On Error Resume Next
    Set GetLabelInFrame = fr.Controls(nm)
    On Error GoTo 0
End Function

Private Sub AttachButton(ByVal b As MSForms.CommandButton)
    Dim h As CCtrlEvents
    Set h = New CCtrlEvents
    h.HookButton b, Me
    gHandlers.Add h
End Sub

Private Sub AttachCombo(ByVal c As MSForms.ComboBox)
    Dim h As CCtrlEvents
    Set h = New CCtrlEvents
    h.HookCombo c, Me
    gHandlers.Add h
End Sub

Private Function HasControl(ByVal Name As String) As Boolean
    Dim dummy As Object
    On Error Resume Next
    Set dummy  = Me.Controls(Name)
    HasControl = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function
