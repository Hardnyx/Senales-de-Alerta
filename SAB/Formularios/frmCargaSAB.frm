'==========================
' UserForm: frmCargaSAB
'==========================
Option Explicit

Private gHandlers       As Collection
Private isRunning       As Boolean
Private gStage          As String
Private gSuppressEvents As Boolean

' ==========================
' API de progreso
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

    Dim lbBg     As MSForms.label
    Dim lbFill   As MSForms.label
    Dim lbPct    As MSForms.label
    Dim lbStatus As MSForms.label

    Set lbBg = GetLabelInFrame(fr, "lblBarBg")
    Set lbFill = GetLabelInFrame(fr, "lblBar")
    Set lbPct = GetLabelInFrame(fr, "lblPct")
    Set lbStatus = GetLabelInFrame(fr, "lblStatus")

    If Not lbBg Is Nothing And Not lbFill Is Nothing Then
        Dim wMax As Single
        wMax = lbBg.width - 2
        If wMax < 0 Then wMax = 0
        lbFill.width = wMax * pct
        If pct > 0 And lbFill.width < 1 Then lbFill.width = 1
    End If

    If Not lbPct Is Nothing Then lbPct.caption = Format$(pct, "0%")
    If Not lbStatus Is Nothing Then lbStatus.caption = msg

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
' Ciclo de vida
' ==========================
Private Sub UserForm_Initialize()
    Set gSABForm = Me

    Set gHandlers = New Collection

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
    Set gSABForm = Nothing
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

    If HasControl("cmdCargar") Then Me.Controls("cmdCargar").Enabled = Not running
    If HasControl("cmdExaminar") Then Me.Controls("cmdExaminar").Enabled = Not running
    If HasControl("cbTipoCarga") Then Me.Controls("cbTipoCarga").Enabled = Not running
    If HasControl("cbOperacion") Then Me.Controls("cbOperacion").Enabled = Not running
    If HasControl("txtMeses") Then Me.Controls("txtMeses").Enabled = Not running
    If HasControl("txtArchivo") Then Me.Controls("txtArchivo").Enabled = Not running

    If HasControl("cmdCancelar") Then
        Me.Controls("cmdCancelar").caption = IIf(running, "Cerrar", "Cancelar")
    End If

    If running Then
        Me.MousePointer = fmMousePointerHourGlass
    Else
        Me.MousePointer = fmMousePointerDefault
    End If

    If Len(Trim$(statusMsg)) > 0 Then
        SetStatusOnly IIf(running, 0.01, 0), statusMsg
    End If
End Sub

' ==========================
' UI: construir controles
' ==========================
Private Sub BuildOrRefreshUI()
    Me.caption = "SAB - Cargar Datos"
    Me.StartUpPosition = 1

    Dim x As Single: x = 12
    Dim y As Single: y = 12

    Dim l  As MSForms.label
    Dim t  As MSForms.TextBox
    Dim cb As MSForms.ComboBox
    Dim b  As MSForms.CommandButton
    Dim fr As MSForms.Frame

    Set l = EnsureLabel(Me, "lblTitulo")
    l.caption = "Cargar datos SAB: Transacciones / Clientes"
    l.Left = x: l.top = y: l.width = 420
    l.Font.Bold = True: l.Font.Size = 12
    y = y + 26

    Set l = EnsureLabel(Me, "lblTipoCarga")
    l.caption = "Tipo de dato:"
    l.Left = x: l.top = y: l.width = 110
    Set cb = EnsureCombo(Me, "cbTipoCarga")
    cb.Left = x + 120: cb.top = y - 3: cb.width = 260
    cb.Style = fmStyleDropDownList
    cb.ControlTipText = "Elige el tipo de dato a cargar."
    AttachCombo cb
    y = y + 28

    Set l = EnsureLabel(Me, "lblOperacion")
    l.caption = "Tipo de operacion:"
    l.Left = x: l.top = y: l.width = 110
    Set cb = EnsureCombo(Me, "cbOperacion")
    cb.Left = x + 120: cb.top = y - 3: cb.width = 260
    cb.Style = fmStyleDropDownList
    cb.ControlTipText = "Elige el tipo de operacion."
    AttachCombo cb
    y = y + 28

    Set l = EnsureLabel(Me, "lblMeses")
    l.caption = "Ultimos meses:"
    l.Left = x: l.top = y: l.width = 110
    Set t = EnsureTextBox(Me, "txtMeses")
    t.Left = x + 120: t.top = y - 3: t.width = 50
    t.ControlTipText = "Cantidad de meses a cargar. Por defecto 6."
    y = y + 28

    Set l = EnsureLabel(Me, "lblArchivo")
    l.caption = "Archivo origen:"
    l.Left = x: l.top = y: l.width = 110
    Set t = EnsureTextBox(Me, "txtArchivo")
    t.Left = x + 120: t.top = y - 3: t.width = 400
    t.ControlTipText = "Ruta del archivo."
    Set b = EnsureButton(Me, "cmdExaminar")
    b.caption = "Examinar...": b.Left = x + 530: b.top = y - 5: b.width = 90
    AttachButton b
    y = y + 34

    ' Indicador de TC cargado (solo informativo, visible cuando hay TC en memoria)
    Set l = EnsureLabel(Me, "lblTCEstado")
    l.caption = ""
    l.Left = x: l.top = y: l.width = 620
    l.ForeColor = RGB(0, 128, 0)
    l.Font.Italic = True
    y = y + 20

    Set fr = EnsureFrame(Me, "fraProg")
    fr.caption = " Progreso"
    fr.Left = x: fr.top = y: fr.width = 630: fr.height = 130
    EnsureProgressControls fr
    y = y + fr.height + 10

    Dim bOK     As MSForms.CommandButton
    Dim bCancel As MSForms.CommandButton

    Set bOK = EnsureButton(Me, "cmdCargar")
    bOK.caption = "Cargar": bOK.Left = x + 390: bOK.top = y: bOK.width = 120
    AttachButton bOK

    Set bCancel = EnsureButton(Me, "cmdCancelar")
    bCancel.caption = "Cancelar": bCancel.Left = x + 520: bCancel.top = y: bCancel.width = 120
    AttachButton bCancel

    Me.width = 670
    Me.height = y + 90
End Sub

Private Sub EnsureProgressControls(ByVal fr As MSForms.Frame)
    Dim txtLog As MSForms.TextBox
    Set txtLog = EnsureTextBox(fr, "txtProgLog")
    txtLog.Left = 10: txtLog.top = 18
    txtLog.width = fr.InsideWidth - 20: txtLog.height = 56
    txtLog.Multiline = True: txtLog.Locked = True
    txtLog.ScrollBars = fmScrollBarsVertical
    txtLog.BackColor = RGB(255, 255, 255)

    Dim lbBg     As MSForms.label
    Dim lbFill   As MSForms.label
    Dim lbPct    As MSForms.label
    Dim lbStatus As MSForms.label

    Set lbBg = EnsureLabel(fr, "lblBarBg")
    lbBg.Left = 10: lbBg.top = txtLog.top + txtLog.height + 10
    lbBg.width = fr.InsideWidth - 70: lbBg.height = 12
    lbBg.BackStyle = fmBackStyleOpaque: lbBg.BackColor = RGB(230, 230, 230)
    lbBg.BorderStyle = fmBorderStyleSingle: lbBg.caption = ""

    Set lbFill = EnsureLabel(fr, "lblBar")
    lbFill.Left = lbBg.Left + 1: lbFill.top = lbBg.top + 1
    lbFill.width = 0: lbFill.height = lbBg.height - 2
    lbFill.BackStyle = fmBackStyleOpaque: lbFill.BackColor = RGB(0, 120, 215)
    lbFill.BorderStyle = fmBorderStyleNone: lbFill.caption = ""

    Set lbPct = EnsureLabel(fr, "lblPct")
    lbPct.Left = lbBg.Left + lbBg.width + 10
    lbPct.top = lbBg.top - 2: lbPct.width = 40: lbPct.height = 14
    lbPct.caption = "0%": lbPct.TextAlign = fmTextAlignRight

    Set lbStatus = EnsureLabel(fr, "lblStatus")
    lbStatus.Left = 10: lbStatus.top = lbBg.top + lbBg.height + 8
    lbStatus.width = fr.InsideWidth - 20: lbStatus.height = 14
    lbStatus.caption = ""
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
            .AddItem "Tipo de Cambio"
            .ListIndex = 0
        End With
    End If

    If HasControl("txtMeses") Then Me.Controls("txtMeses").Value = "6"
    If HasControl("txtArchivo") Then Me.Controls("txtArchivo").Value = ""

    If HasControl("cbOperacion") Then Me.Controls("cbOperacion").Enabled = False
    If HasControl("txtMeses") Then Me.Controls("txtMeses").Enabled = False

    SetOperacionOptions

    gSuppressEvents = False   ' <-- aqui, antes del bloque TC

    On Error Resume Next
    If gTCDict Is Nothing Or gTCDict.count = 0 Then
        modPQ_SAB_MC.TryRebuildTCDictFromSheet
    End If
    On Error GoTo 0
    RefreshTCEstado
End Sub

Private Sub SetOperacionOptions()
    If Not HasControl("cbOperacion") Then Exit Sub
    Dim cbOp As MSForms.ComboBox
    Set cbOp = Me.Controls("cbOperacion")
    cbOp.Clear
    cbOp.AddItem "Seleccionar"
    cbOp.AddItem "Movimiento de Caja - Deposito y Retiro"
    cbOp.AddItem "Movimiento de Caja - Solo Deposito"
    cbOp.AddItem "Movimiento de Caja - Solo Retiro"
    cbOp.AddItem "Cambio de Moneda - Compra y Venta"
    cbOp.AddItem "Cambio de Moneda - Solo Compra"
    cbOp.AddItem "Cambio de Moneda - Solo Venta"
    cbOp.ListIndex = 0
End Sub

Private Sub RefreshTCEstado()
    If Not HasControl("lblTCEstado") Then Exit Sub
    Dim lbl As MSForms.label
    Set lbl = Me.Controls("lblTCEstado")
    If Not gTCDict Is Nothing Then
        If gTCDict.count > 0 Then
            lbl.caption = "Tipo de cambio cargado en memoria (" & gTCDict.count & " registros)."
            lbl.ForeColor = RGB(0, 128, 0)
            Exit Sub
        End If
    End If
    lbl.caption = "Sin tipo de cambio en memoria. Las alertas MC usaran montos originales."
    lbl.ForeColor = RGB(160, 100, 0)
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
' Acciones
' ==========================
Public Sub OnExaminar()
    Dim p As String
    p = PickFileXLS("Selecciona el archivo origen")
    If Len(p) > 0 Then
        If HasControl("txtArchivo") Then Me.Controls("txtArchivo").Value = p
    End If
End Sub

Public Sub OnCargar()
    If isRunning Then Exit Sub

    Dim ruta      As String
    Dim mesesSel  As Long
    Dim op        As String
    Dim opU       As String
    Dim tipoCarga As String
    Dim tipoU     As String

    ruta = CStr(Me.Controls("txtArchivo").Value)
    If Len(Trim$(ruta)) = 0 Then
        MsgBox "Selecciona un archivo origen.", vbExclamation
        Exit Sub
    End If
    If Dir(ruta, vbNormal) = "" Then
        MsgBox "El archivo no existe en la ruta indicada." & vbCrLf & ruta, vbExclamation
        Exit Sub
    End If

    tipoCarga = ""
    If HasControl("cbTipoCarga") Then tipoCarga = CStr(Me.Controls("cbTipoCarga").Value)
    tipoU = UCase$(Trim$(tipoCarga))

    If IsPlaceholder(tipoCarga) Then
        MsgBox "Selecciona el tipo de dato a cargar.", vbExclamation
        Exit Sub
    End If

    SetBusy True, "Iniciando carga..."
    BeginProgressHook
    On Error GoTo fallo

    ' ==========================
    ' Caso: Tipo de Cambio
    ' ==========================
    If tipoU = "TIPO DE CAMBIO" Then
        ProgressToCurrent 0.1, "Cargando tipo de cambio..."
        Set gTCDict = modPQ_SAB_MC.LoadTipoCambioDict(ruta)
        If gTCDict Is Nothing Or gTCDict.count = 0 Then
            MsgBox "No se encontraron datos en la hoja TipoCambio del archivo seleccionado." & _
                   vbCrLf & ruta, vbExclamation, "Tipo de Cambio"
            GoTo salir
        End If
        ProgressToCurrent 1, "Tipo de cambio cargado: " & gTCDict.count & " registros."
        EndProgressHook
        SetBusy False, "Tipo de cambio listo."
        RefreshTCEstado
        ' No cerrar: el usuario puede ahora cargar transacciones en la misma sesion
        Exit Sub
    End If

    ' ==========================
    ' Caso: Clientes
    ' ==========================
    If tipoU = "CLIENTES" Then
        ProgressToCurrent 0.05, "Creando consulta de Clientes SAB..."
        Application.Run "CrearQueryClientesSAB", ruta, True
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
        op = CStr(Me.Controls("cbOperacion").Value)
        If IsPlaceholder(op) Then
            MsgBox "Selecciona el tipo de operacion.", vbExclamation
            GoTo salir
        End If
        mesesSel = val(Me.Controls("txtMeses").Value)
        If mesesSel <= 0 Then mesesSel = 6
        opU = UCase$(Trim$(op))

        ' --- Movimiento de Caja ---
        If InStr(1, opU, "MOVIMIENTO", vbTextCompare) > 0 Then
            Dim mcMode As String
            If InStr(1, opU, "SOLO DEP", vbTextCompare) > 0 Then
                mcMode = "SOLO_DEPOSITO"
            ElseIf InStr(1, opU, "SOLO RET", vbTextCompare) > 0 Then
                mcMode = "SOLO_RETIRO"
            Else
                mcMode = "AMBOS"
            End If

            ' Advertir si no hay TC cargado y hay monedas extranjeras posibles
            If gTCDict Is Nothing Or gTCDict.count = 0 Then
                If MsgBox("No hay tipo de cambio cargado en memoria." & vbCrLf & _
                          "Las operaciones en moneda extranjera no tendran monto en soles." & vbCrLf & vbCrLf & _
                          "Continuar de todos modos?", vbQuestion + vbYesNo, "Tipo de Cambio") = vbNo Then
                    GoTo salir
                End If
            End If

            Application.Run "CrearQuerySAB_MC", ruta, mesesSel, mcMode, True

        ' --- Cambio de Moneda ---
        ElseIf InStr(1, opU, "CAMBIO", vbTextCompare) > 0 Then
            Dim cmMode As String
            If InStr(1, opU, "SOLO COMP", vbTextCompare) > 0 Then
                cmMode = "SOLO_COM"
            ElseIf InStr(1, opU, "SOLO VENTA", vbTextCompare) > 0 Or _
                   InStr(1, opU, "SOLO VEN", vbTextCompare) > 0 Then
                cmMode = "SOLO_VEN"
            Else
                cmMode = "AMBOS"
            End If
            ProgressToCurrent 0.05, "Creando consultas SAB - Cambio de Moneda..."
            Application.Run "CrearQuerySAB_CM", ruta, mesesSel, cmMode, True

        Else
            MsgBox "Operacion no reconocida.", vbExclamation
            GoTo salir
        End If

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

    If StrComp(Name, "cbTipoCarga", vbTextCompare) = 0 Then
        Dim v    As String
        Dim tipo As String
        v = ""
        If HasControl("cbTipoCarga") Then v = CStr(Me.Controls("cbTipoCarga").Value)
        tipo = UCase$(Trim$(v))

        Select Case tipo
            Case "TRANSACCIONES"
                If HasControl("cbOperacion") Then Me.Controls("cbOperacion").Enabled = True
                If HasControl("txtMeses") Then Me.Controls("txtMeses").Enabled = True
                If HasControl("lblOperacion") Then Me.Controls("lblOperacion").Enabled = True
                If HasControl("lblMeses") Then Me.Controls("lblMeses").Enabled = True
            Case "CLIENTES", "TIPO DE CAMBIO"
                If HasControl("cbOperacion") Then
                    Me.Controls("cbOperacion").Enabled = False
                    Me.Controls("cbOperacion").Value = "Seleccionar"
                End If
                If HasControl("txtMeses") Then Me.Controls("txtMeses").Enabled = False
                If HasControl("lblOperacion") Then Me.Controls("lblOperacion").Enabled = False
                If HasControl("lblMeses") Then Me.Controls("lblMeses").Enabled = False
            Case Else
                If HasControl("cbOperacion") Then Me.Controls("cbOperacion").Enabled = False
                If HasControl("txtMeses") Then Me.Controls("txtMeses").Enabled = False
        End Select

        SetStatusOnly 0, "Tipo de dato: " & v
        ClearLog
    End If
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

    Dim lbBg     As MSForms.label
    Dim lbFill   As MSForms.label
    Dim lbPct    As MSForms.label
    Dim lbStatus As MSForms.label

    Set lbBg = GetLabelInFrame(fr, "lblBarBg")
    Set lbFill = GetLabelInFrame(fr, "lblBar")
    Set lbPct = GetLabelInFrame(fr, "lblPct")
    Set lbStatus = GetLabelInFrame(fr, "lblStatus")

    If Not lbBg Is Nothing And Not lbFill Is Nothing Then
        Dim wMax As Single
        wMax = lbBg.width - 2
        If wMax < 0 Then wMax = 0
        lbFill.width = wMax * pct
        If pct > 0 And lbFill.width < 1 Then lbFill.width = 1
    End If

    If Not lbPct Is Nothing Then lbPct.caption = Format$(pct, "0%")
    If Not lbStatus Is Nothing Then lbStatus.caption = msg

    Me.Repaint
    DoEvents
    On Error GoTo 0
End Sub

' ==========================
' Log de progreso
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
    Dim startAt  As Long
    Dim out      As String
    parts = Split(s, vbCrLf)
    If UBound(parts) > 15 Then
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
Private Function EnsureLabel(ByVal parent As Object, ByVal nm As String) As MSForms.label
    Dim lb As MSForms.label
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

Private Function GetLabelInFrame(ByVal fr As MSForms.Frame, ByVal nm As String) As MSForms.label
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
    Set dummy = Me.Controls(Name)
    HasControl = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function

