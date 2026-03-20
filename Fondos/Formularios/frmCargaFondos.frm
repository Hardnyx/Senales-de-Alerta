'==========================
' UserForm: frmCargaFondos
'==========================
Option Explicit

Private gHandlers As Collection
Private isRunning As Boolean
Private gStage As String
Private gSuppressEvents As Boolean

'==========================
' Constantes de UI
'==========================
Private Const TIPO_SEL As String = "Seleccionar"
Private Const TIPO_TRANS As String = "Transacciones"
Private Const TIPO_CLIENTES As String = "Clientes"

Private Const OP_SEL As String = "Seleccionar"
Private Const ORIGEN_FIXED As String = "FONDOS"

Private OP_FON_SUSC As String
Private OP_FON_RESC As String
Private OP_CLI_FON As String

'==========================
' API de progreso (usada por modUF_PollProxy)
' pct: 0 a 1
'==========================
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

    Dim lbBg As MSForms.Label
    Dim lbFill As MSForms.Label
    Dim lbPct As MSForms.Label
    Dim lbStatus As MSForms.Label

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

'==========================
' Callbacks que espera CCtrlEvents
'==========================
Public Sub HandleButtonClick(ByVal ctrlName As String)
    OnButtonClick ctrlName
End Sub

Public Sub HandleComboChange(ByVal ctrlName As String)
    OnComboChanged ctrlName
End Sub

Public Sub OnButtonClick(ByVal ctrlName As String)
    If gSuppressEvents Then Exit Sub

    Select Case ctrlName
        Case "cmdExaminar": OnExaminar
        Case "cmdCargar":   OnCargar
        Case "cmdCancelar": OnCancelar
    End Select
End Sub

Public Sub OnComboChanged(ByVal ctrlName As String)
    If gSuppressEvents Then Exit Sub

    Select Case ctrlName
        Case "cbTipoCarga"
            HandleTipoChanged
        Case "cbOperacion"
            SetStatusOnly 0, "Operaci" & Chr(243) & "n: " & CStr(Me.Controls("cbOperacion").Value)
    End Select
End Sub

'==========================
' Ciclo de vida del form
'==========================
Private Sub UserForm_Initialize()
    OP_FON_SUSC = "Fondos: Suscripci" & Chr(243) & "n"
    OP_FON_RESC = "Fondos: Rescate"
    OP_CLI_FON = "Clientes: Fondos"

    Set gHandlers = New Collection

    gSuppressEvents = True
    BuildOrRefreshUI
    AttachHooksIfNeeded
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

'==========================
' UX: Busy state
'==========================
Private Sub SetBusy(ByVal running As Boolean, Optional ByVal statusMsg As String = "")
    isRunning = running

    If HasControl("cmdCargar") Then Me.Controls("cmdCargar").Enabled = Not running
    If HasControl("cmdExaminar") Then Me.Controls("cmdExaminar").Enabled = Not running
    If HasControl("cbTipoCarga") Then Me.Controls("cbTipoCarga").Enabled = Not running

    If HasControl("cbOperacion") Then
        Me.Controls("cbOperacion").Enabled = Not running And Not IsPlaceholder(CStr(Me.Controls("cbTipoCarga").Value))
    End If

    If HasControl("txtMeses") Then
        Me.Controls("txtMeses").Enabled = Not running And _
            (UCase$(Trim$(CStr(Me.Controls("cbTipoCarga").Value))) = "TRANSACCIONES")
    End If
    If HasControl("lblMeses") Then
        Me.Controls("lblMeses").Enabled = Not running And _
            (UCase$(Trim$(CStr(Me.Controls("cbTipoCarga").Value))) = "TRANSACCIONES")
    End If

    If HasControl("txtArchivo") Then Me.Controls("txtArchivo").Enabled = Not running

    If HasControl("cmdCancelar") Then
        Me.Controls("cmdCancelar").caption = IIf(running, "Cerrar", "Cancelar")
    End If

    Me.MousePointer = IIf(running, fmMousePointerHourGlass, fmMousePointerDefault)

    If Len(Trim$(statusMsg)) > 0 Then
        SetStatusOnly IIf(running, 0.01, 0), statusMsg
    End If
End Sub

'==========================
' UI: crear o reutilizar controles
'==========================
Private Sub BuildOrRefreshUI()
    Me.caption = "Cargar Datos"
    Me.StartUpPosition = 1

    Dim x As Single, y As Single
    Dim lblW As Single, inX As Single, inW As Single
    Dim btnW As Single, gap As Single, pad As Single

    pad = 12
    gap = 10
    x = pad
    y = pad

    lblW = 130
    inX = x + lblW + gap
    inW = 600
    btnW = 120

    Dim L As MSForms.Label
    Dim t As MSForms.TextBox
    Dim cb As MSForms.ComboBox
    Dim b As MSForms.CommandButton
    Dim fr As MSForms.Frame

    Set L = EnsureLabel(Me, "lblTitulo")
    L.caption = "Cargar datos: Transacciones / Clientes"
    L.Left = x
    L.top = y
    L.width = inX + inW + btnW - x
    L.Font.Bold = True
    L.Font.Size = 12
    y = y + 26

    Set L = EnsureLabel(Me, "lblTipoCarga")
    L.caption = "Tipo de dato:"
    L.Left = x
    L.top = y
    L.width = lblW

    Set cb = EnsureCombo(Me, "cbTipoCarga")
    cb.Left = inX
    cb.top = y - 3
    cb.width = inW + btnW
    cb.Style = fmStyleDropDownList
    cb.ControlTipText = "Elige si cargar transacciones o clientes."
    y = y + 30

    Set L = EnsureLabel(Me, "lblOperacion")
    L.caption = "Tipo de operaci" & Chr(243) & "n:"
    L.Left = x
    L.top = y
    L.width = lblW

    Set cb = EnsureCombo(Me, "cbOperacion")
    cb.Left = inX
    cb.top = y - 3
    cb.width = inW + btnW
    cb.Style = fmStyleDropDownList
    cb.ControlTipText = "Operaci" & Chr(243) & "n a cargar."
    y = y + 30

    Set L = EnsureLabel(Me, "lblMeses")
    L.caption = Chr(218) & "ltimos meses:"
    L.Left = x
    L.top = y
    L.width = lblW

    Set t = EnsureTextBox(Me, "txtMeses")
    t.Left = inX
    t.top = y - 3
    t.width = 80
    t.ControlTipText = "Cantidad de meses (por defecto 6)."
    y = y + 30

    Set L = EnsureLabel(Me, "lblArchivo")
    L.caption = "Archivo origen:"
    L.Left = x
    L.top = y
    L.width = lblW

    Set t = EnsureTextBox(Me, "txtArchivo")
    t.Left = inX
    t.top = y - 3
    t.width = inW
    t.ControlTipText = "Ruta del archivo."

    Set b = EnsureButton(Me, "cmdExaminar")
    b.caption = "Examinar..."
    b.Left = inX + inW + gap
    b.top = y - 5
    b.width = btnW
    b.ControlTipText = "Buscar archivo."
    y = y + 36

    Set fr = EnsureFrame(Me, "fraProg")
    fr.caption = " Progreso"
    fr.Left = x
    fr.top = y
    fr.width = inX + inW + btnW - x
    fr.height = 220

    EnsureProgressControls fr
    y = y + fr.height + 12

    Dim bOK As MSForms.CommandButton
    Dim bCancel As MSForms.CommandButton

    Set bOK = EnsureButton(Me, "cmdCargar")
    bOK.caption = "Cargar"
    bOK.width = 120
    bOK.Left = fr.Left + fr.width - (120 + 140)
    bOK.top = y

    Set bCancel = EnsureButton(Me, "cmdCancelar")
    bCancel.caption = "Cancelar"
    bCancel.width = 120
    bCancel.Left = fr.Left + fr.width - 120
    bCancel.top = y

    Me.width = fr.Left + fr.width + pad + 6
    Me.height = y + 90
End Sub

Private Sub AttachHooksIfNeeded()
    On Error Resume Next
    If HasControl("cmdExaminar") Then AttachButton Me.Controls("cmdExaminar")
    If HasControl("cmdCargar") Then AttachButton Me.Controls("cmdCargar")
    If HasControl("cmdCancelar") Then AttachButton Me.Controls("cmdCancelar")
    If HasControl("cbTipoCarga") Then AttachCombo Me.Controls("cbTipoCarga")
    If HasControl("cbOperacion") Then AttachCombo Me.Controls("cbOperacion")
    On Error GoTo 0
End Sub

Private Sub EnsureProgressControls(ByVal fr As MSForms.Frame)
    Dim txtLog As MSForms.TextBox
    Set txtLog = EnsureTextBox(fr, "txtProgLog")
    txtLog.Left = 10
    txtLog.top = 18
    txtLog.width = fr.InsideWidth - 20
    txtLog.height = 110
    txtLog.Multiline = True
    txtLog.Locked = True
    txtLog.ScrollBars = fmScrollBarsVertical
    txtLog.BackColor = RGB(255, 255, 255)

    Dim lbBg As MSForms.Label
    Dim lbFill As MSForms.Label
    Dim lbPct As MSForms.Label
    Dim lbStatus As MSForms.Label

    Set lbBg = EnsureLabel(fr, "lblBarBg")
    lbBg.Left = 10
    lbBg.top = txtLog.top + txtLog.height + 10
    lbBg.width = fr.InsideWidth - 70
    lbBg.height = 12
    lbBg.BackStyle = fmBackStyleOpaque
    lbBg.BackColor = RGB(230, 230, 230)
    lbBg.BorderStyle = fmBorderStyleSingle
    lbBg.caption = ""

    Set lbFill = EnsureLabel(fr, "lblBar")
    lbFill.Left = lbBg.Left + 1
    lbFill.top = lbBg.top + 1
    lbFill.width = 0
    lbFill.height = lbBg.height - 2
    lbFill.BackStyle = fmBackStyleOpaque
    lbFill.BackColor = RGB(0, 120, 215)
    lbFill.BorderStyle = fmBorderStyleNone
    lbFill.caption = ""

    Set lbPct = EnsureLabel(fr, "lblPct")
    lbPct.Left = lbBg.Left + lbBg.width + 10
    lbPct.top = lbBg.top - 2
    lbPct.width = 40
    lbPct.height = 14
    lbPct.caption = "0%"
    lbPct.TextAlign = fmTextAlignRight

    Set lbStatus = EnsureLabel(fr, "lblStatus")
    lbStatus.Left = 10
    lbStatus.top = lbBg.top + lbBg.height + 10
    lbStatus.width = fr.InsideWidth - 20
    lbStatus.height = 14
    lbStatus.caption = ""
End Sub

'==========================
' Inicializacion combos
'==========================
Private Sub InitCombosDefaults()
    gSuppressEvents = True

    If HasControl("cbTipoCarga") Then
        With Me.Controls("cbTipoCarga")
            .Clear
            .AddItem TIPO_SEL
            .AddItem TIPO_TRANS
            .AddItem TIPO_CLIENTES
            .ListIndex = 0
        End With
    End If

    If HasControl("cbOperacion") Then
        With Me.Controls("cbOperacion")
            .Clear
            .AddItem OP_SEL
            .ListIndex = 0
            .Enabled = False
        End With
    End If

    If HasControl("txtMeses") Then
        Me.Controls("txtMeses").Value = "6"
        Me.Controls("txtMeses").Enabled = False
    End If

    If HasControl("lblMeses") Then Me.Controls("lblMeses").Enabled = False
    If HasControl("txtArchivo") Then Me.Controls("txtArchivo").Value = ""

    gSuppressEvents = False
End Sub

Private Sub HandleTipoChanged()
    Dim tipoCarga As String
    Dim tipoU As String
    tipoCarga = CStr(Me.Controls("cbTipoCarga").Value)
    tipoU = UCase$(Trim$(tipoCarga))

    gSuppressEvents = True

    If tipoU = "TRANSACCIONES" Then
        If HasControl("cbOperacion") Then Me.Controls("cbOperacion").Enabled = True
        If HasControl("txtMeses") Then Me.Controls("txtMeses").Enabled = True
        If HasControl("lblMeses") Then Me.Controls("lblMeses").Enabled = True
        SetOperacionOptionsByTipo "TRANSACCIONES"

    ElseIf tipoU = "CLIENTES" Then
        If HasControl("cbOperacion") Then Me.Controls("cbOperacion").Enabled = True
        If HasControl("txtMeses") Then Me.Controls("txtMeses").Enabled = False
        If HasControl("lblMeses") Then Me.Controls("lblMeses").Enabled = False
        SetOperacionOptionsByTipo "CLIENTES"

    Else
        If HasControl("cbOperacion") Then
            With Me.Controls("cbOperacion")
                .Enabled = False
                .Clear
                .AddItem OP_SEL
                .ListIndex = 0
            End With
        End If
        If HasControl("txtMeses") Then Me.Controls("txtMeses").Enabled = False
        If HasControl("lblMeses") Then Me.Controls("lblMeses").Enabled = False
    End If

    gSuppressEvents = False

    SetStatusOnly 0, "Tipo de dato: " & tipoCarga
    ClearLog
End Sub

Private Sub SetOperacionOptionsByTipo(ByVal tipoU As String)
    If Not HasControl("cbOperacion") Then Exit Sub

    Dim cbOp As MSForms.ComboBox
    Set cbOp = Me.Controls("cbOperacion")

    cbOp.Clear
    cbOp.AddItem OP_SEL

    If tipoU = "TRANSACCIONES" Then
        cbOp.AddItem OP_FON_SUSC
        cbOp.AddItem OP_FON_RESC
    ElseIf tipoU = "CLIENTES" Then
        cbOp.AddItem OP_CLI_FON
    End If

    cbOp.ListIndex = 0
End Sub

Private Function IsPlaceholder(ByVal s As String) As Boolean
    Dim u As String
    u = UCase$(Trim$(s))
    IsPlaceholder = (Len(u) = 0) Or (u = UCase$(TIPO_SEL)) Or (u = UCase$(OP_SEL))
End Function

'==========================
' Progreso: inicio/fin
'==========================
Private Sub BeginProgressHook()
    ClearLog
    gStage = vbNullString
    SetStatusOnly 0, "Inicializando..."
End Sub

Private Sub EndProgressHook()
    On Error Resume Next
    modUF_PollProxy.Detach
    Application.StatusBar = False
    On Error GoTo 0
End Sub

'==========================
' Acciones
'==========================
Public Sub OnExaminar()
    Dim p As String
    p = PickFileXLS("Selecciona el archivo origen")
    If Len(p) > 0 Then Me.Controls("txtArchivo").Value = p
End Sub

Public Sub OnCargar()
    If isRunning Then Exit Sub

    Dim ruta As String
    ruta = CStr(Me.Controls("txtArchivo").Value)

    If Len(Trim$(ruta)) = 0 Then
        MsgBox "Selecciona un archivo origen.", vbExclamation
        Exit Sub
    End If
    If Dir(ruta, vbNormal) = "" Then
        MsgBox "El archivo no existe en la ruta indicada." & vbCrLf & ruta, vbExclamation
        Exit Sub
    End If

    Dim dotPos As Long
    dotPos = InStrRev(ruta, ".")
    Dim ext As String
    If dotPos > 0 Then
        ext = LCase$(Trim$(Mid$(ruta, dotPos + 1)))
    Else
        ext = vbNullString
    End If
    If ext <> "xlsx" And ext <> "xls" And ext <> "xlsm" And ext <> "xlsb" Then
        MsgBox "El archivo debe ser un libro de Excel (.xlsx, .xls, .xlsm, .xlsb)." & vbCrLf & ruta, vbExclamation
        Exit Sub
    End If

    Dim tipoCarga As String
    Dim tipoU As String
    tipoCarga = CStr(Me.Controls("cbTipoCarga").Value)
    tipoU = UCase$(Trim$(tipoCarga))

    If IsPlaceholder(tipoCarga) Then
        MsgBox "Selecciona el tipo de dato.", vbExclamation
        Exit Sub
    End If

    Dim op As String
    op = CStr(Me.Controls("cbOperacion").Value)
    If IsPlaceholder(op) Then
        MsgBox "Selecciona el tipo de operaci" & Chr(243) & "n.", vbExclamation
        Exit Sub
    End If

    Dim mesesVal As String
    mesesVal = Trim$(CStr(Me.Controls("txtMeses").Value))
    Dim mesesSel As Long
    If Len(mesesVal) = 0 Or Not IsNumeric(mesesVal) Then
        mesesSel = 6
    Else
        mesesSel = CLng(CDbl(mesesVal))
    End If
    If mesesSel <= 0 Then mesesSel = 6

    On Error GoTo fallo

    SetBusy True, "Iniciando carga..."
    BeginProgressHook

    If tipoU = "CLIENTES" Then
        If StrComp(op, OP_CLI_FON, vbTextCompare) = 0 Then
            ProgressToCurrent 0.05, "Creando consulta de Clientes Fondos..."
            Application.Run "CrearQueryClientesFondos", ruta, True
        Else
            MsgBox "Operaci" & Chr(243) & "n no v" & Chr(225) & "lida para Clientes.", vbExclamation
            GoTo salir
        End If
        ProgressToCurrent 1, "Carga completada."
        GoTo ok
    End If

    If tipoU = "TRANSACCIONES" Then
        If StrComp(op, OP_FON_SUSC, vbTextCompare) = 0 Then
            ProgressToCurrent 0.05, "Creando consultas de Fondos (Suscripci" & Chr(243) & "n)..."
            ' showProg=False: los MsgBox intermedios de etapa se suprimen;
            ' el progreso se muestra en el log y barra del formulario.
            Application.Run "CrearQueryFondos", ruta, mesesSel, False, True, ORIGEN_FIXED, False

        ElseIf StrComp(op, OP_FON_RESC, vbTextCompare) = 0 Then
            ProgressToCurrent 0.05, "Creando consultas de Fondos (Rescate)..."
            Application.Run "CrearQueryFondos", ruta, mesesSel, True, True, ORIGEN_FIXED, False

        Else
            MsgBox "Operaci" & Chr(243) & "n no v" & Chr(225) & "lida para Transacciones.", vbExclamation
            GoTo salir
        End If
        ProgressToCurrent 1, "Carga completada."
        GoTo ok
    End If

    MsgBox "Tipo de dato no reconocido.", vbExclamation
    GoTo salir

ok:
    EndProgressHook
    SetBusy False, "Listo."
    Unload Me
    Exit Sub

salir:
    EndProgressHook
    SetBusy False, "Listo."
    Exit Sub

fallo:
    Dim errNum As Long
    Dim errDesc As String
    Dim errSrc As String

    errNum = Err.Number
    errDesc = Err.Description
    errSrc = Err.Source

    EndProgressHook
    SetBusy False, "Listo."
    SetStatusOnly 0, "Error al cargar."
    ShowErrorDetails ruta, tipoCarga, op, mesesSel, errNum, errDesc, errSrc
End Sub

Public Sub OnCancelar()
    If isRunning Then
        If MsgBox("Hay una operaci" & Chr(243) & "n en progreso. " & Chr(191) & "Deseas cerrar de todos modos?", _
                  vbQuestion + vbYesNo) = vbNo Then Exit Sub
        EndProgressHook
        Unload Me
        Exit Sub
    End If

    EndProgressHook
    Unload Me
End Sub

Private Sub ShowErrorDetails(ByVal ruta As String, ByVal tipoCarga As String, ByVal op As String, ByVal mesesSel As Long, _
                             ByVal errNum As Long, ByVal errDesc As String, ByVal errSrc As String)
    Dim desc As String
    desc = errDesc
    If Len(Trim$(desc)) = 0 Then desc = "(sin descripci" & Chr(243) & "n)"

    Dim src As String
    src = errSrc
    If Len(Trim$(src)) = 0 Then src = "(sin source)"

    Dim msg As String
    msg = "Error " & errNum & vbCrLf & _
          desc & vbCrLf & vbCrLf & _
          "Source: " & src & vbCrLf & _
          "Estado: " & gStage & vbCrLf & _
          "Tipo: " & tipoCarga & vbCrLf & _
          "Operaci" & Chr(243) & "n: " & op & vbCrLf & _
          "Meses: " & CStr(mesesSel) & vbCrLf & _
          "Archivo: " & ruta

    AppendLogLine "ERROR " & errNum & ": " & desc
    MsgBox msg, vbCritical
End Sub

'==========================
' Status sin escribir en el log
'==========================
Private Sub SetStatusOnly(ByVal pct As Double, ByVal msg As String)
    On Error Resume Next
    If pct < 0 Then pct = 0
    If pct > 1 Then pct = 1

    Application.StatusBar = msg

    Dim fr As MSForms.Frame
    Set fr = GetFrameOrNothing("fraProg")
    If fr Is Nothing Then Exit Sub

    Dim lbBg As MSForms.Label
    Dim lbFill As MSForms.Label
    Dim lbPct As MSForms.Label
    Dim lbStatus As MSForms.Label

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

'==========================
' Log de progreso
'==========================
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

    Dim parts() As String
    Dim i As Long
    Dim startAt As Long
    Dim out As String

    parts = Split(s, vbCrLf)
    If UBound(parts) > 25 Then
        startAt = UBound(parts) - 25
        out = vbNullString
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

'==========================
' Helpers: Ensure controls
'==========================
Private Function EnsureLabel(ByVal parent As Object, ByVal nm As String) As MSForms.Label
    Dim lb As MSForms.Label
    On Error Resume Next
    Set lb = parent.Controls(nm)
    On Error GoTo 0
    If lb Is Nothing Then Set lb = parent.Controls.Add("Forms.Label.1", nm, True)
    Set EnsureLabel = lb
End Function

Private Function EnsureTextBox(ByVal parent As Object, ByVal nm As String) As MSForms.TextBox
    Dim tb As MSForms.TextBox
    On Error Resume Next
    Set tb = parent.Controls(nm)
    On Error GoTo 0
    If tb Is Nothing Then Set tb = parent.Controls.Add("Forms.TextBox.1", nm, True)
    Set EnsureTextBox = tb
End Function

Private Function EnsureCombo(ByVal parent As Object, ByVal nm As String) As MSForms.ComboBox
    Dim cb As MSForms.ComboBox
    On Error Resume Next
    Set cb = parent.Controls(nm)
    On Error GoTo 0
    If cb Is Nothing Then Set cb = parent.Controls.Add("Forms.ComboBox.1", nm, True)
    Set EnsureCombo = cb
End Function

Private Function EnsureButton(ByVal parent As Object, ByVal nm As String) As MSForms.CommandButton
    Dim b As MSForms.CommandButton
    On Error Resume Next
    Set b = parent.Controls(nm)
    On Error GoTo 0
    If b Is Nothing Then Set b = parent.Controls.Add("Forms.CommandButton.1", nm, True)
    Set EnsureButton = b
End Function

Private Function EnsureFrame(ByVal parent As Object, ByVal nm As String) As MSForms.Frame
    Dim fr As MSForms.Frame
    On Error Resume Next
    Set fr = parent.Controls(nm)
    On Error GoTo 0
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
    On Error Resume Next
    If LCase$(Trim$(CStr(b.Tag))) = "hooked" Then Exit Sub
    b.Tag = "hooked"
    On Error GoTo 0

    Dim h As CCtrlEvents
    Set h = New CCtrlEvents
    h.HookButton b, Me
    gHandlers.Add h
End Sub

Private Sub AttachCombo(ByVal c As MSForms.ComboBox)
    On Error Resume Next
    If LCase$(Trim$(CStr(c.Tag))) = "hooked" Then Exit Sub
    c.Tag = "hooked"
    On Error GoTo 0

    Dim h As CCtrlEvents
    Set h = New CCtrlEvents
    h.HookCombo c, Me
    gHandlers.Add h
End Sub

Private Function HasControl(ByVal name As String) As Boolean
    Dim dummy As Object
    On Error Resume Next
    Set dummy = Me.Controls(name)
    HasControl = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0
End Function


