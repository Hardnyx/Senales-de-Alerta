'==========================
'  modUtils.bas  (utilidades)
'==========================
Option Explicit

' ==========================
' Diacrítico NBSP
' ==========================
Public Const NBSP_CODE As Long = 160
Public Function NBSP() As String
    NBSP = ChrW$(NBSP_CODE)
End Function

' ==========================
' App / Excel helpers (freeze con restore)
' ==========================
Private mPrevCalc As XlCalculation
Private mPrevScreen As Boolean
Private mPrevEvents As Boolean
Private mPrevAlerts As Boolean
Private mIsFrozen As Boolean

Public Sub UtilsAppFreeze(ByVal freeze As Boolean)
    On Error Resume Next
    With Application
        If freeze Then
            If Not mIsFrozen Then
                mPrevCalc = .Calculation
                mPrevScreen = .ScreenUpdating
                mPrevEvents = .EnableEvents
                mPrevAlerts = .DisplayAlerts
                mIsFrozen = True
            End If
            .ScreenUpdating = False
            .EnableEvents = False
            .DisplayAlerts = False
            .Calculation = xlCalculationManual
        Else
            If mIsFrozen Then
                .Calculation = mPrevCalc
                .ScreenUpdating = mPrevScreen
                .EnableEvents = mPrevEvents
                .DisplayAlerts = mPrevAlerts
                mIsFrozen = False
            Else
                .ScreenUpdating = True
                .EnableEvents = True
                .DisplayAlerts = True
                .Calculation = xlCalculationAutomatic
            End If
        End If
    End With
    On Error GoTo 0
End Sub

' ==========================
' Archivos / Carpetas
' ==========================
Public Function PickFileXLS(Optional ByVal titulo As String = vbNullString) As String
    Dim p As Variant
    Dim picked As String
    picked = vbNullString

    On Error Resume Next
    With Application.FileDialog(msoFileDialogFilePicker)
        .title = IIf(Len(titulo) > 0, titulo, "Selecciona archivo Excel")
        .Filters.Clear
        .Filters.Add "Excel", "*.xlsx;*.xlsm;*.xlsb;*.xls"
        .AllowMultiSelect = False
        If .Show = -1 Then picked = .SelectedItems(1)
    End With
    On Error GoTo 0

    If Len(picked) = 0 Then
        p = Application.GetOpenFilename( _
                "Archivos Excel (*.xlsx;*.xlsm;*.xlsb;*.xls),*.xlsx;*.xlsm;*.xlsb;*.xls", , _
                IIf(Len(titulo) > 0, titulo, "Selecciona archivo Excel"))
        If VarType(p) = vbString Then picked = CStr(p)
    End If

    PickFileXLS = picked
End Function

Public Function PickFolder(Optional ByVal titulo As String = vbNullString) As String
    Dim picked As String
    picked = vbNullString

    On Error Resume Next
    With Application.FileDialog(msoFileDialogFolderPicker)
        .title = IIf(Len(titulo) > 0, titulo, "Selecciona carpeta")
        If .Show = -1 Then picked = .SelectedItems(1)
    End With
    On Error GoTo 0

    PickFolder = picked
End Function

' ==========================
' Hojas
' ==========================
Public Function UtilsEnsureSheet(ByVal nm As String) As Worksheet
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0

    If sh Is Nothing Then
        Set sh = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Sheets(ThisWorkbook.Sheets.count))
        sh.Name = nm
    End If

    Set UtilsEnsureSheet = sh
End Function

Public Sub UtilsClearSheetButKeepName(ByVal sh As Worksheet)
    Dim lo As ListObject, qt As QueryTable, pt As PivotTable, co As ChartObject

    On Error Resume Next
    For Each pt In sh.PivotTables
        pt.TableRange2.Clear
    Next pt

    For Each co In sh.ChartObjects
        co.Delete
    Next co

    For Each lo In sh.ListObjects
        lo.Delete
    Next lo

    For Each qt In sh.QueryTables
        qt.Delete
    Next qt

    sh.Cells.Clear
    On Error GoTo 0
End Sub

Private Function UtilsSheetExists(ByVal nm As String) As Boolean
    Dim sh As Worksheet
    On Error Resume Next
    Set sh = ThisWorkbook.Worksheets(nm)
    UtilsSheetExists = Not sh Is Nothing
    On Error GoTo 0
End Function

Private Function UtilsSanitizeSheetName(ByVal desired As String) As String
    Dim nm As String
    nm = CStr(desired)

    nm = Replace(nm, "[", "(")
    nm = Replace(nm, "]", ")")
    nm = Replace(nm, ":", " - ")
    nm = Replace(nm, "\", " - ")
    nm = Replace(nm, "/", " - ")
    nm = Replace(nm, "?", " - ")
    nm = Replace(nm, "*", " - ")

    nm = UtilsNormalizeText(nm)
    If Len(nm) = 0 Then nm = "Hoja"

    If Len(nm) > 31 Then nm = Left$(nm, 31)
    UtilsSanitizeSheetName = nm
End Function

Public Sub UtilsRenameSheetSafe(ByVal sh As Worksheet, ByVal desired As String)
    Dim base As String
    Dim nm As String
    Dim k As Long

    base = UtilsSanitizeSheetName(desired)
    nm = base

    On Error Resume Next
    If sh.Name = nm Then Exit Sub
    On Error GoTo 0

    If Not UtilsSheetExists(nm) Then
        On Error Resume Next
        sh.Name = nm
        If Err.Number = 0 Then
            On Error GoTo 0
            Exit Sub
        End If
        Err.Clear
        On Error GoTo 0
    End If

    k = 1
    Do
        k = k + 1
        nm = base
        If Len(nm) > 28 Then nm = Left$(nm, 28)
        nm = nm & "_" & CStr(k)

        If Len(nm) > 31 Then nm = Left$(nm, 31)

        If Not UtilsSheetExists(nm) Then
            On Error Resume Next
            sh.Name = nm
            If Err.Number = 0 Then
                On Error GoTo 0
                Exit Sub
            End If
            Err.Clear
            On Error GoTo 0
        End If
    Loop
End Sub

' ==========================
' Texto / Validación simple
' ==========================
Public Function UtilsIsPlaceholder(ByVal s As String) As Boolean
    Dim t As String
    t = UCase$(Trim$(CStr(s)))
    UtilsIsPlaceholder = (Len(t) = 0) Or (t = "SELECCIONAR")
End Function

Public Function UtilsNormalizeText(ByVal s As String) As String
    Dim t As String
    t = Replace(CStr(s), NBSP(), " ")
    t = Replace(t, vbTab, " ")
    UtilsNormalizeText = Trim$(t)
End Function

' ==========================
' Fechas / Nombres
' ==========================
Public Function UtilsMesAbrevES(ByVal dt As Date) As String
    Dim arr
    arr = Array("ENE", "FEB", "MAR", "ABR", "MAY", "JUN", "JUL", "AGO", "SEP", "OCT", "NOV", "DIC")
    UtilsMesAbrevES = arr(Month(dt) - 1)
End Function

Public Function UtilsParseDateAny(ByVal v As Variant) As Variant
    On Error GoTo fin

    If IsDate(v) Then
        UtilsParseDateAny = CDate(v)
        Exit Function
    End If

    Dim s As String
    s = UtilsNormalizeText(CStr(v))
    If Len(s) = 0 Then GoTo fin

    Dim s2 As String
    s2 = Replace(Replace(s, "-", "/"), ".", "/")

    Dim p() As String
    p = Split(s2, "/")

    If UBound(p) = 2 Then
        Dim a As Long, b As Long, y As Long
        Dim d As Long, m As Long

        a = val(p(0))
        b = val(p(1))
        y = val(p(2))

        If y < 100 And y > 0 Then y = 2000 + y

        d = a: m = b
        If d >= 1 And d <= 31 And m >= 1 And m <= 12 Then
            UtilsParseDateAny = DateSerial(y, m, d)
            Exit Function
        End If

        d = b: m = a
        If d >= 1 And d <= 31 And m >= 1 And m <= 12 Then
            UtilsParseDateAny = DateSerial(y, m, d)
            Exit Function
        End If
    End If

fin:
    UtilsParseDateAny = Empty
End Function

Public Function UtilsMonthSpanSuffix(ByVal ini As Date, ByVal fin As Date) As String
    UtilsMonthSpanSuffix = UtilsMesAbrevES(ini) & "_" & UtilsMesAbrevES(fin) & "_" & Year(fin)
End Function

' ==========================
' Números
' ==========================
Public Function UtilsClamp01(ByVal x As Double) As Double
    If x < 0# Then
        UtilsClamp01 = 0#
    ElseIf x > 1# Then
        UtilsClamp01 = 1#
    Else
        UtilsClamp01 = x
    End If
End Function

Public Function UtilsTryCDblLocale(ByVal v As Variant) As Variant
    On Error GoTo fin

    If IsEmpty(v) Or IsNull(v) Then
        UtilsTryCDblLocale = v
        Exit Function
    End If

    If VarType(v) = vbDouble Or VarType(v) = vbSingle Or VarType(v) = vbCurrency Or VarType(v) = vbInteger Or VarType(v) = vbLong Then
        UtilsTryCDblLocale = CDbl(v)
        Exit Function
    End If

    Dim s As String
    s = UtilsNormalizeText(CStr(v))
    If Len(s) = 0 Then GoTo fin

    Dim neg As Boolean
    neg = False
    If Left$(s, 1) = "(" And Right$(s, 1) = ")" Then
        neg = True
        s = Mid$(s, 2, Len(s) - 2)
        s = UtilsNormalizeText(s)
    End If

    s = Replace(s, "S/", vbNullString)
    s = Replace(s, "$", vbNullString)
    s = Replace(s, "USD", vbNullString)
    s = Replace(s, "PEN", vbNullString)
    s = Replace(s, " ", vbNullString)

    Dim posDot As Long, posCom As Long
    posDot = InStrRev(s, ".")
    posCom = InStrRev(s, ",")

    Dim decSep As String, thouSep As String
    If posDot > 0 And posCom > 0 Then
        If posDot > posCom Then
            decSep = "."
            thouSep = ","
        Else
            decSep = ","
            thouSep = "."
        End If
        s = Replace(s, thouSep, vbNullString)
        s = Replace(s, decSep, Application.International(xlDecimalSeparator))
    ElseIf posCom > 0 And posDot = 0 Then
        s = Replace(s, ",", Application.International(xlDecimalSeparator))
    Else
        If Application.International(xlDecimalSeparator) = "," Then
            s = Replace(s, ".", ",")
        End If
    End If

    Dim d As Double
    d = CDbl(s)
    If neg Then d = -d

    UtilsTryCDblLocale = d
    Exit Function

fin:
    UtilsTryCDblLocale = v
End Function

' ==========================
' Progreso estándar (unificado con modUF_PollProxy)
' ==========================
Public Sub UtilsProgress(ByVal pct As Double, ByVal msg As String)
    On Error Resume Next

    pct = UtilsClamp01(pct)
    Application.StatusBar = msg

    Static lastT As Double
    Dim nowT As Double
    nowT = Timer
    If nowT < lastT Then lastT = 0#

    If (nowT - lastT) < 0.15 Then Exit Sub
    lastT = nowT

    modUF_PollProxy.EnsureAttached
    modUF_PollProxy.ProgressToCurrent pct, msg

    DoEvents
    On Error GoTo 0
End Sub

' ==========================
' Hook de progreso para UserForms (compatibilidad)
' ==========================
Public Sub UF_HookProgress(ByVal uf As Object, Optional ByVal intervalSeconds As Double = 0.25)
    On Error Resume Next
    modUF_PollProxy.Attach uf
    On Error GoTo 0
End Sub

Public Sub UF_UnhookProgress()
    On Error Resume Next
    modUF_PollProxy.Detach
    On Error GoTo 0
End Sub