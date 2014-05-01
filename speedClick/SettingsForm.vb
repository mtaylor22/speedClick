Imports System.Runtime.InteropServices
Imports System.Collections.Specialized

Public Class SettingsForm
    Private Const WH_KEYBOARD_LL As Integer = 13
    Private Const WM_KEYUP As Integer = &H101
    Private WM_KEYDOWN As Integer = &H100

    Private Const WM_SYSKEYUP As Integer = &H105
    Private Const WM_SYSKEYDOWN As Integer = &H104
    Private proc As LowLevelKeyboardProcDelegate = AddressOf HookCallback
    Private hookID As IntPtr

    Private Delegate Function LowLevelKeyboardProcDelegate(ByVal nCode As Integer, ByVal wParam As IntPtr, _
        ByVal lParam As IntPtr) As IntPtr

    <DllImport("user32")> _
    Private Shared Function SetWindowsHookEx(ByVal idHook As Integer, ByVal lpfn As LowLevelKeyboardProcDelegate, _
        ByVal hMod As IntPtr, ByVal dwThreadId As UInteger) As IntPtr
    End Function

    <DllImport("user32.dll")> _
    Private Shared Function UnhookWindowsHookEx(ByVal hhk As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("user32.dll")> _
    Private Shared Function CallNextHookEx(ByVal hhk As IntPtr, ByVal nCode As Integer, ByVal wParam As IntPtr, _
        ByVal lParam As IntPtr) As IntPtr
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Unicode)> _
    Private Shared Function GetModuleHandle(ByVal lpModuleName As String) As IntPtr
    End Function

    Sub New()
        InitializeComponent()
        hookID = SetHook(proc)
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As FormClosingEventArgs) Handles Me.FormClosing
        UnhookWindowsHookEx(hookID)
    End Sub

    Private Function SetHook(ByVal proc As LowLevelKeyboardProcDelegate) As IntPtr
        Using curProcess As Process = Process.GetCurrentProcess()
            Using curModule As ProcessModule = curProcess.MainModule
                Return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0)
            End Using
        End Using
    End Function
    Public Shared keycombo As New List(Of Integer)
    Public Shared keycombo_actual As New List(Of Integer)
    Private Sub UpdateKeyCombo()
        If activatedtb IsNot (Nothing) Then
            activatedtb.Text = ""
            Dim kc As New System.Windows.Forms.KeysConverter
            For Each key As Integer In keycombo
                activatedtb.Text += kc.ConvertToString(key) + " "
            Next
        End If
    End Sub
    Private Function setequality(a As List(Of Integer), b As List(Of Integer)) As Boolean
        Dim aset = New HashSet(Of Integer)(a)
        Dim bset = New HashSet(Of Integer)(b)
        Return aset.SetEquals(bset)
    End Function

    Private Function HookCallback(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
        If nCode >= 0 AndAlso (wParam.ToInt32 = WM_KEYDOWN OrElse wParam.ToInt32 = WM_SYSKEYDOWN) Then
            Dim vkCode As Integer = Marshal.ReadInt32(lParam)
            If Not keycombo_actual.Contains(vkCode) Then
                If Not setequality(keycombo, keycombo_actual) Then
                    keycombo.Clear()
                    keycombo_actual.Clear()
                End If
                keycombo_actual.Add(vkCode)
                keycombo.Add(vkCode)
            End If
            UpdateKeyCombo()
        End If
        If nCode >= 0 AndAlso (wParam.ToInt32 = WM_KEYUP OrElse wParam.ToInt32 = WM_SYSKEYUP) Then
            Dim vkCode As Integer = Marshal.ReadInt32(lParam)
            If keycombo_actual.Contains(vkCode) Then keycombo_actual.Remove(vkCode)
            UpdateKeyCombo()
        End If
        Return CallNextHookEx(hookID, nCode, wParam, lParam)
    End Function
    Private Sub SettingsForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim sc As StringCollection = My.Settings.hotkeya
        If Not IsNothing(sc) Then
            Dim sc_list As List(Of String) = sc.Cast(Of String)().ToList()
            Dim sc_intlist As List(Of Integer) = sc_list.ConvertAll(Function(str) Int32.Parse(str))
            keycombo = sc_intlist
            activatedtb = TextBox1
            UpdateKeyCombo()
            activatedtb = Nothing
        End If
        sc = My.Settings.hotkeyb
        If Not IsNothing(sc) Then
            Dim sc_list As List(Of String) = sc.Cast(Of String)().ToList()
            Dim sc_intlist As List(Of Integer) = sc_list.ConvertAll(Function(str) Int32.Parse(str))
            keycombo = sc_intlist
            activatedtb = TextBox2
            UpdateKeyCombo()
            activatedtb = Nothing
        End If
    End Sub
    Private Shared activatedtb As TextBox = Nothing
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Button1.Text.Equals("Set Hotkey (Activate)") Then
            Button1.Text = " (SET) "
            activatedtb = TextBox1
        ElseIf Button1.Text.Equals(" (SET) ") Then
            Dim intList = keycombo.ConvertAll(Function(str) str.ToString)
            Dim sc = New StringCollection
            sc.AddRange(intList.ToArray())
            My.Settings.hotkeya = sc
            My.Settings.Save()
            Button1.Text = "Set Hotkey (Activate)"
            activatedtb = Nothing
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Button2.Text.Equals("Set Hotkey (Reset)") Then
            Button2.Text = " (SET) "
            activatedtb = TextBox2
        ElseIf Button2.Text.Equals(" (SET) ") Then
            Dim intList = keycombo.ConvertAll(Function(str) str.ToString)
            Dim sc = New StringCollection
            sc.AddRange(intList.ToArray())
            My.Settings.hotkeyb = sc
            My.Settings.Save()
            Button2.Text = "Set Hotkey (Reset)"
            activatedtb = Nothing
        End If
    End Sub
End Class