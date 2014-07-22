Option Strict On
Option Explicit On

Imports System.Runtime.InteropServices
Imports System.Drawing.Drawing2D
Imports Microsoft.VisualBasic.PowerPacks
Imports System.IO
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Collections.Specialized


Public Class ActionForm
    <System.Serializable()>
    Public Class dot
        Public type As Integer
        Public e As Point
        Public Sub New(p As Point, t As Integer)
            e = p
            type = t
            oval = DrawCircle(p, If(type = 1, Color.Blue, Color.Black))
        End Sub

        Public Function location() As Point
            ' A sad, makeshift deserialization constructor
            If oval Is Nothing Then oval = DrawCircle(e, If(type = 1, Color.Blue, Color.Black))
            Return oval.Location
        End Function
        <System.NonSerialized>
        Public oval As OvalShape
    End Class
    Public Declare Function SetCursorPos Lib "user32" (ByVal x As UInteger, ByVal y As UInteger) As Long
    Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As UInteger, ByVal dx As UInteger, ByVal dy As UInteger, ByVal cButtons As UInteger, ByVal dwExtraInfo As UInteger)
    Public Const MOUSEEVENTF_LEFTDOWN = &H2
    Public Const MOUSEEVENTF_LEFTUP = &H4

    Public Sub LeftClick()
        LeftDown()
        LeftUp()
    End Sub

    Public Sub LeftDown()
        mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0)
    End Sub
    Public Sub LeftUp()
        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0)
    End Sub

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
        NotifyIcon1.Visible = False
        NotifyIcon1.Dispose()
    End Sub

    Private Function SetHook(ByVal proc As LowLevelKeyboardProcDelegate) As IntPtr
        Using curProcess As Process = Process.GetCurrentProcess()
            Using curModule As ProcessModule = curProcess.MainModule
                Return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0)
            End Using
        End Using
    End Function
    Private Function setequality(a As List(Of Integer), b As List(Of Integer)) As Boolean
        Dim aset = New HashSet(Of Integer)(a)
        Dim bset = New HashSet(Of Integer)(b)
        Return aset.SetEquals(bset)
    End Function
    Private Function HookCallback(ByVal nCode As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
        If nCode >= 0 AndAlso (wParam.ToInt32 = WM_KEYDOWN OrElse wParam.ToInt32 = WM_SYSKEYDOWN) Then
            Dim vkCode As Integer = Marshal.ReadInt32(lParam)
            If Not current_key.Contains(vkCode) Then current_key.Add(vkCode)
            If Not active And setequality(current_key, activate_key) Then
                Me.Visible = True
                mp = MousePosition()
                active = True
            ElseIf Not active And setequality(current_key, reset_key) Then
                Me.Visible = True
                active = False
                For Each vec In pvec
                    vec.oval.Dispose()
                Next
                pvec.Clear()
            End If
        End If
        If nCode >= 0 AndAlso (wParam.ToInt32 = WM_KEYUP OrElse wParam.ToInt32 = WM_SYSKEYUP) Then
            Dim vkCode As Integer = Marshal.ReadInt32(lParam)
            If current_key.Contains(vkCode) Then current_key.Remove(vkCode)
            If active Then
                If activate_key.Contains(vkCode) Then
                    Me.Visible = False
                    active = False
                    If (active_point IsNot Nothing) Then
                        active_point.oval.BorderColor = If(active_point.type = 1, Color.Blue, Color.Black)
                        move_mouse(active_point.location)
                        If active_point.type = 0 Then
                            LeftClick()
                        End If
                    End If
                    active_point = Nothing
                End If
            End If
        End If
        Return CallNextHookEx(hookID, nCode, wParam, lParam)
    End Function

    Public Shared pvec As List(Of dot)
    Public Shared active As Boolean = False
    Public Shared mp As Point
    Public Shared active_point As dot = Nothing
    Public Shared activate_key As New List(Of Integer)
    Public Shared reset_key As New List(Of Integer)
    Public Shared current_key As New List(Of Integer)
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Opacity = 0.2
        Me.Width = My.Computer.Screen.WorkingArea.Width
        Me.Height = Screen.PrimaryScreen.Bounds.Height
        Me.CenterToScreen()
        Me.WindowState = FormWindowState.Maximized
        pvec = New List(Of dot)
        Me.TopMost = True
        ' Get keystroke stuff
        My.Settings.Reload()
        Dim sc As StringCollection = My.Settings.hotkeya
        If Not IsNothing(sc) Then
            Dim sc_list As List(Of String) = sc.Cast(Of String)().ToList()
            Dim sc_intlist As List(Of Integer) = sc_list.ConvertAll(Function(str) Int32.Parse(str))
            activate_key = sc_intlist
        Else
            activate_key = New List(Of Integer)()
            activate_key.Add(Keys.D)
            activate_key.Add(Keys.LControlKey)
        End If
        Dim zc As StringCollection = My.Settings.hotkeyb
        If Not IsNothing(zc) Then
            Dim sc_list As List(Of String) = zc.Cast(Of String)().ToList()
            Dim sc_intlist As List(Of Integer) = sc_list.ConvertAll(Function(str) Int32.Parse(str))
            reset_key = sc_intlist
        Else
            reset_key = New List(Of Integer)()
            reset_key.Add(Keys.R)
            reset_key.Add(Keys.LControlKey)
        End If



    End Sub
    Private Sub Form1_MouseClick(sender As Object, e As MouseEventArgs) Handles Me.MouseClick
        Dim rightclick As Boolean = (e.Button = Windows.Forms.MouseButtons.Right)
        Dim d = New dot(e.Location, If(rightclick, 0, 1))
        pvec.Add(d)
    End Sub
    Private Sub move_mouse(e As Point)
        e = Me.PointToScreen(consider_radius(New Point(e.X, e.Y)))
        SetCursorPos(CUInt(e.X), CUInt(e.Y))
    End Sub
    Private Function consider_radius(e As Point) As Point
        e.X += radius
        e.Y += radius
        Return e
    End Function
    Private Sub Form1_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        If active Then
            Dim min_difference As Double = 361
            Dim min_point As dot = Nothing
            For Each vec As dot In pvec
                If Math.Abs(gen_angle(e.Location, mp) - gen_angle(vec.location, mp)) < min_difference Then
                    min_difference = Math.Abs(gen_angle(e.Location, mp) - gen_angle(consider_radius(vec.location), mp))
                    min_point = vec
                End If
            Next
            If min_point IsNot (Nothing) Then
                If (active_point IsNot Nothing And active_point IsNot (min_point)) Then active_point.oval.BorderColor = If(active_point.type = 1, Color.Blue, Color.Black)
                min_point.oval.BorderColor = Color.Red
                active_point = min_point
            End If
        End If


    End Sub
    Private Function gen_angle(e As Point, mp As Point) As Double
        Return (Math.Atan2(e.Y - mp.Y, e.X - mp.X) / Math.PI * 180) + 180
    End Function

    Private Sub Form1_VisibleChanged(sender As Object, e As EventArgs) Handles Me.VisibleChanged

    End Sub
    Public Shared radius As Integer = 25
    Private Shared Function DrawCircle(pt As Point, color As Color) As Microsoft.VisualBasic.PowerPacks.OvalShape
        Dim canvas As New Microsoft.VisualBasic.PowerPacks.ShapeContainer
        Dim oval1 As New Microsoft.VisualBasic.PowerPacks.OvalShape
        canvas.Parent = ActionForm
        oval1.BorderColor = color
        oval1.Parent = canvas
        oval1.Left = pt.X - radius
        oval1.Top = pt.Y - radius
        oval1.BorderWidth = 7
        oval1.Width = 2 * radius
        oval1.Height = 2 * radius
        Return oval1
    End Function

    Private Sub NotifyIcon1_MouseDoubleClick(sender As Object, e As MouseEventArgs) Handles NotifyIcon1.MouseDoubleClick

    End Sub

    Private Sub Quit_Click(sender As Object, e As EventArgs) Handles Quit.Click
        Me.Close()
    End Sub

    Private Sub Reset_Click(sender As Object, e As EventArgs) Handles Reset.Click
        Me.Visible = True
        active = False
        For Each vec In pvec
            vec.oval.Dispose()
        Next
        pvec.Clear()
    End Sub
    Private Sub save_file()
        Dim fs As FileStream = New FileStream(Directory.GetCurrentDirectory() + "speedClick.bin", FileMode.OpenOrCreate)
        Dim bf As New BinaryFormatter()
        bf.Serialize(fs, pvec)
        fs.Close()
    End Sub
    Private Sub load_file()
        Dim fileExists As Boolean
        fileExists = My.Computer.FileSystem.FileExists(Directory.GetCurrentDirectory() + "speedClick.bin")
        If Not fileExists Then
            MsgBox("No load file")
            Return
        End If
        Dim fs As FileStream = New FileStream(Directory.GetCurrentDirectory() + "speedClick.bin", FileMode.Open)
        Dim bf As New BinaryFormatter()
        pvec = CType(bf.Deserialize(fs), Global.System.Collections.Generic.List(Of Global.speedClick.ActionForm.dot))
        fs.Close()
    End Sub

    Private Sub SaveToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SaveToolStripMenuItem.Click
        save_file()
        MsgBox("File Saved!")
    End Sub

    Private Sub LoadToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LoadToolStripMenuItem.Click
        load_file()
        MsgBox("File Loaded")
    End Sub

    Private Sub SettingsToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SettingsToolStripMenuItem.Click
        SettingsForm.Show()
    End Sub

End Class
