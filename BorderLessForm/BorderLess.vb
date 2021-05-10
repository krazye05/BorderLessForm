Imports System.Drawing
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports BorderLessForm.Borderlessclass

Public Class BorderLess
    Inherits Form

    Private TopLeftCornerPanel As New TransparentPanel
    Private TopRightCornerPanel As New TransparentPanel
    Private BottomLeftCornerPanel As New TransparentPanel
    Private BottomRightCornerPanel As New TransparentPanel
    Private TopBorderPanel As New TransparentPanel
    Private BottomBorderPanel As New TransparentPanel
    Private LeftBorderPanel As New TransparentPanel
    Private RightBorderPanel As New TransparentPanel

    Private panels = {TopLeftCornerPanel, TopRightCornerPanel, BottomLeftCornerPanel, BottomRightCornerPanel, TopBorderPanel, LeftBorderPanel, RightBorderPanel, BottomBorderPanel}
    Protected Overrides Sub WndProc(ByRef m As Message)
        If DesignMode Then
            MyBase.WndProc(m)
            Return
        End If
        Select Case m.Msg
            Case CInt(WindowMessages.WM_NCCALCSIZE)
                'Provides New coordinates for the window client area.
                WmNCCalcSize(m)
                Exit Select
            Case CInt(WindowMessages.WM_NCPAINT)
                'Here should all our painting occur, but...
                WmNCPaint(m)
                Exit Select
            Case CInt(WindowMessages.WM_NCACTIVATE)
                '... WM_NCACTIVATE does some painting directly 
                ' without bothering with WM_NCPAINT ...
                WmNCActivate(m)
                Exit Select
            Case CInt(WindowMessages.WM_SETTEXT)
                '... And some painting is required in here as well
                WmSetText(m)
                Exit Select
            Case CInt(WindowMessages.WM_WINDOWPOSCHANGED)
                WmWindowPosChanged(m)
                Exit Select
            Case 174 ' ignore magic message number
                Exit Select
            Case Else
                MyBase.WndProc(m)
                Exit Select
        End Select
    End Sub
    Private ReadOnly Property MinMaxState As FormWindowState
        Get
            Dim s = NativeMethods.GetWindowLong(Handle, NativeConstants.GWL_STYLE)
            Dim max = (s And CInt(WindowStyle.WS_MAXIMIZE)) > 0
            If max Then Return FormWindowState.Maximized
            Dim min = (s And CInt(WindowStyle.WS_MINIMIZE)) > 0
            If min Then Return FormWindowState.Minimized
            Return FormWindowState.Normal
        End Get
    End Property
    Private Sub WmNCCalcSize(ByRef m As Message)
        Dim r = CType(Marshal.PtrToStructure(m.LParam, GetType(RECT)), RECT)
        Dim max = IIf(MinMaxState = FormWindowState.Maximized, True, False)

        If max Then
            Dim x = NativeMethods.GetSystemMetrics(NativeConstants.SM_CXSIZEFRAME)
            Dim y = NativeMethods.GetSystemMetrics(NativeConstants.SM_CYSIZEFRAME)
            Dim p = NativeMethods.GetSystemMetrics(NativeConstants.SM_CXPADDEDBORDER)
            Dim w = x + p
            Dim h = y + p
            r.left += w
            r.top += h
            r.right -= w
            r.bottom -= h
            Dim appBarData = New APPBARDATA()
            appBarData.cbSize = Marshal.SizeOf(GetType(APPBARDATA))
            Dim autohide = (NativeMethods.SHAppBarMessage(NativeConstants.ABM_GETSTATE, appBarData) And NativeConstants.ABS_AUTOHIDE) <> 0
            If autohide Then r.bottom -= 1
            Marshal.StructureToPtr(r, m.LParam, True)
        Else
            updateResizeBar()
        End If

        For Each panel As Panel In panels
            panel.Visible = Not max
        Next

        m.Result = IntPtr.Zero
    End Sub
    Private Sub WmNCPaint(ByRef msg As Message)
        ' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/gdi/pantdraw_8gdw.asp
        ' example in q. 2.9 on http://www.syncfusion.com/FAQ/WindowsForms/FAQ_c41c.aspx#q1026q

        ' The WParam contains handle to clipRegion or 1 if entire window should be repainted
        'PaintNonClientArea(msg.HWnd, (IntPtr)msg.WParam);

        ' we handled everything
        msg.Result = NativeConstants.TRUE
    End Sub
    Private Sub WmNCActivate(ByRef msg As Message)
        ' http://msdn.microsoft.com/library/default.asp?url=/library/en-us/winui/winui/windowsuserinterface/windowing/windows/windowreference/windowmessages/wm_ncactivate.asp

        Dim active As Boolean = IIf(msg.WParam = NativeConstants.TRUE, True, False)

        If MinMaxState = FormWindowState.Minimized Then
            DefWndProc(msg)
        Else
            ' repaint title bar
            'PaintNonClientArea(msg.HWnd, (IntPtr)1);

            ' allow to deactivate window
            msg.Result = NativeConstants.TRUE
        End If
    End Sub
    Private Sub WmSetText(ByRef msg As Message)
        ' allow the system to receive the new window title
        DefWndProc(msg)

        ' repaint title bar
        'PaintNonClientArea(msg.HWnd, (IntPtr)1);
    End Sub
    Private Sub WmWindowPosChanged(ByRef m As Message)
        DefWndProc(m)
        UpdateBounds()
        Dim pos = CType(Marshal.PtrToStructure(m.LParam, GetType(WINDOWPOS)), WINDOWPOS)
        SetWindowRegion(m.HWnd, 0, 0, pos.cx, pos.cy)
        m.Result = NativeConstants.[TRUE]
    End Sub
    Private Sub SetWindowRegion(ByVal hwnd As IntPtr, ByVal left As Integer, ByVal top As Integer, ByVal right As Integer, ByVal bottom As Integer)
        Dim rgn = NativeMethods.CreateRectRgn(0, 0, 0, 0)
        Dim hrg = New HandleRef(Me, rgn)
        Dim r = NativeMethods.GetWindowRgn(hwnd, hrg.Handle)
        Dim box As RECT
        NativeMethods.GetRgnBox(hrg.Handle, box)

        If box.left <> left OrElse box.top <> top OrElse box.right <> right OrElse box.bottom <> bottom Then
            Dim hr = New HandleRef(Me, NativeMethods.CreateRectRgn(left, top, right, bottom))
            NativeMethods.SetWindowRgn(hwnd, hr.Handle, NativeMethods.IsWindowVisible(hwnd))
        End If

        NativeMethods.DeleteObject(rgn)
    End Sub
    Public Function ToggleMaxMin() As FormWindowState
        Return Me.WindowState = IIf(Me.WindowState = FormWindowState.Maximized, FormWindowState.Normal, FormWindowState.Maximized)
    End Function

    Private Shared titleClickTime As Date = Date.MinValue
    Private Shared titleClickPosition As Point = Point.Empty
    Public Shared Sub inititalizeTaskBar(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then
            Dim clickTime = (Date.Now - titleClickTime).TotalMilliseconds
            If clickTime < SystemInformation.DoubleClickTime AndAlso e.Location = titleClickPosition Then
                sender.TopLevelControl.WindowState = IIf(sender.TopLevelControl.WindowState = FormWindowState.Maximized, FormWindowState.Normal, FormWindowState.Maximized)
            Else
                titleClickTime = Date.Now
                titleClickPosition = e.Location

                NativeMethods.ReleaseCapture()
                Dim pt = New POINTS With {
                    .X = CShort(MousePosition.X),
                    .Y = CShort(MousePosition.Y)
                }
                NativeMethods.SendMessage(sender.TopLevelControl.Handle, CInt(WindowMessages.WM_NCLBUTTONDOWN), CInt(HitTestValues.HTCAPTION), pt)
            End If

        End If
    End Sub

    Private Sub DecorationMouseDown(ByVal hit As HitTestValues, ByVal p As Point)
        NativeMethods.ReleaseCapture()
        Dim pt = New POINTS With {
                .X = CShort(p.X),
                .Y = CShort(p.Y)
            }
        NativeMethods.SendMessage(Handle, CInt(WindowMessages.WM_NCLBUTTONDOWN), CInt(hit), pt)
    End Sub

    Private Sub initilizeForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        initializeResizeBar()
    End Sub

    Private Sub resizeFunc(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then
            Select Case sender.name
                Case "TopLeftCornerPanel"
                    DecorationMouseDown(HitTestValues.HTTOPLEFT, MousePosition)
                    Exit Select
                Case "LeftBorderPanel"
                    DecorationMouseDown(HitTestValues.HTLEFT, MousePosition)
                    Exit Select
                Case "BottomLeftCornerPanel"
                    DecorationMouseDown(HitTestValues.HTBOTTOMLEFT, MousePosition)
                    Exit Select
                Case "BottomBorderPanel"
                    DecorationMouseDown(HitTestValues.HTBOTTOM, MousePosition)
                    Exit Select
                Case "BottomRightCornerPanel"
                    DecorationMouseDown(HitTestValues.HTBOTTOMRIGHT, MousePosition)
                    Exit Select
                Case "RightBorderPanel"
                    DecorationMouseDown(HitTestValues.HTRIGHT, MousePosition)
                    Exit Select
                Case "TopRightCornerPanel"
                    DecorationMouseDown(HitTestValues.HTTOPRIGHT, MousePosition)
                    Exit Select
                Case "TopBorderPanel"
                    DecorationMouseDown(HitTestValues.HTTOP, MousePosition)
                    Exit Select
            End Select
        End If
    End Sub
    Private Sub updateResizeBar()
        Dim screenSize = Screen.FromRectangle(Me.Bounds).WorkingArea.Size
        Dim resizeSize As Integer = (screenSize.Height * 0.001) * 5
        TopLeftCornerPanel.Location = New Point(0, 0)
        TopLeftCornerPanel.Size = New Size(resizeSize, resizeSize)

        LeftBorderPanel.Location = New Point(0, resizeSize)
        LeftBorderPanel.Size = New Size(resizeSize, Me.Height - (resizeSize * 2))

        BottomLeftCornerPanel.Location = New Point(0, Me.Height - resizeSize)
        BottomLeftCornerPanel.Size = New Size(resizeSize, resizeSize)

        BottomBorderPanel.Location = New Point(resizeSize, Me.Height - resizeSize)
        BottomBorderPanel.Size = New Size(Me.Width - (resizeSize * 2), resizeSize)

        BottomRightCornerPanel.Location = New Point(Me.Width - resizeSize, Me.Height - resizeSize)
        BottomRightCornerPanel.Size = New Size(resizeSize, resizeSize)

        RightBorderPanel.Location = New Point(Me.Width - resizeSize, resizeSize)
        RightBorderPanel.Size = New Size(resizeSize, Me.Height - (resizeSize * 2))

        TopRightCornerPanel.Location = New Point(Me.Width - resizeSize, 0)
        TopRightCornerPanel.Size = New Size(resizeSize, resizeSize)

        TopBorderPanel.Location = New Point(resizeSize, 0)
        TopBorderPanel.Size = New Size(Me.Width - (resizeSize * 2), resizeSize)
    End Sub
    Private Sub initializeResizeBar()
        Dim screenSize = Screen.FromRectangle(Me.Bounds).WorkingArea.Size
        Dim resizeSize As Integer = (screenSize.Height * 0.001) * 5

        TopLeftCornerPanel.BackColor = Color.Green
        TopLeftCornerPanel.Name = "TopLeftCornerPanel"

        LeftBorderPanel.Cursor = Cursors.SizeWE
        LeftBorderPanel.Name = "LeftBorderPanel"

        BottomLeftCornerPanel.Cursor = Cursors.SizeNESW
        BottomLeftCornerPanel.Name = "BottomLeftCornerPanel"

        BottomBorderPanel.Cursor = Cursors.SizeNS
        BottomBorderPanel.Name = "BottomBorderPanel"

        BottomRightCornerPanel.Cursor = Cursors.SizeNWSE
        BottomRightCornerPanel.Name = "BottomRightCornerPanel"

        RightBorderPanel.Cursor = Cursors.SizeWE
        RightBorderPanel.Name = "RightBorderPanel"

        TopRightCornerPanel.Cursor = Cursors.SizeNESW
        TopRightCornerPanel.Name = "TopRightCornerPanel"

        TopBorderPanel.Cursor = Cursors.SizeNS
        TopBorderPanel.Name = "TopBorderPanel"

        For Each panel As Panel In panels
            Controls.Add(panel)
            panel.BringToFront()
            AddHandler panel.MouseDown, AddressOf resizeFunc
        Next
        updateResizeBar()
    End Sub
End Class