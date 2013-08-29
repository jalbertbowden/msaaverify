' Copyright (c) Microsoft Corporation.  All rights reserved.

Imports Accessibility
Imports System.Runtime.InteropServices

Namespace MsaaVerify

    Public Class MainForm
        Inherits System.Windows.Forms.Form

#Region "Constants"
        'GetAncestor Constants
        Private Const GaRoot As Integer = 2
        ' needed for GDI drawing
        Const SrcCopy As Integer = &HCC0020
        ' needed for GDI drawing
        Const PenWidth As Integer = 4
        ' status bar messages
        Const StatusBarMessageSearching As String = "Searching..."
#End Region

#Region "Private Variables"
        Private m_ops As Verifier ' class for msaa searches and such
        Private m_blnMouseDown As Boolean ' variable for spy tool methods to know whether or not the mouse is down
        Private m_foundMsaaObject As AccessibleObject = Nothing ' the found MSAA object(s) from the spy tool 
        Dim NewLine As String = System.Environment.NewLine() ' private newline

        ' needed for GDI drawing
        Private m_hDCScreen As IntPtr
        Private m_hDCScreenCompatible As IntPtr
        Private m_hDCOld As IntPtr
        Friend WithEvents GroupBox4 As System.Windows.Forms.GroupBox
        Private m_hScreenCompatibleBmp As IntPtr
#End Region

#Region "Windows Form Designer generated code "

        Public Sub New()
            MyBase.New()

            'This call is required by the Windows Form Designer.
            InitializeComponent()
        End Sub

        'Form overrides dispose to clean up the component list.
        Protected Overloads Overrides Sub Dispose(ByVal disposing As Boolean)
            If disposing Then
                If Not (components Is Nothing) Then
                    components.Dispose()
                End If
            End If
            MyBase.Dispose(disposing)
        End Sub

        'Required by the Windows Form Designer
        Private components As System.ComponentModel.IContainer

        'NOTE: The following procedure is required by the Windows Form Designer
        'It can be modified using the Windows Form Designer.  
        'Do not modify it using the code editor.
        Friend WithEvents Label9 As System.Windows.Forms.Label
        Friend WithEvents ChildCountResults As System.Windows.Forms.TextBox
        Friend WithEvents Label2 As System.Windows.Forms.Label
        Friend WithEvents DescriptionResults As System.Windows.Forms.TextBox
        Friend WithEvents GroupBox3 As System.Windows.Forms.GroupBox
        Friend WithEvents DefaultActionResults As System.Windows.Forms.TextBox
        Friend WithEvents Label8 As System.Windows.Forms.Label
        Friend WithEvents KeyboardShortcutResults As System.Windows.Forms.TextBox
        Friend WithEvents Label10 As System.Windows.Forms.Label
        Friend WithEvents NameResults As System.Windows.Forms.TextBox
        Friend WithEvents Label11 As System.Windows.Forms.Label
        Friend WithEvents ParentResults As System.Windows.Forms.TextBox
        Friend WithEvents Label12 As System.Windows.Forms.Label
        Friend WithEvents RoleResults As System.Windows.Forms.TextBox
        Friend WithEvents Label13 As System.Windows.Forms.Label
        Friend WithEvents StateResults As System.Windows.Forms.TextBox
        Friend WithEvents Label14 As System.Windows.Forms.Label
        Friend WithEvents ValueResults As System.Windows.Forms.TextBox
        Friend WithEvents Label15 As System.Windows.Forms.Label
        Friend WithEvents PictureBox1 As System.Windows.Forms.PictureBox
        Friend WithEvents Label4 As System.Windows.Forms.Label
        Friend WithEvents TextBox2 As System.Windows.Forms.TextBox
        Friend WithEvents Label5 As System.Windows.Forms.Label
        Friend WithEvents TextBox3 As System.Windows.Forms.TextBox
        Friend WithEvents Label6 As System.Windows.Forms.Label
        Friend WithEvents Label7 As System.Windows.Forms.Label
        Friend WithEvents TextBoxPassed As System.Windows.Forms.TextBox
        Friend WithEvents TextBoxFailed As System.Windows.Forms.TextBox
        Friend WithEvents StatusBar1 As System.Windows.Forms.StatusBar
        Friend WithEvents PictureBox2 As System.Windows.Forms.PictureBox
        Friend WithEvents VerifyButton As System.Windows.Forms.Button
        Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
        Friend WithEvents Label1 As System.Windows.Forms.Label
        Friend WithEvents CaptureButton As System.Windows.Forms.Button
        Friend WithEvents HwndTextBox As System.Windows.Forms.TextBox
        Friend WithEvents MenuItemVerifyNow As System.Windows.Forms.MenuItem
        Friend WithEvents MainMenu1 As System.Windows.Forms.MainMenu
        Friend WithEvents MenuItemVerify As System.Windows.Forms.MenuItem
        Friend WithEvents MenuItemHelp As System.Windows.Forms.MenuItem
        Friend WithEvents MenuItemAbout As System.Windows.Forms.MenuItem
        Friend WithEvents MenuItemVerifyMsaa As System.Windows.Forms.MenuItem
        Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
        Friend WithEvents Label3 As System.Windows.Forms.Label
        Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
        Friend WithEvents MenuItemHelpVerifications As System.Windows.Forms.MenuItem

        <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
            Me.components = New System.ComponentModel.Container
            Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
            Me.PictureBox1 = New System.Windows.Forms.PictureBox
            Me.Label4 = New System.Windows.Forms.Label
            Me.TextBox2 = New System.Windows.Forms.TextBox
            Me.Label5 = New System.Windows.Forms.Label
            Me.TextBox3 = New System.Windows.Forms.TextBox
            Me.Label6 = New System.Windows.Forms.Label
            Me.Label7 = New System.Windows.Forms.Label
            Me.TextBoxPassed = New System.Windows.Forms.TextBox
            Me.TextBoxFailed = New System.Windows.Forms.TextBox
            Me.StatusBar1 = New System.Windows.Forms.StatusBar
            Me.VerifyButton = New System.Windows.Forms.Button
            Me.PictureBox2 = New System.Windows.Forms.PictureBox
            Me.GroupBox1 = New System.Windows.Forms.GroupBox
            Me.CaptureButton = New System.Windows.Forms.Button
            Me.HwndTextBox = New System.Windows.Forms.TextBox
            Me.Label1 = New System.Windows.Forms.Label
            Me.GroupBox2 = New System.Windows.Forms.GroupBox
            Me.Label3 = New System.Windows.Forms.Label
            Me.TextBox1 = New System.Windows.Forms.TextBox
            Me.MainMenu1 = New System.Windows.Forms.MainMenu(Me.components)
            Me.MenuItemVerify = New System.Windows.Forms.MenuItem
            Me.MenuItemVerifyMsaa = New System.Windows.Forms.MenuItem
            Me.MenuItemHelp = New System.Windows.Forms.MenuItem
            Me.MenuItemHelpVerifications = New System.Windows.Forms.MenuItem
            Me.MenuItemAbout = New System.Windows.Forms.MenuItem
            Me.Label9 = New System.Windows.Forms.Label
            Me.ChildCountResults = New System.Windows.Forms.TextBox
            Me.Label2 = New System.Windows.Forms.Label
            Me.DescriptionResults = New System.Windows.Forms.TextBox
            Me.GroupBox3 = New System.Windows.Forms.GroupBox
            Me.ValueResults = New System.Windows.Forms.TextBox
            Me.Label15 = New System.Windows.Forms.Label
            Me.StateResults = New System.Windows.Forms.TextBox
            Me.Label14 = New System.Windows.Forms.Label
            Me.RoleResults = New System.Windows.Forms.TextBox
            Me.Label13 = New System.Windows.Forms.Label
            Me.ParentResults = New System.Windows.Forms.TextBox
            Me.Label12 = New System.Windows.Forms.Label
            Me.NameResults = New System.Windows.Forms.TextBox
            Me.Label11 = New System.Windows.Forms.Label
            Me.KeyboardShortcutResults = New System.Windows.Forms.TextBox
            Me.Label10 = New System.Windows.Forms.Label
            Me.DefaultActionResults = New System.Windows.Forms.TextBox
            Me.Label8 = New System.Windows.Forms.Label
            Me.GroupBox4 = New System.Windows.Forms.GroupBox
            CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).BeginInit()
            CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).BeginInit()
            Me.GroupBox1.SuspendLayout()
            Me.GroupBox2.SuspendLayout()
            Me.GroupBox3.SuspendLayout()
            Me.GroupBox4.SuspendLayout()
            Me.SuspendLayout()
            '
            'PictureBox1
            '
            Me.PictureBox1.AccessibleDescription = "drag cross hairs over an image to capture its hwnd"
            Me.PictureBox1.AccessibleName = "Crosshair Image"
            Me.PictureBox1.BackColor = System.Drawing.SystemColors.Control
            Me.PictureBox1.Location = New System.Drawing.Point(43, 29)
            Me.PictureBox1.Name = "PictureBox1"
            Me.PictureBox1.Size = New System.Drawing.Size(24, 24)
            Me.PictureBox1.TabIndex = 2
            Me.PictureBox1.TabStop = False
            '
            'Label4
            '
            Me.Label4.AccessibleName = "AA Name:"
            Me.Label4.Location = New System.Drawing.Point(8, 23)
            Me.Label4.Name = "Label4"
            Me.Label4.Size = New System.Drawing.Size(96, 23)
            Me.Label4.TabIndex = 0
            Me.Label4.Text = "Accessible &Name:"
            '
            'TextBox2
            '
            Me.TextBox2.AccessibleName = "AA Name:"
            Me.TextBox2.Location = New System.Drawing.Point(112, 20)
            Me.TextBox2.Name = "TextBox2"
            Me.TextBox2.ReadOnly = True
            Me.TextBox2.Size = New System.Drawing.Size(272, 20)
            Me.TextBox2.TabIndex = 1
            '
            'Label5
            '
            Me.Label5.AccessibleName = "Class Name:"
            Me.Label5.Location = New System.Drawing.Point(8, 72)
            Me.Label5.Name = "Label5"
            Me.Label5.Size = New System.Drawing.Size(72, 23)
            Me.Label5.TabIndex = 4
            Me.Label5.Text = "C&lass Name:"
            '
            'TextBox3
            '
            Me.TextBox3.AccessibleName = "Class Name:"
            Me.TextBox3.Location = New System.Drawing.Point(112, 72)
            Me.TextBox3.Name = "TextBox3"
            Me.TextBox3.ReadOnly = True
            Me.TextBox3.Size = New System.Drawing.Size(272, 20)
            Me.TextBox3.TabIndex = 5
            '
            'Label6
            '
            Me.Label6.AccessibleName = "Passed:"
            Me.Label6.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
            Me.Label6.ForeColor = System.Drawing.Color.Blue
            Me.Label6.Location = New System.Drawing.Point(228, 451)
            Me.Label6.Name = "Label6"
            Me.Label6.Size = New System.Drawing.Size(48, 16)
            Me.Label6.TabIndex = 30
            Me.Label6.Text = "&Passed:"
            '
            'Label7
            '
            Me.Label7.AccessibleName = "Failed:"
            Me.Label7.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
            Me.Label7.ForeColor = System.Drawing.Color.Red
            Me.Label7.Location = New System.Drawing.Point(320, 450)
            Me.Label7.Name = "Label7"
            Me.Label7.Size = New System.Drawing.Size(40, 16)
            Me.Label7.TabIndex = 32
            Me.Label7.Text = "&Failed:"
            '
            'TextBoxPassed
            '
            Me.TextBoxPassed.AccessibleName = "Passed:"
            Me.TextBoxPassed.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
            Me.TextBoxPassed.Location = New System.Drawing.Point(274, 447)
            Me.TextBoxPassed.Name = "TextBoxPassed"
            Me.TextBoxPassed.ReadOnly = True
            Me.TextBoxPassed.Size = New System.Drawing.Size(40, 20)
            Me.TextBoxPassed.TabIndex = 31
            '
            'TextBoxFailed
            '
            Me.TextBoxFailed.AccessibleName = "Failed:"
            Me.TextBoxFailed.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
            Me.TextBoxFailed.Location = New System.Drawing.Point(366, 447)
            Me.TextBoxFailed.Name = "TextBoxFailed"
            Me.TextBoxFailed.ReadOnly = True
            Me.TextBoxFailed.Size = New System.Drawing.Size(40, 20)
            Me.TextBoxFailed.TabIndex = 33
            '
            'StatusBar1
            '
            Me.StatusBar1.AccessibleName = "MsaaVerify Status Bar"
            Me.StatusBar1.Location = New System.Drawing.Point(0, 473)
            Me.StatusBar1.Name = "StatusBar1"
            Me.StatusBar1.Size = New System.Drawing.Size(419, 22)
            Me.StatusBar1.TabIndex = 35
            Me.StatusBar1.Text = "Ready"
            '
            'VerifyButton
            '
            Me.VerifyButton.AccessibleName = "Verify MSAA"
            Me.VerifyButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
            Me.VerifyButton.Enabled = False
            Me.VerifyButton.Location = New System.Drawing.Point(7, 445)
            Me.VerifyButton.Name = "VerifyButton"
            Me.VerifyButton.Size = New System.Drawing.Size(76, 23)
            Me.VerifyButton.TabIndex = 2
            Me.VerifyButton.Text = "&Verify"
            '
            'PictureBox2
            '
            Me.PictureBox2.AccessibleDescription = "Frame for holding the image"
            Me.PictureBox2.AccessibleName = "Crosshair Image Frame"
            Me.PictureBox2.AccessibleRole = System.Windows.Forms.AccessibleRole.Graphic
            Me.PictureBox2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
            Me.PictureBox2.Location = New System.Drawing.Point(35, 21)
            Me.PictureBox2.Name = "PictureBox2"
            Me.PictureBox2.Size = New System.Drawing.Size(40, 40)
            Me.PictureBox2.TabIndex = 26
            Me.PictureBox2.TabStop = False
            '
            'GroupBox1
            '
            Me.GroupBox1.AccessibleName = "Capture by Point"
            Me.GroupBox1.Controls.Add(Me.PictureBox1)
            Me.GroupBox1.Controls.Add(Me.PictureBox2)
            Me.GroupBox1.Location = New System.Drawing.Point(6, 1)
            Me.GroupBox1.Name = "GroupBox1"
            Me.GroupBox1.Size = New System.Drawing.Size(110, 71)
            Me.GroupBox1.TabIndex = 0
            Me.GroupBox1.TabStop = False
            Me.GroupBox1.Text = "Capture by Point"
            '
            'CaptureButton
            '
            Me.CaptureButton.AccessibleName = "Capture"
            Me.CaptureButton.Enabled = False
            Me.CaptureButton.Location = New System.Drawing.Point(204, 38)
            Me.CaptureButton.Name = "CaptureButton"
            Me.CaptureButton.Size = New System.Drawing.Size(56, 23)
            Me.CaptureButton.TabIndex = 2
            Me.CaptureButton.Text = "C&apture"
            '
            'HwndTextBox
            '
            Me.HwndTextBox.AccessibleName = "Enter the window handle:"
            Me.HwndTextBox.Location = New System.Drawing.Point(9, 41)
            Me.HwndTextBox.Name = "HwndTextBox"
            Me.HwndTextBox.Size = New System.Drawing.Size(189, 20)
            Me.HwndTextBox.TabIndex = 1
            '
            'Label1
            '
            Me.Label1.AccessibleName = "Enter the window handle:"
            Me.Label1.Location = New System.Drawing.Point(6, 19)
            Me.Label1.Name = "Label1"
            Me.Label1.Size = New System.Drawing.Size(192, 16)
            Me.Label1.TabIndex = 0
            Me.Label1.Text = "Enter the window &handle:"
            '
            'GroupBox2
            '
            Me.GroupBox2.Controls.Add(Me.Label3)
            Me.GroupBox2.Controls.Add(Me.TextBox1)
            Me.GroupBox2.Controls.Add(Me.Label4)
            Me.GroupBox2.Controls.Add(Me.Label5)
            Me.GroupBox2.Controls.Add(Me.TextBox2)
            Me.GroupBox2.Controls.Add(Me.TextBox3)
            Me.GroupBox2.Location = New System.Drawing.Point(6, 79)
            Me.GroupBox2.Name = "GroupBox2"
            Me.GroupBox2.Size = New System.Drawing.Size(400, 104)
            Me.GroupBox2.TabIndex = 2
            Me.GroupBox2.TabStop = False
            Me.GroupBox2.Text = "Captured Control"
            '
            'Label3
            '
            Me.Label3.Location = New System.Drawing.Point(8, 48)
            Me.Label3.Name = "Label3"
            Me.Label3.Size = New System.Drawing.Size(88, 23)
            Me.Label3.TabIndex = 2
            Me.Label3.Text = "Accessible &Role:"
            '
            'TextBox1
            '
            Me.TextBox1.Location = New System.Drawing.Point(112, 48)
            Me.TextBox1.Name = "TextBox1"
            Me.TextBox1.ReadOnly = True
            Me.TextBox1.Size = New System.Drawing.Size(272, 20)
            Me.TextBox1.TabIndex = 3
            '
            'MainMenu1
            '
            Me.MainMenu1.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItemVerify, Me.MenuItemHelp})
            '
            'MenuItemVerify
            '
            Me.MenuItemVerify.Index = 0
            Me.MenuItemVerify.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItemVerifyMsaa})
            Me.MenuItemVerify.Text = "Verify"
            '
            'MenuItemVerifyMsaa
            '
            Me.MenuItemVerifyMsaa.Index = 0
            Me.MenuItemVerifyMsaa.Text = "Verify MSAA"
            '
            'MenuItemHelp
            '
            Me.MenuItemHelp.Index = 1
            Me.MenuItemHelp.MenuItems.AddRange(New System.Windows.Forms.MenuItem() {Me.MenuItemHelpVerifications, Me.MenuItemAbout})
            Me.MenuItemHelp.Text = "Help"
            '
            'MenuItemHelpVerifications
            '
            Me.MenuItemHelpVerifications.Index = 0
            Me.MenuItemHelpVerifications.Text = "MSAA Verifications"
            '
            'MenuItemAbout
            '
            Me.MenuItemAbout.Index = 1
            Me.MenuItemAbout.Text = "About"
            '
            'Label9
            '
            Me.Label9.AutoSize = True
            Me.Label9.Location = New System.Drawing.Point(6, 21)
            Me.Label9.Name = "Label9"
            Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
            Me.Label9.Size = New System.Drawing.Size(61, 13)
            Me.Label9.TabIndex = 12
            Me.Label9.Text = "&ChildCount:"
            '
            'ChildCountResults
            '
            Me.ChildCountResults.Location = New System.Drawing.Point(112, 18)
            Me.ChildCountResults.Name = "ChildCountResults"
            Me.ChildCountResults.ReadOnly = True
            Me.ChildCountResults.Size = New System.Drawing.Size(269, 20)
            Me.ChildCountResults.TabIndex = 13
            '
            'Label2
            '
            Me.Label2.AutoSize = True
            Me.Label2.Location = New System.Drawing.Point(6, 45)
            Me.Label2.Name = "Label2"
            Me.Label2.RightToLeft = System.Windows.Forms.RightToLeft.No
            Me.Label2.Size = New System.Drawing.Size(63, 13)
            Me.Label2.TabIndex = 14
            Me.Label2.Text = "&Description:"
            '
            'DescriptionResults
            '
            Me.DescriptionResults.Location = New System.Drawing.Point(112, 42)
            Me.DescriptionResults.Name = "DescriptionResults"
            Me.DescriptionResults.ReadOnly = True
            Me.DescriptionResults.Size = New System.Drawing.Size(269, 20)
            Me.DescriptionResults.TabIndex = 15
            '
            'GroupBox3
            '
            Me.GroupBox3.Controls.Add(Me.ValueResults)
            Me.GroupBox3.Controls.Add(Me.Label15)
            Me.GroupBox3.Controls.Add(Me.StateResults)
            Me.GroupBox3.Controls.Add(Me.Label14)
            Me.GroupBox3.Controls.Add(Me.RoleResults)
            Me.GroupBox3.Controls.Add(Me.Label13)
            Me.GroupBox3.Controls.Add(Me.ParentResults)
            Me.GroupBox3.Controls.Add(Me.Label12)
            Me.GroupBox3.Controls.Add(Me.NameResults)
            Me.GroupBox3.Controls.Add(Me.Label11)
            Me.GroupBox3.Controls.Add(Me.KeyboardShortcutResults)
            Me.GroupBox3.Controls.Add(Me.Label10)
            Me.GroupBox3.Controls.Add(Me.DefaultActionResults)
            Me.GroupBox3.Controls.Add(Me.Label8)
            Me.GroupBox3.Controls.Add(Me.DescriptionResults)
            Me.GroupBox3.Controls.Add(Me.Label2)
            Me.GroupBox3.Controls.Add(Me.ChildCountResults)
            Me.GroupBox3.Controls.Add(Me.Label9)
            Me.GroupBox3.Location = New System.Drawing.Point(6, 189)
            Me.GroupBox3.Name = "GroupBox3"
            Me.GroupBox3.Size = New System.Drawing.Size(400, 252)
            Me.GroupBox3.TabIndex = 3
            Me.GroupBox3.TabStop = False
            Me.GroupBox3.Text = "Test Results"
            '
            'ValueResults
            '
            Me.ValueResults.Location = New System.Drawing.Point(112, 223)
            Me.ValueResults.Name = "ValueResults"
            Me.ValueResults.ReadOnly = True
            Me.ValueResults.Size = New System.Drawing.Size(269, 20)
            Me.ValueResults.TabIndex = 29
            '
            'Label15
            '
            Me.Label15.AutoSize = True
            Me.Label15.Location = New System.Drawing.Point(6, 227)
            Me.Label15.Name = "Label15"
            Me.Label15.RightToLeft = System.Windows.Forms.RightToLeft.No
            Me.Label15.Size = New System.Drawing.Size(37, 13)
            Me.Label15.TabIndex = 28
            Me.Label15.Text = "V&alue:"
            '
            'StateResults
            '
            Me.StateResults.Location = New System.Drawing.Point(112, 197)
            Me.StateResults.Name = "StateResults"
            Me.StateResults.ReadOnly = True
            Me.StateResults.Size = New System.Drawing.Size(269, 20)
            Me.StateResults.TabIndex = 27
            '
            'Label14
            '
            Me.Label14.AutoSize = True
            Me.Label14.Location = New System.Drawing.Point(6, 201)
            Me.Label14.Name = "Label14"
            Me.Label14.RightToLeft = System.Windows.Forms.RightToLeft.No
            Me.Label14.Size = New System.Drawing.Size(35, 13)
            Me.Label14.TabIndex = 26
            Me.Label14.Text = "&State:"
            '
            'RoleResults
            '
            Me.RoleResults.Location = New System.Drawing.Point(112, 171)
            Me.RoleResults.Name = "RoleResults"
            Me.RoleResults.ReadOnly = True
            Me.RoleResults.Size = New System.Drawing.Size(269, 20)
            Me.RoleResults.TabIndex = 25
            '
            'Label13
            '
            Me.Label13.AutoSize = True
            Me.Label13.Location = New System.Drawing.Point(6, 175)
            Me.Label13.Name = "Label13"
            Me.Label13.RightToLeft = System.Windows.Forms.RightToLeft.No
            Me.Label13.Size = New System.Drawing.Size(32, 13)
            Me.Label13.TabIndex = 24
            Me.Label13.Text = "&Role:"
            '
            'ParentResults
            '
            Me.ParentResults.Location = New System.Drawing.Point(112, 145)
            Me.ParentResults.Name = "ParentResults"
            Me.ParentResults.ReadOnly = True
            Me.ParentResults.Size = New System.Drawing.Size(269, 20)
            Me.ParentResults.TabIndex = 23
            '
            'Label12
            '
            Me.Label12.AutoSize = True
            Me.Label12.Location = New System.Drawing.Point(6, 149)
            Me.Label12.Name = "Label12"
            Me.Label12.RightToLeft = System.Windows.Forms.RightToLeft.No
            Me.Label12.Size = New System.Drawing.Size(41, 13)
            Me.Label12.TabIndex = 22
            Me.Label12.Text = "&Parent:"
            '
            'NameResults
            '
            Me.NameResults.Location = New System.Drawing.Point(112, 120)
            Me.NameResults.Name = "NameResults"
            Me.NameResults.ReadOnly = True
            Me.NameResults.Size = New System.Drawing.Size(269, 20)
            Me.NameResults.TabIndex = 21
            '
            'Label11
            '
            Me.Label11.AutoSize = True
            Me.Label11.Location = New System.Drawing.Point(6, 123)
            Me.Label11.Name = "Label11"
            Me.Label11.RightToLeft = System.Windows.Forms.RightToLeft.No
            Me.Label11.Size = New System.Drawing.Size(38, 13)
            Me.Label11.TabIndex = 20
            Me.Label11.Text = "&Name:"
            '
            'KeyboardShortcutResults
            '
            Me.KeyboardShortcutResults.Location = New System.Drawing.Point(112, 94)
            Me.KeyboardShortcutResults.Name = "KeyboardShortcutResults"
            Me.KeyboardShortcutResults.ReadOnly = True
            Me.KeyboardShortcutResults.Size = New System.Drawing.Size(269, 20)
            Me.KeyboardShortcutResults.TabIndex = 19
            '
            'Label10
            '
            Me.Label10.AutoSize = True
            Me.Label10.Location = New System.Drawing.Point(6, 97)
            Me.Label10.Name = "Label10"
            Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
            Me.Label10.Size = New System.Drawing.Size(95, 13)
            Me.Label10.TabIndex = 18
            Me.Label10.Text = "&KeyboardShortcut:"
            '
            'DefaultActionResults
            '
            Me.DefaultActionResults.Location = New System.Drawing.Point(112, 68)
            Me.DefaultActionResults.Name = "DefaultActionResults"
            Me.DefaultActionResults.ReadOnly = True
            Me.DefaultActionResults.Size = New System.Drawing.Size(269, 20)
            Me.DefaultActionResults.TabIndex = 17
            '
            'Label8
            '
            Me.Label8.AutoSize = True
            Me.Label8.Location = New System.Drawing.Point(6, 71)
            Me.Label8.Name = "Label8"
            Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
            Me.Label8.Size = New System.Drawing.Size(74, 13)
            Me.Label8.TabIndex = 16
            Me.Label8.Text = "D&efaultAction:"
            '
            'GroupBox4
            '
            Me.GroupBox4.Controls.Add(Me.HwndTextBox)
            Me.GroupBox4.Controls.Add(Me.CaptureButton)
            Me.GroupBox4.Controls.Add(Me.Label1)
            Me.GroupBox4.Location = New System.Drawing.Point(140, 1)
            Me.GroupBox4.Name = "GroupBox4"
            Me.GroupBox4.Size = New System.Drawing.Size(266, 71)
            Me.GroupBox4.TabIndex = 1
            Me.GroupBox4.TabStop = False
            Me.GroupBox4.Text = "Capture by Window Handle"
            '
            'MainForm
            '
            Me.AccessibleName = "MsaaVerifyForm"
            Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
            Me.AutoScroll = True
            Me.ClientSize = New System.Drawing.Size(419, 495)
            Me.Controls.Add(Me.GroupBox4)
            Me.Controls.Add(Me.StatusBar1)
            Me.Controls.Add(Me.VerifyButton)
            Me.Controls.Add(Me.TextBoxFailed)
            Me.Controls.Add(Me.TextBoxPassed)
            Me.Controls.Add(Me.Label7)
            Me.Controls.Add(Me.Label6)
            Me.Controls.Add(Me.GroupBox1)
            Me.Controls.Add(Me.GroupBox2)
            Me.Controls.Add(Me.GroupBox3)
            Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
            Me.Menu = Me.MainMenu1
            Me.Name = "MainForm"
            Me.RightToLeftLayout = True
            Me.Text = "MsaaVerify"
            CType(Me.PictureBox1, System.ComponentModel.ISupportInitialize).EndInit()
            CType(Me.PictureBox2, System.ComponentModel.ISupportInitialize).EndInit()
            Me.GroupBox1.ResumeLayout(False)
            Me.GroupBox2.ResumeLayout(False)
            Me.GroupBox2.PerformLayout()
            Me.GroupBox3.ResumeLayout(False)
            Me.GroupBox3.PerformLayout()
            Me.GroupBox4.ResumeLayout(False)
            Me.GroupBox4.PerformLayout()
            Me.ResumeLayout(False)
            Me.PerformLayout()

        End Sub

#End Region

#Region "Spy Tool Methods"

        Private Sub PictureBox1_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseDown
            PictureBox1.Invalidate()
            ResetForm()
            m_blnMouseDown = True
            Me.Cursor = New Cursor(New System.IO.MemoryStream(My.Resources.CrosshairsCursor))
            Me.StatusBar1.Text = StatusBarMessageSearching
        End Sub

        Private Sub PictureBox1_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseMove
            If m_blnMouseDown Then
                PictureBox1.Invalidate()
            End If

            ClearRect()

            HandleMouse()
        End Sub

        Private Sub HandleMouse()
            If m_blnMouseDown Then
                ' we are searching like accExplorer
                Dim x As Integer = Windows.Forms.Cursor.Position.X
                Dim y As Integer = Windows.Forms.Cursor.Position.Y

                ' create the accessible object by point
                Try
                    m_foundMsaaObject = New AccessibleObject(x, y)
                Catch ex As FailedToCreateAccessibleObject
                    Me.StatusBar1.Text = ex.Message
                End Try

                ' update UI
                Me.UpdateCapturedControlDisplay()

                ' draw the rectangle around the object
                DrawRect(m_foundMsaaObject.X, m_foundMsaaObject.Y, m_foundMsaaObject.Width, m_foundMsaaObject.Height)
            End If
        End Sub

        Private Sub PictureBox1_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox1.MouseUp
            ' change the cursor to a wait cursor
            Me.Cursor = Cursors.WaitCursor
            m_blnMouseDown = False

            PictureBox1.Refresh()
            ClearRect() ' no need to highlight the captured object anymore

            ' update the UI with the information from the captured accessible object
            CaptureAccessibleObject()

            ' change cursor back to default
            Me.Cursor = Cursors.Default
            Me.StatusBar1.Text = "Ready"
            Me.VerifyButton.Enabled = True
        End Sub

        Private Sub CaptureAccessibleObject()

            ' check for valid aa object.  BUG: Clicking on the crosshairs throws exception
            If Me.m_foundMsaaObject Is Nothing Then
                Me.StatusBar1.Text = "Failed to find valid accessible object"
            Else

                Try
                    ' filter out the special cases, where the captured accessible object isn't necessarily
                    ' the one that we want to do verifications upon
                    Select Case m_foundMsaaObject.Role

                        Case MsaaRoles.StatusBar, MsaaRoles.ListItem
                            ' for these roles, it is the parent we're interested in
                            m_foundMsaaObject = m_foundMsaaObject.Parent

                        Case MsaaRoles.Text, MsaaRoles.PushButton
                            ' check whether we captured a combo box, since both text and PushButtons are children of combo boxes
                            If m_foundMsaaObject.Parent.Role = MsaaRoles.ComboBox Then
                                ' yep, we've captured a PushButton that belongs to a combo box, so add the combo box
                                m_foundMsaaObject = m_foundMsaaObject.Parent
                            ElseIf m_foundMsaaObject.Parent.Parent.Role = MsaaRoles.ComboBox Then
                                ' yep, we've captured an edit box that belongs to a combo box
                                ' add the grandparent, because there's a Window wrapper between the combo box and the edit box
                                m_foundMsaaObject = m_foundMsaaObject.Parent.Parent
                            End If
                        Case MsaaRoles.Window
                            ' in hopes that we can identify the child object in the verifications...
                            m_foundMsaaObject = m_foundMsaaObject.Child(3)
                    End Select

                Catch e As FailedToCreateAccessibleObject
                Catch e As ObjectIsElementException
                Catch e As ChildNotFoundException
                    ' while trying to identify the AA object, one of its methods threw an exception
                    ' since we can't further identify it, we'll try our best to verify it
                End Try

                ' update UI
                UpdateCapturedControlDisplay()

            End If
        End Sub

        Private Sub UpdateCapturedControlDisplay()
            ' show the accessible name on the form
            Me.TextBox2.Text = m_foundMsaaObject.Name

            ' show the class name
            Me.TextBox3.Text = m_foundMsaaObject.ClassName

            ' show the role type 
            ' the role text for static text is "text", which I confuse every time with a textbox that has a role text of "editable text"
            If m_foundMsaaObject.Role = MsaaRoles.StaticText Then
                Me.TextBox1.Text = "[static text]"
            Else
                Me.TextBox1.Text = m_foundMsaaObject.RoleText
            End If

            ' show the hwnd
            Me.HwndTextBox.Text = m_foundMsaaObject.Hwnd.ToString
        End Sub

        Private Sub PictureBox1_Paint(ByVal sender As Object, ByVal e As System.Windows.Forms.PaintEventArgs) Handles PictureBox1.Paint
            If m_blnMouseDown = False Then
                ' Create image.
                Dim imageFile As Image = My.Resources.Crosshairs
                ' Draw image to screen.
                e.Graphics.DrawImage(imageFile, New PointF(0.0F, 0.0F))
            End If
        End Sub

        Private Sub ResetForm()
            ClearRect()
            ClearVerificationResults()
            Me.StatusBar1.Text = "Ready"
            Me.m_foundMsaaObject = Nothing
        End Sub

        Private Sub ClearRect()
            If Not m_foundMsaaObject Is Nothing Then
                NativeMethods.BitBlt(m_hDCScreen, m_foundMsaaObject.X - PenWidth, m_foundMsaaObject.Y - PenWidth, m_foundMsaaObject.Width + PenWidth * 2, m_foundMsaaObject.Height + PenWidth * 2, m_hDCScreenCompatible, 0, 0, SrcCopy)
                NativeMethods.SelectObject(m_hDCScreenCompatible, m_hDCOld)
                If Not m_hDCScreen.Equals(IntPtr.Zero) Then
                    NativeMethods.DeleteDC(m_hDCScreen)
                End If

                If Not m_hDCScreenCompatible.Equals(IntPtr.Zero) Then
                    NativeMethods.DeleteDC(m_hDCScreenCompatible)
                End If

                If Not m_hScreenCompatibleBmp.Equals(IntPtr.Zero) Then
                    NativeMethods.DeleteObject(m_hScreenCompatibleBmp)
                End If
            End If
        End Sub

        Private Sub DrawRect(ByVal Left As Integer, ByVal Top As Integer, ByVal Width As Integer, ByVal Height As Integer)
            Left = Left - PenWidth
            Top = Top - PenWidth
            Width = Width + PenWidth * 2
            Height = Height + PenWidth * 2

            m_hDCScreen = NativeMethods.GetDC(Nothing)
            Dim formGraphics As Graphics = Graphics.FromHdc(m_hDCScreen)

            m_hDCScreenCompatible = NativeMethods.CreateCompatibleDC(m_hDCScreen)
            m_hScreenCompatibleBmp = NativeMethods.CreateCompatibleBitmap(m_hDCScreen, Width, Height)
            m_hDCOld = NativeMethods.SelectObject(m_hDCScreenCompatible, m_hScreenCompatibleBmp)

            NativeMethods.BitBlt(m_hDCScreenCompatible, 0, 0, Width, Height, m_hDCScreen, Left, Top, SrcCopy)

            ' Create Pen and Rectangle 
            Dim blackPen As New Pen(Color.FromArgb(128, 0, 0, 0), PenWidth)
            Dim rect As New Rectangle(Left + PenWidth, Top + PenWidth, Width - PenWidth * 2, Height - PenWidth * 2)
            formGraphics.CompositingMode = Drawing2D.CompositingMode.SourceOver

            formGraphics.DrawRectangle(blackPen, rect)
        End Sub

#End Region

#Region "Private Methods"

        Private Sub VerifyNow()
            Me.Cursor = Cursors.WaitCursor
            Me.StatusBar1.Text = "Verifying... Please wait."
            Verifier.VerifyMsaaObjects(Me.m_foundMsaaObject)
            Me.StatusBar1.Text = "Ready"
            Me.Cursor = Cursors.Default
            Me.VerifyButton.Enabled = False
        End Sub

        Private Sub ClearVerificationResults()
            Me.ChildCountResults.Clear()
            Me.DescriptionResults.Clear()
            Me.DefaultActionResults.Clear()
            Me.KeyboardShortcutResults.Clear()
            Me.NameResults.Clear()
            Me.ParentResults.Clear()
            Me.RoleResults.Clear()
            Me.StateResults.Clear()
            Me.ValueResults.Clear()
            Me.TextBoxPassed.Text = ""
            Me.TextBoxFailed.Text = ""
        End Sub
#End Region

#Region "Event Handlers"
        Private Sub MenuItemAbout_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemAbout.Click
            MessageBox.Show("MsaaVerify 1.1", "About", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1)
        End Sub

        Private Sub MenuItemHelpVerifications_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemHelpVerifications.Click
            Dim helpVerificationsDlg As New HelpVerifications
            helpVerificationsDlg.ShowDialog()
        End Sub

        Private Sub MenuItemVerifyMsaa_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MenuItemVerifyMsaa.Click
            VerifyNow()
        End Sub

        Private Sub VerifyNow_OnClick(ByVal sender As System.Object, ByVal e As System.EventArgs)
            VerifyNow()
        End Sub

        Private Sub VerifyButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles VerifyButton.Click
            VerifyNow()
        End Sub

        Private Sub CrosshairsRadioButton_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Me.HwndTextBox.ReadOnly = True
            Me.CaptureButton.Enabled = False
        End Sub

        Private Sub HwndRadioButton_CheckedChanged(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Me.HwndTextBox.ReadOnly = False
            Me.CaptureButton.Enabled = True
        End Sub

        Private Sub CaptureButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles CaptureButton.Click
            Dim hwnd As IntPtr
            Try
                hwnd = New IntPtr(CType(Me.HwndTextBox.Text, Integer))
                If IsWindowValid(hwnd) Then
                    ' clear the UI
                    ResetForm()
                    ' create the accessible object by window
                    m_foundMsaaObject = New AccessibleObject(hwnd)
                    ' update the UI
                    CaptureAccessibleObject()
                Else
                    MsgBox("Could not find a valid window with HWND: " & Me.HwndTextBox.Text)
                End If
            Catch ex As System.InvalidCastException
                MsgBox("Hwnd must contain only integer values.")
            Catch exc As COMException
                MsgBox("Could not find a valid window with HWND: " & Me.HwndTextBox.ToString)
            End Try
        End Sub

        Private Sub ResetFormButton_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
            Me.ResetForm()
        End Sub

        ' determine whether a window is valid
        Public Function IsWindowValid(ByVal hwnd As IntPtr) As Boolean
            Return NativeMethods.IsWindow(hwnd) <> 0
        End Function
#End Region

        Private Sub HwndTextBox_TextChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles HwndTextBox.TextChanged
            If Me.HwndTextBox.Text = "" Then
                Me.CaptureButton.Enabled = False
            Else
                Me.CaptureButton.Enabled = True
            End If
        End Sub
    End Class
End Namespace
