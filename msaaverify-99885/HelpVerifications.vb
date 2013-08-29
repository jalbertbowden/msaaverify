' Copyright (c) Microsoft Corporation.  All rights reserved.

Public Class HelpVerifications
    Inherits System.Windows.Forms.Form

#Region " Windows Form Designer generated code "

    Public Sub New()
        MyBase.New()

        'This call is required by the Windows Form Designer.
        InitializeComponent()

        'Add any initialization after the InitializeComponent() call
        Me.ListBox1.SetSelected(0, True)
        Me.ListBox2.SetSelected(0, True)
        Me.LinkLabel1.Text = Me.GetSelectedPropertyURL(MsaaProperties.ChildCount)
        Me.LinkLabel2.Text = Me.GetSelectedRoleURL(MsaaRoles.Button)
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
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ListBox2 As System.Windows.Forms.ListBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents LinkLabel1 As System.Windows.Forms.LinkLabel
    Friend WithEvents LinkLabel2 As System.Windows.Forms.LinkLabel
    <System.Diagnostics.DebuggerStepThrough()> Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.Label2 = New System.Windows.Forms.Label
        Me.ListBox2 = New System.Windows.Forms.ListBox
        Me.Label3 = New System.Windows.Forms.Label
        Me.Label4 = New System.Windows.Forms.Label
        Me.LinkLabel1 = New System.Windows.Forms.LinkLabel
        Me.LinkLabel2 = New System.Windows.Forms.LinkLabel
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.Location = New System.Drawing.Point(8, 8)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(136, 24)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Verified properties:"
        '
        'ListBox1
        '
        Me.ListBox1.Items.AddRange(New Object() {"ChildCount", "DefaultAction", "Description", "KeyboardShortcut", "Name", "Parent", "Role", "State", "Value"})
        Me.ListBox1.Location = New System.Drawing.Point(8, 32)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(128, 134)
        Me.ListBox1.TabIndex = 1
        '
        'Label2
        '
        Me.Label2.Location = New System.Drawing.Point(8, 184)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(192, 23)
        Me.Label2.TabIndex = 4
        Me.Label2.Text = "Additional info on selected property:"
        '
        'ListBox2
        '
        Me.ListBox2.Items.AddRange(New Object() {"Button", "Check box", "Combo box", "Edit box", "List box", "List view", "Progress bar", "Radio button", "Static text", "Status bar"})
        Me.ListBox2.Location = New System.Drawing.Point(208, 32)
        Me.ListBox2.Name = "ListBox2"
        Me.ListBox2.Size = New System.Drawing.Size(136, 134)
        Me.ListBox2.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.Location = New System.Drawing.Point(208, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(88, 23)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Verified Roles:"
        '
        'Label4
        '
        Me.Label4.Location = New System.Drawing.Point(8, 240)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(168, 23)
        Me.Label4.TabIndex = 6
        Me.Label4.Text = "Additional info on selected role:"
        '
        'LinkLabel1
        '
        Me.LinkLabel1.Location = New System.Drawing.Point(8, 208)
        Me.LinkLabel1.Name = "LinkLabel1"
        Me.LinkLabel1.Size = New System.Drawing.Size(352, 23)
        Me.LinkLabel1.TabIndex = 5
        '
        'LinkLabel2
        '
        Me.LinkLabel2.Location = New System.Drawing.Point(8, 264)
        Me.LinkLabel2.Name = "LinkLabel2"
        Me.LinkLabel2.Size = New System.Drawing.Size(352, 23)
        Me.LinkLabel2.TabIndex = 7
        '
        'HelpVerifications
        '
        Me.AutoScaleBaseSize = New System.Drawing.Size(5, 13)
        Me.ClientSize = New System.Drawing.Size(360, 298)
        Me.Controls.Add(Me.LinkLabel2)
        Me.Controls.Add(Me.LinkLabel1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.ListBox2)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.ListBox1)
        Me.Controls.Add(Me.Label1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.Name = "HelpVerifications"
        Me.Text = "Help on MSAA Verifications"
        Me.ResumeLayout(False)

    End Sub

#End Region

    Private Sub ListBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged
        Dim selectedProperty As String = Me.ListBox1.SelectedItem
        Me.LinkLabel1.Text = GetSelectedPropertyURL(selectedProperty)
    End Sub

    Private Sub ListBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ListBox2.SelectedIndexChanged
        Dim selectedRole As String = Me.ListBox2.SelectedItem
        Me.LinkLabel2.Text = GetSelectedRoleURL(selectedRole)
    End Sub

    Private Function GetSelectedPropertyURL(ByVal selectedProperty As String) As String
        Select Case selectedProperty
            Case MsaaProperties.ChildCount
                Return "http://msdn.microsoft.com/library/en-us/msaa/msaaccrf_1w1g.asp"
            Case MsaaProperties.DefaultAction
                Return "http://msdn.microsoft.com/library/en-us/msaa/msaaccrf_3mb2.asp"
            Case MsaaProperties.Description
                Return "http://msdn.microsoft.com/library/en-us/msaa/msaaccrf_4y7i.asp"
            Case MsaaProperties.KeyboardShortcut
                Return "http://msdn.microsoft.com/library/en-us/msaa/msaaccrf_91o4.asp"
            Case MsaaProperties.Name
                Return "http://msdn.microsoft.com/library/en-us/msaa/msaaccrf_6asl.asp"
            Case MsaaProperties.Parent
                Return "http://msdn.microsoft.com/library/en-us/msaa/msaaccrf_0e0k.asp"
            Case MsaaProperties.Role
                Return "http://msdn.microsoft.com/library/en-us/msaa/msaaccrf_002t.asp"
            Case MsaaProperties.State
                Return "http://msdn.microsoft.com/library/en-us/msaa/msaaccrf_6vs5.asp"
            Case MsaaProperties.Value
                Return "http://msdn.microsoft.com/library/en-us/msaa/msaaccrf_9vl1.asp"
            Case Else
                Return "Unable to find URL for selected property"
        End Select

    End Function

    Private Function GetSelectedRoleURL(ByVal selectedRole As String) As String
        Select Case selectedRole
            Case MsaaRoles.Button
                Return "http://msdn.microsoft.com/library/en-us/msaa/msaapndx_6ypa.asp"
            Case MsaaRoles.CheckBox
                Return "http://msdn.microsoft.com/library/en-us/msaa/msaapndx_4erc.asp"
            Case MsaaRoles.ComboBox
                Return "http://msdn.microsoft.com/library/en-us/msaa/msaapndx_53nc.asp"
            Case MsaaRoles.EditBox
                Return "http://msdn.microsoft.com/library/en-us/msaa/msaapndx_7bu4.asp"
            Case MsaaRoles.ListBox
                Return "http://msdn.microsoft.com/library/en-us/msaa/msaapndx_1jy0.asp"
            Case MsaaRoles.ListView
                Return "http://msdn.microsoft.com/library/en-us/msaa/msaapndx_3y5o.asp"
            Case MsaaRoles.ProgressBar
                Return "http://msdn.microsoft.com/library/en-us/msaa/msaapndx_5u7g.asp"
            Case MsaaRoles.RadioButton
                Return "http://msdn.microsoft.com/library/en-us/msaa/msaapndx_7ulq.asp"
            Case MsaaRoles.StaticText
                Return "http://msdn.microsoft.com/library/en-us/msaa/msaapndx_4ws4.asp"
            Case MsaaRoles.StatusBar
                Return "http://msdn.microsoft.com/library/en-us/msaa/msaapndx_823w.asp"
            Case Else
                Return "Unable to find URL for selected property"
        End Select

    End Function

    Private Class MsaaProperties
        Public Const ChildCount As String = "ChildCount"
        Public Const DefaultAction As String = "DefaultAction"
        Public Const Description As String = "Description"
        Public Const KeyboardShortcut As String = "KeyboardShortcut"
        Public Const Name As String = "Name"
        Public Const Parent As String = "Parent"
        Public Const Role As String = "Role"
        Public Const State As String = "State"
        Public Const Value As String = "Value"
    End Class

    Private Class MsaaRoles
        Public Const Button As String = "Button"
        Public Const CheckBox As String = "Check box"
        Public Const ComboBox As String = "Combo box"
        Public Const EditBox As String = "Edit box"
        Public Const ListBox As String = "List box"
        Public Const ListView As String = "List view"
        Public Const ProgressBar As String = "Progress bar"
        Public Const RadioButton As String = "Radio button"
        Public Const StaticText As String = "Static text"
        Public Const StatusBar As String = "Status bar"
    End Class

    Private Sub LinkLabel1_LinkClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Process.Start("IExplore.exe")
    End Sub
End Class

