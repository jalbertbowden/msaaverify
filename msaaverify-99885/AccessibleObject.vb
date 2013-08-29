' Copyright (c) Microsoft Corporation.  All rights reserved.

Imports Accessibility
Imports System.Text
Imports System.Runtime.InteropServices

Namespace MsaaVerify

#Region "Enums"
    ' determines whether the object is a list view or a list box
    Public Enum ListType
        Neither = 0
        ListBox = 1
        ListView = 2
    End Enum

    ' determines what win32 state the checkbox is in
    Public Enum CheckBoxState
        UnChecked = 0
        Checked = 1
        InDeterminate = 2
    End Enum

    ' represents the possible combinations of states a control can be in
    ' these are taken from oleacc.h
    Public Enum MsaaStates
        Unavailable = &H1
        Selected = &H2
        Focused = &H4
        Pressed = &H8
        Checked = &H10
        Mixed = &H20
        [ReadOnly] = &H40
        HotTracked = &H80
        [Default] = &H100
        Expanded = &H200
        Collapsed = &H400
        Busy = &H800
        Floating = &H1000
        Marqueed = &H2000
        Animated = &H4000
        Invisible = &H8000
        OffScreen = &H10000
        Sizeable = &H20000
        Moveable = &H40000
        SelfVoicing = &H80000
        Focusable = &H100000
        Selectable = &H200000
        Linked = &H400000
        Traversed = &H800000
        MultiSelectable = &H1000000
        ExtSelectable = &H2000000
        AlertLow = &H4000000
        AlertMedium = &H8000000
        AlertHigh = &H10000000
        Valid = &H1FFFFFFF
        HasPopUp = &H40000000
        [Protected] = &H20000000
        Normal = 0
    End Enum

    ' represents the type of control or role
    ' these are taken from oleacc.h
    Public Enum MsaaRoles
        TitleBar = &H1
        MenuBar = &H2
        ScrollBar = &H3
        Grip = &H4
        Sound = &H5
        Cursor = &H6
        Caret = &H7
        Alert = &H8
        Window = &H9
        Client = &HA
        MenuPopup = &HB
        MenuItem = &HC
        MenuButton = &H39
        ToolTip = &HD
        Application = &HE
        Document = &HF
        Pane = &H10
        Chart = &H11
        Dialog = &H12
        Border = &H13
        Grouping = &H14
        Separator = &H15
        ToolBar = &H16
        StatusBar = &H17
        Table = &H18
        ColumnHeader = &H19
        RowHeader = &H1A
        Column = &H1B
        Row = &H1C
        Cell = &H1D
        Link = &H1E
        HelpBalloon = &H1F
        Character = &H20
        List = &H21
        ListItem = &H22
        Outline = &H23
        OutlineItem = &H24
        PageTab = &H25
        PropertyPage = &H26
        Indicator = &H27
        Graphic = &H28
        StaticText = &H29
        Text = &H2A
        PushButton = &H2B
        CheckButton = &H2C
        RadioButton = &H2D
        ComboBox = &H2E
        DropList = &H2F
        ProgressBar = &H30
        Dial = &H31
        HotKeyField = &H32
        Slider = &H33
        SpinButton = &H34
        Diagram = &H35
        Animation = &H36
        Equation = &H37
        ButtonDropDown = &H38
        ButtonMenu = &H39
        ButtonDropDownGrid = &H3A
        Whitespace = &H3B
        PageTabList = &H3C
        Clock = &H3D
        None = 0
    End Enum

#End Region

    Public Class AccessibleObject

#Region "Private Const and Variables"
        ' the accessible object's interface
        Private m_msaa As Accessibility.IAccessible
        ' the accessible object's child id 
        Private m_childId As Integer
        ' the hwnd for the accessible object's control
        Private m_hwnd As IntPtr

        ' used for string builder
        Private Const MaxRole As Integer = 128
        ' used for string builder
        Private Const MaxState As Integer = 256
        ' if it doesn't have a child or is the parent object, the id is 0
        Public Const ChildId_Self As Integer = &H0
        ' we're only interested in the client, not the window or native object model or whatever else
        Private Const MsaaObjectIdClient As Integer = &HFFFFFFFC
#End Region

#Region "Constructors"

        ' Create an ActiveAccessibility object from point
        ' See http://msdn2.microsoft.com/en-gb/library/ms696163.aspx
        Public Sub New(ByVal x As Integer, ByVal y As Integer)
            Dim childIdTemp As Object
            Dim hresult As Integer = NativeMethods.AccessibleObjectFromPoint(lx:=x, ly:=y, ppoleAcc:=m_msaa, pvarElement:=childIdTemp)
            If hresult = 0 Then
                m_childId = childIdTemp
            Else
                Throw New FailedToCreateAccessibleObject("Failed to create accessible object from point.  Got hresult " + hresult.ToString)
            End If
        End Sub

        ' Create an ActiveAccessibility object from an existing IAccessible interface
        Public Sub New(ByVal accessibleObject As IAccessible)
            m_msaa = accessibleObject
            m_childId = ChildID
        End Sub

        ' Create an ActiveAccessibility object from an existing IAccessible interface for a specified child object
        Public Sub New(ByVal accessibleObject As IAccessible, ByVal childID As Integer)
            m_msaa = accessibleObject
            m_childId = childID
        End Sub

        ' Create an ActiveAccessibility object from a known hwnd
        Public Sub New(ByVal windowHandle As IntPtr)
            'This defines the riid for the IAccessible interface
            Dim riidGUID As New Guid("{618736E0-3C3D-11CF-810C-00AA00389B71}")

            ' We pass in MSAA Object ID client, since we're most likely never going to use Office Object Model
            ' If we were to use the Office Object Model, we would pass in MsaaObjectID.NativeOM
            ' read more at http://msdn.microsoft.com/library/en-us/msaa/msaaccrf_2r7b.asp
            Dim hresult As Integer = NativeMethods.AccessibleObjectFromWindow(windowHandle, MsaaObjectIdClient, riidGUID, m_msaa)
            If hresult = 0 Then
                Me.SetHwnd(windowHandle)
                m_childId = AccessibleObject.ChildId_Self ' AccessibleObjectFromWindow returns the object itself and not one of its simple elements or children
            Else
                Throw New FailedToCreateAccessibleObject("Failed to create accessible object from window.  Got hresult " + hresult.ToString)
            End If
        End Sub
#End Region

#Region "IAccessible Properties and Methods Used"

        ' get a pointer to the requested child object, if possible
        Public ReadOnly Property Child(ByVal index As Integer) As AccessibleObject
            Get
                If Me.ChildCount > 0 Then
                    If index < 0 OrElse index >= Me.ChildCount Then
                        Throw New ChildNotFoundException("Index provided is less than 0 or greater than the number of children")
                    End If

                    ' return the child object
                    Return Me.Children(index)
                Else
                    Throw New ObjectIsElementException("Cannot access Child property since this object is an element")
                End If
            End Get
        End Property

        ' get the child id we cached when we constructed the object
        Public ReadOnly Property ChildID() As Integer
            Get
                Return m_childId
            End Get
        End Property

        ' get the number of children the object has
        Public Overloads ReadOnly Property ChildCount() As Integer
            Get
                If m_childId = 0 Then
                    Return m_msaa.accChildCount
                Else
                    ' we got a child element, so call its parent for the actual count
                    Return CType(m_msaa.accChild(m_childId), IAccessible).accChildCount
                End If
            End Get
        End Property

        ' get the control's default action
        Public Overloads ReadOnly Property DefaultAction() As String
            Get
                Return Me.GetNonNullString(m_msaa.accDefaultAction(Me.ChildID))
            End Get
        End Property

        ' get the control's description
        Public Overloads ReadOnly Property Description() As String
            Get
                Return Me.GetNonNullString(m_msaa.accDescription(Me.ChildID))
            End Get
        End Property

        ' get the contorl's keyboard shortcut
        Public Overloads ReadOnly Property KeyboardShortcut() As String
            Get
                Return Me.GetNonNullString(m_msaa.accKeyboardShortcut(Me.ChildID))
            End Get
        End Property

        ' get the control's name
        Public Overloads ReadOnly Property Name() As String
            Get
                Return Me.GetNonNullString(m_msaa.accName(Me.ChildID))
            End Get
        End Property

        ' get the control's parent
        Public ReadOnly Property Parent() As AccessibleObject
            Get
                If (m_childId = 0) Then
                    Return New AccessibleObject(AccessibleObject:=CType(m_msaa.accParent, IAccessible))
                Else
                    Return New AccessibleObject(AccessibleObject:=m_msaa)
                End If
            End Get
        End Property

        ' get the control's role
        Public Overloads ReadOnly Property Role() As Integer
            Get
                Return CInt(m_msaa.accRole(m_childId))
            End Get
        End Property

        ' get the control's role as a string
        Public ReadOnly Property RoleText() As String
            Get
                Dim roleID As Integer = CInt(m_msaa.accRole(m_childId))
                Dim roleString As New StringBuilder(MaxRole)
                NativeMethods.GetRoleText(roleID, roleString, MaxRole)
                Return roleString.ToString()
            End Get
        End Property

        ' get the control's state.
        Public ReadOnly Property State() As Integer
            Get
                Return CInt(m_msaa.accState(Me.ChildID))
            End Get
        End Property

        ' get the control's state as a string
        Public ReadOnly Property StateText() As String
            Get
                Dim stateBit As Integer
                Dim tmpStateText As String = ""

                Dim state As Integer = CInt(m_msaa.accState(Me.ChildID))
                Dim stateString As New StringBuilder(MaxState)

                ' get the state text for the normal state separately
                If state = 0 Then
                    NativeMethods.GetStateText(stateBit, stateString, MaxState)
                Else
                    For Each stateBit In [Enum].GetValues(GetType(MsaaStates))
                        If (state And stateBit) = stateBit Then
                            NativeMethods.GetStateText(stateBit, stateString, MaxState)
                            tmpStateText = tmpStateText + "," + stateString.ToString()
                        End If
                    Next stateBit
                End If

                Return tmpStateText.TrimStart(","c)
            End Get
        End Property

        ' get the control's value
        Public ReadOnly Property Value() As String
            Get
                Return Me.GetNonNullString(m_msaa.accValue(Me.ChildID))
            End Get
        End Property
#End Region

#Region "Public Methods"

        ' get the control's hwnd
        Public ReadOnly Property Hwnd() As IntPtr
            Get
                If Me.m_hwnd.ToInt32 = 0 Then
                    Dim handle As IntPtr

                    ' get a handle from the accessible object we have
                    NativeMethods.WindowFromAccessibleObject(m_msaa, handle)

                    ' create an accessible object
                    Dim msaa As New AccessibleObject(Me.IAccessible)

                    ' cache the hwnd for future use
                    m_hwnd = handle

                    Return handle
                Else
                    Return Me.m_hwnd
                End If
            End Get
        End Property

        ' get the pointer to the control's msaa interface
        Public ReadOnly Property IAccessible() As IAccessible
            Get
                Return m_msaa
            End Get
        End Property

        ' get the control's x coordinate
        Public ReadOnly Property X() As Integer
            Get
                Try
                    Dim x0, y0, width, height As Integer
                    m_msaa.accLocation(x0, y0, width, height, m_childId)
                    Return x0
                Catch ex As COMException
                    Return 0
                End Try
            End Get
        End Property

        ' get the control's y coordinate
        Public ReadOnly Property Y() As Integer
            Get
                Try
                    Dim x0, y0, width, height As Integer
                    m_msaa.accLocation(x0, y0, width, height, m_childId)
                    Return y0
                Catch ex As COMException
                    Return 0
                End Try
            End Get
        End Property

        ' get the control's width
        Public ReadOnly Property Width() As Integer
            Get
                Try
                    Dim x0, y0, w, h As Integer
                    m_msaa.accLocation(x0, y0, w, h, m_childId)
                    Return w
                Catch ex As COMException
                    Return 0
                End Try
            End Get
        End Property

        ' get the control's height
        Public ReadOnly Property Height() As Integer
            Get
                Try
                    Dim x0, y0, w, h As Integer
                    m_msaa.accLocation(x0, y0, w, h, m_childId)
                    Return h
                Catch ex As COMException
                    Return 0
                End Try
            End Get
        End Property

        ' get the classname for the control
        Public ReadOnly Property ClassName() As String
            Get
                Dim windowClassString As New String("0"c, 256)
                Dim windowClassPtr As IntPtr = Marshal.StringToHGlobalAnsi(windowClassString)
                Dim count As Integer = NativeMethods.GetClassName(Me.Hwnd, windowClassPtr, 255)
                windowClassString = Marshal.PtrToStringAnsi(windowClassPtr, count)
                Marshal.FreeHGlobal(windowClassPtr)
                Return windowClassString
            End Get
        End Property

        ' get the control's caption
        Public ReadOnly Property Caption() As String
            Get
                Dim chararray() As Char

                Dim length As Integer = CInt(NativeMethods.GetWindowTextLength(m_hwnd) And &HFFFF&)
                ReDim chararray(length + 1)
                length = NativeMethods.GetWindowText(m_hwnd, chararray, length + 1)

                Return New String(chararray, 0, length)
            End Get
        End Property

        ' get the collection of msaa children for the control
        Public ReadOnly Property Children() As ActiveAccessibilityCollection
            Get

                Dim childrenReturned As Integer
                Dim theChildren() As Object
                Dim childIAcc As IAccessible
                Dim index As Integer

                Dim childCount As Integer = Me.ChildCount
                Dim result As New ActiveAccessibilityCollection

                If Me.ChildCount > 0 Then
                    ReDim theChildren(Me.ChildCount - 1)
                    Try
                        NativeMethods.AccessibleChildren(m_msaa, 0, childCount, theChildren, childrenReturned)
                    Catch ex As COMException
                    End Try

                    For index = 0 To childrenReturned - 1
                        If TypeOf theChildren(index) Is IAccessible Then
                            childIAcc = CType(theChildren(index), IAccessible)
                        Else
                            childIAcc = Nothing
                        End If

                        If Not childIAcc Is Nothing Then
                            ' the cast to IAccessible was successful, so just add it to the collection
                            result.Add(New AccessibleObject(childIAcc))
                        Else
                            ' wouldn't cast, so it's a child ID
                            result.Add(New AccessibleObject(m_msaa, CInt(theChildren(index))))
                        End If
                    Next
                End If

                Return result
            End Get
        End Property

        ' get the number of items in a list box
        Public ReadOnly Property ListboxCount() As Integer
            Get
                Dim result As IntPtr = NativeMethods.SendMessage(Me.Hwnd, NativeMethods.LB_GETCOUNT, IntPtr.Zero, IntPtr.Zero)
                If result.ToInt32() = NativeMethods.LB_ERR Then
                    Throw New NativeMethodCallFailedException("Failed to get the count of ListBox items")
                End If

                Return result.ToInt32()
            End Get
        End Property

        ' get the text in the list box
        Public ReadOnly Property ListboxText(ByVal listItemIndex As Integer) As String
            Get
                Dim lParam As New StringBuilder(NativeMethods.NumBytes)
                Dim length As Integer = NativeMethods.SendMessageByStringBuilder(Me.Hwnd, NativeMethods.LB_GETTEXT, New IntPtr(listItemIndex), lParam).ToInt32()
                If length = NativeMethods.LB_ERR Then
                    Throw New NativeMethodCallFailedException("Failed to get the text in the list box")
                End If

                Return lParam.ToString()
            End Get
        End Property

        ' get the text in the edit box
        Public ReadOnly Property TextboxText() As String
            Get
                Dim lParam As New StringBuilder(NativeMethods.NumBytes)
                Dim result As IntPtr = NativeMethods.SendMessageByStringBuilder(Me.Hwnd, NativeMethods.WM_GETTEXT, New IntPtr(NativeMethods.NumBytes), lParam)
                If result.ToInt32 > 0 Then
                    TextboxText = lParam.ToString()
                Else
                    Throw New NativeMethodCallFailedException("Failed to get the text in the edit box")
                End If
            End Get
        End Property

        ' get the check box's state
        Public ReadOnly Property CheckBoxState() As CheckBoxState
            Get
                ' todo: Are there any messages we can send to the checkbox to determine whether it is really checked or not?
            End Get
        End Property

        ' get the number of items in the combo box
        Public ReadOnly Property ComboBoxCount() As Integer
            Get
                Dim result As IntPtr = NativeMethods.SendMessage(Me.Hwnd, NativeMethods.CB_GETCOUNT, IntPtr.Zero, IntPtr.Zero)
                If result.ToInt32() = NativeMethods.CB_ERR Then
                    Throw New NativeMethodCallFailedException("Failed to get the count of ComboBox items")
                Else
                    Return result.ToInt32()
                End If
            End Get
        End Property

        ' get the text in the combo box
        Public Overloads ReadOnly Property ComboBoxText() As String
            Get
                Return ComboBoxText(-1)
            End Get
        End Property

        ' get the text of the specified combo box list item
        Public Overloads ReadOnly Property ComboBoxText(ByVal index As Integer) As String
            Get
                Dim result As IntPtr
                Dim lParam As New StringBuilder(NativeMethods.NumBytes \ 2)

                If index = -1 Then
                    index = Me.ComboBoxSelectedIndex
                End If

                If index < -1 Or index >= ComboBoxCount Then
                    Throw New System.ArgumentException("Requested ComboBox item index is invalid")
                End If

                If index = -1 Then
                    Throw New NativeMethodCallFailedException("Failed to get the text from the combo box")
                Else
                    result = NativeMethods.SendMessageByStringBuilder(Me.Hwnd, NativeMethods.CB_GETLBTEXT, New IntPtr(index), lParam)
                    If result.ToInt64() > 0 Then
                        Return lParam.ToString
                    Else
                        Throw New NativeMethodCallFailedException("Failed to get the text from the combo box")
                    End If
                End If
            End Get
        End Property
#End Region

#Region "Private Methods"

        ' cache the control's hwnd
        Private Sub SetHwnd(ByVal windowHandle As IntPtr)
            m_hwnd = windowHandle
        End Sub

        ' get the combo box's selected index
        Private ReadOnly Property ComboBoxSelectedIndex() As Integer
            Get
                Return NativeMethods.SendMessage(Me.Hwnd, NativeMethods.CB_GETCURSEL, IntPtr.Zero, IntPtr.Zero).ToInt32()
            End Get
        End Property

        ' return an empty string for any null values returned
        Private Function GetNonNullString(ByVal str As String) As String
            If str Is Nothing Then
                Return ""
            Else
                Return str
            End If
        End Function
#End Region

    End Class

#Region "ActiveAccessibilityCollection"
    Public Class ActiveAccessibilityCollection : Inherits CollectionBase
        Public Sub New()
            MyBase.New()
        End Sub

        Default Public ReadOnly Property Item(ByVal index As Integer) As AccessibleObject
            Get
                Return CType(List(index), AccessibleObject)
            End Get
        End Property

        Public Function Add(ByVal value As AccessibleObject) As Integer
            Return List.Add(value)
        End Function
    End Class
#End Region

End Namespace