' Copyright (c) Microsoft Corporation.  All rights reserved.

Imports System.Text
Imports System.IO
Imports System.CodeDom
Imports System.CodeDom.Compiler
Imports System.Runtime.InteropServices

Namespace MsaaVerify

    ' verifier
    Public MustInherit Class Verifier

        Public Shared Sub VerifyMsaaObjects(ByVal msaaObject As AccessibleObject)

            ' we test based on the accessible role - if the role is wrong, expect a high failure rate
            Select Case msaaObject.Role

                Case MsaaRoles.PushButton ' basic button control
                    ButtonVerifier.Verify(msaaObject)

                Case MsaaRoles.Text ' this is for an edit box control
                    EditBoxVerifier.Verify(msaaObject)
            End Select
        End Sub

        Protected MustInherit Class BaseVerifier

            Private Const nullDisplay As String = "<null>"
            Private Const nonNullDisplay As String = "<non-null value>"
            Private Const Passed As String = "Passed"
            Private Const Failed As String = "Failed"

            Protected Overloads Shared Sub ChildCount(ByVal aaObject As AccessibleObject, ByVal methodSupported As Boolean, ByVal expectedChildCount As Integer)
                Dim actualChildCount As Integer = 0 ' the value we get from MSAA
                Dim errorCode As Integer = 0 ' the return value we get back from the MSAA call.  Assume S_OK
                Dim message As String = "" ' any information regarding the test

                ' get the actual MSAA childcount
                Try
                    actualChildCount = aaObject.ChildCount()
                Catch ex As COMException
                    errorCode = ex.ErrorCode  ' didn't get back S_OK.  Record error code
                End Try

                ' verify and update message box as appropriate
                Dim results As Boolean = Verify(actualChildCount, expectedChildCount, methodSupported, errorCode, message)
                DisplayResults(results, expectedChildCount, actualChildCount, MainForm.ChildCountResults)
            End Sub

            Protected Overloads Shared Sub Description(ByVal aaObject As AccessibleObject, ByVal methodSupported As Boolean, ByVal expectedDescription As String)
                Dim actualDescription As String = 0 ' the value we get from MSAA
                Dim errorCode As Integer = 0 ' the return value we get back from the MSAA call.  Assume S_OK
                Dim message As String = "" ' any information regarding the test

                ' get the actual MSAA description
                Try
                    actualDescription = aaObject.Description()
                Catch ex As COMException
                    errorCode = ex.ErrorCode  ' didn't get back S_OK.  Record error code
                End Try

                ' verify and update message box as appropriate
                Dim results As Boolean = Verify(actualDescription, expectedDescription, methodSupported, errorCode, message)
                DisplayResults(results, expectedDescription, actualDescription, MainForm.DescriptionResults)
            End Sub

            Protected Overloads Shared Sub DefaultAction(ByVal aaObject As AccessibleObject, ByVal methodSupported As Boolean, ByVal expectedDefaultAction As String)
                Dim actualDefaultAction As String = 0 ' the value we get from MSAA
                Dim errorCode As Integer = 0 ' the return value we get back from the MSAA call.  Assume S_OK
                Dim message As String = "" ' any information regarding the test

                ' get the actual MSAA DefaultAction
                Try
                    actualDefaultAction = aaObject.DefaultAction()
                Catch ex As COMException
                    errorCode = ex.ErrorCode  ' didn't get back S_OK.  Record error code
                End Try

                ' verify and update message box as appropriate
                Dim results As Boolean = Verify(actualDefaultAction, expectedDefaultAction, methodSupported, errorCode, message)
                DisplayResults(results, expectedDefaultAction, actualDefaultAction, MainForm.DefaultActionResults)
            End Sub

            Protected Overloads Shared Sub KeyboardShortcut(ByVal aaObject As AccessibleObject, ByVal methodSupported As Boolean, ByVal expectedKeyboardShortcut As String)
                Dim actualKeyboardShortcut As String = 0 ' the value we get from MSAA
                Dim errorCode As Integer = 0 ' the return value we get back from the MSAA call.  Assume S_OK
                Dim message As String = "" ' any information regarding the test

                ' get the actual MSAA KeyboardShortcut
                Try
                    actualKeyboardShortcut = aaObject.KeyboardShortcut()
                Catch ex As COMException
                    errorCode = ex.ErrorCode  ' didn't get back S_OK.  Record error code
                End Try

                ' verify and update message box as appropriate
                Dim results As Boolean = Verify(actualKeyboardShortcut, expectedKeyboardShortcut, methodSupported, errorCode, message)
                DisplayResults(results, expectedKeyboardShortcut, actualKeyboardShortcut, MainForm.KeyboardShortcutResults)
            End Sub

            Protected Overloads Shared Sub KeyboardShortcut(ByVal aaObject As AccessibleObject, ByVal methodSupported As Boolean)
                Dim actualKeyboardShortcut As String = 0 ' the value we get from MSAA
                Dim errorCode As Integer = 0 ' the return value we get back from the MSAA call.  Assume S_OK
                Dim message As String = "" ' any information regarding the test

                ' get the actual MSAA KeyboardShortcut
                Try
                    actualKeyboardShortcut = aaObject.KeyboardShortcut()
                Catch ex As COMException
                    errorCode = ex.ErrorCode  ' didn't get back S_OK.  Record error code
                End Try

                ' verify and update message box as appropriate
                Dim results As Boolean = Verify(actualKeyboardShortcut, True, errorCode, message)
                DisplayResults(results, nonNullDisplay, actualKeyboardShortcut, MainForm.KeyboardShortcutResults)
            End Sub

            Protected Overloads Shared Sub Name(ByVal aaObject As AccessibleObject, ByVal methodSupported As Boolean, ByVal expectedName As String)
                Dim actualName As String = 0 ' the value we get from MSAA
                Dim errorCode As Integer = 0 ' the return value we get back from the MSAA call.  Assume S_OK
                Dim message As String = "" ' any information regarding the test

                ' get the actual MSAA Name
                Try
                    actualName = aaObject.Name()
                Catch ex As COMException
                    errorCode = ex.ErrorCode  ' didn't get back S_OK.  Record error code
                End Try

                ' verify and update message box as appropriate
                Dim results As Boolean = Verify(actualName, expectedName, True, errorCode, message)
                DisplayResults(results, expectedName, actualName, MainForm.NameResults)
            End Sub

            Protected Overloads Shared Sub Name(ByVal aaObject As AccessibleObject, ByVal methodSupported As Boolean)
                Dim actualName As String = 0 ' the value we get from MSAA
                Dim errorCode As Integer = 0 ' the return value we get back from the MSAA call.  Assume S_OK
                Dim message As String = "" ' any information regarding the test

                ' get the actual MSAA Name
                Try
                    actualName = aaObject.Name()
                Catch ex As COMException
                    errorCode = ex.ErrorCode  ' didn't get back S_OK.  Record error code
                End Try

                ' verify and update message box as appropriate
                Dim results As Boolean = Verify(actualName, True, errorCode, message)
                DisplayResults(results, nonNullDisplay, actualName, MainForm.NameResults)
            End Sub

            Protected Overloads Shared Sub Parent(ByVal aaObject As AccessibleObject)
                Dim errorCode As Integer = 0 ' the return value we get back from the MSAA call.  Assume S_OK
                Dim message As String = "" ' any information regarding the test

                ' get the actual MSAA Parent
                Try
                    Dim aaParent As AccessibleObject = aaObject.Parent
                Catch ex As COMException
                    errorCode = ex.ErrorCode  ' didn't get back S_OK.  Record error code
                End Try

                ' verify and update message box as appropriate
                Dim results As Boolean = CType(errorCode = 0, Boolean)
                DisplayResults(results, "Return Code 0", "Return Code " + errorCode.ToString, MainForm.ParentResults)
            End Sub

            Protected Overloads Shared Sub Role(ByVal aaObject As AccessibleObject)
                Dim errorCode As Integer = 0 ' the return value we get back from the MSAA call.  Assume S_OK
                Dim message As String = "" ' any information regarding the test

                ' get the actual MSAA Role
                Try
                    Dim role As Integer = aaObject.Role
                Catch ex As COMException
                    errorCode = ex.ErrorCode  ' didn't get back S_OK.  Record error code
                End Try

                ' verify and update message box as appropriate
                Dim results As Boolean = CType(errorCode = 0, Boolean)
                DisplayResults(results, "Return Code 0", "Return Code " + errorCode.ToString, MainForm.RoleResults)
            End Sub

            Protected Overloads Shared Sub State(ByVal aaObject As AccessibleObject, ByVal expectedStates As Integer)
                Dim actualState As Integer = 0
                Dim stateBit As Integer = 0
                Dim errorCode As Integer = 0 ' the return value we get back from the MSAA call.  Assume S_OK
                Dim message As String = "" ' any information regarding the test

                ' get the actual MSAA State
                Try
                    actualState = aaObject.State
                Catch ex As COMException
                    errorCode = ex.ErrorCode  ' didn't get back S_OK.  Record error code
                End Try

                '' for each bit, for now, we're just going to check whether this control should support it as a state.
                '' TODO:  get the actual state from a user32.dll call (if possible) and then test the state
                '' if the above isn't possible, export the expected state to an xml file
                'For Each stateBit In [Enum].GetValues(GetType(MsaaStates))
                '    If (actualState And stateBit) = stateBit Then
                '        Select Case stateBit
                '            Case MsaaStates.Normal
                '            Case MsaaStates.ReadOnly
                '            Case MsaaStates.Invisible
                '            Case Else ' unexpected state
                '                message += aaObject.StateText + " "
                '        End Select
                '    End If
                'Next stateBit

                ' verify and update message box as appropriate
                Dim results As Boolean = CType(errorCode = 0, Boolean)
                DisplayResults(results, "Return Code 0", "Return Code " + errorCode.ToString, MainForm.StateResults)
            End Sub

            Protected Overloads Shared Sub Value(ByVal aaObject As AccessibleObject, ByVal methodSupported As Boolean, ByVal expectedValue As String)
                Dim actualValue As String = 0 ' the value we get from MSAA
                Dim errorCode As Integer = 0 ' the return value we get back from the MSAA call.  Assume S_OK
                Dim message As String = "" ' any information regarding the test

                ' get the actual MSAA Value
                Try
                    actualValue = aaObject.Value()
                Catch ex As COMException
                    errorCode = ex.ErrorCode  ' didn't get back S_OK.  Record error code
                End Try

                ' verify and update message box as appropriate
                Dim results As Boolean = Verify(actualValue, expectedValue, methodSupported, errorCode, message)
                DisplayResults(results, expectedValue, actualValue, MainForm.ValueResults)
            End Sub

            Protected Shared Function RemoveNmemonicAmpersand(ByVal caption As String) As String
                ' check whether the caption exists
                If Not caption = "" Then
                    ' BUG:  Yeah, there could be multiple & in the string. In that case the very first & wins...
                    caption = caption.Replace("&", "")
                End If
                Return caption
            End Function

            Protected Shared Function GetNmemonic(ByVal caption As String) As String
                Dim nmemonic As String = ""
                If Not caption = "" Then
                    Dim index As Integer = caption.LastIndexOf("&")
                    If index > -1 Then
                        ' in the case that the first letter is the nmemonic, lower case it
                        ' BUG:  Need to localize this
                        nmemonic = "Alt+" + caption.Substring(index + 1, 1).ToLower
                    End If
                End If

                Return nmemonic
            End Function

#Region "BaseVerifier Private Methods"
            Private Shared Sub DisplayResults(ByVal success As Boolean, ByVal expected As String, ByVal actual As String, ByVal txtBox As TextBox)
                If expected = "" Then
                    expected = nullDisplay
                End If

                If actual = "" Then
                    actual = nullDisplay
                End If

                Dim message As New StringBuilder(" - Expected: " + expected + " - " + " Actual: " + actual)

                If success Then
                    message.Insert(0, Passed)
                    txtBox.Text = message.ToString
                Else
                    message.Insert(0, Failed)
                    txtBox.Text = message.ToString
                    txtBox.ForeColor = Color.Red
                End If
            End Sub

            Private Shared Function Verify(ByVal actual As String, ByVal expected As String, ByVal methodSupported As Boolean, ByVal errorCode As Integer, ByRef message As String) As Boolean
                Dim hrSuccess As Boolean = False ' always fail by default
                hrSuccess = VerifyHResult(methodSupported, errorCode, message)

                ' verify the expected value against the actual value
                ' TODO:  think about how this will work for keyboard shortcuts...
                Return CType(hrSuccess AndAlso CType(actual = expected, Boolean), Boolean)
            End Function

            ' verify non-null values
            Private Shared Function Verify(ByVal actual As String, ByVal methodSupported As Boolean, ByVal errorCode As Integer, ByRef message As String) As Boolean
                Dim hrSuccess As Boolean = False ' always fail by default
                hrSuccess = VerifyHResult(methodSupported, errorCode, message)

                ' verify non-null value
                If actual Is Nothing OrElse actual = "" Then
                    Return CType(hrSuccess AndAlso False, Boolean)
                Else
                    Return CType(hrSuccess AndAlso True, Boolean)
                End If

            End Function

            Private Shared Function VerifyHResult(ByVal methodSupported As Boolean, ByVal errorCode As Integer, ByRef message As String) As Boolean
                ' If the control type supports the property, the return code is S_OK
                ' If the control type does not support child count, either S_OK or DISP_E_MEMBERNOTFOUND can be returned
                ' but any other COM HResult returned is a failure
                If (methodSupported AndAlso errorCode = 0) OrElse _
                    (Not methodSupported AndAlso (errorCode = NativeMethods.DISP_E_MEMBERNOTFOUND OrElse errorCode = 0)) Then
                    Return True
                Else
                    Return False
                    message = "Unexpected HResult.  Expected DISP_E_MEMBERNOTFOUND or S_OK, but got error code number " + errorCode.ToString + " instead"
                End If
            End Function

#End Region
        End Class

        Protected MustInherit Class ButtonVerifier : Inherits BaseVerifier

            Public Shared Sub Verify(ByVal aaObject As AccessibleObject)
                ' a button does not have any children
                BaseVerifier.ChildCount(aaObject, False, 0)

                ' a button does not have a description
                BaseVerifier.Description(aaObject, True, "")

                ' a button has a default action of Press
                ' BUG:  Use ResourceRefactoring tool to grab the localized text from oleacc.dll
                BaseVerifier.DefaultAction(aaObject, True, "Press")

                ' TODO:  The method GetWindowText() that we call to abstract the nmemonic is the same
                ' method that MSAA calls to set the nmemonic, so this isn't much of a test right now, unless the keyboard shortcut was overridden
                ' ideally, we'll have an xml file that the user specifies values such as this that we pull from...
                ' a button's keyboard shortcut is its nmemonic or access key, i.e. the underlined letter.
                Dim nmemonic As String = GetNmemonic(aaObject.Caption())
                BaseVerifier.KeyboardShortcut(aaObject, True, nmemonic)

                ' TODO:  The method GetWindowText() that we call to get the accessible name is thesame
                ' method that MSAA calls to set the name, so this isn't much of a test right now, unless the MSAA name was overridden 
                ' ideally, we'll have an xml file that the user specifies values such as this that we pull from...
                ' a button's name is the same as its caption
                Dim nameWithoutAmpersand As String = RemoveNmemonicAmpersand(aaObject.Caption)
                BaseVerifier.Name(aaObject, True, nameWithoutAmpersand)

                ' only thing I want to verify right now is that we get a successful return code on the Parent invocation
                BaseVerifier.Parent(aaObject)

                ' only thing I want to verify right now is that we get a successful return code when we retrieve the role
                BaseVerifier.Role(aaObject)

                ' TODO: to be completed
                BaseVerifier.State(aaObject, MsaaStates.Unavailable)

                ' buttons don't have a value
                BaseVerifier.Value(aaObject, False, "")
            End Sub
        End Class

        Protected MustInherit Class EditBoxVerifier : Inherits BaseVerifier

            Public Shared Sub Verify(ByVal aaObject As AccessibleObject)
                ' a text box does not have any children
                BaseVerifier.ChildCount(aaObject, True, 0)

                ' a text box does not have a description
                BaseVerifier.Description(aaObject, True, "")

                ' a text box does not have a default actions, so verify DISP_E_MEMBERNOTFOUND is returned
                BaseVerifier.DefaultAction(aaObject, False, "")

                'TODO:  Either visual identification or grabbing a control based on its proximity to the edit box
                'is required to correctly identify the keyboardShortcut.  For now, we'll verify the keyboard shortcut is non-null
                BaseVerifier.KeyboardShortcut(aaObject, True)

                'TODO:  Either visual identification or grabbing a control based on its proximity to the edit box
                'is required to correctly identify the name.  For now, we'll verify the name is non-null
                'a text box name comes from the label that immediately preceeds it in tab order
                BaseVerifier.Name(aaObject, True)

                ' only thing I want to verify right now is that we get a successful return code on the Parent invocation
                BaseVerifier.Parent(aaObject)

                ' only thing I want to verify right now is that we get a successful return code when we retrieve the role
                BaseVerifier.Role(aaObject)

                ' TODO: to be completed
                BaseVerifier.State(aaObject, MsaaStates.Unavailable)

                ' a text box value is its contents
                ' TODO:  this textbox call should live somewhere else, like a User32 wrapper class
                Dim txtBoxContents As String = aaObject.TextboxText()
                BaseVerifier.Value(aaObject, True, txtBoxContents)
            End Sub
        End Class
    End Class
End Namespace