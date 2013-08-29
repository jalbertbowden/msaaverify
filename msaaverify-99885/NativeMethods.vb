' Copyright (c) Microsoft Corporation.  All rights reserved.

Imports System.Text
Imports System.Runtime.InteropServices

Namespace MsaaVerify
    MustInherit Class NativeMethods

#Region "Constants"

        Public Const NumChars As Integer = 32512
        Public Const NumBytes As Integer = NumChars * 2
        Public Const WM_USER As Integer = &H400
        Public Const WM_GETTEXT As Integer = &HD

        ' combo box constants
        Public Const CB_GETCOUNT As Integer = &H146
        Public Const CB_ERR As Integer = (-1)
        Public Const CB_GETLBTEXT As Integer = &H148
        Public Const CB_GETCURSEL As Integer = &H147

        ' list box constants
        Public Const LB_GETCOUNT As Integer = &H18B
        Public Const LB_ERR As Integer = (-1)
        Public Const LB_GETTEXT As Integer = &H189

        ' COM Results Constant from winerror.h - needed for MsaaVerify for certain WinControls
        Public Const DISP_E_MEMBERNOTFOUND As Integer = &H80020003
#End Region

        ' user32
        Public Declare Unicode Function GetWindowText Lib "user32" Alias "GetWindowTextW" (ByVal HWnd As IntPtr, ByVal byteArray As Char(), ByVal cch As Integer) As Integer
        Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthW" (ByVal HWnd As IntPtr) As Integer
        Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
        Public Declare Unicode Function SendMessageByStringBuilder Lib "user32" Alias "SendMessageW" (ByVal HWnd As IntPtr, ByVal wMsg As Integer, ByVal wParam As IntPtr, ByVal lParam As StringBuilder) As IntPtr
        Public Declare Function IsWindow Lib "user32" Alias "IsWindow" (ByVal HWnd As IntPtr) As Integer

        ' needed for Spy tools
        Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal HWnd As Integer, ByVal lpString As System.Text.StringBuilder, ByVal cch As Integer) As Integer
        Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal HWnd As IntPtr, ByVal lpClassName As IntPtr, ByVal nMaxCount As Integer) As Integer

        ' GDI32 - needed to draw the little box around the AA object found while search like AccExplorer
        Public Declare Function BitBlt Lib "Gdi32" Alias "BitBlt" (ByVal hdcDest As IntPtr, ByVal nXDest As Integer, ByVal nYDest As Integer, ByVal nWidth As Integer, ByVal nHeight As Integer, ByVal hdcSrc As IntPtr, ByVal nXSrc As Integer, ByVal nYSrc As Integer, ByVal dwRop As Integer) As Boolean
        Public Declare Function CreateCompatibleDC Lib "Gdi32" Alias "CreateCompatibleDC" (ByVal hDC As IntPtr) As IntPtr
        Public Declare Function CreateCompatibleBitmap Lib "Gdi32" Alias "CreateCompatibleBitmap" (ByVal hDC As IntPtr, ByVal Width As Integer, ByVal Height As Integer) As IntPtr
        Public Declare Function SelectObject Lib "Gdi32" Alias "SelectObject" (ByVal hDC As IntPtr, ByVal hgdiobj As IntPtr) As IntPtr
        Public Declare Function DeleteDC Lib "gdi32.dll" Alias "DeleteDC" (ByVal hdc As IntPtr) As Boolean
        Public Declare Function DeleteObject Lib "gdi32.dll" Alias "DeleteObject" (ByVal handle As IntPtr) As Boolean
        Public Declare Function GetDC Lib "user32" Alias "GetDC" (ByVal hWnd As Integer) As IntPtr

        ' oleacc.dll
        Public Declare Function AccessibleChildren Lib "oleacc" Alias "AccessibleChildren" (ByVal paccContainer As Accessibility.IAccessible, ByVal iChildStart As Integer, ByVal cChildren As Integer, <MarshalAs(UnmanagedType.LPArray, SizeParamIndex:=2), InAttribute(), Out()> ByVal rgvarChildren() As Object, ByRef pcObtained As Integer) As Integer
        Public Declare Function AccessibleObjectFromPoint Lib "oleacc" Alias "AccessibleObjectFromPoint" (ByVal lx As Integer, ByVal ly As Integer, ByRef ppoleAcc As Accessibility.IAccessible, ByRef pvarElement As Object) As Integer
        Public Declare Function AccessibleObjectFromWindow Lib "oleacc" Alias "AccessibleObjectFromWindow" (ByVal hwnd As IntPtr, ByVal dwId As Integer, ByRef riid As Guid, ByRef ppvObject As Accessibility.IAccessible) As Integer
        Public Declare Function GetStateText Lib "oleacc" Alias "GetStateTextA" (ByVal dwStateBit As Integer, ByVal szState As StringBuilder, ByVal cchStateBitMax As Short) As Integer
        Public Declare Function WindowFromAccessibleObject Lib "oleacc" Alias "WindowFromAccessibleObject" (ByVal paccIdentity As Accessibility.IAccessible, ByRef hwnd As IntPtr) As Integer
        Public Declare Function GetRoleText Lib "oleacc" Alias "GetRoleTextA" (ByVal dwRole As Integer, ByVal szRole As StringBuilder, ByVal cchRoleMax As Short) As Integer
    End Class
End Namespace
