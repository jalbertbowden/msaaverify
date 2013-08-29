' Copyright (c) Microsoft Corporation.  All rights reserved.
' hello world

Namespace MsaaVerify

    Public Class ChildNotFoundException : Inherits Exception
        Public Sub New(ByVal message As String)
            MyBase.New(message)
        End Sub
    End Class

    Public Class NativeMethodCallFailedException : Inherits Exception
        Public Sub New(ByVal message As String)
            MyBase.New(message)
        End Sub
    End Class

    Public Class ObjectIsElementException : Inherits Exception
        Public Sub New(ByVal message As String)
            MyBase.New(message)
        End Sub
    End Class

    Public Class FailedToCreateAccessibleObject : Inherits Exception
        Public Sub New(ByVal message As String)
            MyBase.New(message)
        End Sub
    End Class
End Namespace