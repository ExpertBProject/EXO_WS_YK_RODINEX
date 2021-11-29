Imports System.IdentityModel
Imports System.IdentityModel.Selectors
Imports System.Security.Principal
Imports System.ServiceModel
Public Class Authenticator
    Inherits UserNamePasswordValidator

    Public Overrides Sub Validate(userName As String, password As String)
        If userName <> "YKRODINEX" OrElse password <> "YK982RNwh0" Then
            Throw New FaultException("Invalid user and/or password")
        End If
    End Sub
End Class
