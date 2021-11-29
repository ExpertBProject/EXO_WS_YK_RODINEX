Option Strict On

Imports System.Data.SqlClient
Imports System.Xml
Imports System.IO
Imports System.Reflection
Imports SAPbobsCOM
Imports System.Security.Cryptography

Public MustInherit Class SDEXO

#Region "Attributes"

    Private _connectionString As String
    Private _databaseName As String
    Private _eXO_db As SqlConnection
    Private _company As SAPbobsCOM.Company
    Private _configPath As String = System.AppDomain.CurrentDomain.BaseDirectory() & "\config"

#End Region

#Region "Properties"

    Protected Property ConnectionString() As String
        Get
            Return _connectionString
        End Get
        Set(ByVal value As String)
            _connectionString = value
        End Set
    End Property

    Protected Property EXO_db() As SqlConnection
        Get
            Return _eXO_db
        End Get
        Set(ByVal value As SqlConnection)
            _eXO_db = value
        End Set
    End Property

    Protected Property Company() As SAPbobsCOM.Company
        Get
            Return _company
        End Get
        Set(ByVal value As SAPbobsCOM.Company)
            _company = value
        End Set
    End Property

#End Region

#Region "Constructors"

    Protected Sub New()

    End Sub

    Protected Sub New(ByVal connectionString As String)
        Me.ConnectionString = connectionString
    End Sub

#End Region

#Region "Methods"

    Protected Sub ConnectSQLServer()
        Dim ConfigXMLDocument As XmlDocument = Nothing

        Try
            If Me.EXO_db Is Nothing OrElse Me.EXO_db.State = ConnectionState.Closed Then
                ConfigXMLDocument = New XmlDocument
                ConfigXMLDocument.Load(_configPath & "\RODINEX.config")

                If Not ConfigXMLDocument Is Nothing Then
                    If Not ConfigXMLDocument.DocumentElement Is Nothing Then
                        If ConfigXMLDocument.DocumentElement.Name.ToLower = "configuration" Then
                            If ConfigXMLDocument.DocumentElement.HasChildNodes Then
                                For i As Integer = 0 To ConfigXMLDocument.DocumentElement.ChildNodes.Count - 1
                                    If ConfigXMLDocument.DocumentElement.ChildNodes(i).Name.ToLower = "setup" Then
                                        If ConfigXMLDocument.DocumentElement.ChildNodes(i).HasChildNodes Then
                                            For j As Integer = 0 To ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes.Count - 1
                                                If ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Name.ToUpper = "SQL" Then
                                                    For k As Integer = 0 To ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Attributes.Count - 1
                                                        If ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Attributes(k).Name.ToLower = "connectionstring" Then
                                                            Me.ConnectionString = ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Attributes(k).InnerText
                                                        End If
                                                    Next
                                                End If
                                            Next
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                End If

                Me.EXO_db = New SqlConnection(Me.ConnectionString)

                Me.EXO_db.Open()
            End If

        Catch ex As Exception
            Throw ex
        Finally
            ConfigXMLDocument = Nothing
        End Try
    End Sub

    Protected Sub DisconnectSQLServer()
        Try
            If Not Me.EXO_db Is Nothing AndAlso Me.EXO_db.State = ConnectionState.Open Then
                Me.EXO_db.Close()
            End If

        Catch ex As Exception
            Throw ex
        Finally
            If Not Me.EXO_db Is Nothing Then
                Me.EXO_db.Dispose()
                Me.EXO_db = Nothing
            End If
        End Try
    End Sub

    Protected Sub ConnectSAP()
        Dim ConfigXMLDocument As XmlDocument = Nothing

        Try
            If Me.Company Is Nothing OrElse Me.Company.Connected = False Then
                ConfigXMLDocument = New XmlDocument
                ConfigXMLDocument.Load(_configPath & "\RODINEX.config")

                If Not ConfigXMLDocument Is Nothing Then
                    If Not ConfigXMLDocument.DocumentElement Is Nothing Then
                        If ConfigXMLDocument.DocumentElement.Name.ToLower = "configuration" Then
                            If ConfigXMLDocument.DocumentElement.HasChildNodes Then
                                For i As Integer = 0 To ConfigXMLDocument.DocumentElement.ChildNodes.Count - 1
                                    If ConfigXMLDocument.DocumentElement.ChildNodes(i).Name.ToLower = "setup" Then
                                        If ConfigXMLDocument.DocumentElement.ChildNodes(i).HasChildNodes Then
                                            For j As Integer = 0 To ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes.Count - 1
                                                If ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Name.ToUpper = "DI" Then
                                                    Me.Company = New SAPbobsCOM.Company
                                                    Me.Company.UseTrusted = False

                                                    For k As Integer = 0 To ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Attributes.Count - 1
                                                        If ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Attributes(k).Name.ToLower = "server" Then
                                                            Me.Company.Server = ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Attributes(k).InnerText
                                                        End If
                                                        If ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Attributes(k).Name.ToLower = "database" Then
                                                            Me.Company.CompanyDB = ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Attributes(k).InnerText
                                                        End If
                                                        If ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Attributes(k).Name.ToLower = "user" Then
                                                            Me.Company.UserName = ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Attributes(k).InnerText
                                                        End If
                                                        If ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Attributes(k).Name.ToLower = "pass" Then
                                                            Me.Company.Password = ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Attributes(k).InnerText
                                                        End If
                                                        If ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Attributes(k).Name.ToLower = "dbuser" Then
                                                            Me.Company.DbUserName = ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Attributes(k).InnerText
                                                        End If
                                                        If ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Attributes(k).Name.ToLower = "dbpass" Then
                                                            Me.Company.DbPassword = ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Attributes(k).InnerText
                                                        End If
                                                        If ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Attributes(k).Name.ToLower = "idioma" Then
                                                            Select Case ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Attributes(k).InnerText
                                                                Case "IN"
                                                                    Me.Company.language = SAPbobsCOM.BoSuppLangs.ln_English
                                                                Case "ES"
                                                                    Me.Company.language = SAPbobsCOM.BoSuppLangs.ln_Spanish
                                                                Case "FR"
                                                                    Me.Company.language = SAPbobsCOM.BoSuppLangs.ln_French
                                                                Case "AL"
                                                                    Me.Company.language = SAPbobsCOM.BoSuppLangs.ln_German
                                                            End Select
                                                        End If

                                                        If ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Attributes(k).Name.ToLower = "versionsql" Then
                                                            If ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Attributes(k).InnerText = "SQL2005" Then
                                                                Me.Company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2005
                                                            ElseIf ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Attributes(k).InnerText = "SQL2008" Then
                                                                Me.Company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2008
                                                            ElseIf ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Attributes(k).InnerText = "SQL2012" Then
                                                                Me.Company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2012
                                                            ElseIf ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Attributes(k).InnerText = "SQL2016" Then
                                                                Me.Company.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_MSSQL2016
                                                            End If
                                                        End If

                                                        If ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Attributes(k).Name.ToLower = "licenseserver" Then
                                                            Me.Company.LicenseServer = ConfigXMLDocument.DocumentElement.ChildNodes(i).ChildNodes(j).Attributes(k).InnerText
                                                        End If
                                                    Next
                                                End If
                                            Next
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                End If

                If Me.Company.Connect <> 0 Then
                    Throw New Exception(Me.Company.GetLastErrorCode.ToString + "/" + Me.Company.GetLastErrorDescription)
                End If
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            ConfigXMLDocument = Nothing
        End Try
    End Sub

    Protected Sub DisconnectSAP()
        Try
            If Me.Company IsNot Nothing AndAlso Me.Company.Connected Then
                Me.Company.Disconnect()
            End If

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            If Me.Company IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.FinalReleaseComObject(Me.Company)
                Me.Company = Nothing
            End If
        End Try
    End Sub

#End Region

End Class


