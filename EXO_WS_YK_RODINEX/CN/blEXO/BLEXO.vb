Option Explicit On
Option Strict On

Imports SAPbobsCOM
Imports System.Data
Imports System.IO
Imports System.Data.SqlClient
Imports System.Xml
Imports System.Xml.XPath

Public Class BLEXO
    Inherits SDEXO

#Region " Constructor"

    Public Sub New()
        'constructor por defecto
    End Sub

#End Region

#Region "Métodos SQL Server"
    Public Function GetValueDB(ByRef sTable As String, ByRef sField As String, ByRef sCondition As String) As String
        Dim cmd As SqlCommand = Nothing
        Dim da As SqlDataAdapter = Nothing
        Dim dt As System.Data.DataTable = Nothing
        Dim sSQL As String = ""

        Try
            MyBase.ConnectSQLServer()

            If sCondition = "" Then
                sSQL = "SELECT " & sField & " FROM " & sTable
            Else
                sSQL = "SELECT " & sField & " FROM " & sTable & " WHERE " & sCondition
            End If

            cmd = New SqlCommand(sSQL, Me.EXO_db)
            cmd.CommandTimeout = 0

            da = New SqlDataAdapter

            da.SelectCommand = cmd
            dt = New System.Data.DataTable
            da.Fill(dt)

            If dt.Rows.Count <= 0 Then
                Return ""
            Else
                If Not IsDBNull(dt.Rows(0).Item(0).ToString) Then
                    Return dt.Rows(0).Item(0).ToString
                Else
                    Return ""
                End If
            End If

        Catch ex As Exception
            Throw ex
        Finally
            If Not dt Is Nothing Then
                dt.Dispose()
                dt = Nothing
            End If

            If Not cmd Is Nothing Then
                cmd.Dispose()
                cmd = Nothing
            End If

            If Not da Is Nothing Then
                da.Dispose()
                da = Nothing
            End If

            MyBase.DisconnectSQLServer()
        End Try
    End Function
    Public Sub FillDtDB(ByRef dt As System.Data.DataTable, ByVal sSQL As String)
        Dim cmd As SqlCommand = Nothing
        Dim da As SqlDataAdapter = Nothing

        Try
            MyBase.ConnectSQLServer()

            cmd = New SqlCommand(sSQL, Me.EXO_db)
            cmd.CommandTimeout = 0

            da = New SqlDataAdapter

            da.SelectCommand = cmd
            da.Fill(dt)

        Catch ex As Exception
            Throw ex
        Finally
            If Not cmd Is Nothing Then
                cmd.Dispose()
                cmd = Nothing
            End If

            If Not da Is Nothing Then
                da.Dispose()
                da = Nothing
            End If

            MyBase.DisconnectSQLServer()
        End Try
    End Sub
    Public Function ExecuteSqlDB(ByVal sSQL As String) As Boolean
        Dim cmd As SqlCommand = Nothing

        ExecuteSqlDB = False

        Try
            MyBase.ConnectSQLServer()

            cmd = New SqlCommand(sSQL, Me.EXO_db)
            cmd.ExecuteNonQuery()

            ExecuteSqlDB = True

        Catch ex As Exception
            Throw ex
        Finally
            If Not cmd Is Nothing Then
                cmd.Dispose()
                cmd = Nothing
            End If

            MyBase.DisconnectSQLServer()
        End Try
    End Function
#End Region

#Region "Métodos SAP"
    'Crear Pedido
    Public Function AddUpdateORDR(ByRef oxml As XDocument, ByRef oLog As EXO_Log.EXO_Log) As String
        Dim oORDR As SAPbobsCOM.Documents = Nothing
        Dim sDocEntry As String = ""
        Dim strNumSerie As String = ""

        Dim strExiste As String = ""
        Dim strCliente As String = ""
        Dim blEXO As BLEXO = Nothing
        Dim strFecha As String = ""
        Dim strFechaEnvio As String = ""

        Dim sNumDoc As String = ""
        Dim strRef As String = ""
        Dim strCodArt As String = ""
        Dim strComprobar As String = ""
        Dim strFP As String = ""
        Dim strEstado As String = ""
        Dim strDestino As String = ""
        Dim sAnno As String = ""


        Dim EarliestDeliveryDate As String = ""
        Dim TransportPriority As String = ""
        Dim Campaign As String = ""
        Dim DeliveryRemarks As String = ""
        Dim BlanketOrderReference As String = ""
        Dim CustomerReference_DocumentID As String = ""
        Dim BuyerParty_PartyID As String = ""
        Dim BuyerParty_AgencyCode As String = ""
        Dim OrderingParty_PartyID As String = ""
        Dim OrderingParty_AgencyCode As String = ""
        Dim Consignee_PartyID As String = ""
        Dim Consignee_AgencyCode As String = ""
        Dim PaymentTerms_PaymentMethod As String = ""
        Dim intLin As Integer = 0
        Dim strNomDir As String = ""
        Dim strDireccion As String = ""
        Dim strCP As String = ""
        Dim strProvincia As String = ""
        Dim strPais As String = ""
        Dim strTlfno As String = ""
        Dim aryTextFile() As String

        'EcoTasa
        Dim sSQL As String = ""
        'Dim sCeCo As String = ""
        Dim dtTipoEcotasa As System.Data.DataTable = Nothing
        Dim sEcotasa As String = ""
        Dim sTipoEcotasa As String = "0"
        Dim cPrecEcotasa As Double = 0
        Dim dtEcotasa As System.Data.DataTable = Nothing
        Dim cTotalesEcotasa As Double = 0
        Dim sPais As String = ""

        Dim dtEmpleado As System.Data.DataTable = Nothing
        Dim sEmpVentas As String = "0"
        Dim dtCeCo As System.Data.DataTable = Nothing
        Dim sCeCo As String = ""

        AddUpdateORDR = ""

        Try

            Try
                EarliestDeliveryDate = oxml.Descendants("EarliestDeliveryDate").Value.ToString
            Catch ex As Exception
                EarliestDeliveryDate = ""
            End Try

            Try
                TransportPriority = oxml.Descendants("TransportPriority").Value.ToString
            Catch ex As Exception
                TransportPriority = ""
            End Try

            Try
                Campaign = oxml.Descendants("Campaign").Value.ToString
            Catch ex As Exception
                Campaign = ""
            End Try

            Try
                DeliveryRemarks = oxml.Descendants("DeliveryRemarks").Value.ToString
            Catch ex As Exception
                DeliveryRemarks = ""
            End Try

            Try
                BlanketOrderReference = oxml.Descendants("BlanketOrderReference").Descendants("DocumentID").Value.ToString
            Catch ex As Exception
                BlanketOrderReference = ""
            End Try

            Try
                CustomerReference_DocumentID = oxml.Descendants("CustomerReference").Descendants("DocumentID").Value.ToString
            Catch ex As Exception
                CustomerReference_DocumentID = ""
            End Try

            Try
                BuyerParty_PartyID = oxml.Descendants("BuyerParty").Descendants("PartyID").Value.ToString
            Catch ex As Exception
                BuyerParty_PartyID = ""
            End Try

            Try
                BuyerParty_AgencyCode = oxml.Descendants("BuyerParty").Descendants("AgencyCode").Value.ToString
            Catch ex As Exception
                BuyerParty_AgencyCode = ""
            End Try

            Try
                OrderingParty_PartyID = oxml.Descendants("OrderingParty").Descendants("PartyID").Value.ToString
            Catch ex As Exception
                OrderingParty_PartyID = ""
            End Try

            Try
                OrderingParty_AgencyCode = oxml.Descendants("OrderingParty").Descendants("AgencyCode").Value.ToString
            Catch ex As Exception
                OrderingParty_AgencyCode = ""
            End Try

            Try
                Consignee_PartyID = oxml.Descendants("Consignee").Descendants("PartyID").Value.ToString
            Catch ex As Exception
                Consignee_PartyID = ""
            End Try

            Try
                Consignee_AgencyCode = oxml.Descendants("Consignee").Descendants("AgencyCode").Value.ToString
            Catch ex As Exception
                Consignee_AgencyCode = ""
            End Try

            Try
                PaymentTerms_PaymentMethod = oxml.Descendants("PaymentTerms").Descendants("PaymentMethod").Value.ToString
            Catch ex As Exception
                PaymentTerms_PaymentMethod = ""
            End Try

            oLog.escribeMensaje("Conectando a SAP...", EXO_Log.EXO_Log.Tipo.informacion)
            MyBase.ConnectSAP()
            oLog.escribeMensaje("Conectado con SAP...", EXO_Log.EXO_Log.Tipo.informacion)
            blEXO = New BLEXO
            sAnno = Right(Year(Now).ToString, 2)
            oORDR = CType(Me.Company.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oOrders), SAPbobsCOM.Documents)
            strNumSerie = GetValueDB("""NNM1""", """Series""", """ObjectCode"" = 17 And ""SeriesName""='D1" & sAnno & "'")
            oLog.escribeMensaje("Nº Serie:  " & strNumSerie & " .", EXO_Log.EXO_Log.Tipo.advertencia)
            If strNumSerie = "" Then
                oLog.escribeMensaje("La serie " & strNumSerie & " no existe en la base de datos", EXO_Log.EXO_Log.Tipo.error)
                Throw New Exception("La serie " & strNumSerie & " no existe en la base de datos")
            End If
            strCliente = BuyerParty_PartyID
            strComprobar = GetValueDB("""OCRD""", """CardCode""", """U_EXO_BUYERPARTY"" ='" & strCliente & "'")
            If strComprobar = "" Then
                'No existe el cliente
                oLog.escribeMensaje("El cliente " & strCliente & " no existe en la base de datos", EXO_Log.EXO_Log.Tipo.error)
                Throw New Exception("El cliente " & strCliente & " no existe en la base de datos")
            Else
                strCliente = strComprobar
            End If

            'Hacer consulta getvaluedb para saber si existe el pedido ORDR
            strRef = CustomerReference_DocumentID
            strExiste = GetValueDB("""ORDR""", """DocEntry""", """CardCode""='" & strCliente & "' and ""NumAtCard""='" & strRef & "'")
            If strExiste <> "" Then
                'MODIFICAR
                'Antes de modificar, comprobar el estado, y si es diferente de C, dejo continuar, sino mensaje y me salgo
                strEstado = GetValueDB("""ORDR""", """DocStatus""", """DocEntry""=" & CInt(strExiste) & "")
                If strEstado <> "C" Then
                    oORDR.GetByKey(CInt(strExiste))
                    'recorrer lineas y borrar y las hago de nuevo
                    For i = 0 To oORDR.Lines.Count - 1
                        oORDR.Lines.SetCurrentLine(0)
                        oORDR.Lines.Delete()
                    Next
                Else
                    oLog.escribeMensaje("El estado del documento: " & strRef & " es cerrado o cancelado, no se puede modificar", EXO_Log.EXO_Log.Tipo.error)
                    Throw New Exception("El estado del documento: " & strRef & " es cerrado o cancelado, no se puede modificar")
                End If
            Else
                'Nuevo datos cabecera
                oORDR.Series = CInt(strNumSerie)
                oORDR.CardCode = strCliente
                oORDR.UserFields.Fields.Item("U_EXO_EDIE").Value = "Y"
            End If

            'Datos cabecera para crear o modificar
            'Formato para fechas
            strFecha = Year(Now).ToString & "-" & Month(Now).ToString & "-" & Day(Now)

            oORDR.TaxDate = CDate(strFecha)
            oORDR.DocDate = CDate(strFecha)
            oORDR.DocDueDate = CDate(strFecha)
            'Fecha de envio
            If EarliestDeliveryDate <> "" Then

            End If
            strFechaEnvio = EarliestDeliveryDate
            If strFechaEnvio <> "" Then
                oORDR.DocDueDate = CDate(strFechaEnvio)
            End If


            'Referencia del Cliente
            oORDR.NumAtCard = strRef

            'Forma de pago
            strFP = PaymentTerms_PaymentMethod
            If strFP <> "" Then
                strComprobar = GetValueDB("""OPYM""", """PayMethCod""", """PayMethCod"" ='" & strFP & "'")
                If strComprobar = "" Then
                    ''no existe la forma de pago, excepcion y me salgo
                    oLog.escribeMensaje("La forma de pago " & strFP & " no existe en la base de datos", EXO_Log.EXO_Log.Tipo.error)
                    Throw New Exception("La forma de pago " & strFP & " no existe en la base de datos")
                Else
                    oORDR.PaymentMethod = strFP
                End If
            End If

            'Comentarios
            oORDR.Comments = "RODINEX WS" '" Pedido RODINEX creado a través de WEB SERVICE. Agency Code " & BuyerParty_AgencyCode

            If OrderingParty_PartyID <> "" Then
                aryTextFile = OrderingParty_PartyID.ToString.Split(CChar("|"))

                For i = 0 To UBound(aryTextFile)
                    Select Case i
                        Case 1
                            strNomDir = aryTextFile(i)
                        Case 2
                            strDireccion = aryTextFile(i)
                        Case 3
                            strCP = aryTextFile(i)
                        Case 4
                            strProvincia = aryTextFile(i)
                        Case 5
                            strPais = aryTextFile(i)
                        Case 6
                            strTlfno = aryTextFile(i)
                    End Select
                Next i
            End If

            'Direccion
            strDestino = GetValueDB("""CRD1""", """Address""", " ""CardCode""='" & strCliente & "' and ""U_EXO_CODGS"" ='" & Consignee_PartyID & "'")
            If strDestino = "" Then
                'No existe la dirección 
                oLog.escribeMensaje("La dirección del cliente " & strCliente & " con Código " & Consignee_PartyID & " no existe en la base de datos", EXO_Log.EXO_Log.Tipo.error)
                Throw New Exception("La dirección del cliente " & strCliente & " con Código " & Consignee_PartyID & " no existe en la base de datos")
            End If

            oORDR.ShipToCode = strDestino

            sSQL = "SELECT T0.U_EXO_BonusApl, T0.U_EXO_Desligar, " &
                   "ISNULL((SELECT ISNULL(TDir.U_SEI_EMPLE, '0') " &
                   "FROM CRD1 TDir WITH (NOLOCK) WHERE TDir.AdresType = 'S' AND TDir.CardCode = T0.CardCode AND TDir.Address = '" & strDestino & "' ), 0) AS 'Emp', " &
                   "isnull(T0.SlpCode, 0) AS 'DefecSlp' " &
                   "FROM OCRD T0 WITH (NOLOCK) " &
                   "WHERE T0.CardCode = '" & strCliente & "'"

            dtEmpleado = New System.Data.DataTable()
            FillDtDB(dtEmpleado, sSQL)

            If dtEmpleado.Rows.Count > 0 Then
                'Empleado del depto. de ventas. Si no esta el de la dir, cojo el de cabecera
                sEmpVentas = dtEmpleado.Rows.Item(0).Item("Emp").ToString
                If sEmpVentas = "0" Then
                    sEmpVentas = dtEmpleado.Rows.Item(0).Item("DefecSlp").ToString
                End If

                If sEmpVentas <> "0" Then
                    sSQL = "SELECT isnull(T1.U_CCusto, '') as CeCo FROM OSLP T1 WITH (NOLOCK) WHERE T1.SlpCode = " & sEmpVentas
                    dtCeCo = New System.Data.DataTable()
                    FillDtDB(dtCeCo, sSQL)

                    If dtCeCo.Rows.Count > 0 Then
                        sCeCo = dtCeCo.Rows.Item(0).Item("CeCo").ToString
                    End If
                End If
            End If

            sSQL = "SELECT TOP 1 T0.ExpnsCode, T0.ExpnsName FROM OEXD T0 WITH (NOLOCK) WHERE U_EXO_EcoTasa = 'Y'"
            dtTipoEcotasa = New System.Data.DataTable()
            FillDtDB(dtTipoEcotasa, sSQL)
            If dtTipoEcotasa.Rows.Count > 0 Then
                sTipoEcotasa = dtTipoEcotasa.Rows.Item(0).Item("ExpnsCode").ToString
                oLog.escribeMensaje("Tipo Ecotasa " & sTipoEcotasa, EXO_Log.EXO_Log.Tipo.informacion)
            End If
            sPais = GetValueDB("CRD1 WITH (NOLOCK)", "Country", "CardCode='" & strCliente & "' and Address='" & strDestino & "' and AdresType='S'")
            dtEcotasa = New System.Data.DataTable()

            'recorro la parte de articulos
            For Each report As XElement In oxml.Descendants("OrderLine")
                Dim sPrecio As String = "0"
                If intLin <> 0 Then
                    oORDR.Lines.Add()
                End If
                'Comprobar que exista el articulo
                If report.Element("OrderedArticle").Element("ArticleIdentification").Element("EANUCCArticleID").Value.ToString <> "" Then
                    strCodArt = GetValueDB("""OITM""", """ItemCode""", """U_SEI_JANCODE"" ='" & report.Element("OrderedArticle").Element("ArticleIdentification").Element("EANUCCArticleID").Value.ToString & "'")
                    If strCodArt = "" Then
                        'No existe el articulo, excepcion y me salgo
                        oLog.escribeMensaje("El artículo con el EAN " & strCodArt & " no existe en la base de datos", EXO_Log.EXO_Log.Tipo.error)
                        Throw New Exception("El artículo con el EAN " & strCodArt & " no existe en la base de datos")
                    End If
                Else
                    strCodArt = GetValueDB("""OITM""", """ItemCode""", """ItemCode"" ='" & report.Element("OrderedArticle").Element("ArticleIdentification").Element("ManufacturersArticleID").Value.ToString & "'")
                    If strCodArt = "" Then
                        'No existe el articulo, excepcion y me salgo
                        oLog.escribeMensaje("El artículo " & strCodArt & " no existe en la base de datos", EXO_Log.EXO_Log.Tipo.error)
                        Throw New Exception("El artículo " & strCodArt & " no existe en la base de datos")
                    End If
                End If

                If strCodArt <> "" Then
                    Dim sEstado As String = "" : Dim dtStock As System.Data.DataTable = Nothing
                    Dim sLPrecio As String = GetValueDB("""@EXO_OGEN1""", """U_EXO_INFV""", """U_EXO_NOMV"" ='TarifaRODINEX'")
                    oLog.escribeMensaje("La lista de precio " & sLPrecio & ".", EXO_Log.EXO_Log.Tipo.advertencia)
                    sEstado = GetValueDB("""OITM""", """U_EXO_Estado""", """ItemCode"" ='" & strCodArt & "'")
                    'sPrecio = GetValueDB("""ITM1""", """Price""", """ItemCode"" ='" & strCodArt & "' and ""PriceList""=" & sLPrecio)
                    sSQL = "Select dbo.[EXOPrecioClientePed]('" + strCliente + "' , '" + strCodArt + "', " + strNumSerie + ") as ""Precio"" "
                    Dim dtprecio = New System.Data.DataTable("Precio")
                    blEXO.FillDtDB(dtprecio, sSQL)
                    'oLog.escribeMensaje(" - EXOPrecioClientePed. " & sSQL, EXO_Log.EXO_Log.Tipo.advertencia)
                    If dtprecio.Rows.Count > 0 Then
                        sPrecio = (dtprecio.Rows(0).Item("Precio").ToString)
                        'oLog.escribeMensaje("entra Precio:  " & sPrecio & " - EXOPrecioClientePed. " & sSQL, EXO_Log.EXO_Log.Tipo.advertencia)
                    Else
                        sPrecio = GetValueDB("""ITM1""", """Price""", """ItemCode"" ='" & strCodArt & "' and ""PriceList""=" & sLPrecio)
                        'oLog.escribeMensaje("etnra Precio:  " & sPrecio & " - ITM1.", EXO_Log.EXO_Log.Tipo.advertencia)
                    End If
                    sSQL = "SELECT ""ItemCode"" ,""ItemName"", ""U_SEI_JANCODE"",  " _
                    & " Case  " _
                    & " When sum(OnHand) - sum(IsCommited)< 0 Then '0' " _
                    & " WHEN sum(OnHand) - sum(IsCommited)> 50 THEN '>50' " _
                    & " Else  cast(cast(sum(OnHand) - sum(IsCommited) as int) as varchar)   " _
                    & " End   ""AVAILABILITY"",  " _
                    & " Case   " _
                    & " WHEN sum(OnHand) - sum(IsCommited)< 0 THEN '0' " _
                    & " WHEN sum(OnHand) - sum(IsCommited)> 50 THEN '>50' " _
                    & " Else cast(cast(sum(OnHand) - sum(IsCommited) as int) as varchar)  " _
                    & " End   ""QUANTITY VALUE""  " _
                    & " from (  " _
                    & " Select t1.ItemCode, t1.ItemName,t1.U_SEI_JANCODE, t3.OnHand , t3.IsCommited  " _
                    & " from oitm t1  " _
                    & " inner join [Yokohama_prod].dbo.[OITW] t3 With (NOLOCK) On t1.ItemCode = t3.ItemCode and t3.WhsCode='01' " _
                    & " WHERE ItmsGrpCod='108'  AND ISNUMERIC(U_SEI_JANCODE)=1 and (t1.U_SEI_CATEGORY1='Yokohama-Tires' or t1.U_SEI_CATEGORY1='Alliance-Tires') AND t1.U_SEICATEGORY2 in ('PCR/VAN','TBS' )  " _
                    & " And (ISNULL(t1.U_EXO_Estado, '') = 'A' OR (ISNULL(t1.U_EXO_Estado, '') = 'D'  " _
                    & " And ([Yokohama_prod].dbo.EXOStockB2BES(t1.ItemCode) + [Yokohama_prod].dbo.EXOStockB2BPT(t1.ItemCode)) > 0))  " _
                    & " union all " _
                    & " select t1.ItemCode,t1.ItemName, t1.U_SEI_JANCODE, t4.OnHand ,t4.IsCommited  " _
                    & " from oitm t1   " _
                    & " inner join [Yokohama_PT].dbo.[OITW] t4 WITH (NOLOCK) ON t1.ItemCode = t4.ItemCode and t4.WhsCode='PT01' " _
                    & " WHERE ItmsGrpCod='108'  AND ISNUMERIC(U_SEI_JANCODE)=1 and (t1.U_SEI_CATEGORY1='Yokohama-Tires' or t1.U_SEI_CATEGORY1='Alliance-Tires') AND t1.U_SEICATEGORY2 in ('PCR/VAN','TBS' )  " _
                    & " And (ISNULL(t1.U_EXO_Estado, '') = 'A' OR (ISNULL(t1.U_EXO_Estado, '') = 'D'  " _
                    & " And ([Yokohama_prod].dbo.EXOStockB2BES(t1.ItemCode) + [Yokohama_prod].dbo.EXOStockB2BPT(t1.ItemCode)) > 0)) " _
                    & " ) as DatosCompletos " _
                    & " WHERE ""ItemCode""='" & strCodArt & "' " _
                    & " group by ItemCode,ItemName,U_SEI_JANCODE " _
                    & " order by ItemCode"
                    dtStock = New System.Data.DataTable("Articulos")
                    blEXO.FillDtDB(dtStock, sSQL)
                    Dim iDisponible As Integer = 0
                    If dtStock.Rows.Count > 0 Then
                        If dtStock.Rows(0).Item("AVAILABILITY").ToString <> ">50" Then
                            iDisponible = CInt(dtStock.Rows(0).Item("AVAILABILITY").ToString)
                        Else
                            iDisponible = 50
                        End If
                    End If
                    If sEstado = "B" Then
                        'Estado B2B
                        oLog.escribeMensaje("El artículo " & strCodArt & " tiene como estado -B2B-. No se puede usar.", EXO_Log.EXO_Log.Tipo.error)
                        Throw New Exception("El artículo " & strCodArt & " tiene como estado -B2B-. No se puede usar.")
                    ElseIf sEstado = "D" Then
                        If iDisponible = 0 Then
                            'Estado Discontinuo y Stock=0
                            oLog.escribeMensaje("El artículo " & strCodArt & " tiene como estado -Discontinuo- y su stock es 0.", EXO_Log.EXO_Log.Tipo.error)
                            sSQL = "SELECT ItemCode as [Artigo], Dscription as [Descrição], U_Qtd as [Qtd], U_Barco as [Barco], U_DtPrevIn as [Data prevista de entrega], U_ShipingN as [Shipping Nº] FROM ("
                            sSQL &= " SELECT     Yokohama_PT.dbo.PQT1.ItemCode, Yokohama_PT.dbo.PQT1.Dscription, lin.U_Qtd, lin.U_Barco, lin.U_DtPrevIn, lin.U_ShipingN "
                            sSQL &= " FROM         Yokohama_PT.dbo.OPQT INNER JOIN"
                            sSQL &= " Yokohama_PT.dbo.PQT1 ON Yokohama_PT.dbo.OPQT.DocEntry = Yokohama_PT.dbo.PQT1.DocEntry INNER JOIN "
                            sSQL &= " Yokohama_PT.dbo.[@EBR_ODETMOV] ON Yokohama_PT.dbo.PQT1.DocEntry = Yokohama_PT.dbo.[@EBR_ODETMOV].U_DocEntry AND Yokohama_PT.dbo.PQT1.LineNum = Yokohama_PT.dbo.[@EBR_ODETMOV].U_LineNum INNER JOIN "
                            sSQL &= " Yokohama_PT.dbo.[@EBR_DETMOV1] AS lin ON Yokohama_PT.dbo.[@EBR_ODETMOV].DocEntry = lin.DocEntry "
                            sSQL &= " WHERE     (lin.U_InvoiceN IS NOT NULL) AND (lin.U_DtPrevIn IS NOT NULL) AND (lin.U_Destino = 3) AND (NOT EXISTS "
                            sSQL &= " (SELECT     U_Qtd FROM          Yokohama_PT.dbo.[@EBR_DETMOV1] "
                            sSQL &= " WHERE      (U_Origem = 3) AND (DocEntry = lin.DocEntry) AND (U_ShipingN = lin.U_ShipingN) AND (U_InvoiceN = lin.U_InvoiceN) AND  "
                            sSQL &= "  (U_DDocEntry = lin.U_DDocEntry) AND (U_DLineNum = lin.U_DLineNum))) "
                            sSQL &= " UNION ALL "
                            sSQL &= " SELECT     Yokohama_PT.dbo.PQT1.ItemCode, Yokohama_PT.dbo.PQT1.Dscription, lin.U_Qtd, lin.U_Barco, lin.U_DtPrevIn, lin.U_ShipingN "
                            sSQL &= " FROM         Yokohama_PT.dbo.OPQT INNER JOIN "
                            sSQL &= " Yokohama_PT.dbo.PQT1 ON Yokohama_PT.dbo.OPQT.DocEntry = Yokohama_PT.dbo.PQT1.DocEntry INNER JOIN "
                            sSQL &= " Yokohama_PT.dbo.[@EBR_ODETMOV] ON Yokohama_PT.dbo.PQT1.DocEntry = Yokohama_PT.dbo.[@EBR_ODETMOV].U_DocEntry AND Yokohama_PT.dbo.PQT1.LineNum = Yokohama_PT.dbo.[@EBR_ODETMOV].U_LineNum INNER JOIN "
                            sSQL &= " Yokohama_PT.dbo.[@EBR_DETMOV1] AS lin ON Yokohama_PT.dbo.[@EBR_ODETMOV].DocEntry = lin.DocEntry "
                            sSQL &= " WHERE     (lin.U_InvoiceN IS NULL) AND (lin.U_DtPrevIn IS NOT NULL) AND (lin.U_Destino = 3) AND (NOT EXISTS "
                            sSQL &= " (SELECT     U_Qtd FROM          Yokohama_PT.dbo.[@EBR_DETMOV1] "
                            sSQL &= "  WHERE      (U_Origem = 3) AND (DocEntry = lin.DocEntry) AND (U_ShipingN = lin.U_ShipingN))) ) B"
                            sSQL &= " WHERE B.ItemCode = N'" & strCodArt & "' ORDER BY U_DtPrevIn"
                            dtStock = New System.Data.DataTable("Prevision")
                            blEXO.FillDtDB(dtStock, sSQL)

                            If dtStock.Rows.Count = 0 Then
                                Throw New Exception("El artículo " & strCodArt & " tiene como estado -Discontinuo- y su stock es 0 y la previsión es 0.")
                            Else
                                Dim iPrevisión As Integer = 0
                                iPrevisión = CInt(dtStock.Rows(0).Item("Qtd").ToString)
                                If iPrevisión > 0 Then
                                    oLog.escribeMensaje("El artículo " & strCodArt & " tiene como previsión " & iPrevisión.ToString & ". Se permite generar el pedido.", EXO_Log.EXO_Log.Tipo.advertencia)
                                Else
                                    Throw New Exception("El artículo " & strCodArt & " tiene como estado -Discontinuo- y su stock es 0 y la previsión es 0.")
                                End If
                            End If
                        Else
                            If CDbl(sPrecio) = 0 Then
                                'Estado Discontinuo y Stock<>0 y precio =0
                                oLog.escribeMensaje("El artículo " & strCodArt & " tiene como estado -Discontinuo- y su stock es distinto de 0 pero su precio es 0.", EXO_Log.EXO_Log.Tipo.error)
                                Throw New Exception("El artículo " & strCodArt & " tiene como estado -Discontinuo- y su stock es distinto de 0 pero su precio es 0.")
                            End If
                        End If
                    End If
                End If
                'ecostasa para portes
                cPrecEcotasa = 0
                sEcotasa = ""

                'Como sacar tipo ecotasa España (Portugal lo mismo pero con su campo)
                If sPais = "ES" Then
                    dtEcotasa.Clear()

                    sSQL = "SELECT isnull(T0.U_TXECOVATORSP, '') AS 'TipoEcotasa', isnull(T0.U_SEI_TXECOVATORSPE, 0) AS 'ValorEcotasa' " &
                           "FROM OITM T0 WITH (NOLOCK) WHERE T0.ItemCode = '" & strCodArt & "'"

                    FillDtDB(dtEcotasa, sSQL)
                    If dtEcotasa.Rows.Count > 0 Then
                        sEcotasa = dtEcotasa.Rows(0).Item("TipoEcotasa").ToString
                        cPrecEcotasa = CDbl(dtEcotasa.Rows(0).Item("ValorEcotasa").ToString.Replace(".", ","))
                        oLog.escribeMensaje("Ecotasa:" & sEcotasa & " y PrecioEcoTasa:" & CStr(cPrecEcotasa), EXO_Log.EXO_Log.Tipo.informacion)
                    End If
                ElseIf sPais = "PT" Then
                    dtEcotasa.Clear()
                    sSQL = "SELECT isnull(T0.U_SEI_TXECOTPT, '') AS 'TipoEcotasa', isnull(T0.U_SEI_TX_ECOPTE, 0) AS 'ValorEcotasa' " &
                           "FROM OITM T0 WITH (NOLOCK) WHERE T0.ItemCode = '" & strCodArt & "'"
                    FillDtDB(dtEcotasa, sSQL)
                    If dtEcotasa.Rows.Count > 0 Then
                        sEcotasa = dtEcotasa.Rows(0).Item("TipoEcotasa").ToString
                        cPrecEcotasa = CDbl(dtEcotasa.Rows(0).Item("ValorEcotasa").ToString.Replace(".", ","))
                        oLog.escribeMensaje("Ecotasa:" & sEcotasa & " y PrecioEcoTasa:" & CStr(cPrecEcotasa), EXO_Log.EXO_Log.Tipo.informacion)
                    End If
                End If

                oORDR.Lines.ItemCode = strCodArt
                oORDR.Lines.Quantity = CDbl(report.Element("OrderedArticle").Element("RequestedQuantity").Element("QuantityValue").Value.ToString.Replace(".", ","))
                oORDR.Lines.UnitPrice = CDbl(sPrecio)
                If sEcotasa <> "" Then
                    oORDR.Lines.UserFields.Fields.Item("U_EXO_EcoCod").Value = sEcotasa
                    oORDR.Lines.UserFields.Fields.Item("U_EXO_EcoPric").Value = cPrecEcotasa
                    oORDR.Lines.UserFields.Fields.Item("U_EXO_EcoImp").Value = CDbl(report.Element("OrderedArticle").Element("RequestedQuantity").Element("QuantityValue").Value.ToString.Replace(".", ",")) * cPrecEcotasa
                    cTotalesEcotasa += CDbl(report.Element("OrderedArticle").Element("RequestedQuantity").Element("QuantityValue").Value.ToString.Replace(".", ",")) * cPrecEcotasa
                End If

                oORDR.Lines.SalesPersonCode = CInt(sEmpVentas)
                If sCeCo <> "" Then
                    oORDR.Lines.CostingCode = sCeCo
                    oORDR.Lines.COGSCostingCode = sCeCo
                End If

                intLin = intLin + 1
            Next

            If cTotalesEcotasa <> 0 Then
                oORDR.Expenses.SetCurrentLine(0)
                oORDR.Expenses.ExpenseCode = CInt(sTipoEcotasa)
                oORDR.Expenses.LineTotal = cTotalesEcotasa
                oLog.escribeMensaje("TotalEcoTasa:" & CStr(cTotalesEcotasa), EXO_Log.EXO_Log.Tipo.informacion)
            End If

            If strExiste <> "" Then
                If oORDR.Update() <> 0 Then
                    oLog.escribeMensaje(Me.Company.GetLastErrorCode & " / " & Me.Company.GetLastErrorDescription, EXO_Log.EXO_Log.Tipo.error)
                    Throw New Exception(Me.Company.GetLastErrorCode & " / " & Me.Company.GetLastErrorDescription)
                Else
                    oLog.escribeMensaje("Pedido Actualizado", EXO_Log.EXO_Log.Tipo.advertencia)
                End If
            Else
                If oORDR.Add() <> 0 Then
                    oLog.escribeMensaje(Me.Company.GetLastErrorCode & " / " & Me.Company.GetLastErrorDescription, EXO_Log.EXO_Log.Tipo.error)
                    Throw New Exception(Me.Company.GetLastErrorCode & " / " & Me.Company.GetLastErrorDescription)
                Else
                    oLog.escribeMensaje("Pedido Creado", EXO_Log.EXO_Log.Tipo.advertencia)
                End If
            End If
            Me.Company.GetNewObjectCode(sDocEntry)
            sNumDoc = GetValueDB("""ORDR""", """DocNum""", """DocEntry"" = " & CInt(sDocEntry) & "")

            AddUpdateORDR = CStr(sNumDoc)

            'Creación de la alerta.
            EnviarAlerta(sDocEntry, sNumDoc, strExiste)
            oLog.escribeMensaje("Pedido Nº" & sNumDoc & " Y pedido Interno nº" & sDocEntry, EXO_Log.EXO_Log.Tipo.advertencia)

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            MyBase.DisconnectSAP()
            oLog.escribeMensaje("Desconexión con SAP", EXO_Log.EXO_Log.Tipo.informacion)
            If oORDR IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oORDR)
        End Try

    End Function

    'crear alerta sap
    Public Sub EnviarAlerta(ByVal strDocEntry As String, ByVal strPedido As String, ByVal strExiste As String)
        Dim oCompany As SAPbobsCOM.Company = Nothing
        Dim pMessageDataColumns As SAPbobsCOM.MessageDataColumns = Nothing
        Dim pMessageDataColumn As SAPbobsCOM.MessageDataColumn = Nothing
        Dim oLines As SAPbobsCOM.MessageDataLines = Nothing
        Dim oLine As SAPbobsCOM.MessageDataLine = Nothing
        Dim oRecipientCollection As SAPbobsCOM.RecipientCollection = Nothing

        Dim oMessageService As SAPbobsCOM.MessagesService = Nothing
        Dim oMessage As SAPbobsCOM.Message = Nothing

        Dim oDBSAP As SqlConnection = Nothing
        Dim sSQL As String = ""
        Dim oDtSAP As System.Data.DataTable = Nothing

        Dim blEXO As BLEXO = Nothing


        Try
            MyBase.ConnectSAP()
            blEXO = New BLEXO

            sSQL = "Select t1.USER_CODE " &
                       "FROM OUSR t1 With (NOLOCK) " &
                       "WHERE ISNULL(t1.U_EXO_RESWS, 'N') = 'Y'"
            ' Como no tengo el campo y tengo uno parecido he probado con él
            'sSQL = "Select t1.USER_CODE " &
            '           "FROM OUSR t1 With (NOLOCK) " &
            '           "WHERE ISNULL(t1.U_EXO_ENVJOB, 'N') = 'Y'"
            oDtSAP = New System.Data.DataTable
            FillDtDB(oDtSAP, sSQL)
            If oDtSAP.Rows.Count > 0 Then 'Si hay usuarios con esta alerta activada, enviamos alertas
                oMessageService = CType(Me.Company.GetCompanyService.GetBusinessService(SAPbobsCOM.ServiceTypes.MessagesService), SAPbobsCOM.MessagesService)
                oMessage = CType(oMessageService.GetDataInterface(SAPbobsCOM.MessagesServiceDataInterfaces.msdiMessage), SAPbobsCOM.Message)
                If strExiste <> "" Then
                    oMessage.Subject = "Pedido RODINEX Número " & strPedido & " modificado por Web Serive correctamente"
                    oMessage.Text = "Se ha modificado el Pedido RODINEX Número " & strPedido & " correctamente "
                Else
                    oMessage.Subject = "Pedido RODINEX Número " & strPedido & " creado por Web Serive correctamente"
                    oMessage.Text = "Se ha creado el Pedido RODINEX Número " & strPedido & " correctamente "
                End If
                oRecipientCollection = oMessage.RecipientCollection

                For j As Integer = 0 To oDtSAP.Rows.Count - 1
                    oRecipientCollection.Add()
                    oRecipientCollection.Item(j).SendInternal = SAPbobsCOM.BoYesNoEnum.tYES
                    oRecipientCollection.Item(j).UserCode = oDtSAP.Rows.Item(j).Item("USER_CODE").ToString
                Next

                pMessageDataColumns = oMessage.MessageDataColumns

                pMessageDataColumn = pMessageDataColumns.Add()
                pMessageDataColumn.ColumnName = "Número interno"
                pMessageDataColumn.Link = SAPbobsCOM.BoYesNoEnum.tYES
                oLines = pMessageDataColumn.MessageDataLines
                oLine = oLines.Add()
                oLine.Value = strDocEntry
                oLine.Object = "17"
                oLine.ObjectKey = strDocEntry

                pMessageDataColumn = pMessageDataColumns.Add()
                pMessageDataColumn.ColumnName = "Número Pedido"
                oLines = pMessageDataColumn.MessageDataLines
                oLine = oLines.Add()
                oLine.Value = strPedido

                oMessageService.SendMessage(oMessage)

            End If
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally

            If oDtSAP IsNot Nothing Then oDtSAP.Dispose()
            If pMessageDataColumns IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pMessageDataColumns)
            If pMessageDataColumn IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(pMessageDataColumn)
            If oLines IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oLines)
            If oLine IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oLine)
            If oRecipientCollection IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oRecipientCollection)
            'If oCmpSrv IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oCmpSrv)
            If oMessageService IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oMessageService)
            If oMessage IsNot Nothing Then System.Runtime.InteropServices.Marshal.FinalReleaseComObject(oMessage)

            MyBase.DisconnectSAP()
        End Try
    End Sub
#End Region

#Region "CrearXML"
    Private Sub createNode(ByVal writer As XmlTextWriter)
        writer.WriteStartElement("DocumentID")
        writer.WriteString("A2")
        writer.WriteEndElement()
        writer.WriteStartElement("Variant")
        writer.WriteString("5")
        writer.WriteEndElement()
        writer.WriteStartElement("ErrorHead")
        writer.WriteStartElement("ErrorCode")
        writer.WriteString("0")
        writer.WriteEndElement()
        writer.WriteEndElement()

        'writer.WriteStartElement("Product_id")
        'writer.WriteString(pID)
        'writer.WriteEndElement()
        'writer.WriteStartElement("Product_name")
        'writer.WriteString(pName)
        'writer.WriteEndElement()
        'writer.WriteStartElement("Product_price")
        'writer.WriteString(pPrice)
        'writer.WriteEndElement()
        'writer.WriteEndElement()
    End Sub
#End Region

End Class


