' NOTA: puede usar el comando "Cambiar nombre" del menú contextual para cambiar el nombre de clase "Service1" en el código, en svc y en el archivo de configuración.
' NOTA: para iniciar el Cliente de prueba WCF para probar este servicio, seleccione Service1.svc o Service1.svc.vb en el Explorador de soluciones e inicie la depuración.
Imports System.Web.Script.Serialization
Imports CN
Imports System.IO
Imports System.Xml
Imports System.Xml.Serialization
Imports System.Collections.ObjectModel
Imports System.Reflection
Public Class YokoRODINEXService

    Implements IYokoRODINEXService
    Public Sub New()
    End Sub
    Function CrearPedido(ByVal dxml As XmlElement) As XmlElement Implements IYokoRODINEXService.CrearPedido

        Dim clsErrores As New Errores()
        Dim sError As String = ""
        Dim pedido As New PEDIDOSAP
        Dim blEXO As BLEXO = Nothing

        Dim oXml As XDocument

        Dim xmlSerializador As XmlSerializer = Nothing
        Dim ns As XmlSerializerNamespaces = New XmlSerializerNamespaces()
        Dim sw As StringWriter = New StringWriter

        Dim clsPedidoCreado As New PEDIDOSAP()
        Dim sPedido As String = ""

        Dim clsResponseOrder As New Order_response_A2
        Dim oLog As EXO_Log.EXO_Log = Nothing

        Try
            oLog = New EXO_Log.EXO_Log("C:\inetpub\logs\ExpertOne\logWS_EXO_WS_YK_RODINEX_", 50, EXO_Log.EXO_Log.Nivel.todos, 4, "", EXO_Log.EXO_Log.GestionFichero.dia)

            blEXO = New BLEXO
            oLog.escribeMensaje("Llamada a ""Crear Pedido""...", EXO_Log.EXO_Log.Tipo.advertencia)
            oLog.escribeMensaje("Convirtiendo a string Información...", EXO_Log.EXO_Log.Tipo.informacion)
            'oXml = XDocument.Parse(sxml)
            oXml = XDocument.Parse(dxml.OuterXml)

            oLog.escribeMensaje("Tratando Información...", EXO_Log.EXO_Log.Tipo.informacion)

            sPedido = blEXO.AddUpdateORDR(oXml, oLog)


        Catch exCOM As System.Runtime.InteropServices.COMException
            If exCOM.InnerException IsNot Nothing AndAlso exCOM.InnerException.Message <> "" Then
                sError = exCOM.InnerException.Message
            Else
                sError = exCOM.Message
            End If
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            If ex.InnerException IsNot Nothing AndAlso ex.InnerException.Message <> "" Then
                sError = ex.InnerException.Message
            Else
                sError = ex.Message
            End If
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Finally
            ns = New XmlSerializerNamespaces()
            ns.Add("ew", "http://www.reifen.net")

            sw = New StringWriter
            xmlSerializador = New XmlSerializer(GetType(Order_response_A2))

            clsResponseOrder = CargarDatos(oXml, sError, sPedido)

            xmlSerializador.Serialize(sw, clsResponseOrder, ns)
            Dim strRespuesta As String

            strRespuesta = sw.ToString
            strRespuesta = Replace(strRespuesta, "<Articulos>", "", , , CompareMethod.Text)
            strRespuesta = Replace(strRespuesta, "</Articulos>", "", , , CompareMethod.Text)
            strRespuesta = Replace(strRespuesta, "Error1", "Error", , , CompareMethod.Text)
            strRespuesta = Replace(strRespuesta, "utf-16", "utf-8")
            'CrearPedido = strRespuesta

            oLog.escribeMensaje("Conviertiendo Información...", EXO_Log.EXO_Log.Tipo.informacion)
            Dim dResp As New XmlDocument
            dResp.LoadXml(strRespuesta)
            oLog.escribeMensaje("Enviando Información...", EXO_Log.EXO_Log.Tipo.informacion)
            CrearPedido = dResp.DocumentElement
            oLog.escribeMensaje("Enviada Información...", EXO_Log.EXO_Log.Tipo.informacion)

            If oXml IsNot Nothing Then
                oXml = Nothing
            End If

            If sw IsNot Nothing Then
                sw.Close()
                sw.Dispose()
                sw = Nothing
            End If

            'xmlSerializador = Nothing
            clsErrores = Nothing
            ns = Nothing
            oLog.escribeMensaje("Fin Llamada a ""Crear Pedido""...", EXO_Log.EXO_Log.Tipo.advertencia)
        End Try
    End Function
    Function ConsultaStock(ByVal dxml As XmlElement) As XmlElement Implements IYokoRODINEXService.ConsultaStock

        Dim clsErrores As New Errores()
        Dim sError As String = ""
        Dim blEXO As BLEXO = Nothing

        Dim oXml As XDocument

        Dim xmlSerializador As XmlSerializer = Nothing
        Dim ns As XmlSerializerNamespaces = New XmlSerializerNamespaces()
        Dim sw As StringWriter = New StringWriter

        Dim clsResponseOrder As New Stock_response_A2

        Dim oLog As EXO_Log.EXO_Log = Nothing

        Try
            oLog = New EXO_Log.EXO_Log("C:\inetpub\logs\ExpertOne\logWS_EXO_WS_YK_RODINEX_", 50, EXO_Log.EXO_Log.Nivel.todos, 4, "", EXO_Log.EXO_Log.GestionFichero.dia)

            blEXO = New BLEXO
            oLog.escribeMensaje("Llamada a ""Consulta Stock""...", EXO_Log.EXO_Log.Tipo.advertencia)
            'Dim sxml As String = dxml.InnerXml
            oLog.escribeMensaje("Convirtiendo a string Información...", EXO_Log.EXO_Log.Tipo.informacion)
            'oXml = XDocument.Parse(sxml)
            oXml = XDocument.Parse(dxml.OuterXml)

            oLog.escribeMensaje("Tratando Información...", EXO_Log.EXO_Log.Tipo.informacion)
        Catch exCOM As System.Runtime.InteropServices.COMException
            If exCOM.InnerException IsNot Nothing AndAlso exCOM.InnerException.Message <> "" Then
                sError = exCOM.InnerException.Message
            Else
                sError = exCOM.Message
            End If
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Catch ex As Exception
            If ex.InnerException IsNot Nothing AndAlso ex.InnerException.Message <> "" Then
                sError = ex.InnerException.Message
            Else
                sError = ex.Message
            End If
            oLog.escribeMensaje(sError, EXO_Log.EXO_Log.Tipo.error)
        Finally
            ns = New XmlSerializerNamespaces()
            ns.Add("ew", "http://www.reifen.net")

            sw = New StringWriter
            xmlSerializador = New XmlSerializer(GetType(Stock_response_A2))
            oLog.escribeMensaje("Consultando Información...", EXO_Log.EXO_Log.Tipo.informacion)
            clsResponseOrder = CargarDatosStock(oXml, sError, oLog)

            xmlSerializador.Serialize(sw, clsResponseOrder, ns)
            Dim strRespuesta As String

            strRespuesta = sw.ToString
            strRespuesta = Replace(strRespuesta, "<Articulos>" & Chr(13) & Chr(10), "", , , CompareMethod.Text)
            strRespuesta = Replace(strRespuesta, "</Articulos>" & Chr(13) & Chr(10), "", , , CompareMethod.Text)
            strRespuesta = Replace(strRespuesta, "Error1", "Error", , , CompareMethod.Text)
            strRespuesta = Replace(strRespuesta, "Variant1", "Variant", , , CompareMethod.Text)
            strRespuesta = Replace(strRespuesta, "utf-16", "utf-8")
            oLog.escribeMensaje("Conviertiendo Información...", EXO_Log.EXO_Log.Tipo.informacion)
            Dim dResp As New XmlDocument
            dResp.LoadXml(strRespuesta)
            oLog.escribeMensaje("Enviando Información...", EXO_Log.EXO_Log.Tipo.informacion)
            ConsultaStock = dResp.DocumentElement
            'ConsultaStock = strRespuesta
            oLog.escribeMensaje("Enviada Información...", EXO_Log.EXO_Log.Tipo.informacion)
            If oXml IsNot Nothing Then
                oXml = Nothing
            End If

            If sw IsNot Nothing Then
                sw.Close()
                sw.Dispose()
                sw = Nothing
            End If
            clsErrores = Nothing
            ns = Nothing
            oLog.escribeMensaje("Fin Llamada a ""Consulta Stock""...", EXO_Log.EXO_Log.Tipo.advertencia)
        End Try
    End Function
    Function CargarDatos(ByRef oxml As XDocument, ByRef strError As String, ByVal sPedido As String) As Order_response_A2
        Dim blEXO As BLEXO = Nothing
        Try
            Dim Datos As New Order_response_A2
            blEXO = New BLEXO
            Datos.DocumentID = oxml.Descendants("DocumentID").Value.ToString
            Datos.Variant1 = oxml.Descendants("Variant").Value.ToString
            Datos.ErrorHead = New ErrorHead
            If strError <> "" Then
                Datos.ErrorHead.ErrorCode = strError
            Else
                Datos.ErrorHead.ErrorCode = 0
            End If
            Datos.CustomerReference = New CustomerReference
            Datos.CustomerReference.DocumentID = oxml.Descendants("CustomerReference").Elements("DocumentID").Value.ToString
            Datos.BuyerParty = New BuyerParty
            Datos.BuyerParty.PartyID = oxml.Descendants("BuyerParty").Elements("PartyID").Value.ToString
            Datos.BuyerParty.AgencyCode = oxml.Descendants("BuyerParty").Elements("AgencyCode").Value.ToString
            'Datos.OrderingParty = New OrderingParty
            'Datos.OrderingParty.PartyID = oxml.Descendants("OrderingParty").Elements("PartyID").Value.ToString
            'Datos.OrderingParty.AgencyCode = oxml.Descendants("OrderingParty").Elements("AgencyCode").Value.ToString
            'Datos.Consignee = New Consignee
            'Datos.Consignee.PartyID = oxml.Descendants("Consignee").Elements("PartyID").Value.ToString
            'Datos.Consignee.AgencyCode = oxml.Descendants("Consignee").Elements("AgencyCode").Value.ToString

            Datos.Articulos = New List(Of OrderLine)
            For Each report As XElement In oxml.Descendants("OrderLine")
                Dim item As New OrderLine

                item.LineID = report.Element("LineID").Value.ToString
                item.SuppliersOrderLineID = report.Element("LineID").Value.ToString

                'item.AdditionalCustomerReferenceNumber = New AdditionalCustomerReferenceNumber
                'item.AdditionalCustomerReferenceNumber.DocumentID = report.Element("AdditionalCustomerReferenceNumber").Element("DocumentID").Value.ToString
                item.OrderedArticle = New OrderedArticle
                item.OrderedArticle.ArticleIdentification = New ArticleIdentification
                item.OrderedArticle.ArticleIdentification.ManufacturersArticleID = report.Element("OrderedArticle").Element("ArticleIdentification").Element("ManufacturersArticleID").Value.ToString
                item.OrderedArticle.ArticleIdentification.EANUCCArticleID = report.Element("OrderedArticle").Element("ArticleIdentification").Element("EANUCCArticleID").Value.ToString
                item.OrderedArticle.ArticleDescription = New ArticleDescription
                Dim strCodArt As String = ""
                If report.Element("OrderedArticle").Element("ArticleIdentification").Element("EANUCCArticleID").Value.ToString <> "" Then
                    strCodArt = blEXO.GetValueDB("""OITM""", """ItemCode""", """U_SEI_JANCODE"" ='" & report.Element("OrderedArticle").Element("ArticleIdentification").Element("EANUCCArticleID").Value.ToString & "'")
                Else
                    strCodArt = blEXO.GetValueDB("""OITM""", """ItemCode""", """ItemCode"" ='" & report.Element("OrderedArticle").Element("ArticleIdentification").Element("ManufacturersArticleID").Value.ToString & "'")
                End If
                item.OrderedArticle.ArticleDescription.ArticleDescriptionText = blEXO.GetValueDB("""OITM""", """ItemName""", """ItemCode"" ='" & strCodArt & "'")
                item.OrderedArticle.RequestedDeliveryDate = report.Element("OrderedArticle").Element("RequestedDeliveryDate").Value.ToString
                item.OrderedArticle.ArticleComment = ""
                item.OrderedArticle.OrderReference = New OrderReference
                item.OrderedArticle.OrderReference.DocumentID = sPedido 'Nº de SAP
                item.OrderedArticle.Error1 = New Error1
                If strError <> "" Then
                    If strError.Contains("-4008") Then
                        item.OrderedArticle.Error1.ErrorCode = "301"
                    Else
                        item.OrderedArticle.Error1.ErrorCode = "1"
                    End If
                Else
                    item.OrderedArticle.Error1.ErrorCode = "0"
                End If
                item.OrderedArticle.Error1.ErrorText = strError
                item.OrderedArticle.ScheduleDetails = New ScheduleDetails
                item.OrderedArticle.ScheduleDetails.DeliveryDate = report.Element("OrderedArticle").Element("RequestedDeliveryDate").Value.ToString
                item.OrderedArticle.ScheduleDetails.ConfirmedQuantity = New ConfirmedQuantity
                item.OrderedArticle.ScheduleDetails.ConfirmedQuantity.QuantityValue = report.Element("OrderedArticle").Element("RequestedQuantity").Element("QuantityValue").Value
                item.OrderedArticle.OrderedQuantity = New OrderedQuantity
                item.OrderedArticle.OrderedQuantity.QuantityValue = report.Element("OrderedArticle").Element("RequestedQuantity").Element("QuantityValue").Value
                Datos.Articulos.Add(item)
            Next

            CargarDatos = Datos

        Catch ex As Exception
            Throw ex
        End Try

    End Function
    Function CargarDatosStock(ByRef oxml As XDocument, ByRef strError As String, ByRef oLog As EXO_Log.EXO_Log) As Stock_response_A2
        Try
            Dim Datos As New Stock_response_A2
            Dim sSQL As String = ""
            Dim dtStock As System.Data.DataTable = Nothing
            Dim blEXO As BLEXO = Nothing
            blEXO = New BLEXO

            Datos.DocumentID = oxml.Descendants("DocumentID").Value.ToString
            Datos.Variant1 = oxml.Descendants("Variant").Value.ToString
            Datos.ErrorHead = New ErrorHead
            If strError = "" Then
                Datos.ErrorHead.ErrorCode = "0"
            Else
                Datos.ErrorHead.ErrorCode = strError
            End If
            Datos.CustomerReference = New CustomerReference
            Datos.CustomerReference.DocumentID = oxml.Descendants("CustomerReference").Elements("DocumentID").Value.ToString
            Datos.BuyerParty = New BuyerParty
            Datos.BuyerParty.PartyID = oxml.Descendants("BuyerParty").Elements("PartyID").Value.ToString
            Datos.BuyerParty.AgencyCode = oxml.Descendants("BuyerParty").Elements("AgencyCode").Value.ToString
            Datos.Consignee = New Consignee
            Datos.Consignee.PartyID = oxml.Descendants("Consignee").Elements("PartyID").Value.ToString
            Datos.Consignee.AgencyCode = oxml.Descendants("Consignee").Elements("AgencyCode").Value.ToString

            Datos.Articulos = New List(Of OrderLine)
            For Each report As XElement In oxml.Descendants("OrderLine")
                Dim item As New OrderLine

                item.LineID = report.Element("LineID").Value.ToString
                item.OrderedArticle = New OrderedArticle
                item.OrderedArticle.ArticleIdentification = New ArticleIdentification
                Dim sCodBarras As String = report.Element("OrderedArticle").Element("ArticleIdentification").Element("ManufacturersArticleID").Value.ToString
                item.OrderedArticle.ArticleIdentification.ManufacturersArticleID = sCodBarras

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
            & " WHERE U_SEI_JANCODE='" & sCodBarras & "' " _
            & " group by ItemCode,ItemName,U_SEI_JANCODE " _
            & " order by ItemCode"
                dtStock = New System.Data.DataTable("Articulos")
                blEXO.FillDtDB(dtStock, sSQL)
                Dim iDisponible As Integer = 0
                If dtStock.Rows.Count > 0 Then
                    item.OrderedArticle.ArticleDescription = New ArticleDescription
                    item.OrderedArticle.ArticleDescription.ArticleDescriptionText = dtStock.Rows(0).Item("ItemName").ToString
                    item.OrderedArticle.Availability = CStr(dtStock.Rows(0).Item("AVAILABILITY").ToString)
                    If dtStock.Rows(0).Item("AVAILABILITY").ToString <> ">50" Then
                        iDisponible = CInt(dtStock.Rows(0).Item("AVAILABILITY").ToString)
                    Else
                        iDisponible = 50
                    End If

                End If
                item.OrderedArticle.RequestedQuantity = New RequestedQuantity
                item.OrderedArticle.RequestedQuantity.QuantityValue = report.Element("OrderedArticle").Element("RequestedQuantity").Element("QuantityValue").Value.ToString
                item.OrderedArticle.Error1 = New Error1
                If strError <> "" Then
                    If strError.Contains("-4008") Then
                        item.OrderedArticle.Error1.ErrorCode = "301"
                    End If
                Else
                    item.OrderedArticle.Error1.ErrorCode = "0"
                End If
                item.OrderedArticle.ScheduleDetails = New ScheduleDetails
                If iDisponible = 0 Then
                    item.OrderedArticle.ScheduleDetails.DeliveryDate = "9999-12-31"
                Else
                    Dim dFecha As Date = Now
                    Dim sHora As String = Format(Now, "HH:mm:ss")
                    If Left(sHora, 2) >= 18 Then
                        dFecha = DateAdd(DateInterval.Day, 2, dFecha)
                    Else
                        dFecha = DateAdd(DateInterval.Day, 1, dFecha)
                    End If
                    item.OrderedArticle.ScheduleDetails.DeliveryDate = Year(dFecha).ToString & "-" & Month(dFecha).ToString & "-" & Day(dFecha)
                End If
                'item.OrderedArticle.ScheduleDetails.DeliveryDate = Year(Now).ToString & "-" & Month(Now).ToString & "-" & Day(Now)
                item.OrderedArticle.ScheduleDetails.AvailableQuantity = New AvailableQuantity
                Dim iCPed As Integer = CInt(report.Element("OrderedArticle").Element("RequestedQuantity").Element("QuantityValue").Value.ToString)
                If iCPed <= iDisponible Then
                    item.OrderedArticle.ScheduleDetails.AvailableQuantity.QuantityValue = CInt(report.Element("OrderedArticle").Element("RequestedQuantity").Element("QuantityValue").Value.ToString)
                Else
                    item.OrderedArticle.ScheduleDetails.AvailableQuantity.QuantityValue = iDisponible.ToString
                End If

                Datos.Articulos.Add(item)
            Next
            CargarDatosStock = Datos
        Catch ex As Exception
            Throw ex
        End Try
    End Function
#Region "CrearXML"
    Private Sub createNode(ByVal writer As XmlTextWriter, ByVal datos As Order_response_A2, ByVal strError As String)
        Try

            writer.WriteStartElement("DocumentID")
            writer.WriteString("A2")
            writer.WriteEndElement()
            writer.WriteStartElement("Variant")
            writer.WriteString("5")
            writer.WriteEndElement()
            writer.WriteStartElement("ErrorHead")
            writer.WriteStartElement("ErrorCode")
            If strError <> "" Then
                If strError.Contains("-4008") Then
                    writer.WriteString("301")
                End If
            Else
                writer.WriteString("0")
            End If
            writer.WriteEndElement()
            writer.WriteEndElement()

            writer.WriteStartElement("CustomerReference")
            writer.WriteStartElement("DocumentID")
            writer.WriteString(datos.CustomerReference.DocumentID.ToString)
            writer.WriteEndElement()
            writer.WriteEndElement()

            writer.WriteStartElement("BuyerParty")

            writer.WriteStartElement("PartyID")
            writer.WriteString(datos.BuyerParty.PartyID.ToString)
            writer.WriteEndElement()
            writer.WriteStartElement("AgencyCode")
            writer.WriteString(datos.BuyerParty.AgencyCode.ToString)
            writer.WriteEndElement()


            writer.WriteEndElement()

            'recorrer lineas
            Dim i As Integer = 0
            For Each linea As OrderLine In datos.Articulos

                'If i = 1 Then
                writer.WriteStartElement("OrderLine") 'abro OrderLine

                writer.WriteStartElement("LineID")  'abro  LineID
                writer.WriteString(linea.LineID.ToString)
                writer.WriteEndElement() 'cierreLineID
                writer.WriteStartElement("SuppliersOrderLineID") 'abro  SuppliersOrderLineID
                writer.WriteString(linea.LineID.ToString)
                writer.WriteEndElement() 'cierre SuppliersOrderLineID

                writer.WriteStartElement("OrderedArticle") 'abro  OrderedArticle

                writer.WriteStartElement("ArticleIdentification")  'abro  ArticleIdentification
                writer.WriteStartElement("ManufacturersArticleID") 'abro  ManufacturersArticleID
                writer.WriteString(linea.OrderedArticle.ArticleIdentification.ManufacturersArticleID.ToString)
                writer.WriteEndElement() 'cierre ManufacturersArticleID
                writer.WriteStartElement("EANUCCArticleID")
                writer.WriteString(linea.OrderedArticle.ArticleIdentification.EANUCCArticleID.ToString)
                writer.WriteEndElement() 'cierre EANUCCArticleID
                writer.WriteEndElement() 'cierre ArticleIdentification

                writer.WriteStartElement("ArticleDescription")
                writer.WriteStartElement("ArticleDescriptionText")
                writer.WriteString(linea.OrderedArticle.ArticleDescription.ArticleDescriptionText)
                writer.WriteEndElement() 'cierre ArticleDescriptionText
                writer.WriteEndElement() 'cierre ArticleDescription

                writer.WriteStartElement("OrderReference")
                writer.WriteStartElement("DocumentID")
                writer.WriteString("")
                writer.WriteEndElement() 'cierre DocumentID
                writer.WriteEndElement() 'cierre OrderReference

                writer.WriteStartElement("Error") ' abro Error
                writer.WriteStartElement("ErrorCode") ' abro ErrorCode
                writer.WriteString("0")
                writer.WriteEndElement() 'cierre ErrorCode
                writer.WriteEndElement() 'cierr eError

                writer.WriteStartElement("ScheduleDetails")
                writer.WriteStartElement("DeliveryDate")
                writer.WriteString(linea.OrderedArticle.RequestedDeliveryDate.ToString)
                writer.WriteEndElement() 'cierre DeliveryDate
                writer.WriteStartElement("ConfirmedQuantity")
                writer.WriteStartElement("QuantityValue")
                writer.WriteString(linea.OrderedArticle.RequestedQuantity.QuantityValue.ToString)
                writer.WriteEndElement() 'cierre QuantityValue
                writer.WriteEndElement() 'cierre ConfirmedQuantity
                writer.WriteEndElement() 'cierre ScheduleDetails

                writer.WriteStartElement("OrderedQuantity")
                writer.WriteStartElement("QuantityValue")
                writer.WriteString(linea.OrderedArticle.RequestedQuantity.QuantityValue.ToString)
                writer.WriteEndElement() 'cierre QuantityValue
                writer.WriteEndElement() ' fin OrderedQuantity

                writer.WriteEndElement() ' fin OrderLine
                writer.WriteEndElement() ' fin OrderedArticle
                'End If
                i = +1
            Next

        Catch ex As Exception
            writer.WriteEndElement()
            writer.WriteEndDocument()
            writer.Close()
        End Try
    End Sub
#End Region
End Class

