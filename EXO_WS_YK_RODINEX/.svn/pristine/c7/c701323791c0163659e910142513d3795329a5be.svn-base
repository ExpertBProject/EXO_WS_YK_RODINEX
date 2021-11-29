Imports System.Net
Imports System.IO
Imports System.Xml
Public Class Pruebas
    'Declaración de método necesario para poder acceder al servicio web a través del protocolo https
    Private Function ValidarCertificado(ByVal sender As Object, ByVal certificate As System.Security.Cryptography.X509Certificates.X509Certificate,
                                        ByVal chain As System.Security.Cryptography.X509Certificates.X509Chain,
                                        ByVal sslPolicyErrors As System.Net.Security.SslPolicyErrors) As Boolean
        Return True
    End Function

    Private Sub btnPedido_Click(sender As Object, e As EventArgs) Handles btnPedido.Click
        Try

            Dim oYokoWS As YokoRODINEXService.YokoRODINEXServiceClient = New YokoRODINEXService.YokoRODINEXServiceClient
            Dim sResultado As String = ""

            System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)

            'Llamada necesaria para poder acceder al servicio web por el protocolo https

            oYokoWS.ClientCredentials.UserName.UserName = "YKRODINEX"
            oYokoWS.ClientCredentials.UserName.Password = "YK982RNwh0"
            Dim strCadena As String = "<?xml version=""1.0"" encoding=""UTF-8""?><ew:order_A2 xmlns:ew=""http://www.reifen.net""><DocumentID>A2</DocumentID><Variant>5</Variant><CustomerReference><DocumentID>PCN50360</DocumentID></CustomerReference><BuyerParty><PartyID>12585</PartyID><AgencyCode>91</AgencyCode></BuyerParty><Consignee><PartyID>ADM-NEX</PartyID><AgencyCode>91</AgencyCode></Consignee><OrderLine><LineID>1</LineID><OrderedArticle><ArticleIdentification><ManufacturersArticleID/><EANUCCArticleID>4968814938239</EANUCCArticleID></ArticleIdentification><RequestedDeliveryDate>2019-10-25</RequestedDeliveryDate><RequestedQuantity><QuantityValue>1</QuantityValue></RequestedQuantity></OrderedArticle></OrderLine><OrderLine><LineID>2</LineID><OrderedArticle><ArticleIdentification><ManufacturersArticleID/><EANUCCArticleID>4968814938239</EANUCCArticleID></ArticleIdentification><RequestedDeliveryDate>2019-10-25</RequestedDeliveryDate><RequestedQuantity><QuantityValue>1</QuantityValue></RequestedQuantity></OrderedArticle></OrderLine></ew:order_A2>"
            Dim dDocumento As New System.Xml.XmlDocument
            Dim dResultado As XElement
            dDocumento.LoadXml(strCadena)
            Dim xE As XElement = XElement.Load(New XmlNodeReader(dDocumento))
            'sResultado = oYokoWS.CrearPedido(strCadena)
            dResultado = oYokoWS.CrearPedido(xE)
            sResultado = dResultado.ToString
            MessageBox.Show(sResultado)
            Dim file As System.IO.StreamWriter
            file = My.Computer.FileSystem.OpenTextFileWriter("C:\Desarrollo\Yokohama\EXO_WS_YK_RODINEX\Ficheros\Order_Response.xml", False)
            file.WriteLine(sResultado)
            file.Close()
            oYokoWS.Close()

        Catch ex As Exception
            MessageBox.Show(ex.InnerException.Message)
        End Try
    End Sub

    Private Sub btnConsulta_Click(sender As Object, e As EventArgs) Handles btnConsulta.Click
        Try

            Dim oYokoWS As YokoRODINEXService.YokoRODINEXServiceClient = New YokoRODINEXService.YokoRODINEXServiceClient
            Dim sResultado As String = ""

            System.Net.ServicePointManager.ServerCertificateValidationCallback = New System.Net.Security.RemoteCertificateValidationCallback(AddressOf ValidarCertificado)

            'Llamada necesaria para poder acceder al servicio web por el protocolo https

            oYokoWS.ClientCredentials.UserName.UserName = "YKRODINEX"
            oYokoWS.ClientCredentials.UserName.Password = "YK982RNwh0"
            'Dim strCadena As String = "<?xml version=""1.0"" encoding=""UTF-8""?><ew:inquiry_A2 xmlns:ew=""http://www.reifen.net""><DocumentID>A2</DocumentID><Variant>5</Variant><CustomerReference><DocumentID>ALOPEZ</DocumentID></CustomerReference><BuyerParty><PartyID>12585</PartyID><AgencyCode>91</AgencyCode></BuyerParty><Consignee><PartyID>NLLE</PartyID><AgencyCode>91</AgencyCode></Consignee><OrderLine><LineID>10</LineID><OrderedArticle><ArticleIdentification><ManufacturersArticleID>4968814925499</ManufacturersArticleID></ArticleIdentification><RequestedDeliveryDate/><RequestedQuantity><QuantityValue>4</QuantityValue></RequestedQuantity></OrderedArticle></OrderLine><OrderLine><LineID>20</LineID><OrderedArticle><ArticleIdentification><ManufacturersArticleID>4968814784355</ManufacturersArticleID></ArticleIdentification><RequestedDeliveryDate/><RequestedQuantity><QuantityValue>4</QuantityValue></RequestedQuantity></OrderedArticle></OrderLine><OrderLine><LineID>30</LineID><OrderedArticle><ArticleIdentification><ManufacturersArticleID>4968814759759</ManufacturersArticleID></ArticleIdentification><RequestedDeliveryDate/><RequestedQuantity><QuantityValue>4</QuantityValue></RequestedQuantity></OrderedArticle></OrderLine><OrderLine><LineID>40</LineID><OrderedArticle><ArticleIdentification><ManufacturersArticleID>4968814855840</ManufacturersArticleID></ArticleIdentification><RequestedDeliveryDate/><RequestedQuantity><QuantityValue>4</QuantityValue></RequestedQuantity></OrderedArticle></OrderLine><OrderLine><LineID>50</LineID><OrderedArticle><ArticleIdentification><ManufacturersArticleID>4968814840501</ManufacturersArticleID></ArticleIdentification><RequestedDeliveryDate/><RequestedQuantity><QuantityValue>4</QuantityValue></RequestedQuantity></OrderedArticle></OrderLine><OrderLine><LineID>60</LineID><OrderedArticle><ArticleIdentification><ManufacturersArticleID>4968814803667</ManufacturersArticleID></ArticleIdentification><RequestedDeliveryDate/><RequestedQuantity><QuantityValue>4</QuantityValue></RequestedQuantity></OrderedArticle></OrderLine><OrderLine><LineID>70</LineID><OrderedArticle><ArticleIdentification><ManufacturersArticleID>4968814801427</ManufacturersArticleID></ArticleIdentification><RequestedDeliveryDate/><RequestedQuantity><QuantityValue>4</QuantityValue></RequestedQuantity></OrderedArticle></OrderLine><OrderLine><LineID>80</LineID><OrderedArticle><ArticleIdentification><ManufacturersArticleID>4968814926373</ManufacturersArticleID></ArticleIdentification><RequestedDeliveryDate/><RequestedQuantity><QuantityValue>4</QuantityValue></RequestedQuantity></OrderedArticle></OrderLine></ew:inquiry_A2>"
            Dim strCadena As String = "<?xml version=""1.0"" encoding=""UTF-8""?><ew:inquiry_A2 xmlns:ew=""http://www.reifen.net""><DocumentID>A2</DocumentID><Variant>5</Variant><CustomerReference><DocumentID>JASANCHO</DocumentID></CustomerReference><BuyerParty><PartyID/><AgencyCode>91</AgencyCode></BuyerParty><Consignee><PartyID/><AgencyCode>91</AgencyCode></Consignee><OrderLine><LineID>10</LineID><OrderedArticle><ArticleIdentification><ManufacturersArticleID>0210840345</ManufacturersArticleID></ArticleIdentification><RequestedDeliveryDate>2019-08-29</RequestedDeliveryDate><RequestedQuantity><QuantityValue>4</QuantityValue></RequestedQuantity></OrderedArticle></OrderLine></ew:inquiry_A2>"
            Dim dDocumento As New System.Xml.XmlDocument
            Dim dResultado As XElement
            dDocumento.LoadXml(strCadena)
            Dim xE As XElement = XElement.Load(New XmlNodeReader(dDocumento))
            'sResultado = oYokoWS.ConsultaStock(strCadena)
            dResultado = oYokoWS.ConsultaStock(xE)
            sResultado = dResultado.ToString
            MessageBox.Show(sResultado)
            Dim file As System.IO.StreamWriter
            file = My.Computer.FileSystem.OpenTextFileWriter("C:\Desarrollo\Yokohama\EXO_WS_YK_RODINEX\Ficheros\Stock_Response.xml", False)
            file.WriteLine(sResultado)
            file.Close()
            oYokoWS.Close()

        Catch ex As Exception
            MessageBox.Show(ex.InnerException.Message)
        End Try
    End Sub
End Class
