' NOTA: puede usar el comando "Cambiar nombre" del menú contextual para cambiar el nombre de interfaz "IService1" en el código y en el archivo de configuración a la vez.
Imports System.Xml
<ServiceContract()>
Public Interface IYokoRODINEXService

    ' TODO: Add your service operations here
    <OperationContract()>
    Function CrearPedido(ByVal xOrder As XmlElement) As XmlElement

    <OperationContract()>
    Function ConsultaStock(ByVal xStock As XmlElement) As XmlElement
End Interface

' Utilice un contrato de datos, como se ilustra en el ejemplo siguiente, para agregar tipos compuestos a las operaciones de servicio.

<DataContract()>
<Serializable()>
Public Class Errores

#Region "Atributos"

    Private _TextError As String

#End Region

#Region "Propiedades"

    <DataMember()>
    Public Property TextError() As String
        Get
            Return _TextError
        End Get
        Set(ByVal Value As String)
            _TextError = Value
        End Set
    End Property

#End Region

End Class
<DataContract()>
<Serializable()>
Public Class PEDIDOSAP
#Region "Atributos"
    Private _DocumentID As String

#End Region
#Region "Propiedades"
    <DataMember()>
    Public Property DocumentID() As String
        Get
            Return _DocumentID
        End Get
        Set(ByVal Value As String)
            _DocumentID = Value
        End Set
    End Property
#End Region
End Class

<DataContract()>
<Serializable()>
Public Class Order_response_A2
#Region "Atributos"

    Private _DocumentID As String
    Private _Variant As String
    Private _Campaign As String
    Private _ErrorHead As ErrorHead
    Private _CustomerReference As CustomerReference
    Private _BuyerParty As BuyerParty
    Private _OrderingParty As OrderingParty
    Private _Consignee As Consignee
    Private _Articulos As List(Of OrderLine)
    Private _OrderedArticle As OrderedArticle


#End Region
#Region "Propiedades"

    <DataMember()>
    Public Property DocumentID() As String
        Get
            Return _DocumentID
        End Get
        Set(ByVal Value As String)
            _DocumentID = Value
        End Set
    End Property

    <DataMember()>
    Public Property Variant1() As String
        Get
            Return _Variant
        End Get
        Set(ByVal Value As String)
            _Variant = Value
        End Set
    End Property
    <DataMember()>
    Public Property Campaign() As String
        Get
            Return _Campaign
        End Get
        Set(ByVal Value As String)
            _Campaign = Value
        End Set
    End Property

    <DataMember()>
    Public Property ErrorHead() As ErrorHead
        Get
            Return _ErrorHead
        End Get
        Set(ByVal Value As ErrorHead)
            _ErrorHead = Value
        End Set
    End Property

    <DataMember()>
    Public Property CustomerReference() As CustomerReference
        Get
            Return _CustomerReference
        End Get
        Set(ByVal Value As CustomerReference)
            _CustomerReference = Value
        End Set
    End Property

    <DataMember()>
    Public Property BuyerParty() As BuyerParty
        Get
            Return _BuyerParty
        End Get
        Set(ByVal Value As BuyerParty)
            _BuyerParty = Value
        End Set
    End Property
    <DataMember()>
    Public Property OrderingParty() As OrderingParty
        Get
            Return _OrderingParty
        End Get
        Set(ByVal Value As OrderingParty)
            _OrderingParty = Value
        End Set
    End Property
    <DataMember()>
    Public Property Consignee() As Consignee
        Get
            Return _Consignee
        End Get
        Set(ByVal Value As Consignee)
            _Consignee = Value
        End Set
    End Property
    <DataMember()>
    Public Property Articulos As List(Of OrderLine)
        Get
            Return _Articulos
        End Get
        Set(ByVal Value As List(Of OrderLine))
            _Articulos = Value
        End Set
    End Property


#End Region
End Class
<DataContract()>
<Serializable()>
Public Class Stock_response_A2
#Region "Atributos"

    Private _DocumentID As String
    Private _Variant As String
    Private _Campaign As String
    Private _ErrorHead As ErrorHead
    Private _CustomerReference As CustomerReference
    Private _BuyerParty As BuyerParty
    Private _OrderingParty As OrderingParty
    Private _Consignee As Consignee
    Private _Articulos As List(Of OrderLine)
    Private _OrderedArticle As OrderedArticle


#End Region
#Region "Propiedades"

    <DataMember()>
    Public Property DocumentID() As String
        Get
            Return _DocumentID
        End Get
        Set(ByVal Value As String)
            _DocumentID = Value
        End Set
    End Property

    <DataMember()>
    Public Property Variant1() As String
        Get
            Return _Variant
        End Get
        Set(ByVal Value As String)
            _Variant = Value
        End Set
    End Property
    <DataMember()>
    Public Property Campaign() As String
        Get
            Return _Campaign
        End Get
        Set(ByVal Value As String)
            _Campaign = Value
        End Set
    End Property

    <DataMember()>
    Public Property ErrorHead() As ErrorHead
        Get
            Return _ErrorHead
        End Get
        Set(ByVal Value As ErrorHead)
            _ErrorHead = Value
        End Set
    End Property

    <DataMember()>
    Public Property CustomerReference() As CustomerReference
        Get
            Return _CustomerReference
        End Get
        Set(ByVal Value As CustomerReference)
            _CustomerReference = Value
        End Set
    End Property

    <DataMember()>
    Public Property BuyerParty() As BuyerParty
        Get
            Return _BuyerParty
        End Get
        Set(ByVal Value As BuyerParty)
            _BuyerParty = Value
        End Set
    End Property
    <DataMember()>
    Public Property OrderingParty() As OrderingParty
        Get
            Return _OrderingParty
        End Get
        Set(ByVal Value As OrderingParty)
            _OrderingParty = Value
        End Set
    End Property
    <DataMember()>
    Public Property Consignee() As Consignee
        Get
            Return _Consignee
        End Get
        Set(ByVal Value As Consignee)
            _Consignee = Value
        End Set
    End Property
    <DataMember()>
    Public Property Articulos As List(Of OrderLine)
        Get
            Return _Articulos
        End Get
        Set(ByVal Value As List(Of OrderLine))
            _Articulos = Value
        End Set
    End Property


#End Region
End Class
<DataContract()>
<Serializable()>
Public Class ErrorHead
#Region "Atributos"
    <DataMember()>
    Private _ErrorCode As String
#End Region
#Region "Propiedades"
    <DataMember()>
    Public Property ErrorCode() As String
        Get
            Return _ErrorCode
        End Get
        Set(ByVal Value As String)
            _ErrorCode = Value
        End Set
    End Property
#End Region
End Class

<DataContract()>
<Serializable()>
Public Class CustomerReference
#Region "Atributos"
    <DataMember()>
    Private _DocumentID As String
#End Region
#Region "Propiedades"
    <DataMember()>
    Public Property DocumentID() As String
        Get
            Return _DocumentID
        End Get
        Set(ByVal Value As String)
            _DocumentID = Value
        End Set
    End Property
#End Region
End Class

<DataContract()>
<Serializable()>
Public Class BuyerParty
#Region "Atributos"
    <DataMember()>
    Private _PartyID As String
    <DataMember()>
    Private _AgencyCode As String
#End Region
#Region "Propiedades"
    <DataMember()>
    Public Property PartyID() As String
        Get
            Return _PartyID
        End Get
        Set(ByVal Value As String)
            _PartyID = Value
        End Set
    End Property

    <DataMember()>
    Public Property AgencyCode() As String
        Get
            Return _AgencyCode
        End Get
        Set(ByVal Value As String)
            _AgencyCode = Value
        End Set
    End Property
#End Region
End Class
<DataContract()>
<Serializable()>
Public Class OrderingParty
#Region "Atributos"
    <DataMember()>
    Private _PartyID As String
    <DataMember()>
    Private _AgencyCode As String
#End Region
#Region "Propiedades"
    <DataMember()>
    Public Property PartyID() As String
        Get
            Return _PartyID
        End Get
        Set(ByVal Value As String)
            _PartyID = Value
        End Set
    End Property

    <DataMember()>
    Public Property AgencyCode() As String
        Get
            Return _AgencyCode
        End Get
        Set(ByVal Value As String)
            _AgencyCode = Value
        End Set
    End Property
#End Region
End Class
<DataContract()>
<Serializable()>
Public Class Consignee
#Region "Atributos"
    <DataMember()>
    Private _PartyID As String
    <DataMember()>
    Private _AgencyCode As String
#End Region
#Region "Propiedades"
    <DataMember()>
    Public Property PartyID() As String
        Get
            Return _PartyID
        End Get
        Set(ByVal Value As String)
            _PartyID = Value
        End Set
    End Property

    <DataMember()>
    Public Property AgencyCode() As String
        Get
            Return _AgencyCode
        End Get
        Set(ByVal Value As String)
            _AgencyCode = Value
        End Set
    End Property
#End Region
End Class
<DataContract()>
<Serializable()>
Public Class OrderLine
#Region "Atributos"
    <DataMember()>
    Private _LineID As String
    <DataMember()>
    Private _SuppliersOrderLineID As String
    <DataMember()>
    Private _AdditionalCustomerReferenceNumber As AdditionalCustomerReferenceNumber
    <DataMember()>
    Private _OrderedArticle As OrderedArticle

#End Region
#Region "Propiedades"
    <DataMember()>
    Public Property LineID() As String
        Get
            Return _LineID
        End Get
        Set(ByVal Value As String)
            _LineID = Value
        End Set
    End Property

    <DataMember()>
    Public Property SuppliersOrderLineID() As String
        Get
            Return _SuppliersOrderLineID
        End Get
        Set(ByVal Value As String)
            _SuppliersOrderLineID = Value
        End Set
    End Property
    <DataMember()>
    Public Property AdditionalCustomerReferenceNumber() As AdditionalCustomerReferenceNumber
        Get
            Return _AdditionalCustomerReferenceNumber
        End Get
        Set(ByVal Value As AdditionalCustomerReferenceNumber)
            _AdditionalCustomerReferenceNumber = Value
        End Set
    End Property
    <DataMember()>
    Public Property OrderedArticle() As OrderedArticle
        Get
            Return _OrderedArticle
        End Get
        Set(ByVal Value As OrderedArticle)
            _OrderedArticle = Value
        End Set
    End Property
#End Region
End Class

<DataContract()>
<Serializable()>
Public Class AdditionalCustomerReferenceNumber
#Region "Atributos"
    <DataMember()>
    Private _DocumentID As String
#End Region
#Region "Propiedades"
    <DataMember()>
    Public Property DocumentID() As String
        Get
            Return _DocumentID
        End Get
        Set(ByVal Value As String)
            _DocumentID = Value
        End Set
    End Property
#End Region
End Class

<DataContract()>
<Serializable()>
Public Class OrderedArticle
#Region "Atributos"
    <DataMember()>
    Private _ArticleIdentification As ArticleIdentification
    <DataMember()>
    Private _ArticleDescription As ArticleDescription
    <DataMember()>
    Private _Availability As String
    <DataMember()>
    Private _RequestedDeliveryDate As String
    <DataMember()>
    Private _ArticleComment As String
    <DataMember()>
    Private _OrderReference As OrderReference
    <DataMember()>
    Private _RequestedQuantity As RequestedQuantity
    <DataMember()>
    Private _Error As Error1
    <DataMember()>
    Private _OrderedQuantity As OrderedQuantity
    <DataMember()>
    Private _ScheduleDetails As ScheduleDetails
#End Region
#Region "Propiedades"

    <DataMember()>
    Public Property RequestedDeliveryDate() As String
        Get
            Return _RequestedDeliveryDate
        End Get
        Set(ByVal Value As String)
            _RequestedDeliveryDate = Value
        End Set
    End Property
    <DataMember()>
    Public Property ArticleComment() As String
        Get
            Return _ArticleComment
        End Get
        Set(ByVal Value As String)
            _ArticleComment = Value
        End Set
    End Property

    <DataMember()>
    Public Property ArticleIdentification() As ArticleIdentification
        Get
            Return _ArticleIdentification
        End Get
        Set(ByVal Value As ArticleIdentification)
            _ArticleIdentification = Value
        End Set
    End Property
    <DataMember()>
    Public Property ArticleDescription() As ArticleDescription
        Get
            Return _ArticleDescription
        End Get
        Set(ByVal Value As ArticleDescription)
            _ArticleDescription = Value
        End Set
    End Property
    <DataMember()>
    Public Property Availability() As String
        Get
            Return _Availability
        End Get
        Set(ByVal Value As String)
            _Availability = Value
        End Set
    End Property
    <DataMember()>
    Public Property OrderReference() As OrderReference
        Get
            Return _OrderReference
        End Get
        Set(ByVal Value As OrderReference)
            _OrderReference = Value
        End Set
    End Property

    <DataMember()>
    Public Property RequestedQuantity() As RequestedQuantity
        Get
            Return _RequestedQuantity
        End Get
        Set(ByVal Value As RequestedQuantity)
            _RequestedQuantity = Value
        End Set
    End Property
    <DataMember()>
    Public Property Error1() As Error1
        Get
            Return _Error
        End Get
        Set(ByVal Value As Error1)
            _Error = Value
        End Set
    End Property

    <DataMember()>
    Public Property OrderedQuantity() As OrderedQuantity
        Get
            Return _OrderedQuantity
        End Get
        Set(ByVal Value As OrderedQuantity)
            _OrderedQuantity = Value
        End Set
    End Property
    <DataMember()>
    Public Property ScheduleDetails() As ScheduleDetails
        Get
            Return _ScheduleDetails
        End Get
        Set(ByVal Value As ScheduleDetails)
            _ScheduleDetails = Value
        End Set
    End Property
#End Region
End Class
<DataContract()>
<Serializable()>
Public Class ScheduleDetails
#Region "Atributos"
    <DataMember()>
    Private _DeliveryDate As String
    <DataMember()>
    Private _AvailableQuantity As AvailableQuantity
    <DataMember()>
    Private _ConfirmedQuantity As ConfirmedQuantity
#End Region
#Region "Propiedades"
    <DataMember()>
    Public Property DeliveryDate() As String
        Get
            Return _DeliveryDate
        End Get
        Set(ByVal Value As String)
            _DeliveryDate = Value
        End Set
    End Property
    <DataMember()>
    Public Property AvailableQuantity() As AvailableQuantity
        Get
            Return _AvailableQuantity
        End Get
        Set(ByVal Value As AvailableQuantity)
            _AvailableQuantity = Value
        End Set
    End Property
    <DataMember()>
    Public Property ConfirmedQuantity() As ConfirmedQuantity
        Get
            Return _ConfirmedQuantity
        End Get
        Set(ByVal Value As ConfirmedQuantity)
            _ConfirmedQuantity = Value
        End Set
    End Property
#End Region
End Class
<DataContract()>
<Serializable()>
Public Class AvailableQuantity
#Region "Atributos"
    <DataMember()>
    Private _QuantityValue As String

#End Region
#Region "Propiedades"
    <DataMember()>
    Public Property QuantityValue() As String
        Get
            Return _QuantityValue
        End Get
        Set(ByVal Value As String)
            _QuantityValue = Value
        End Set
    End Property

#End Region
End Class
<DataContract()>
<Serializable()>
Public Class ConfirmedQuantity
#Region "Atributos"
    <DataMember()>
    Private _QuantityValue As String

#End Region
#Region "Propiedades"
    <DataMember()>
    Public Property QuantityValue() As String
        Get
            Return _QuantityValue
        End Get
        Set(ByVal Value As String)
            _QuantityValue = Value
        End Set
    End Property

#End Region
End Class
<DataContract()>
<Serializable()>
Public Class ArticleIdentification
#Region "Atributos"
    <DataMember()>
    Private _ManufacturersArticleID As String
    <DataMember()>
    Private _EANUCCArticleID As String
#End Region
#Region "Propiedades"
    <DataMember()>
    Public Property ManufacturersArticleID() As String
        Get
            Return _ManufacturersArticleID
        End Get
        Set(ByVal Value As String)
            _ManufacturersArticleID = Value
        End Set
    End Property

    <DataMember()>
    Public Property EANUCCArticleID() As String
        Get
            Return _EANUCCArticleID
        End Get
        Set(ByVal Value As String)
            _EANUCCArticleID = Value
        End Set
    End Property
#End Region
End Class
<DataContract()>
<Serializable()>
Public Class RequestedQuantity
#Region "Atributos"
    <DataMember()>
    Private _QuantityValue As String

#End Region
#Region "Propiedades"
    <DataMember()>
    Public Property QuantityValue() As String
        Get
            Return _QuantityValue
        End Get
        Set(ByVal Value As String)
            _QuantityValue = Value
        End Set
    End Property

#End Region
End Class
<DataContract()>
<Serializable()>
Public Class OrderedQuantity
#Region "Atributos"
    <DataMember()>
    Private _QuantityValue As String

#End Region
#Region "Propiedades"
    <DataMember()>
    Public Property QuantityValue() As String
        Get
            Return _QuantityValue
        End Get
        Set(ByVal Value As String)
            _QuantityValue = Value
        End Set
    End Property

#End Region
End Class
<DataContract()>
<Serializable()>
Public Class ArticleDescription
#Region "Atributos"
    <DataMember()>
    Private _ArticleDescriptionText As String

#End Region
#Region "Propiedades"
    <DataMember()>
    Public Property ArticleDescriptionText() As String
        Get
            Return _ArticleDescriptionText
        End Get
        Set(ByVal Value As String)
            _ArticleDescriptionText = Value
        End Set
    End Property

#End Region
End Class

<DataContract()>
<Serializable()>
Public Class OrderReference
#Region "Atributos"
    <DataMember()>
    Private _DocumentID As String

#End Region
#Region "Propiedades"
    <DataMember()>
    Public Property DocumentID() As String
        Get
            Return _DocumentID
        End Get
        Set(ByVal Value As String)
            _DocumentID = Value
        End Set
    End Property

#End Region
End Class

<DataContract()>
<Serializable()>
Public Class Error1
#Region "Atributos"
    <DataMember()>
    Private _ErrorCode As String
    <DataMember()>
    Private _ErrorText As String
#End Region
#Region "Propiedades"
    <DataMember()>
    Public Property ErrorCode() As String
        Get
            Return _ErrorCode
        End Get
        Set(ByVal Value As String)
            _ErrorCode = Value
        End Set
    End Property
    <DataMember()>
    Public Property ErrorText() As String
        Get
            Return _ErrorText
        End Get
        Set(ByVal Value As String)
            _ErrorText = Value
        End Set
    End Property

#End Region
End Class
