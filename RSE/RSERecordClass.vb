Imports FileHelpers
Imports RSE.FileHelpers

<IgnoreFirst, DelimitedRecord(";")>
Public Class RSERecord

    <FieldOrder(1), FieldTitle("Week")>
    Public Week As String

    <FieldOrder(2), FieldTitle("DealerNo")>
    Public DealerNo As String

    <FieldOrder(3), FieldTitle("RSE_Active")>
    Public RSE_Active As String

    <FieldOrder(4), FieldTitle("Dealer Name")>
    Public DealerName As String

    <FieldOrder(5), FieldTitle("SalesmanNo")>
    Public SalesmanNo As Integer

    <FieldOrder(6), FieldTitle("Salesman_Name")>
    Public Salesman_Name As String

    <FieldOrder(7), FieldTitle("Brands")>
    Public Brands As String

    <FieldOrder(8), FieldTitle("Email contacts")>
    Public Emailcontacts As Integer

    <FieldOrder(9), FieldTitle("Telephone contacts")>
    Public Telephonecontacts As Integer

    <FieldOrder(10), FieldTitle("Internet")>
    Public Internet As Integer

    <FieldOrder(11), FieldTitle("Acquisition")>
    Public Acquisition As Integer

    <FieldOrder(12), FieldTitle("Acquisition before service")>
    Public AcquisitionBeforeService As Integer

    <FieldOrder(13), FieldTitle("Acquired customer comes")>
    Public AcquiredCustomerComes As Integer

    <FieldOrder(14), FieldTitle("Spontaneous customer")>
    Public SpontaneousCustomer As Integer

    <FieldOrder(15), FieldTitle("FollowUp customer")>
    Public FollowUpCustomer As Integer

    <FieldOrder(16), FieldTitle("Sum personal contacts")>
    Public SumPersonalContacts As Integer

    <FieldOrder(17), FieldTitle("UV tests")>
    Public UVTests As Integer

    <FieldOrder(18), FieldTitle("Test drives")>
    Public TestDrives As Integer

    <FieldOrder(19), FieldTitle("Handover offers")>
    Public HandoverOffers As Integer

    <FieldOrder(20), FieldTitle("thereof financial offers")>
    Public ThereofFinancialOffers As Integer

    <FieldOrder(21), FieldTitle("Offers incl extended warranty")>
    Public OffersIncl_ExtendedWarranty As Integer

    <FieldOrder(22), FieldTitle("Followup offers")>
    Public FollowupOffers As Integer

    <FieldOrder(23), FieldTitle("Conclusion contracts NV")>
    Public ConclusionContractsNV As Integer

    <FieldOrder(24), FieldTitle("Conclusion contracts UV")>
    Public ConclusionContractsUV As Integer

    <FieldOrder(25), FieldTitle("Conversion rate"), FieldConverter(ConverterKind.Double, ",")>
    Public ConversionRate As Double

    <FieldOrder(26), FieldTitle("Conclusion financial contracts")>
    Public ConclusionFinancialContracts As Integer

    <FieldOrder(27), FieldTitle("Conclusion insurance contracts")>
    Public ConclusionInsuranceContracts As Integer

    <FieldOrder(28), FieldTitle("Conclusions  extd warranty")>
    Public ConclusionsExtdWarranty As Integer

    <FieldOrder(29), FieldTitle("Deliveries")>
    Public Deliveries As Integer

    <FieldOrder(30), FieldTitle("Follup deliveries")>
    Public FollupDeliveries As Integer

    <FieldOrder(31), FieldTitle("Ongoing servicing")>
    Public OngoingServicing As Integer

    <FieldOrder(32), FieldTitle("Working days")>
    Public WorkingDays As Integer

End Class
