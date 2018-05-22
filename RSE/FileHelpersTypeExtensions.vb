Imports System.Reflection
Imports FileHelpers

Namespace FileHelpers
    ' Thanks to Richard Dingwall @ https://gist.github.com/rdingwall/1391429

    Public Class FieldTitleAttribute
        Inherits Attribute

        Public Sub New(_name As String)
            If _name Is Nothing Then Throw New ArgumentNullException("name")
            Name = _name
        End Sub

        Public Property Name As String

    End Class


    Public Module FileHelpersTypeExtensions
        <Runtime.CompilerServices.Extension>
        Public Function GetFieldTitles(type As Type) As IEnumerable(Of String)
            Dim fields = From field In type.GetFields(
                BindingFlags.GetField Or
                BindingFlags.Public Or
                BindingFlags.NonPublic Or
                BindingFlags.Instance)
                         Where field.IsFileHelpersField()
                         Select field

            Return From field In fields
                   Let attrs = field.GetCustomAttributes(True)
                   Let order = attrs.OfType(Of FieldOrderAttribute)().Single().Order
                   Let title = attrs.OfType(Of FieldTitleAttribute)().Single().Name
                   Order By order
                   Select title
        End Function

        <Runtime.CompilerServices.Extension>
        Public Function GetCsvHeader(type As Type) As String
            Return String.Join(Type.Delimiter, type.GetFieldTitles())
        End Function

        <Runtime.CompilerServices.Extension>
        Private Function IsFileHelpersField(field As FieldInfo) As Boolean
            Return field.GetCustomAttributes(True).OfType(Of FieldOrderAttribute)().Any()
        End Function

    End Module
End Namespace
