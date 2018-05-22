Imports System.Text
Imports CommandLine

Public Class Options
    <[Option]("i", "input", Required:=True, HelpText:="Input file to read.")>
    Public Property InputFile As String
    <[Option]("o", "output", Required:=False, HelpText:="Output directory.")>
    Public Property Output As String
    <[Option]("c", "clean", Required:=False, HelpText:="Clear contents of destination folder before extraction")>
    Public Property Clean As Boolean
    <[Option]("r", "rename", Required:=False, HelpText:="Rename source file")>
    Public Property Rename As Boolean

    <HelpOption>
    Public Function GetUsage() As String
        Dim usage As New StringBuilder

        usage.Append(Text.HelpText.AutoBuild(Me))

        Return usage.ToString
    End Function
End Class
