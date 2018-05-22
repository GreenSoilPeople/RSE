Imports FH = FileHelpers
Imports RSE.FileHelpers
Imports System.Reflection
Imports System.Reflection.Emit

Module Module1

    Private Options As Options
    Private fileNameBrand As Dictionary(Of String, String) = New Dictionary(Of String, String) From {{"A", "Audi"}, {"V", "VW PKW"}, {"C", "Skoda"}, {"S", "Seat"}, {"P", "Porsche"}, {"G", "Weltauto"}, {"L", "LNF"}}
    Private header As String
    Private week As String

    Sub Main()

        Options = New Options

        If CommandLine.Parser.Default.ParseArguments(CommandLineHelper.CreateArgs(Environment.CommandLine), Options) Then
            If Not IO.File.Exists(Options.InputFile) Then
                Console.WriteLine($"File not found: {Options.InputFile}")
                Console.ReadLine()
                Return
            End If
        Else
            Return
        End If


        If ParseFile() Then
            If Options.Rename Then
                FileIO.FileSystem.RenameFile(Options.InputFile, $"{IO.Path.GetFileName(Options.InputFile.Replace(IO.Path.GetExtension(Options.InputFile), ""))}_{week}_procesat{IO.Path.GetExtension(Options.InputFile)}")
            End If
            Console.WriteLine("Done!")
        End If

        Console.ReadLine()
    End Sub

    Private Sub Autodetect()
        Dim detector As New FH.Detection.SmartFormatDetector
        Dim formats = detector.DetectFileFormat("D:\RAPORT RSE\RSE_SUMMARY.csv")



        'For Each f In formats
        '    Console.WriteLine($"Format Detected, confidence: {f.Confidence}%")
        '    Dim delimited = f.ClassBuilderAsDelimited

        '    Console.WriteLine($"    Delimiter: {delimited.Delimiter}")
        '    Console.WriteLine($"    Fields:")

        '    For Each field In delimited.Fields
        '        Console.WriteLine($"        {field.FieldName}: {field.FieldType}")
        '    Next
        'Next
        Dim InputFile As String = "D:\RAPORT RSE\RSE_SUMMARY.csv"

        Dim exportPath As String = "D:\RAPORT RSE\RSE_SUMMARY_auto.xls"
        Dim delimited = formats(0).ClassBuilderAsDelimited

        Dim recClass = delimited.CreateRecordClass()



        Dim provider As New FH.ExcelNPOIStorage.ExcelNPOIStorage(recClass, exportPath, 0, 0)

        Dim engine As New FH.FileHelperEngine(recClass)


        Dim records = engine.ReadFile(InputFile)

        'provider.ColumnsHeaders.AddRange(header.Split(records(0).GetType.Delimiter))

        provider.InsertRecords(records.ToArray)

    End Sub

    Private Function ParseFile() As Boolean
        Dim detector As New FH.Detection.SmartFormatDetector
        Dim formats = detector.DetectFileFormat(Options.InputFile)

        Dim delimited = formats(0).ClassBuilderAsDelimited

        Dim recClass = delimited.CreateRecordClass()


        Dim engine As New FH.FileHelperEngine(recClass)
        Dim records = engine.ReadFile(Options.InputFile)

        week = Right(records.OrderByDescending(Function(r) r.Week).First.Week, 2)

        'header = recClass.GetCsvHeader
        header = New IO.StreamReader(Options.InputFile).ReadLine



        Console.WriteLine($"Processign file {Options.InputFile}")


        For Each brand As KeyValuePair(Of String, String) In fileNameBrand

            Console.Write($"Exporting brand {brand.Value}...")
            Dim filtered = From record In records
                           Where record.Brands.Contains(brand.Key)
                           Select record

            Dim exportDir As String = $"{Options.Output}\{brand.Value}"
            Dim exportPath As String = $"{Options.Output}\{brand.Value}\Raport_RSE_sapt_{week}.xls"

            If Not IO.Directory.Exists(exportDir) Then
                Try
                    IO.Directory.CreateDirectory(exportDir)
                Catch ex As Exception
                    Console.WriteLine($"Could not create output: {ex.Message}")
                    Console.ReadLine()
                    Return False
                End Try
            End If

            'Check if -c flag and clear output folder
            If Options.Clean Then
                For Each f As String In IO.Directory.GetFiles(exportDir)
                    IO.File.Delete(f)
                Next
            End If

            Dim provider As New FH.ExcelNPOIStorage.ExcelNPOIStorage(recClass, exportPath, 0, 0)

            provider.ColumnsHeaders.AddRange(header.Split(delimited.Delimiter))

            provider.InsertRecords(filtered.ToArray)
            Console.Write($"OK{vbNewLine}")

        Next
        Return True

    End Function

End Module
