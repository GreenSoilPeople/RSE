Imports FH = FileHelpers
Imports RSE.FileHelpers
Imports System.Reflection
Imports System.Reflection.Emit

Module Module1

    Private Options As Options
    Private records As IEnumerable(Of RSERecord)
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

        Dim engine As New FH.FileHelperEngine(Of RSERecord)
        records = engine.ReadFile(Options.InputFile)

        week = Right(records.OrderByDescending(Function(r) r.Week).First.Week, 2)

        header = GetType(RSERecord).GetCsvHeader

        Console.WriteLine($"Processign file {Options.InputFile}")

        If ParseFile() Then
            If Options.Rename Then
                FileIO.FileSystem.RenameFile(Options.InputFile, $"{IO.Path.GetFileName(Options.InputFile.Replace(IO.Path.GetExtension(Options.InputFile), ""))}_{week}_procesat{IO.Path.GetExtension(Options.InputFile)}")
            End If
            Console.WriteLine("Done!")
        End If

        Console.ReadLine()
    End Sub

    Private Function ParseFile() As Boolean

        For Each brand As KeyValuePair(Of String, String) In fileNameBrand

            Console.Write($"Exporting brand {brand.Value}...")
            Dim filtered = From record As RSERecord In records
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

            Dim provider As New FH.ExcelNPOIStorage.ExcelNPOIStorage(GetType(RSERecord), exportPath, 0, 0)

            provider.ColumnsHeaders.AddRange(header.Split(records(0).GetType.Delimiter))

            provider.InsertRecords(filtered.ToArray)
            Console.Write($"OK{vbNewLine}")

        Next
        Return True

    End Function

    Private Function AutoGenerateClass() As Type

    End Function

    Public Function CreateClass(ByVal className As String, ByVal properties As Dictionary(Of String, Type)) As Type

        Dim myDomain As AppDomain = AppDomain.CurrentDomain
        Dim myAsmName As New AssemblyName("MyAssembly")
        Dim myAssembly As AssemblyBuilder = myDomain.DefineDynamicAssembly(myAsmName, AssemblyBuilderAccess.Run)

        Dim myModule As ModuleBuilder = myAssembly.DefineDynamicModule("MyModule")

        Dim myType As TypeBuilder = myModule.DefineType(className, TypeAttributes.Public)

        myType.DefineDefaultConstructor(MethodAttributes.Public)

        For Each o In properties

            Dim prop As PropertyBuilder = myType.DefineProperty(o.Key, PropertyAttributes.HasDefault, o.Value, Nothing)

            Dim field As FieldBuilder = myType.DefineField("_" + o.Key, o.Value, FieldAttributes.[Private])

            Dim getter As MethodBuilder = myType.DefineMethod("get_" + o.Key, MethodAttributes.[Public] Or MethodAttributes.SpecialName Or MethodAttributes.HideBySig, o.Value, Type.EmptyTypes)
            Dim getterIL As ILGenerator = getter.GetILGenerator()
            getterIL.Emit(OpCodes.Ldarg_0)
            getterIL.Emit(OpCodes.Ldfld, field)
            getterIL.Emit(OpCodes.Ret)

            Dim setter As MethodBuilder = myType.DefineMethod("set_" + o.Key, MethodAttributes.[Public] Or MethodAttributes.SpecialName Or MethodAttributes.HideBySig, Nothing, New Type() {o.Value})
            Dim setterIL As ILGenerator = setter.GetILGenerator()
            setterIL.Emit(OpCodes.Ldarg_0)
            setterIL.Emit(OpCodes.Ldarg_1)
            setterIL.Emit(OpCodes.Stfld, field)
            setterIL.Emit(OpCodes.Ret)

            prop.SetGetMethod(getter)
            prop.SetSetMethod(setter)

        Next

        Return myType.CreateType()

    End Function
End Module
