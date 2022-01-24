Imports System
Imports System.CodeDom.Compiler
Imports System.Collections.Generic
Imports System.IO
Imports System.Reflection
Imports System.Text
Imports System.Text.RegularExpressions

Namespace SpreadsheetDocServerAPIPart2

    Public MustInherit Class ExampleCodeEvaluator

        Protected MustOverride ReadOnly Property CodeStart As String

        Protected MustOverride ReadOnly Property CodeEnd As String

        Protected MustOverride Function GetCodeDomProvider() As CodeDomProvider

        Protected MustOverride Function GetModuleAssembly() As String

        Protected MustOverride Function GetExampleClassName() As String

        Public Function ExcecuteCodeAndGenerateDocument(ByVal args As SpreadsheetDocServerAPIPart2.CodeEvaluationEventArgs) As Boolean
            Dim theCode As String = System.[String].Concat(Me.CodeStart, args.Code, Me.CodeEnd)
            Dim linesOfCode As String() = New String() {theCode}
            Return Me.CompileAndRun(linesOfCode, args.EvaluationParameter)
        End Function

        Protected Friend Function CompileAndRun(ByVal linesOfCode As String(), ByVal evaluationParameter As Object) As Boolean
            Dim CompilerParams As System.CodeDom.Compiler.CompilerParameters = New System.CodeDom.Compiler.CompilerParameters()
            CompilerParams.GenerateInMemory = True
            CompilerParams.TreatWarningsAsErrors = False
            CompilerParams.GenerateExecutable = False
            Dim referencesSystem As String() = New String() {"System.dll", "System.Windows.Forms.dll", "System.Data.dll", "System.Xml.dll", "System.Drawing.dll"}
            Dim referencesDX As String() = New String() {AssemblyInfo.SRAssemblyData, Me.GetModuleAssembly(), AssemblyInfo.SRAssemblyOfficeCore, AssemblyInfo.SRAssemblyPrintingCore, AssemblyInfo.SRAssemblyPrinting, AssemblyInfo.SRAssemblyDocs, AssemblyInfo.SRAssemblyUtils}
            Dim references As String() = New String(referencesSystem.Length + referencesDX.Length - 1) {}
            For referenceIndex As Integer = 0 To referencesSystem.Length - 1
                references(referenceIndex) = referencesSystem(referenceIndex)
            Next

            Dim i As Integer = 0, initial As Integer = referencesSystem.Length
            While i < referencesDX.Length
                Dim assembly As System.Reflection.Assembly = System.Reflection.Assembly.Load(referencesDX(i) & AssemblyInfo.FullAssemblyVersionExtension)
                If assembly IsNot Nothing Then references(i + initial) = assembly.Location
                i += 1
            End While

            CompilerParams.ReferencedAssemblies.AddRange(references)
            Dim provider As System.CodeDom.Compiler.CodeDomProvider = Me.GetCodeDomProvider()
            Dim compile As System.CodeDom.Compiler.CompilerResults = provider.CompileAssemblyFromSource(CompilerParams, linesOfCode)
            If compile.Errors.HasErrors Then
                'string text = "Compile error: ";
                'foreach(CompilerError ce in compile.Errors) {
                '    text += "rn" + ce.ToString();
                '}
                'MessageBox.Show(text);
                Return False
            End If

            Dim [module] As System.Reflection.[Module] = Nothing
            Try
                [module] = compile.CompiledAssembly.GetModules()(0)
            Catch
            End Try

            Dim moduleType As System.Type = Nothing
            If [module] Is Nothing Then
                Return False
            End If

            moduleType = [module].[GetType](Me.GetExampleClassName())
            Dim methInfo As System.Reflection.MethodInfo = Nothing
            If moduleType Is Nothing Then
                Return False
            End If

            methInfo = moduleType.GetMethod("Process")
            If methInfo IsNot Nothing Then
                Try
                    methInfo.Invoke(Nothing, New Object() {evaluationParameter})
                Catch __unusedException1__ As System.Exception
                    Return False ' an error in Spreadsheet
                End Try

                Return True
            End If

            Return False
        End Function
    End Class

    Public Class CodeExampleGroup

        Public Sub New()
        End Sub

        Public Property Name As String

        Public Property Examples As List(Of SpreadsheetDocServerAPIPart2.CodeExample)

        Public Property Id As Integer

        Friend Sub Merge(ByVal n As SpreadsheetDocServerAPIPart2.CodeExampleGroup)
            Me.Examples.AddRange(n.Examples)
        End Sub
    End Class

    Public Class CodeExample

        Public Property CodeCS As String

        Public Property CodeVB As String

        Public Property RegionName As String

        Public Property HumanReadableGroupName As String

        Public Property ExampleGroup As String

        Public Property Id As Integer

        Friend Sub Merge(ByVal n As SpreadsheetDocServerAPIPart2.CodeExample)
            Me.CodeCS += System.Environment.NewLine
            Me.CodeCS += n.CodeCS
            Me.CodeVB += System.Environment.NewLine
            Me.CodeVB += n.CodeVB
        End Sub
    End Class

    Public Enum ExampleLanguage
        Csharp = 0
        VB = 1
    End Enum

#Region "CodeExampleDemoUtils"
    Public Module CodeExampleDemoUtils

        Public Function GatherExamplesFromProject(ByVal examplesPath As String, ByVal language As SpreadsheetDocServerAPIPart2.ExampleLanguage) As Dictionary(Of String, System.IO.FileInfo)
            Dim result As System.Collections.Generic.Dictionary(Of String, System.IO.FileInfo) = New System.Collections.Generic.Dictionary(Of String, System.IO.FileInfo)()
            For Each fileName As String In System.IO.Directory.GetFiles(examplesPath, "*" & SpreadsheetDocServerAPIPart2.CodeExampleDemoUtils.GetCodeExampleFileExtension(language))
                result.Add(System.IO.Path.GetFileNameWithoutExtension(fileName), New System.IO.FileInfo(fileName))
            Next

            Return result
        End Function

        Public Function GetCodeExampleFileExtension(ByVal language As SpreadsheetDocServerAPIPart2.ExampleLanguage) As String
            If language = SpreadsheetDocServerAPIPart2.ExampleLanguage.VB Then Return ".vb"
            Return ".cs"
        End Function

        Public Function DeleteLeadingWhiteSpaces(ByVal lines As String(), ByVal stringToDelete As System.[String]) As String()
            Dim result As String() = New String(lines.Length - 1) {}
            Dim stringToDeleteLength As Integer = stringToDelete.Length
            For i As Integer = 0 To lines.Length - 1
                Dim index As Integer = lines(CInt((i))).IndexOf(stringToDelete)
                result(i) = If((index >= 0), lines(CInt((i))).Substring(index + stringToDeleteLength), lines(i))
            Next

            Return result
        End Function

        Public Function ConvertStringToMoreHumanReadableForm(ByVal exampleName As String) As String
            Dim result As String = SpreadsheetDocServerAPIPart2.CodeExampleDemoUtils.SplitCamelCase(exampleName)
            result = result.Replace(" In ", " in ")
            result = result.Replace(" And ", " and ")
            result = result.Replace(" To ", " to ")
            result = result.Replace(" From ", " from ")
            result = result.Replace(" With ", " with ")
            result = result.Replace(" By ", " by ")
            Return result
        End Function

        Private Function SplitCamelCase(ByVal exampleName As String) As String
            Dim length As Integer = exampleName.Length
            If length = 1 Then Return exampleName
            Dim result As System.Text.StringBuilder = New System.Text.StringBuilder(length * 2)
            For position As Integer = 0 To length - 1 - 1
                Dim current As Char = exampleName(position)
                Dim [next] As Char = exampleName(position + 1)
                result.Append(current)
                If Char.IsLower(current) AndAlso Char.IsUpper([next]) Then
                    result.Append(" "c)
                End If
            Next

            result.Append(exampleName(length - 1))
            Return result.ToString()
        End Function

        Public Function GetExamplePath(ByVal exampleFolderName As String) As String '"CodeExamples"
            Dim examplesPath2 As String = System.IO.Path.Combine(System.IO.Directory.GetCurrentDirectory() & "\..\..\", exampleFolderName)
            If System.IO.Directory.Exists(examplesPath2) Then Return examplesPath2
            Dim examplesPathInInsallation As String = SpreadsheetDocServerAPIPart2.CodeExampleDemoUtils.GetRelativeDirectoryPath(exampleFolderName)
            Return examplesPathInInsallation
        End Function

        'public static string GetExamplePath() {
        '    string examplesPath2 = Path.Combine(Directory.GetCurrentDirectory() + "\\..\\..\\", "CodeExamples");
        '    if (Directory.Exists(examplesPath2))
        '        return examplesPath2;
        '    string examplesPathInInsallation = GetRelativeDirectoryPath("CodeExamples");
        '    return examplesPathInInsallation;
        '}
        Public Function GetRelativeDirectoryPath(ByVal name As String) As String
            name = "Data\" & name
            Dim path As String = System.Windows.Forms.Application.StartupPath
            Dim s As String = "\"
            For i As Integer = 0 To 10
                If System.IO.Directory.Exists(path & s & name) Then
                    Return(path & s & name)
                Else
                    s += "..\"
                End If
            Next

            Return ""
        End Function

        Public Function FindExamples(ByVal examplePath As String, ByVal examplesCS As System.Collections.Generic.Dictionary(Of String, System.IO.FileInfo), ByVal examplesVB As System.Collections.Generic.Dictionary(Of String, System.IO.FileInfo)) As List(Of SpreadsheetDocServerAPIPart2.CodeExampleGroup)
            Dim result As System.Collections.Generic.List(Of SpreadsheetDocServerAPIPart2.CodeExampleGroup) = New System.Collections.Generic.List(Of SpreadsheetDocServerAPIPart2.CodeExampleGroup)()
            Dim current As System.Collections.Generic.Dictionary(Of String, System.IO.FileInfo) = Nothing
            Dim csExampleFinder As SpreadsheetDocServerAPIPart2.ExampleFinder
            Dim vbExampleFinder As SpreadsheetDocServerAPIPart2.ExampleFinder
            If examplesCS.Count = 0 Then
                current = examplesVB
                csExampleFinder = Nothing
                vbExampleFinder = New SpreadsheetDocServerAPIPart2.ExampleFinderVB()
            ElseIf examplesVB.Count = 0 Then
                current = examplesCS
                csExampleFinder = New SpreadsheetDocServerAPIPart2.ExampleFinderCSharp()
                vbExampleFinder = Nothing
            Else
                current = examplesCS
                csExampleFinder = New SpreadsheetDocServerAPIPart2.ExampleFinderCSharp()
                vbExampleFinder = New SpreadsheetDocServerAPIPart2.ExampleFinderVB()
            End If

            For Each sourceCodeItem As System.Collections.Generic.KeyValuePair(Of String, System.IO.FileInfo) In current
                Dim key As String = sourceCodeItem.Key
                Dim foundExamplesCS As System.Collections.Generic.List(Of SpreadsheetDocServerAPIPart2.CodeExample) = New System.Collections.Generic.List(Of SpreadsheetDocServerAPIPart2.CodeExample)()
                If csExampleFinder IsNot Nothing Then foundExamplesCS = csExampleFinder.Process(examplesCS(key))
                Dim foundExamplesVB As System.Collections.Generic.List(Of SpreadsheetDocServerAPIPart2.CodeExample) = New System.Collections.Generic.List(Of SpreadsheetDocServerAPIPart2.CodeExample)()
                If vbExampleFinder IsNot Nothing Then foundExamplesVB = vbExampleFinder.Process(examplesVB(key))
                Dim mergedExamples As System.Collections.Generic.List(Of SpreadsheetDocServerAPIPart2.CodeExample) = New System.Collections.Generic.List(Of SpreadsheetDocServerAPIPart2.CodeExample)()
                If foundExamplesCS.Count <> 0 AndAlso foundExamplesVB.Count = 0 Then
                    mergedExamples = foundExamplesCS
                ElseIf foundExamplesCS.Count = 0 AndAlso foundExamplesVB.Count <> 0 Then
                    mergedExamples = foundExamplesVB
                ElseIf(foundExamplesCS.Count = foundExamplesVB.Count) Then
                    mergedExamples = SpreadsheetDocServerAPIPart2.CodeExampleDemoUtils.MergeExamples(foundExamplesCS, foundExamplesVB)
                End If

                If mergedExamples.Count = 0 Then Continue For
                Dim group As SpreadsheetDocServerAPIPart2.CodeExampleGroup = New SpreadsheetDocServerAPIPart2.CodeExampleGroup() With {.Name = mergedExamples(CInt((0))).HumanReadableGroupName, .Examples = mergedExamples}
                result.Add(group)
            Next

            Return result
        End Function

        Private Function MergeExamples(ByVal foundExamplesCS As System.Collections.Generic.List(Of SpreadsheetDocServerAPIPart2.CodeExample), ByVal foundExamplesVB As System.Collections.Generic.List(Of SpreadsheetDocServerAPIPart2.CodeExample)) As List(Of SpreadsheetDocServerAPIPart2.CodeExample)
            Dim result As System.Collections.Generic.List(Of SpreadsheetDocServerAPIPart2.CodeExample) = New System.Collections.Generic.List(Of SpreadsheetDocServerAPIPart2.CodeExample)()
            Dim count As Integer = foundExamplesCS.Count
            For i As Integer = 0 To count - 1
                Dim itemCS As SpreadsheetDocServerAPIPart2.CodeExample = foundExamplesCS(i)
                Dim itemVB As SpreadsheetDocServerAPIPart2.CodeExample = foundExamplesVB(i)
                If Equals(itemCS.HumanReadableGroupName, itemVB.HumanReadableGroupName) AndAlso Equals(itemCS.RegionName, itemVB.RegionName) Then
                    Dim merged As SpreadsheetDocServerAPIPart2.CodeExample = New SpreadsheetDocServerAPIPart2.CodeExample()
                    merged.RegionName = itemCS.RegionName
                    merged.HumanReadableGroupName = itemCS.HumanReadableGroupName
                    merged.CodeCS = itemCS.CodeCS
                    merged.CodeVB = itemVB.CodeVB
                    result.Add(merged)
                Else
                    Throw New System.InvalidOperationException()
                End If
            Next

            Return result
        End Function

        Public Function DetectExampleLanguage(ByVal solutionFileNameWithoutExtenstion As String) As ExampleLanguage
            Dim projectPath As String = System.IO.Directory.GetCurrentDirectory() & "\..\..\"
            Dim csproject As String() = System.IO.Directory.GetFiles(projectPath, "*.csproj")
            If csproject.Length <> 0 AndAlso csproject(CInt((0))).EndsWith(solutionFileNameWithoutExtenstion & ".csproj") Then Return SpreadsheetDocServerAPIPart2.ExampleLanguage.Csharp
            Dim vbproject As String() = System.IO.Directory.GetFiles(projectPath, "*.vbproj")
            If vbproject.Length <> 0 AndAlso vbproject(CInt((0))).EndsWith(solutionFileNameWithoutExtenstion & ".vbproj") Then Return SpreadsheetDocServerAPIPart2.ExampleLanguage.VB
            Return SpreadsheetDocServerAPIPart2.ExampleLanguage.Csharp
        End Function
    End Module

#End Region
#Region "ExampleFinder"
    Public MustInherit Class ExampleFinder

        Public MustOverride ReadOnly Property RegexRegionPattern As String

        Public MustOverride ReadOnly Property RegionStarts As String

        Public Function Process(ByVal fileWithExample As System.IO.FileInfo) As List(Of SpreadsheetDocServerAPIPart2.CodeExample)
            If fileWithExample Is Nothing Then Return New System.Collections.Generic.List(Of SpreadsheetDocServerAPIPart2.CodeExample)()
            Dim groupName As String = System.IO.Path.GetFileNameWithoutExtension(CStr((fileWithExample.Name))).TrimEnd(New Char() {"1"c, "2"c, "3"c})
            Dim code As String
            Using stream As System.IO.FileStream = System.IO.File.Open(fileWithExample.FullName, System.IO.FileMode.Open, System.IO.FileAccess.Read)
                Dim sr As System.IO.StreamReader = New System.IO.StreamReader(stream)
                code = sr.ReadToEnd()
            End Using

            Dim foundExamples As System.Collections.Generic.List(Of SpreadsheetDocServerAPIPart2.CodeExample) = Me.ParseSourceFileAndFindRegionsWithExamples(groupName, code)
            Return foundExamples
        End Function

        ' todo: remove example group
        Public Function ParseSourceFileAndFindRegionsWithExamples(ByVal groupName As String, ByVal sourceCode As String) As List(Of SpreadsheetDocServerAPIPart2.CodeExample)
            Dim result As System.Collections.Generic.List(Of SpreadsheetDocServerAPIPart2.CodeExample) = New System.Collections.Generic.List(Of SpreadsheetDocServerAPIPart2.CodeExample)()
            Dim matches = System.Text.RegularExpressions.Regex.Matches(sourceCode, Me.RegexRegionPattern, System.Text.RegularExpressions.RegexOptions.Singleline)
            For Each match In matches
                Dim matchString As String = match.ToString()
                Dim splitter As String = If(matchString.IndexOf(Global.Microsoft.VisualBasic.Constants.vbCrLf) >= 0, Global.Microsoft.VisualBasic.Constants.vbCrLf, Global.Microsoft.VisualBasic.Constants.vbLf)
                Dim lines As String() = match.ToString().Split(New String() {splitter}, System.StringSplitOptions.None)
                If lines.Length <= 2 Then Continue For
                'string endRegion = lines[lines.Length - 1];
                lines = Me.DeleteLeadingWhiteSpacesFromSourceCode(lines)
                Dim regionName As String = System.[String].Empty
                Dim regionIsValid As Boolean = Me.ValidateRegionName(lines, regionName)
                If Not regionIsValid Then Continue For
                Dim exampleCode As String = String.Join(Global.Microsoft.VisualBasic.Constants.vbCrLf, lines, 1, lines.Length - 2)
                result.Add(Me.CreateSpreadsheetExample(groupName, regionName, exampleCode))
            Next

            Return result
        End Function

        Protected Function CreateSpreadsheetExample(ByVal exampleGroup As String, ByVal regionName As String, ByVal exampleCode As String) As CodeExample
            Dim result As SpreadsheetDocServerAPIPart2.CodeExample = New SpreadsheetDocServerAPIPart2.CodeExample()
            Me.SetExampleCode(exampleCode, result)
            result.RegionName = regionName
            result.HumanReadableGroupName = SpreadsheetDocServerAPIPart2.CodeExampleDemoUtils.ConvertStringToMoreHumanReadableForm(exampleGroup)
            Return result
        End Function

        Protected MustOverride Sub SetExampleCode(ByVal exampleCode As String, ByVal newExample As SpreadsheetDocServerAPIPart2.CodeExample)

        Protected Overridable Function DeleteLeadingWhiteSpacesFromSourceCode(ByVal lines As String()) As String()
            Return SpreadsheetDocServerAPIPart2.CodeExampleDemoUtils.DeleteLeadingWhiteSpaces(lines, "            ")
        End Function

        Protected Overridable Function ValidateRegionName(ByVal lines As String(), ByRef regionName As String) As Boolean
            Dim region As String = lines(0)
            Dim regionIndex As Integer = region.IndexOf(Me.RegionStarts)
            If regionIndex < 0 Then
                regionName = System.[String].Empty
                Return False
            End If

            Dim keepHashMark As Integer = 0 ' "#example" if value is -1 or "example" if value will be 0
            regionName = SpreadsheetDocServerAPIPart2.CodeExampleDemoUtils.ConvertStringToMoreHumanReadableForm(region.Substring(regionIndex + Me.RegionStarts.Length + keepHashMark))
            Return True
        End Function
    End Class

#End Region
#Region "ExampleFinderVB"
    Public Class ExampleFinderVB
        Inherits SpreadsheetDocServerAPIPart2.ExampleFinder

        'public ExampleFinderVB() {
        '}
        Public Overrides ReadOnly Property RegexRegionPattern As String
            Get
                Return "#Region.*?#End Region"
            End Get
        End Property

        Public Overrides ReadOnly Property RegionStarts As String
            Get
                Return "#Region ""#"
            End Get
        End Property

        Protected Overrides Function DeleteLeadingWhiteSpacesFromSourceCode(ByVal lines As String()) As String()
            Dim result As String() = MyBase.DeleteLeadingWhiteSpacesFromSourceCode(lines)
            Return SpreadsheetDocServerAPIPart2.CodeExampleDemoUtils.DeleteLeadingWhiteSpaces(result, Global.Microsoft.VisualBasic.Constants.vbTab & Global.Microsoft.VisualBasic.Constants.vbTab & Global.Microsoft.VisualBasic.Constants.vbTab)
        End Function

        Protected Overrides Function ValidateRegionName(ByVal lines As String(), ByRef regionName As String) As Boolean
            Dim result As Boolean = MyBase.ValidateRegionName(lines, regionName)
            If Not result Then Return result
            regionName = regionName.TrimEnd(""""c)
            Return True
        End Function

        Protected Overrides Sub SetExampleCode(ByVal code As String, ByVal newExample As SpreadsheetDocServerAPIPart2.CodeExample)
            newExample.CodeVB = code
        End Sub
    End Class

#End Region
#Region "ExampleFinderCSharp"
    Public Class ExampleFinderCSharp
        Inherits SpreadsheetDocServerAPIPart2.ExampleFinder

        Public Overrides ReadOnly Property RegexRegionPattern As String
            Get
                Return "#region.*?#endregion"
            End Get
        End Property

        Public Overrides ReadOnly Property RegionStarts As String
            Get
                Return "#region #"
            End Get
        End Property

        Protected Overrides Sub SetExampleCode(ByVal code As String, ByVal newExample As SpreadsheetDocServerAPIPart2.CodeExample)
            newExample.CodeCS = code
        End Sub
    End Class

#End Region
#Region "LeakSafeCompileEventRouter"
    Public Class LeakSafeCompileEventRouter

        Private ReadOnly weakControlRef As System.WeakReference

        Public Sub New(ByVal [module] As SpreadsheetDocServerAPIPart2.ExampleEvaluatorByTimer)
            'Guard.ArgumentNotNull(module, "module");
            Me.weakControlRef = New System.WeakReference([module])
        End Sub

        Public Sub OnCompileExampleTimerTick(ByVal sender As Object, ByVal e As System.EventArgs)
            Dim [module] As SpreadsheetDocServerAPIPart2.ExampleEvaluatorByTimer = CType(Me.weakControlRef.Target, SpreadsheetDocServerAPIPart2.ExampleEvaluatorByTimer)
            If [module] IsNot Nothing Then [module].CompileExample(sender, e)
        End Sub
    End Class

    Public Class CodeEvaluationEventArgs
        Inherits System.EventArgs

        Public Property Result As Boolean

        Public Property Code As String

        Public Property Language As ExampleLanguage

        Public Property EvaluationParameter As Object

        Public Property RegionName As String
    End Class

    Public Delegate Sub CodeEvaluationEventHandler(ByVal sender As Object, ByVal e As SpreadsheetDocServerAPIPart2.CodeEvaluationEventArgs)

    Public Class OnAfterCompileEventArgs
        Inherits System.EventArgs

        Public Property Result As Boolean
    End Class

    Public Delegate Sub OnAfterCompileEventHandler(ByVal sender As Object, ByVal e As SpreadsheetDocServerAPIPart2.OnAfterCompileEventArgs)

#End Region
    Public MustInherit Class ExampleEvaluatorByTimer
        Implements System.IDisposable

        Private leakSafeCompileEventRouter As SpreadsheetDocServerAPIPart2.LeakSafeCompileEventRouter

        Private compileExampleTimer As System.Windows.Forms.Timer

        Private compileComplete As Boolean = True

        Const CompileTimeIntervalInMilliseconds As Integer = 2000

        Public Sub New(ByVal enableTimer As Boolean)
            Me.leakSafeCompileEventRouter = New SpreadsheetDocServerAPIPart2.LeakSafeCompileEventRouter(Me)
            'this.compileExampleTimer = new System.Windows.Forms.Timer();
            If enableTimer Then
                Me.compileExampleTimer = New System.Windows.Forms.Timer()
                Me.compileExampleTimer.Interval = SpreadsheetDocServerAPIPart2.ExampleEvaluatorByTimer.CompileTimeIntervalInMilliseconds
                AddHandler Me.compileExampleTimer.Tick, New System.EventHandler(AddressOf Me.leakSafeCompileEventRouter.OnCompileExampleTimerTick) 'OnCompileTimerTick
                Me.compileExampleTimer.Enabled = True
            End If
        End Sub

        Public Sub New()
            Me.New(True)
        End Sub

#Region "Events"
        Public Event QueryEvaluate As SpreadsheetDocServerAPIPart2.CodeEvaluationEventHandler

        'public event CodeEvaluationEventHandler QueryEvaluateEvent {
        '    add { onQeuryEvaluate += value; }
        '    remove { onQeuryEvaluate -= value; }
        '}
        Protected Friend Overridable Function RaiseQueryEvaluate() As CodeEvaluationEventArgs
            If QueryEvaluateEvent IsNot Nothing Then
                Dim args As SpreadsheetDocServerAPIPart2.CodeEvaluationEventArgs = New SpreadsheetDocServerAPIPart2.CodeEvaluationEventArgs()
                RaiseEvent QueryEvaluate(Me, args)
                Return args
            End If

            Return Nothing
        End Function

        Public Event OnBeforeCompile As System.EventHandler

        'public event EventHandler OnBeforeCompileEvent { add { onBeforeCompile += value; } remove { onBeforeCompile -= value; } }
        Private Sub RaiseOnBeforeCompile()
            RaiseEvent OnBeforeCompile(Me, New System.EventArgs())
        End Sub

        Public Event OnAfterCompile As SpreadsheetDocServerAPIPart2.OnAfterCompileEventHandler

        'public event OnAfterCompileEventHandler OnAfterCompileEvent { add { onAfterCompile += value; } remove { onAfterCompile -= value; } }
        Private Sub RaiseOnAfterCompile(ByVal result As Boolean)
            RaiseEvent OnAfterCompile(Me, New SpreadsheetDocServerAPIPart2.OnAfterCompileEventArgs() With {.Result = result})
        End Sub

#End Region
        Public Sub CompileExample(ByVal sender As Object, ByVal e As System.EventArgs)
            If Not Me.compileComplete Then Return
            Dim args As SpreadsheetDocServerAPIPart2.CodeEvaluationEventArgs = Me.RaiseQueryEvaluate()
            If Not args.Result Then Return
            Me.ForceCompile(args)
        End Sub

        Public Sub ForceCompile(ByVal args As SpreadsheetDocServerAPIPart2.CodeEvaluationEventArgs)
            Me.compileComplete = False
            If Not System.[String].IsNullOrEmpty(args.Code) Then Me.CompileExampleAndShowPrintPreview(args)
            Me.compileComplete = True
        End Sub

        Private Sub CompileExampleAndShowPrintPreview(ByVal args As SpreadsheetDocServerAPIPart2.CodeEvaluationEventArgs)
            Dim evaluationSucceed As Boolean = False
            Try
                Me.RaiseOnBeforeCompile()
                evaluationSucceed = Me.Evaluate(args)
            Finally
                Me.RaiseOnAfterCompile(evaluationSucceed)
            End Try
        End Sub

        Public Function Evaluate(ByVal args As SpreadsheetDocServerAPIPart2.CodeEvaluationEventArgs) As Boolean
            Dim spreadsheetExampleCodeEvaluator As SpreadsheetDocServerAPIPart2.ExampleCodeEvaluator = Me.GetExampleCodeEvaluator(args.Language)
            Return spreadsheetExampleCodeEvaluator.ExcecuteCodeAndGenerateDocument(args)
        End Function

        Protected MustOverride Function GetExampleCodeEvaluator(ByVal language As SpreadsheetDocServerAPIPart2.ExampleLanguage) As ExampleCodeEvaluator

        Public Sub Dispose() Implements Global.System.IDisposable.Dispose
            If Me.compileExampleTimer IsNot Nothing Then
                Me.compileExampleTimer.Enabled = False
                If Me.leakSafeCompileEventRouter IsNot Nothing Then RemoveHandler Me.compileExampleTimer.Tick, New System.EventHandler(AddressOf Me.leakSafeCompileEventRouter.OnCompileExampleTimerTick) 'OnCompileTimerTick
                Me.compileExampleTimer.Dispose()
                Me.compileExampleTimer = Nothing
            End If
        End Sub
    End Class
End Namespace
