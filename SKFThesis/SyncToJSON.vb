Imports System.Data.SQLite
Imports System.IO
Imports System.Windows.Forms
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq


Public Class SyncToJSON
    Public Class VisibilityEntry
        Public Property ChildBlockName As String
        Public Property VisibilityState As String
        Public Property VisibilityKey As String ' Optional: for reverse mapping
    End Class

    Public Shared Sub SyncParserJSON(jsonpath As String)
        Dim filepath As String
        Dim DrawingName As String
        Dim openFile As New OpenFileDialog()
        openFile.Title = "Select an AutoCAD Drawing"
        openFile.Filter = "AutoCAD Files (*.dwg)|*.dwg|All Files (*.*)|*.*"

        If openFile.ShowDialog() = DialogResult.OK Then
            filepath = openFile.FileName
            DrawingName = Path.GetFileNameWithoutExtension(filepath)
        Else
            Return
        End If

        Dim parentBlockName As String = DrawingName & "_Block"

        ' Load or initialize JSON
        Dim root As JObject
        If Not File.Exists(jsonpath) OrElse String.IsNullOrWhiteSpace(File.ReadAllText(jsonpath)) Then
            root = New JObject(New JProperty("components", New JArray()))
        Else
            Try
                root = JObject.Parse(File.ReadAllText(jsonpath))
            Catch ex As JsonReaderException
                root = New JObject(New JProperty("components", New JArray()))
            End Try
        End If

        ' Ensure components array exists
        Dim components As JArray = TryCast(root("components"), JArray)
        If components Is Nothing Then
            components = New JArray()
            root("components") = components
        End If

        ' Build lookup
        Dim compLookup As New Dictionary(Of String, JObject)(StringComparer.OrdinalIgnoreCase)
        For Each comp As JObject In components
            Dim name As String = comp("blockName")?.ToString()
            If Not String.IsNullOrEmpty(name) Then
                compLookup(name) = comp
            End If
        Next

        ' Ensure parent component entry exists
        Dim compObj As JObject = Nothing
        If Not compLookup.TryGetValue(parentBlockName, compObj) Then
            compObj = New JObject(
            New JProperty("type", DrawingName),
            New JProperty("pattern", ""),
            New JProperty("blockName", parentBlockName),
            New JProperty("features", New JArray())
        )
            components.Add(compObj)
            compLookup(parentBlockName) = compObj
        End If

        Dim features As JArray = CType(compObj("features"), JArray)

        ' Loop through each child and fetch visibility states
        Dim childBlocks As List(Of String) = DatabaseFunctions.GetChildBlockDefinitions(parentBlockName)

        For Each childBlock In childBlocks
            Dim visStates As List(Of String) = DatabaseFunctions.GetChildVisibilityStatesFromDatabase(childBlock)
            If visStates Is Nothing OrElse visStates.Count = 0 Then Continue For

            ' Find or create feature for this child
            Dim childFeat As JObject = features _
            .OfType(Of JObject)() _
            .FirstOrDefault(Function(f) f("blockName")?.ToString() = childBlock)

            If childFeat Is Nothing Then
                childFeat = New JObject(
                New JProperty("blockName", childBlock),
                New JProperty("map", New JObject())
            )
                features.Add(childFeat)
            End If

            Dim mapObj As JObject = CType(childFeat("map"), JObject)

            ' Add all visibility states with keys = "", or you can use the state itself as the key
            For Each visState In visStates
                ' Optional: Use visibility state as key, or "" if you want reverse mapping
                If Not mapObj.Properties().Any(Function(p) p.Value.ToString() = visState) Then
                    mapObj.Add(visState, visState)
                End If
            Next
        Next

        ' Save updated JSON
        File.WriteAllText(jsonpath, JsonConvert.SerializeObject(root, Formatting.Indented))
        Console.WriteLine($"Updated '{jsonpath}' with visibility states for '{parentBlockName}'.")
    End Sub

End Class
