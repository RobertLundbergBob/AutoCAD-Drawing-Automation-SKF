Imports System.IO
Imports System.Text.RegularExpressions
Imports Newtonsoft.Json


Public Class IDCodeParser
    '--- 1. Config classes ---
    Public Class RootConfig
        Public Property components As List(Of ComponentDef)
    End Class

    Public Class ComponentDef
        Public Property type As String
        Public Property pattern As String
        Public Property blockName As String
        Public Property features As List(Of FeatureMap)
        Public Property binaryFlags As BinaryFlagsConfig
    End Class

    Public Class FeatureMap
        Public Property blockName As String
        Public Property value As String        ' optional
        Public Property valueFromGroup As String    'Optional
        Public Property defaultValue As String      'optional
        Public Property multiple As Boolean 'optional
        Public Property map As Dictionary(Of String, Object) 'optional
        Public Property PositionNumber As String
    End Class

    Public Class BinaryFlagsConfig
        Public Property groupName As String
        Public Property AttributeName As String
        Public Property onValue As String
    End Class

    '--- 2. Load JSON ---
    Public Shared Function LoadConfig(path As String) As RootConfig
        Return JsonConvert.DeserializeObject(Of RootConfig)(File.ReadAllText(path))
    End Function

    '--- 3. Parse code generically ---
    Public Shared Function ParseCode(code As String, cfg As RootConfig) _
            As List(Of Tuple(Of String, Object))

        Dim out As New List(Of Tuple(Of String, Object))()

        For Each comp In cfg.components
            Dim m = Regex.Match(code, comp.pattern, RegexOptions.IgnoreCase)
            If Not m.Success Then Continue For

            ' map of groupName → mappedValue
            Dim groupValues As New Dictionary(Of String, Object)()   ' ← NEW

            ' 3.1 → Parent block only
            out.Add(Tuple.Create(comp.blockName, CType(Nothing, Object)))

            For Each feat In comp.features
                Dim grp = m.Groups(feat.blockName)
                Dim mapLookupKey As String = If(grp.Success, grp.Value, "")


                Dim valObj As Object = Nothing

                ' A) valueFromGroup takes precedence if specified
                If Not String.IsNullOrEmpty(feat.valueFromGroup) Then
                    Dim valueSourceGroup = m.Groups(feat.valueFromGroup)
                    If valueSourceGroup.Success Then
                        valObj = valueSourceGroup.Value
                    ElseIf feat.defaultValue IsNot Nothing Then ' Fallback for valueFromGroup if its source group fails
                        valObj = feat.defaultValue
                    End If

                    ' B) Multiple-output feature (e.g., screws)
                ElseIf feat.multiple AndAlso feat.map IsNot Nothing Then
                    Dim caps = m.Groups(feat.blockName).Captures
                    Dim items As New List(Of String)
                    If caps.Count > 1 Then
                        ' e.g. connectors/accessories: each Capture is one char
                        For Each c As Capture In caps
                            items.Add(c.Value)
                        Next
                    ElseIf caps.Count = 1 AndAlso m.Groups(feat.blockName).Value.Length > 1 Then
                        ' e.g. pump element "600" or "606": split the single capture into chars
                        For Each ch As Char In m.Groups(feat.blockName).Value
                            items.Add(ch.ToString())
                        Next
                    End If
                    For i As Integer = 0 To items.Count - 1
                        Dim letter = items(i)
                        If feat.map.ContainsKey(letter) Then
                            Dim mapped = feat.map(letter)
                            ' record first one for placeholders
                            If Not groupValues.ContainsKey(feat.blockName) Then
                                groupValues(feat.blockName) = mapped
                            End If
                            Dim targetName = feat.blockName & (i + 1).ToString()
                            out.Add(Tuple.Create(targetName, mapped))
                        End If
                    Next
                    Continue For
                    ' C) Map-based single feature , 
                ElseIf feat.map IsNot Nothing Then
                    If feat.map.ContainsKey(mapLookupKey) Then
                        valObj = feat.map(mapLookupKey)
                    ElseIf feat.defaultValue IsNot Nothing Then
                        ' Fallback if mapLookupKey (e.g. a rare captured value) isn't in map,
                        ' and map doesn't contain "" for a general default.
                        valObj = feat.defaultValue
                    End If

                    ' D) Static value or defaultValue (if no map and no valueFromGroup)
                Else
                    If grp.Success AndAlso feat.value IsNot Nothing Then
                        valObj = feat.value
                    ElseIf (Not grp.Success OrElse String.IsNullOrEmpty(grp.Value)) AndAlso feat.defaultValue IsNot Nothing Then
                        valObj = feat.defaultValue
                    End If
                End If
                ' --- END OF REVISED LOGIC ORDER ---


                ' Record the mapped value for this group (if not multiple and valObj determined)
                ' for placeholder replacement.
                If valObj IsNot Nothing AndAlso Not feat.multiple AndAlso Not groupValues.ContainsKey(feat.blockName) Then
                    groupValues(feat.blockName) = valObj
                End If

                If valObj Is Nothing Then Continue For ' If no value could be determined, skip output for this feature

                ' Check if this feature has a PositionNumber → special case
                If Not String.IsNullOrEmpty(feat.PositionNumber) AndAlso groupValues.ContainsKey(feat.PositionNumber) Then
                    ' Get the monitor value (e.g., "P2 Piston...") and the mapped position (e.g., "4")
                    Dim monitorVal = groupValues(feat.PositionNumber)
                    Dim posVal = valObj.ToString()
                    Dim compositeKey = feat.PositionNumber & posVal ' e.g., "VPK_Monitoring4"
                    'REMOVE the previous plain entry (e.g., VPK_Monitoring)
                    out.RemoveAll(Function(t) String.Equals(t.Item1, feat.PositionNumber, StringComparison.OrdinalIgnoreCase))
                    ' Add only the position-qualified version
                    out.Add(Tuple.Create(compositeKey, monitorVal))
                Else
                    ' Generic placeholder replacement for standard features
                    Dim currentBlockName As String = feat.blockName
                    For Each mph As Match In Regex.Matches(currentBlockName, "\{(\w+)\}")
                        Dim keyName = mph.Groups(1).Value
                        If groupValues.ContainsKey(keyName) Then
                            currentBlockName = currentBlockName.Replace(
                "{" & keyName & "}",
                groupValues(keyName).ToString()
            )
                        End If
                    Next

                    ' Add normal block name → value tuple
                    out.Add(Tuple.Create(currentBlockName, valObj))
                End If
            Next
            ' Found a matching component, no need to check others
            Return out
        Next

        Return out ' Return empty list if no component pattern matched
    End Function




End Class
