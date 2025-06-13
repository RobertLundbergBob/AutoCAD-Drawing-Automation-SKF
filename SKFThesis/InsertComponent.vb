Imports System.Security.Cryptography
Imports System.Windows.Forms
Imports System.IO
Imports Newtonsoft.Json
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Internal.PreviousInput
Imports System.Drawing
Imports Autodesk.AutoCAD.DatabaseServices.Filters
Imports Autodesk.AutoCAD.Windows.Data
Imports System.Windows.Forms.VisualStyles
Imports System.ComponentModel
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar
Imports Newtonsoft.Json.Linq
Imports Autodesk.AutoCAD.GraphicsSystem
Imports System.Text.RegularExpressions

Public Class InsertComponent

    'Setup of windows forms and the UI elements within 
    Inherits Form
    Private IDtxtbox As New TextBox
    Private PosCheckbox As New CheckBox
    Private PosLabel As New Label
    Private IDButton As New Button
    Private CBTier1 As New ComboBox
    Private b_insert As New Button
    Private b_modify As New Button
    Private b_back As New Button
    Private xoffset As Integer = 200
    Private Column1_X As Integer = 80
    Private CBTier2 As New ComboBox
    Private CBTier3 As New ComboBox
    Private OutletCBVisibility As New ComboBox
    Private UIElementsTier1 As New List(Of Object)
    Private UIElementsTier2 As New List(Of Object)
    Private UIElementsTier3 As New List(Of Object)
    Private IDCodeGRP As New GroupBox()
    Private SelectComponentGRP As New GroupBox()
    Private CompParamGRP As New GroupBox()
    Private mainscrollpanel As New Panel()
    Private InsertGRP As New GroupBox()

    'Define folderpath to Design Library
    Dim folderPath As String = Functions.GetRad1Path() & "\"

    Public Sub New() ' Initialize Form
        Me.Text = "Insert Component"
        Me.Size = New Drawing.Size(1200, 1000)
        'creates a panel that scrolls if the UI is too big
        mainscrollpanel.AutoScroll = True
        mainscrollpanel.Dock = DockStyle.Fill
        mainscrollpanel.Name = "mainscrollpanel"
        Me.Controls.Add(mainscrollpanel)

        'UI GROUP FOR INSERT W/ ID CODE
        IDCodeGRP.Text = "Insert w/ IdCode"
        IDCodeGRP.Location = New Point(Column1_X, 20)
        IDCodeGRP.AutoSize = True
        IDCodeGRP.AutoSizeMode = AutoSizeMode.GrowAndShrink
        IDCodeGRP.Padding = New Padding(10)
        mainscrollpanel.Controls.Add(IDCodeGRP)
        ' Add ID Code textbox at given location and add Or label under it
        IDtxtbox.Location = New Drawing.Point(15, 45)
        IDtxtbox.Size = New Drawing.Size(200, 30)
        IDButton.Location = New Drawing.Point(IDtxtbox.Right + 20, 40)
        IDButton.Size = New Drawing.Size(120, 30)
        IDButton.Text = "Insert ID"
        IDButton.BackColor = Color.LimeGreen
        AddHandler IDButton.Click, AddressOf IDButton_Click
        IDCodeGRP.Controls.Add(IDButton)
        Dim lbltxt As New Windows.Forms.Label() With {.Text = "Enter ID Code", .Location = New Drawing.Point(IDtxtbox.Location.X, IDtxtbox.Location.Y - 25)}
        IDCodeGRP.Controls.Add(lbltxt)
        IDCodeGRP.Controls.Add(IDtxtbox)


        Dim ortxt As New Windows.Forms.Label() With {.Text = "Or", .Location = New Drawing.Point(IDCodeGRP.Left, IDCodeGRP.Bottom + 30)}
        mainscrollpanel.Controls.Add(ortxt)

        'UI GROUP FOR SELECTING COMPONENT TYPE
        SelectComponentGRP.Text = "Select Component"
        SelectComponentGRP.Location = New Point(Column1_X, IDCodeGRP.Bottom + 60)
        SelectComponentGRP.AutoSize = True
        SelectComponentGRP.AutoSizeMode = AutoSizeMode.GrowAndShrink
        SelectComponentGRP.Padding = New Padding(10)  ' leave some breathing room

        mainscrollpanel.Controls.Add(SelectComponentGRP)

        ' Configure Initial ComboBox
        CBTier1.Location = New Drawing.Point(15, 35)
        CBTier1.Size = New Drawing.Size(175, 30)
        SelectComponentGRP.Controls.Add(CBTier1)


        If Not Directory.Exists(folderPath) Then
            MessageBox.Show("The required folder path does not exist: " & folderPath, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Exit Sub
        End If
        Dim folders As String() = Directory.GetDirectories(folderPath)

        ' Add folder names to ComboBox
        For Each folder As String In folders
            CBTier1.Items.Add(Path.GetFileName(folder))
        Next
        CBTier1.Text = "Select Component Type"
        AddHandler CBTier1.SelectedIndexChanged, AddressOf CBTier1_SelectedIndexChanged

        'UI GROUP FOR INSERT/MODIFY and POSITION COMPONENT
        InsertGRP.Text = "Insert/Modify"
        InsertGRP.Location = New Point(950, 20)
        InsertGRP.AutoSize = True
        InsertGRP.AutoSizeMode = AutoSizeMode.GrowAndShrink
        InsertGRP.Padding = New Padding(10)  ' leave some breathing room
        mainscrollpanel.Controls.Add(InsertGRP)

        b_insert.Location = New Drawing.Point(15, 150)
        b_insert.Size = New Drawing.Size(150, 60)
        b_insert.Text = "Insert"
        b_insert.BackColor = Color.LimeGreen
        AddHandler b_insert.Click, AddressOf b_insert_Click

        PosCheckbox.Location = New Drawing.Point(b_insert.Left, b_insert.Top - 35)
        PosCheckbox.Size = New Drawing.Size(20, 20)
        PosLabel.Location = New Drawing.Point(PosCheckbox.Right + 40, PosCheckbox.Top)
        PosLabel.Size = New Drawing.Size(90, 30)
        PosLabel.Text = "Choose Insert Position"

        b_modify.Location = New Drawing.Point(b_insert.Location.X, b_insert.Location.Y - 120)
        b_modify.Size = New Drawing.Size(150, 60)
        b_modify.Text = "Modify"
        b_modify.BackColor = Color.White
        AddHandler b_modify.Click, AddressOf b_modify_click

        b_back.Location = New Drawing.Point(10, 10)
        b_back.Size = New Drawing.Size(50, 20)
        b_back.Text = "Back"
        b_back.BackColor = Color.Firebrick
        AddHandler b_back.Click, AddressOf b_back_click

        ' Add Controls to Form
        InsertGRP.Controls.Add(PosCheckbox)
        InsertGRP.Controls.Add(PosLabel)
        mainscrollpanel.Controls.Add(b_back)
        InsertGRP.Controls.Add(b_modify)
        InsertGRP.Controls.Add(b_insert)


        'LOAD METERING DEVICE JSON
        Dim Jsonpath = folderPath & "Metering Devices\MeteringDevices.JSON"
        LoadDeviceConfigs(Jsonpath)
        'LOAD GENERATE ID CODE JSON
        LoadVisibilityToArticleMap(folderPath)
    End Sub

    Public Shared visibilityMap As New Dictionary(Of String, Dictionary(Of String, Dictionary(Of String, DualMapEntry)))(StringComparer.OrdinalIgnoreCase)
    Public Shared SingleUseComponentsMap As New Dictionary(Of String, HashSet(Of String))
    Public Shared ComponentFormatMap As New Dictionary(Of String, String)()

    Public Shared Sub LoadVisibilityToArticleMap(folderpath As String)
        Dim jsonText = File.ReadAllText(folderpath & "GenerateID.json")
        Dim fullMap = JObject.Parse(jsonText)
        Dim components = fullMap("Components") _
                    .ToObject(Of Dictionary(Of String, JObject))()

        For Each compKey In components.Keys
            Dim compObj = components(compKey)
            visibilityMap(compKey) = New Dictionary(Of String, Dictionary(Of String, DualMapEntry))(StringComparer.OrdinalIgnoreCase)

            For Each paramProp In compObj.Properties()
                Dim layerKey = paramProp.Name
                Dim token = paramProp.Value

                If layerKey.Equals("Format", StringComparison.OrdinalIgnoreCase) Then
                    ComponentFormatMap(compKey) = token.ToString()
                    Continue For
                End If

                If layerKey.Equals("SingleUseComponents", StringComparison.OrdinalIgnoreCase) Then
                    Dim list = token.ToObject(Of List(Of String))()
                    SingleUseComponentsMap(compKey) = New HashSet(Of String)(list, StringComparer.OrdinalIgnoreCase)
                    Continue For
                End If

                ' Build a state→DualMapEntry map for this parameter
                Dim stateMap = New Dictionary(Of String, DualMapEntry)(StringComparer.OrdinalIgnoreCase)
                visibilityMap(compKey)(layerKey) = stateMap

                Select Case token.Type
                    Case JTokenType.Object
                        ' each child is either "state": string  OR  "state": { article:, codeValue: }
                        For Each st In CType(token, JObject).Properties()
                            Dim entry As DualMapEntry
                            If st.Value.Type = JTokenType.String Then
                                ' legacy single-string → treat as codeValue
                                entry = New DualMapEntry With {
                            .article = "",
                            .codeValue = st.Value.Value(Of String)()
                             }
                            Else
                                ' parse the object into our DualMapEntry
                                entry = st.Value.ToObject(Of DualMapEntry)()
                                ' ensure nulls become ""
                                entry.article = If(entry.article, "")
                                entry.codeValue = If(entry.codeValue, "")
                            End If
                            stateMap(st.Name) = entry
                        Next

                    Case JTokenType.String
                        'IF only a string the Value is treated as IDCode
                        Dim singlevalue As String = token.Value(Of String)()
                        stateMap("") = New DualMapEntry With {
                        .article = "",
                        .codeValue = singlevalue
                            }

                    Case Else
                        ' ignore arrays or other unexpected types
                End Select
            Next
        Next
    End Sub

    Private Sub b_back_click()
        'Open mainform
        Dim MainForm As New MainForm()
        MainForm.Show()

        'Close Insert Form
        Me.Close()

    End Sub

    Private Sub b_modify_click(sender As Object, e As EventArgs)
        'Open Modify Form
        Dim blockSelectorForm As New ModifyComponent()
        blockSelectorForm.Show()

        'Close Insert Form
        Me.Close()
    End Sub

    Private Sub IDButton_Click(sender As Object, e As EventArgs)

        Dim cfg = IDCodeParser.LoadConfig(folderPath & "InsertID.json")
        Dim code = IDtxtbox.Text
        Dim list = IDCodeParser.ParseCode(code, cfg)

        ' ── Guard: is the code valid? ─────────────────────────────
        If list Is Nothing OrElse list.Count = 0 Then
            MessageBox.Show(
            $"Invalid Identification Code: '{code}'",
            "Parse Error",
            MessageBoxButtons.OK,
            MessageBoxIcon.Warning
        )
            Exit Sub
        End If

        Dim cfgdevice As DeviceConfig = Nothing
        Dim ParentBlockName As String = DatabaseFunctions.GetMatchingParentBlock(code)
        Dim componenttype As String = DatabaseFunctions.GetParentAttribute(ParentBlockName, "COMPONENTTYPE")
        If componenttype IsNot Nothing AndAlso componenttype.IndexOf("Metering Device", StringComparison.OrdinalIgnoreCase) >= 0 Then
            If Not Nothing AndAlso Not _deviceConfigs.TryGetValue(ParentBlockName, cfgdevice) Then
                Throw New InvalidOperationException($"No configuration found for {ParentBlockName}")
            End If
        End If
        Dim Filename As String = ParentBlockName.Replace("_Block", "")
        Functions.InsertBlockWithParameters(folderPath, Filename, ParentBlockName, list, PosCheckbox, cfgdevice)
        t += 1
    End Sub

    Private Function CBTier1_SelectedIndexChanged(sender As Object, e As EventArgs)
        ClearFlipCheckBoxes(FlipCheckboxes, FlipLabels, CompParamGRP)
        Functions.ClearUIElementList(UIElementsTier1, SelectComponentGRP) ' Will dispose CBTier3 if it's in there
        Functions.ClearUIElementList(UIElementsTier2, CompParamGRP) ' Will dispose OutletCBVisibility, etc.
        Functions.ClearUIElementList(UIElementsTier3, CompParamGRP) ' Will dispose panel, checkboxall, etc.

        Dim selectedItem As String = CBTier1.SelectedItem.ToString()
        If CBTier2.Items.Count > 0 Then
            ' Clear existing items and text
            CBTier2.Items.Clear()
            CBTier2.Text = ""
        Else
            ' Create new ComboBox if it doesn't exist
            CBTier2 = New ComboBox()
            CBTier2.Location = New Drawing.Point(CBTier1.Location.X + xoffset, CBTier1.Location.Y)
            CBTier2.Size = New Drawing.Size(175, 30)
            SelectComponentGRP.Controls.Add(CBTier2)
        End If

        ' Update ComboBox items and the text based on the selected item
        UpdateComboBoxItems(CBTier2, selectedItem)
        Return selectedItem
    End Function

    'Updates the Tier2 Combobox based on the selected item from Tier1
    Private Sub UpdateComboBoxItems(comboBox As ComboBox, selectedItem As String)
        'Folder path to the DWG Files within the selected component type
        Dim folderPathtier1 As String = folderPath & selectedItem

        Dim folders As String() = Directory.GetDirectories(folderPathtier1)

        ' Add folder names to ComboBox
        For Each folder As String In folders
            comboBox.Items.Add(Path.GetFileName(folder))
        Next

        comboBox.Text = "Select " & selectedItem
        AddHandler comboBox.SelectedIndexChanged, AddressOf CBTier2_SelectedIndexChanged

    End Sub

    'Adds Tier3 combobox that adds DWG files as items
    Private Function CBTier2_SelectedIndexChanged(sender As Object, e As EventArgs)
        ClearFlipCheckBoxes(FlipCheckboxes, FlipLabels, CompParamGRP)
        Functions.ClearUIElementList(UIElementsTier1, SelectComponentGRP) ' Will dispose CBTier3 if it's in there
        Functions.ClearUIElementList(UIElementsTier2, CompParamGRP) ' Will dispose OutletCBVisibility, etc.
        Functions.ClearUIElementList(UIElementsTier3, CompParamGRP) ' Will dispose panel, checkboxall, etc.

        Dim selecteditem1 As String = CBTier1.SelectedItem.ToString()
        Dim selecteditem2 As String = CBTier2.SelectedItem.ToString()
        Dim folderpathtier2 As String = folderPath & selecteditem1 & "\" & selecteditem2

        If Directory.Exists(folderpathtier2) Then
            Dim dwgFiles As String() = Directory.GetFiles(folderpathtier2, "*.dwg")

            'If only one dwg file in the folder then create the new forms elements from the dwg file in the folder
            If dwgFiles.Length = 1 Then
                Dim filepath As String = dwgFiles(0)
                ProcessDwgFiles(filepath, selecteditem2)
                ' IF there are more dwg files then create a tier3 combobox where user chooses the dwg file to build the rest of the forms with 
                Functions.InsertHelperImage(folderpathtier2, selecteditem2, CompParamGRP, mainscrollpanel)
            Else
                If CBTier3.Items.Count > 0 Then
                    CBTier3.Items.Clear()
                    CBTier3.Text = ""
                Else
                    ' Create new ComboBox if it doesn't exist
                    CBTier3 = New ComboBox()
                    CBTier3.Location = New Drawing.Point(CBTier2.Location.X + xoffset, CBTier2.Location.Y)
                    CBTier3.Size = New Drawing.Size(150, 30)
                    SelectComponentGRP.Controls.Add(CBTier3)
                End If

                For Each file As String In dwgFiles
                    CBTier3.Items.Add(Path.GetFileNameWithoutExtension(file))
                Next
                CBTier3.Text = "Select" & selecteditem2
                UIElementsTier1.Add(CBTier3)
                AddHandler CBTier3.SelectedIndexChanged, AddressOf CBTier3_SelectedIndexChanged
            End If

        End If

    End Function

    Private Sub ProcessDwgFiles(filepath As String, selecteditem As String)
        Dim blockname As String = selecteditem & "_Block"
        Dim visibilitystates As List(Of String) = DatabaseFunctions.GetVisibilityStatesFromDatabase(blockname)

        Dim DoubleNestedAttribute As String = DatabaseFunctions.GetParentAttribute(selecteditem, "DOUBLENESTED")
        If DoubleNestedAttribute = "True" Then

        End If

        'Modify Component parameters
        CompParamGRP.Text = "Select Component Parameters"
        CompParamGRP.Location = New Point(Column1_X, SelectComponentGRP.Bottom + 40)
        CompParamGRP.AutoSize = True
        CompParamGRP.AutoSizeMode = AutoSizeMode.GrowAndShrink
        CompParamGRP.Padding = New Padding(10) ' leave some breathing room
        mainscrollpanel.Controls.Add(CompParamGRP)

        Dim basepointx As Integer = 15
        Dim basepointy As Integer = 25

        OutletCBVisibility = Functions.CreateComboBoxes(CompParamGRP, visibilitystates, blockname, basepointx, basepointy, UIElementsTier2, AddressOf OutletCBVisibilitys_SelectedIndexChanged, AddressOf ChildCB_VisibilityChanged)

    End Sub

    Private Function CBTier3_SelectedIndexChanged(sender As Object, e As EventArgs)
        ClearFlipCheckBoxes(FlipCheckboxes, FlipLabels, CompParamGRP)
        Functions.ClearUIElementList(UIElementsTier2, CompParamGRP)
        Functions.ClearUIElementList(UIElementsTier3, CompParamGRP)

        Dim selecteditem1 As String = CBTier1.SelectedItem.ToString()
        Dim selecteditem2 As String = CBTier2.SelectedItem.ToString()
        Dim selecteditem3 As String = CBTier3.SelectedItem.ToString()

        Dim filepath As String = folderPath & selecteditem1 & "\" & selecteditem2 & "\" & selecteditem3 & ".dwg"

        ProcessDwgFiles(filepath, selecteditem3)
    End Function

    Public Shared Sub ClearFlipCheckBoxes(FlipCheckBoxes As Dictionary(Of ComboBox, CheckBox), FlipLabels As Dictionary(Of ComboBox, Label), parent As Control)
        For Each kvp In FlipCheckBoxes
            parent.Controls.Remove(kvp.Value)
        Next
        For Each kvp In FlipLabels
            parent.Controls.Remove(kvp.Value)
        Next
        FlipCheckBoxes.Clear()
        FlipLabels.Clear()
    End Sub

    Private Shared FlipCheckboxes As New Dictionary(Of ComboBox, CheckBox)
    Private Shared FlipLabels As New Dictionary(Of ComboBox, Windows.Forms.Label)
    Public Shared ChildToDoubleBindingMap As New Dictionary(Of ComboBox, DoubleNestedBinding)


    Private Shared Sub ChildCB_VisibilityChanged(sender As Object, e As EventArgs)
        Dim cb As ComboBox = CType(sender, ComboBox)
        Dim selectedState As String = cb.SelectedItem?.ToString()

        If String.IsNullOrWhiteSpace(selectedState) Then Return
        Dim fulltag As String = cb.Tag.ToString()
        Dim visibilityMarkerIndex As Integer = fulltag.IndexOf("_Visibility_")
        Dim childBlockName As String = If(visibilityMarkerIndex >= 0, fulltag.Substring(0, visibilityMarkerIndex), fulltag)
        Dim ChildBlockID As Integer? = UploadToDatabase.GetChildBlockID(childBlockName)
        Dim ParameterID As Integer? = UploadToDatabase.GetParameterID(ChildBlockID, Nothing, Nothing, "Visibility")
        ' Get the ParameterValueID for the selected state
        Dim parameterValueID As Integer? = DatabaseFunctions.GetParameterValueID(ParameterID, selectedState)

        Functions.PlaceFlipCheckBox(parameterValueID, FlipCheckboxes, cb, FlipLabels)

        'FOR DOUBLENESTED BLOCKS 
        If ChildToDoubleBindingMap.ContainsKey(cb) Then
            Dim binding = ChildToDoubleBindingMap(cb)
            Dim doubleCombo = binding.DoubleCombo
            Dim candidates = binding.DoubleBlockCandidates

            ' Match the selected state to one of the double child block names
            selectedState = cb.SelectedItem.ToString()
            Dim matchedBlock As String = candidates.FirstOrDefault(Function(name) name.StartsWith(selectedState & "_", StringComparison.OrdinalIgnoreCase))

            If Not String.IsNullOrEmpty(matchedBlock) Then
                ' Get and load visibility states
                Dim visibilityStates As List(Of String) = DatabaseFunctions.GetDoubleChildVisibilityStatesFromDatabase(matchedBlock)
                doubleCombo.Items.Clear()
                For Each state In visibilityStates
                    doubleCombo.Items.Add(state)
                Next
                doubleCombo.SelectedIndex = -1
                doubleCombo.Text = "Select " & matchedBlock.Replace("_", " ")
                doubleCombo.Tag = matchedBlock & "_Visibility_" & "1"
            Else
                ' Optional: clear combo if no match found
                doubleCombo.Items.Clear()
                doubleCombo.Text = "(No Match)"
            End If
        End If

    End Sub



    Public meteringdevicepanel As Panel
    Public CheckBoxAll As Label
    Private Shared _deviceConfigs As Dictionary(Of String, DeviceConfig)
    Public Shared Sub LoadDeviceConfigs(jsonPath As String)
        Dim json As String = File.ReadAllText(jsonPath)
        _deviceConfigs = JsonConvert.DeserializeObject(Of Dictionary(Of String, DeviceConfig))(json)
    End Sub

    '------JSON METERING DEVICE PANEL BUILD------------
    Private Sub OutletCBVisibilitys_SelectedIndexChanged(sender As Object, e As EventArgs)
        Functions.ClearUIElementList(UIElementsTier3, CompParamGRP)

        Dim parentkey As String
        If CBTier3 IsNot Nothing AndAlso CBTier3.SelectedItem IsNot Nothing Then
            parentkey = CBTier3.SelectedItem.ToString() & "_Block"
        Else
            parentkey = CBTier2.SelectedItem.ToString() & "_Block"
        End If

        Dim cfg As DeviceConfig = Nothing
        If Not _deviceConfigs.TryGetValue(parentkey, cfg) Then
            MessageBox.Show("No configuration for " & parentkey)
            Return
        End If

        Dim outletCount As Integer
        If Not Integer.TryParse(OutletCBVisibility.SelectedItem?.ToString(), outletCount) Then
            MessageBox.Show("Invalid outlet count")
            Return
        End If

        CheckBoxAll = Functions.BuildMeteringDeviceControls(CompParamGRP, OutletCBVisibility, cfg, UIElementsTier3, parentkey, AddressOf ChildCB_VisibilityChanged, AddressOf CheckBoxChecked, outletCount)
    End Sub

    Public Sub CheckBoxChecked(sender As Object, e As EventArgs)
        Functions.CheckBoxChecked(sender, UIElementsTier3)
    End Sub
    Private Shared t As Integer = 1

    'Insert button click
    Public Async Sub b_insert_Click(sender As Object, e As EventArgs)

        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Autodesk.AutoCAD.DatabaseServices.Database = acDoc.Database
        Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor

        Using docLock As DocumentLock = acDoc.LockDocument()
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim bt As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)

                Dim inspoint As Point3d = New Point3d(0, 0, 0)
                ' If checkbox is checked, ask user for insert point
                If PosCheckbox.Checked Then
                    Dim ppo As New PromptPointOptions(vbLf & "Specify insertion point:")
                    Dim ppr As PromptPointResult = ed.GetPoint(ppo)
                    If ppr.Status = PromptStatus.OK Then
                        inspoint = ppr.Value
                    Else
                        ed.WriteMessage(vbLf & "Insertion point not selected. Command canceled.")
                        Return
                    End If
                End If

                Dim selectedtype = CBTier1.SelectedItem?.ToString()
                Dim selectedcomponent As String
                Dim selectedcomponenttier2 As String
                If SelectComponentGRP.Controls.Contains(CBTier3) Then
                    selectedcomponent = CBTier3.SelectedItem?.ToString()
                    selectedcomponenttier2 = CBTier2.SelectedItem?.ToString()
                Else
                    selectedcomponent = CBTier2.SelectedItem?.ToString()
                    selectedcomponenttier2 = CBTier2.SelectedItem?.ToString()
                End If

                Dim selectedcomponentblock = selectedcomponent & "_Block"
                If String.IsNullOrEmpty(selectedcomponent) Then
                    ed.WriteMessage(vbLf & "Please select a component type.")
                    Return
                End If

                'Insert cloned block reference and moving the hatch to appropriate location for correct draw order. 
                Dim br As BlockReference = Functions.InsertCloneBR(folderPath, selectedcomponent, selectedcomponentblock, acCurDb, inspoint)
                Dim ms As BlockTableRecord = acTrans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Functions.MoveHatchToBack(acTrans, br)
                Functions.MoveBROrder(acTrans, br, selectedcomponentblock)
                ms.AppendEntity(br)
                acTrans.AddNewlyCreatedDBObject(br, True)
                Functions.AppendAttribute(acTrans, br)

                ed.WriteMessage(vbLf & "Block '" & selectedcomponent & "' inserted successfully at " & inspoint.ToString())

                'Sync Attributes so that ID Code can be updated later
                Dim blockName As String = String.Empty
                If br.DynamicBlockTableRecord <> ObjectId.Null Then
                    Dim dynamicBlockTableRecord As BlockTableRecord = CType(acTrans.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                    blockName = dynamicBlockTableRecord.Name
                Else
                    Dim btrId As ObjectId = br.BlockTableRecord
                    Dim btr As BlockTableRecord = CType(acTrans.GetObject(btrId, OpenMode.ForRead), BlockTableRecord)
                    blockName = btr.Name
                End If
                Dim attSyncTask = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.ExecuteInCommandContextAsync(
Async Function()
    Await ed.CommandAsync(New Object() {"_.ATTSYNC", "_Name", blockName})
End Function, Nothing)
                Await attSyncTask

                '-----UPDATE ALL PARAMETERS ACCORDING TO CHOICES MADE IN THE UI--------
                Functions.UpdateParentVisibility(CompParamGRP, br)
                Dim nestedblocks As New Dictionary(Of BlockReference, String)
                nestedblocks = Functions.GetRuntimeNestedBlocks(acTrans, br)

                Functions.AddDoubleChildToBRList(selectedcomponentblock, nestedblocks, acTrans)

                Dim groupedNested = From kvp In nestedblocks
                                    Group kvp By blockDefName = kvp.Value Into Group

                For Each group In groupedNested
                    Dim instances = group.Group.ToList()
                    For i As Integer = 1 To instances.Count
                        ' build the tag name for this nested block.                   
                        ' for each nested block reference in this group, update its dynamic property.
                        Dim Instancecount As Integer = instances.Count
                        Dim Name As String = Functions.NormalizeBlockName(group.blockDefName)
                        Dim childblkref As BlockReference = instances(i - 1).Key
                        Dim deviceKey As String = selectedcomponentblock
                        Dim cfg As DeviceConfig = Nothing
                        'Update Child block parameters to match the corresponding comboboxes. Logic for finding the correct combobox is inside GetTagName
                        If CBTier1.SelectedItem.ToString().IndexOf("Metering Devices", StringComparison.OrdinalIgnoreCase) >= 0 Then
                            If Not _deviceConfigs.TryGetValue(deviceKey, cfg) Then
                                Throw New InvalidOperationException($"No configuration found for {deviceKey}")
                            End If
                        End If
                        Dim tagname As String = Functions.GetTagName(OutletCBVisibility, Name, acTrans, childblkref, cfg, Name)
                        Functions.UpdateParameters(childblkref, CompParamGRP, tagname, FlipCheckboxes)
                    Next
                Next
                Functions.AssignIdCode(acTrans, br, selectedcomponentblock)
                acDoc.Editor.Regen()
                acTrans.Commit()
                t += 1

            End Using
        End Using
    End Sub
End Class

Public Class DeviceConfig
    Public Property TotalOutletCount As Integer
    Public Property ChildBlockName As String
    Public Property NumberOfSides As Integer
    Public Property Sides As List(Of String)
    Public Property StartAt As String    ' "top" or "bottom"
    Public Property MiddleElements As List(Of MiddleElementConfig)
    Public Property SubElements As List(Of SubElementsConfig)
End Class

Public Class MiddleElementConfig
    Public Property Type As String
    Public Property CountExpression As String
    Public Property ChildBlockName As String
    Public Property AttributeName As String
    Public Property TagPrefix As String
    Public Property PromptText As String
End Class

Public Class SubElementsConfig
    Public Property ChildBlockName As String
    Public Property PromptText As String
    Public Property TagSuffix As String
End Class

Public Class DoubleNestedBinding
    Public Property DoubleCombo As ComboBox
    Public Property DoubleBlockCandidates As List(Of String)
End Class