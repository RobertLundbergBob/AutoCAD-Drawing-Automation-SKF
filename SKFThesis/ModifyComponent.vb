Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Newtonsoft.Json
Imports System.Drawing
Imports System.Reflection.Emit
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports System.IO


Public Class ModifyComponent
    Private MainCBVisibility As New ComboBox
    Private lstBlocks As New ListBox()
    Private b_back As New Button
    Private btnApply As New Button() With {.Text = "Apply Visibility"}
    Private visibilityControls As New Dictionary(Of ObjectId, ComboBox) ' Stores visibility ComboBoxes for each block
    Private blockDict As New Dictionary(Of String, ObjectId) ' Stores block reference IDs
    Private comboboxes As New List(Of ComboBox)
    Private comboboxeslvl2 As New List(Of ComboBox)
    Dim labels As New List(Of System.Windows.Forms.Label)
    Private CheckBoxAll As New CheckBox
    Dim ChildBlocks As New Dictionary(Of BlockReference, String)
    Private UIElementsTier1 As New List(Of Object)
    Private UIElementsTier2 As New List(Of Object)
    Private UIElementsTier3 As New List(Of Object)
    Dim CompParamGRP As New GroupBox()
    Dim mainscrollpanel As New Panel()

    Public Sub New()
        ' Initialize Form
        Me.Text = "Modify Drawing"
        Me.Size = New Drawing.Size(1450, 800)
        mainscrollpanel.AutoScroll = True
        mainscrollpanel.Dock = DockStyle.Fill
        mainscrollpanel.Name = "mainscrollpanel"
        Me.Controls.Add(mainscrollpanel)

        ' Configure ListBox
        lstBlocks.Location = New Drawing.Point(80, 20)
        lstBlocks.Size = New Drawing.Size(230, 320)
        AddHandler lstBlocks.SelectedIndexChanged, AddressOf lstblocks_selectedindexchanged

        ' Configure Apply Button
        Dim midpointlstblock As Integer = lstBlocks.Left + lstBlocks.Width / 2
        btnApply.Location = New Drawing.Point(midpointlstblock - 75, lstBlocks.Bottom + 20)
        btnApply.Size = New Drawing.Size(150, 50)
        btnApply.BackColor = Color.LimeGreen
        AddHandler btnApply.Click, AddressOf btnApply_Click

        b_back.Location = New Drawing.Point(20, 20)
        b_back.Size = New Drawing.Size(50, 20)
        b_back.Text = "Back"
        AddHandler b_back.Click, AddressOf b_back_click

        ' Add Controls to Form
        mainscrollpanel.Controls.Add(b_back)
        mainscrollpanel.Controls.Add(lstBlocks)
        mainscrollpanel.Controls.Add(btnApply)

        Dim Jsonpath = Functions.GetRad1Path() & "\Metering Devices\MeteringDevices.JSON"
        LoadDeviceConfigs(Jsonpath)
        ' Load Blocks on Form Load
        LoadBlockReferences()
    End Sub

    Private Sub b_back_click()
        'Open mainform
        Dim MainForm As New MainForm()
        MainForm.Show()

        'Close Insert Form
        Me.Close()
    End Sub

    Private Sub LoadBlockReferences()
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Autodesk.AutoCAD.DatabaseServices.Database = doc.Database

        Using tr As Transaction = db.TransactionManager.StartTransaction()
            Dim bt As BlockTable = DirectCast(tr.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
            Dim ms As BlockTableRecord = DirectCast(tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)

            lstBlocks.Items.Clear()
            blockDict.Clear()

            Dim i As Integer = 1
            ' Iterate through Model Space to find block references
            For Each entId As ObjectId In ms
                Dim ent As Entity = TryCast(tr.GetObject(entId, OpenMode.ForRead), Entity)
                If TypeOf ent Is BlockReference Then
                    Dim blkRef As BlockReference = DirectCast(ent, BlockReference)
                    Dim blockDef As BlockTableRecord = DirectCast(tr.GetObject(blkRef.DynamicBlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                    Dim blkName As String = blockDef.Name ' Use the block definition name

                    'Remove clone from name
                    Dim pattern As String = "Clone\d+$"
                    Dim regex As New Regex(pattern, RegexOptions.IgnoreCase)
                    Dim newName As String = regex.Replace(blkName, "")

                    Dim Attribute As String = Functions.GetAttribute(tr, blkRef, "IDENTIFICATION_CODE")
                    Dim ListBoxName As String = newName & " " & Attribute & i

                    ' Add block reference to ListBox
                    lstBlocks.Items.Add(ListBoxName)
                    blockDict.Add(ListBoxName, entId)
                    i += 1
                End If
            Next

            tr.Commit()
        End Using
    End Sub

    '' <summary>
    '' when a block is selected, Create UI for that block with active visibility states of the chosen block 
    '' </summary>
    Private Sub lstblocks_selectedindexchanged(sender As Object, e As EventArgs)
        ' clear previous visibility comboboxes
        InsertComponent.ClearFlipCheckBoxes(FlipCheckboxes, FlipLabels, Me)
        Functions.ClearUIElementList(UIElementsTier1, CompParamGRP)
        Functions.ClearUIElementList(UIElementsTier2, CompParamGRP)
        Functions.ClearUIElementList(UIElementsTier3, CompParamGRP)


        If lstBlocks.SelectedItem Is Nothing Then Exit Sub

        Dim selectedblock As String = lstBlocks.SelectedItem.ToString()
        Dim BlockName As String = selectedblock.Split(" "c)(0)
        Dim ComponentName = BlockName.Replace("_Block", "")
        Dim visibilitystates As List(Of String) = DatabaseFunctions.GetVisibilityStatesFromDatabase(BlockName)
        If Not blockDict.ContainsKey(selectedblock) Then Exit Sub

        CompParamGRP.Text = "Modify Parameters"
        CompParamGRP.Location = New Point(lstBlocks.Right + 50, 20)
        CompParamGRP.AutoSize = True
        CompParamGRP.AutoSizeMode = AutoSizeMode.GrowAndShrink
        CompParamGRP.Padding = New Padding(10)  ' leave some breathing room
        mainscrollpanel.Controls.Add(CompParamGRP)

        Dim basepointx As Integer = 15
        Dim basepointy As Integer = 20
        MainCBVisibility = Functions.CreateComboBoxes(CompParamGRP, visibilitystates, BlockName, basepointx, basepointy, UIElementsTier2, AddressOf OutletCBVisibilitys_SelectedIndexChanged, AddressOf ChildCB_VisibilityChanged)

        'Show Current Visibility State in ComboBoxes
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Autodesk.AutoCAD.DatabaseServices.Database = doc.Database
        Dim ed As Editor = doc.Editor

        Dim folderpath2 = Functions.GetRad1Path() & "\" & ComponentName & "\" & ComponentName
        Functions.InsertHelperImage(folderpath2, ComponentName, CompParamGRP, mainscrollpanel)

        Using docLock As DocumentLock = doc.LockDocument()
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Try
                    ' Try to get the entity
                    Dim entID As ObjectId = blockDict(selectedblock)
                    Dim ent As Entity = TryCast(tr.GetObject(entID, OpenMode.ForRead), Entity)

                    If TypeOf ent Is BlockReference Then
                        Dim blkRef As BlockReference = CType(ent, BlockReference)
                        Dim CurrentParentState As String = Functions.GetCurrentVisibilityState(blkRef)
                        If MainCBVisibility.Items.Contains(CurrentParentState) Then
                            MainCBVisibility.SelectedItem = CurrentParentState
                        Else
                            ' Optional: handle case where the visibility state isn't found
                            ed.WriteMessage(vbLf & "Visibility state not found in combo box.")
                        End If

                        Dim currentdevicekey As String = BlockName
                        Dim cfg As DeviceConfig = Nothing ' Initialize cfg to Nothing
                        Dim deviceConfigFound As Boolean = False

                        ' Try to get the device configuration if a currentDeviceKey is available
                        If Not String.IsNullOrEmpty(currentdevicekey) AndAlso _deviceConfigs IsNot Nothing Then
                            deviceConfigFound = _deviceConfigs.TryGetValue(currentdevicekey, cfg)
                        End If

                        ChildBlocks = Functions.GetRuntimeNestedBlocks(tr, blkRef)
                        Dim groupedChildren = From kvp In ChildBlocks
                                              Group kvp By blockDefName = kvp.Value Into Group

                        For Each group In groupedChildren
                            Dim instances = group.Group.ToList()
                            For i As Integer = 1 To instances.Count
                                Dim childbr As BlockReference = instances(i - 1).Key
                                Dim childname As String = group.blockDefName
                                Dim instancecount As Integer = instances.Count
                                If childbr.Visible Then
                                    Dim CurrentChildState As String = Functions.GetCurrentVisibilityState(childbr)
                                    Dim dynDef As BlockTableRecord = tr.GetObject(childbr.DynamicBlockTableRecord, OpenMode.ForRead)
                                    Dim Name As String = Functions.NormalizeBlockName(group.blockDefName)
                                    Dim tagname As String = Functions.GetTagName(MainCBVisibility, Name, tr, childbr, cfg, BlockName)
                                    Dim CB As ComboBox = Functions.FindComboBoxByTag(CompParamGRP, tagname)
                                    If CB IsNot Nothing Then
                                        For Each item As Object In CB.Items
                                            If String.Equals(item.ToString(), CurrentChildState, StringComparison.OrdinalIgnoreCase) Then
                                                CB.SelectedItem = item
                                                Exit For
                                            End If
                                        Next
                                    End If
                                End If
                            Next
                        Next
                        ed.WriteMessage(vbLf & "BlockReference found: " & BlockName)
                    Else
                        ed.WriteMessage(vbLf & "Selected object is not a BlockReference.")
                    End If

                    tr.Commit()
                Catch ex As Exception
                    ed.WriteMessage(vbLf & "Error: " & ex.Message)
                End Try
            End Using
        End Using
    End Sub

    Private Shared FlipCheckboxes As New Dictionary(Of ComboBox, CheckBox)
    Private Shared FlipLabels As New Dictionary(Of ComboBox, System.Windows.Forms.Label)

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

        ' Call your database function
        Functions.PlaceFlipCheckBox(parameterValueID, FlipCheckboxes, cb, FlipLabels)

        'FOR DOUBLENESTED BLOCKS 
        If InsertComponent.ChildToDoubleBindingMap.ContainsKey(cb) Then
            Dim binding = InsertComponent.ChildToDoubleBindingMap(cb)
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

    Dim meteringdevicepanel As New Panel
    Private Shared _deviceConfigs As Dictionary(Of String, DeviceConfig)
    Public Shared Sub LoadDeviceConfigs(jsonPath As String)
        Dim json As String = File.ReadAllText(jsonPath)
        _deviceConfigs = JsonConvert.DeserializeObject(Of Dictionary(Of String, DeviceConfig))(json)
    End Sub

    '------JSON METERING DEVICE PANEL BUILD------------
    Private Sub OutletCBVisibilitys_SelectedIndexChanged(sender As Object, e As EventArgs)
        Functions.ClearUIElementList(UIElementsTier3, CompParamGRP)

        Dim selectedblock As String = lstBlocks.SelectedItem.ToString()
        Dim ParentKey As String = selectedblock.Split(" "c)(0)

        Dim cfg As DeviceConfig = Nothing
        If Not _deviceConfigs.TryGetValue(ParentKey, cfg) Then
            MessageBox.Show("No configuration for " & ParentKey)
            Return
        End If

        Dim outletCount As Integer
        If Not Integer.TryParse(MainCBVisibility.SelectedItem?.ToString(), outletCount) Then
            MessageBox.Show("Invalid outlet count")
            Return
        End If

        CheckBoxAll = Functions.BuildMeteringDeviceControls(CompParamGRP, MainCBVisibility, cfg, UIElementsTier3, ParentKey, AddressOf ChildCB_VisibilityChanged, AddressOf CheckBoxChecked, outletCount)

        'Show Current Visibility State in ComboBoxes
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Autodesk.AutoCAD.DatabaseServices.Database = doc.Database
        Dim ed As Editor = doc.Editor

        Using docLock As DocumentLock = doc.LockDocument()
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Try
                    ' Try to get the entity
                    Dim entID As ObjectId = blockDict(selectedblock)
                    Dim ent As Entity = TryCast(tr.GetObject(entID, OpenMode.ForRead), Entity)

                    If TypeOf ent Is BlockReference Then
                        Dim blkRef As BlockReference = CType(ent, BlockReference)

                        Dim currentdevicekey As String = ParentKey
                        cfg = Nothing ' Initialize cfg to Nothing
                        Dim deviceConfigFound As Boolean = False

                        ' Try to get the device configuration if a currentDeviceKey is available
                        If Not String.IsNullOrEmpty(currentdevicekey) AndAlso _deviceConfigs IsNot Nothing Then
                            deviceConfigFound = _deviceConfigs.TryGetValue(currentdevicekey, cfg)
                        End If

                        ChildBlocks = Functions.GetRuntimeNestedBlocks(tr, blkRef)
                        Dim groupedChildren = From kvp In ChildBlocks
                                              Group kvp By blockDefName = kvp.Value Into Group

                        For Each group In groupedChildren
                            Dim instances = group.Group.ToList()
                            For i As Integer = 1 To instances.Count
                                Dim childbr As BlockReference = instances(i - 1).Key
                                Dim childname As String = group.blockDefName
                                Dim instancecount As Integer = instances.Count
                                If childbr.Visible Then
                                    Dim CurrentChildState As String = Functions.GetCurrentVisibilityState(childbr)
                                    Dim dynDef As BlockTableRecord = tr.GetObject(childbr.DynamicBlockTableRecord, OpenMode.ForRead)
                                    Dim Name As String = Functions.NormalizeBlockName(group.blockDefName)
                                    Dim tagname As String = Functions.GetTagName(MainCBVisibility, Name, tr, childbr, cfg, ParentKey)
                                    Dim CB As ComboBox = Functions.FindComboBoxByTag(CompParamGRP, tagname)
                                    If CB IsNot Nothing Then
                                        For Each item As Object In CB.Items
                                            If String.Equals(item.ToString(), CurrentChildState, StringComparison.OrdinalIgnoreCase) Then
                                                CB.SelectedItem = item
                                                Exit For
                                            End If
                                        Next
                                    End If

                                End If
                            Next
                        Next
                        ed.WriteMessage(vbLf & "BlockReference found: " & ParentKey)
                    Else
                        ed.WriteMessage(vbLf & "Selected object is not a BlockReference.")
                    End If

                    tr.Commit()
                Catch ex As Exception
                    ed.WriteMessage(vbLf & "Error: " & ex.Message)
                End Try
            End Using
        End Using
    End Sub

    Public Sub CheckBoxChecked(sender As Object, e As EventArgs)
        Functions.CheckBoxChecked(sender, UIElementsTier3)
    End Sub

    ''' <summary>
    ''' Apply the selected visibility state to the block reference.
    ''' </summary>
    Private Sub btnApply_Click(sender As Object, e As EventArgs)
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Autodesk.AutoCAD.DatabaseServices.Database = doc.Database
        Dim ed As Editor = doc.Editor

        Using lock As DocumentLock = doc.LockDocument
            Using tr As Transaction = db.TransactionManager.StartTransaction()

                Dim selectedblock As String = lstBlocks.SelectedItem.ToString()
                Dim BlockName As String = selectedblock.Split(" "c)(0)

                Dim entID As ObjectId = blockDict(selectedblock)
                Dim ent As Entity = TryCast(tr.GetObject(entID, OpenMode.ForRead), Entity)
                Dim blkRef As BlockReference
                If TypeOf ent Is BlockReference Then
                    blkRef = CType(ent, BlockReference)
                    'Outlet Count
                    Functions.UpdateParentVisibility(CompParamGRP, blkRef)
                End If

                Dim currentdevicekey As String = BlockName
                Dim cfg As DeviceConfig = Nothing ' Initialize cfg to Nothing
                Dim deviceConfigFound As Boolean = False

                ' Try to get the device configuration if a currentDeviceKey is available
                If Not String.IsNullOrEmpty(currentdevicekey) AndAlso _deviceConfigs IsNot Nothing Then
                    deviceConfigFound = _deviceConfigs.TryGetValue(currentdevicekey, cfg)
                End If

                ChildBlocks.Clear()
                ChildBlocks = Functions.GetRuntimeNestedBlocks(tr, blkRef)

                Functions.AddDoubleChildToBRList(BlockName, ChildBlocks, tr)

                Dim groupedNested = From kvp In ChildBlocks
                                    Group kvp By blockDefName = kvp.Value Into Group

                For Each group In groupedNested
                    Dim instances = group.Group.ToList()
                    For i As Integer = 1 To instances.Count
                        Dim ChildBr As BlockReference = instances(i - 1).Key
                        Dim instancecount As Integer = instances.Count
                        Dim Name As String = Functions.NormalizeBlockName(group.blockDefName)
                        Dim tagname As String = Functions.GetTagName(MainCBVisibility, Name, tr, ChildBr, cfg, BlockName)
                        If ChildBr.Visible Then
                            Functions.UpdateParameters(ChildBr, CompParamGRP, tagname, FlipCheckboxes)
                        End If
                    Next
                Next
                Functions.AssignIdCode(tr, blkRef, selectedblock)
                doc.Editor.Regen()
                tr.Commit()
                ed.WriteMessage(vbLf & "Block visibility updated successfully!")
            End Using
        End Using
    End Sub

End Class