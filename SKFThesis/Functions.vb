Imports System.Configuration
Imports System.Data.Entity.Core.Mapping
Imports System.Data.SQLite
Imports System.Diagnostics.PerformanceData
Imports System.Drawing
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text.RegularExpressions
Imports System.Windows.Forms
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.GraphicsInterface
Imports Autodesk.AutoCAD.PlottingServices
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.Colors
Imports Microsoft.SqlServer.Server
Imports Microsoft.VisualBasic.Devices
Imports Newtonsoft.Json
Imports Newtonsoft.Json.Linq
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application

Public Class Functions
    Public Shared folderPath As String = GetRad1Path() & "\"
    <CommandMethod("UploadToDatabase")>
    Public Sub UploadDatabase()
        UploadToDatabase.UploadBlockLibraryData()
        'Dim frm As New UploadToDatabase
        '
        'frm.Show()
    End Sub

    <CommandMethod("Component")>
    Public Sub Component()
        Dim frm As New MainForm

        frm.Show()
    End Sub

    <CommandMethod("UpdateInsertIDJSON")>
    Public Sub SyncParserJSON()
        SyncToJSON.SyncParserJSON(folderPath & "TrialParserUpload.JSON")
    End Sub

    <CommandMethod("ChangeColorToWhite")>
    Public Sub ChangeLayerWhiteColorCommand()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Dim lt As LayerTable = trans.GetObject(db.LayerTableId, OpenMode.ForRead)

            Dim layerName As String = "mask_hatch"

            Dim layerId As ObjectId = ObjectId.Null
            Dim ltr As LayerTableRecord
            For Each id As ObjectId In lt
                ltr = trans.GetObject(id, OpenMode.ForRead)
                If String.Equals(ltr.Name, layerName, StringComparison.OrdinalIgnoreCase) Then
                    layerId = id
                    Exit For
                End If
            Next

            If layerId = ObjectId.Null Then
                ed.WriteMessage(vbLf & "Layer '" & layerName & "' not found (case-insensitive).")
                Return
            End If

            Dim ltrId As ObjectId = lt(layerName)
            ltr = trans.GetObject(ltrId, OpenMode.ForWrite)

            'Set color to white
            ltr.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(255, 255, 255)

            ed.WriteMessage(vbLf & "Color of layer '" & layerName & "' changed to white.")

            trans.Commit()
        End Using
    End Sub

    <CommandMethod("ChangeColorToBlue")>
    Public Sub ChangeLayerColorCommand()
        Dim doc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Using trans As Transaction = db.TransactionManager.StartTransaction()
            Dim lt As LayerTable = trans.GetObject(db.LayerTableId, OpenMode.ForRead)

            Dim layerName As String = "mask_hatch"

            Dim layerId As ObjectId = ObjectId.Null
            Dim ltr As LayerTableRecord
            For Each id As ObjectId In lt
                ltr = trans.GetObject(id, OpenMode.ForRead)
                If String.Equals(ltr.Name, layerName, StringComparison.OrdinalIgnoreCase) Then
                    layerId = id
                    Exit For
                End If
            Next

            If layerId = ObjectId.Null Then
                ed.WriteMessage(vbLf & "Layer '" & layerName & "' not found (case-insensitive).")
                Return
            End If

            Dim ltrId As ObjectId = lt(layerName)
            ltr = trans.GetObject(ltrId, OpenMode.ForWrite)

            'Set color to blue
            ltr.Color = Autodesk.AutoCAD.Colors.Color.FromRgb(33, 40, 48)

            ed.WriteMessage(vbLf & "Color of layer '" & layerName & "' changed to blue.")

            trans.Commit()
        End Using
    End Sub

    Public Shared Function GetRad1Path() As String
        Dim configPath As String = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "Rad1Tool\config.txt")

        ' Check if config file exists
        If File.Exists(configPath) Then
            Dim savedPath As String = File.ReadAllText(configPath).Trim()
            If Directory.Exists(savedPath) Then
                Return savedPath
            End If
        End If

        ' Ask user to select folder on first launch or if path is invalid
        Using folderDlg As New FolderBrowserDialog()
            folderDlg.Description = "Select the RAD1 folder (e.g., Examensarbete\RAD1)"
            folderDlg.ShowNewFolderButton = False

            If folderDlg.ShowDialog() = DialogResult.OK Then
                Dim selectedPath = folderDlg.SelectedPath
                ' Save it for future use
                Directory.CreateDirectory(Path.GetDirectoryName(configPath))
                File.WriteAllText(configPath, selectedPath)
                Return selectedPath
            Else
                MessageBox.Show("No folder selected. Application cannot continue.", "Setup Required", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return Nothing
            End If
        End Using
    End Function

    Public Shared Function InsertHelperImage(folderpathtier2, selecteditem, compparamgrp, mainscrollpanel)
        'Insert Picture if exists
        Dim possibleExts As String() = {"png", "jpg", "jpeg"}
        Dim foundImagePath As String = Nothing

        ' Try each extension until we find one that exists:
        For Each ext In possibleExts
            Dim candidate As String = Path.Combine(folderpathtier2, selecteditem & "." & ext)
            If File.Exists(candidate) Then
                foundImagePath = candidate
                Exit For
            End If
        Next

        If Not String.IsNullOrEmpty(foundImagePath) Then
            Dim pb As PictureBox = Nothing

            ' Check if we've already created it before (so we don't stack multiple PictureBoxes)
            ' Here, we search your container (e.g. SelectComponentGRP) for any control with Tag = "dynamicPreview"
            For Each c As Control In compparamgrp.Controls
                If TypeOf c Is PictureBox AndAlso c.Tag IsNot Nothing AndAlso c.Tag.ToString() = "dynamicPreview" Then
                    pb = CType(c, PictureBox)
                    Exit For
                End If
            Next
            If pb Is Nothing Then
                ' Not created yet, so make a brand-new one:
                pb = New PictureBox()
                'Place below parameters combo boxes
                Dim xPos As Integer = compparamgrp.Right + 20
                Dim yPos As Integer = compparamgrp.Top
                pb.Location = New Drawing.Point(xPos, yPos)
                ' Give it a fixed size (e.g. 200×200)
                pb.Size = New Drawing.Size(500, 500)
                pb.BorderStyle = BorderStyle.FixedSingle
                ' Set SizeMode to Zoom so the image scales properly
                pb.SizeMode = PictureBoxSizeMode.Zoom
                ' Finally, add it to the same container that holds your other dynamic controls
                compparamgrp.Controls.Add(pb)
            End If

            ' Now load the image (if it was previously loaded, Dispose the old one first):
            If pb.Image IsNot Nothing Then
                pb.Image.Dispose()
            End If
            pb.Image = System.Drawing.Image.FromFile(foundImagePath)
        End If
    End Function

    Public Shared Function GetVisibilityStates(trans As Transaction, bt As BlockTable, blkref As BlockReference) As List(Of String)
        Dim visibilitystates As New List(Of String)

        ' Open the DWG file
        Dim ms As BlockTableRecord = DirectCast(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)


        ' Iterate through Model Space to § block references
        For Each entId As ObjectId In ms
            Dim ent As Entity = TryCast(trans.GetObject(entId, OpenMode.ForRead), Entity)
            If TypeOf ent Is BlockReference Then
                blkref = DirectCast(ent, BlockReference)
                Dim blockDef As BlockTableRecord = DirectCast(trans.GetObject(blkref.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                Dim blkName As String = blockDef.Name ' Use the block definition name
            End If
        Next

        Dim dynProps As DynamicBlockReferencePropertyCollection = blkref.DynamicBlockReferencePropertyCollection

        For Each prop As DynamicBlockReferenceProperty In dynProps
            If prop.PropertyName.Contains("Visibility") Then
                ' Populate ComboBox with available visibility states
                For Each allowedValue As Object In prop.GetAllowedValues
                    visibilitystates.Add(allowedValue.ToString())
                Next
                Exit For
            End If
        Next

        Return visibilitystates
    End Function

    Public Shared Function GetBlockReference(trans, bt, blockname) As BlockReference

        Dim ms As BlockTableRecord = DirectCast(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)

        ' Iterate through Model Space to find block references
        For Each entId As ObjectId In ms
            Dim ent As Entity = TryCast(trans.GetObject(entId, OpenMode.ForRead), Entity)
            If TypeOf ent Is BlockReference Then
                Dim br As BlockReference = DirectCast(ent, BlockReference)
                Dim btr As BlockTableRecord = DirectCast(trans.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                Dim Name As String = btr.Name
                If String.Compare(btr.Name, blockname, True) = 0 Then
                    Return br
                End If
            End If
        Next
    End Function

    Public Shared Function GetChildVisibilityStates(trans As Transaction, childblkref As BlockReference) As List(Of String)
        Dim ChildVisibilityStates As New List(Of String)

        Dim dynBlockId As ObjectId = childblkref.DynamicBlockTableRecord
        Dim dynBlockDef As BlockTableRecord = DirectCast(trans.GetObject(dynBlockId, OpenMode.ForRead), BlockTableRecord)
        Dim nestedBlockName As String = dynBlockDef.Name
        Dim dynProps As DynamicBlockReferencePropertyCollection = childblkref.DynamicBlockReferencePropertyCollection
        For Each prop As DynamicBlockReferenceProperty In dynProps
            If prop.PropertyName.Contains("Visibility") Then
                ' Populate List with available visibility states
                For Each allowedValue As Object In prop.GetAllowedValues
                    ChildVisibilityStates.Add(allowedValue.ToString())
                Next
                Exit For
            End If
        Next

        Return ChildVisibilityStates
    End Function

    Public Shared Function GetNestedVisibilityStates(trans As Transaction, blkref As BlockReference) As Dictionary(Of BlockReference, Tuple(Of String, List(Of String)))
        Dim nestedvisibilitystates As New Dictionary(Of BlockReference, Tuple(Of String, List(Of String)))

        ' Access nested blocks inside the selected block
        Dim blockDef As BlockTableRecord = DirectCast(trans.GetObject(blkref.DynamicBlockTableRecord, OpenMode.ForRead), BlockTableRecord)

        For Each entId As ObjectId In blockDef


            Dim ent As Entity = TryCast(trans.GetObject(entId, OpenMode.ForRead), Entity)
            If TypeOf ent Is BlockReference Then
                Dim nestedBlkRef As BlockReference = DirectCast(ent, BlockReference)
                Dim dynBlockId As ObjectId = nestedBlkRef.DynamicBlockTableRecord
                Dim dynBlockDef As BlockTableRecord = DirectCast(trans.GetObject(dynBlockId, OpenMode.ForRead), BlockTableRecord)
                Dim nestedBlockName As String = dynBlockDef.Name

                Dim dynProps As DynamicBlockReferencePropertyCollection = nestedBlkRef.DynamicBlockReferencePropertyCollection
                For Each prop As DynamicBlockReferenceProperty In dynProps
                    If prop.PropertyName.Contains("Visibility") Then
                        ' Check if the block reference already exists in the dictionary
                        If Not nestedvisibilitystates.ContainsKey(nestedBlkRef) Then
                            nestedvisibilitystates(nestedBlkRef) = Tuple.Create(nestedBlockName, New List(Of String))
                        End If

                        ' Populate ComboBox with available visibility states
                        For Each allowedValue As Object In prop.GetAllowedValues
                            nestedvisibilitystates(nestedBlkRef).Item2.Add(allowedValue.ToString())
                        Next
                        Exit For
                    End If
                Next
            End If
        Next

        Return nestedvisibilitystates
    End Function

    Public Shared Function VisibilityToArticleNumber(componentName As String, childblockname As String, state As String) As String
        ' 0) if the caller accidentally gave us a null/empty key, just return empty
        If String.IsNullOrWhiteSpace(childblockname) Then
            Return ""
        End If

        Dim compDict As Dictionary(Of String, Dictionary(Of String, DualMapEntry)) = Nothing
        If Not InsertComponent.visibilityMap.TryGetValue(componentName, compDict) Then
            Return ""
        End If

        Dim paramDict As Dictionary(Of String, DualMapEntry) = Nothing
        If Not compDict.TryGetValue(childblockname, paramDict) Then
            Return ""
        End If

        Dim entry As DualMapEntry = Nothing
        If Not paramDict.TryGetValue(state, entry) Then
            Return ""
        End If

        Return entry.article
    End Function

    Public Shared Function VisibilityToCodeValue(componentName As String, childblockname As String, state As String) As String
        If String.IsNullOrWhiteSpace(childblockname) Then
            Return ""
        End If

        ' 1) Try grab the component → (parameter → state) map
        Dim compDict As Dictionary(Of String, Dictionary(Of String, DualMapEntry)) = Nothing
        If Not InsertComponent.visibilityMap.TryGetValue(componentName, compDict) Then
            Return ""   ' unknown component
        End If

        ' 2) Try grab the parameter → state map
        Dim paramDict As Dictionary(Of String, DualMapEntry) = Nothing
        If Not compDict.TryGetValue(childblockname, paramDict) Then
            Return ""   ' this parameter isn't defined for that component
        End If

        ' 3) Try grab the actual entry for this visibility state
        Dim entry As DualMapEntry = Nothing
        If Not paramDict.TryGetValue(state, entry) Then
            Return ""   ' that state wasn't in your JSON
        End If

        ' 4) Finally return the codeValue (may be "" if JSON omitted it)
        Return entry.codeValue
    End Function
    Public Shared Function GenerateIdCode(tr As Transaction, Br As BlockReference, parentName As String) As String
        Dim idCodeParts As New Dictionary(Of String, List(Of String))
        Dim ParentVisibilityState As String = Functions.GetCurrentVisibilityState(Br)
        Dim componentName = parentName.Split("_"c)(0) 'parentName.Replace("_Block", "")
        Dim No_of_plugg As Integer = 0
        Dim formatRule As String = ""


        If Not InsertComponent.ComponentFormatMap.TryGetValue(componentName, formatRule) Then
            Return "FORMAT_NOT_DEFINED"
        End If

        idCodeParts("Component") = New List(Of String) From {componentName}

        Dim childblocks As Dictionary(Of BlockReference, String) = GetRuntimeNestedBlocks(tr, Br)
        Dim groupedchildren = From kvp In childblocks
                              Group kvp By blockDefName = kvp.Value Into Group

        For Each group In groupedchildren
            Dim expectedSpecial = componentName & "_Monitoring"
            If String.Equals(group.blockDefName, expectedSpecial, StringComparison.OrdinalIgnoreCase) Then
                Dim monitorstate = group.Group _
            .Select(Function(kvp) New With {Key .br = kvp.Key, Key .state = Functions.GetCurrentVisibilityState(kvp.Key)}) _
            .FirstOrDefault(Function(x) x.state <> "Without")
                If monitorstate IsNot Nothing Then
                    ' map visibility → article number (e.g. "2" or "3")
                    Dim Codeval As String = VisibilityToCodeValue(componentName, group.blockDefName, monitorstate.state)
                    idCodeParts(group.blockDefName) = New List(Of String) From {Codeval}

                    ' grab the POSITION attribute and map that, too
                    Dim posAttr As String = GetAttribute(tr, monitorstate.br, "POSITION")
                    Dim CodeValPos As String = VisibilityToCodeValue(componentName, group.blockDefName & "_Position", posAttr)
                    idCodeParts(group.blockDefName & "_Position") = New List(Of String) From {CodeValPos}
                End If
                Continue For
            End If
            Dim instances = group.Group.ToList()
            For i As Integer = 1 To instances.Count
                Dim childbr As BlockReference = instances(i - 1).Key
                ' Remove suffix from blockreference to match the translation json file
                If childbr.Visible Then
                    Dim childstate As String = Functions.GetCurrentVisibilityState(childbr)
                    Dim baseName As String = NormalizeBlockName(group.blockDefName)

                    If childstate = "Plugg" Then
                        No_of_plugg += 1
                    End If
                    Dim codeval As String = VisibilityToCodeValue(componentName, baseName, childstate)
                    Dim posvalue As String
                    If InsertComponent.visibilityMap(componentName).ContainsKey(group.blockDefName & "_Position") Then
                        Dim pos As String = GetAttribute(tr, childbr, "POSITION")
                        posvalue = VisibilityToCodeValue(componentName, baseName & "_Position", pos)
                    End If
                    'If Not String.IsNullOrEmpty(artnr) Then
                    Dim position_attribute As String = GetAttribute(tr, childbr, "POSITION")
                    Dim logicalComponentKey As String
                    If Not String.IsNullOrEmpty(position_attribute) Then
                        logicalComponentKey = $"{baseName}_{position_attribute}"  'Add attribute position number to blockDefName
                    Else
                        logicalComponentKey = baseName
                    End If
                    Dim singleUseList As HashSet(Of String) = Nothing
                    If InsertComponent.SingleUseComponentsMap.TryGetValue(componentName, singleUseList) AndAlso
singleUseList.Contains(baseName) Then
                        If Not idCodeParts.ContainsKey(baseName) Then
                            idCodeParts(baseName) = New List(Of String)
                        End If
                        If Not idCodeParts(baseName).Contains(codeval) Then
                            idCodeParts(baseName).Add(codeval)
                        End If
                    Else
                        If Not idCodeParts.ContainsKey(logicalComponentKey) Then
                            idCodeParts(logicalComponentKey) = New List(Of String)
                        End If
                        idCodeParts(logicalComponentKey).Add(codeval)
                    End If

                End If
            Next
        Next
        Dim doublenested As String = DatabaseFunctions.GetParentAttribute(componentName & "_Block", "DOUBLENESTED")
        If doublenested = "True" Then
            For Each kvp In childblocks.ToList()
                ' Get the nested visibility states for this inner block reference.
                Dim innerblocks As Dictionary(Of BlockReference, String) = Functions.GetRuntimeNestedBlocks(tr, kvp.Key)
                For Each innerKvp As KeyValuePair(Of BlockReference, String) In innerblocks
                    If innerKvp.Key.Visible Then
                        Dim nestedState As String = GetCurrentVisibilityState(innerKvp.Key)
                        For Each param In InsertComponent.visibilityMap(componentName)
                            If InsertComponent.visibilityMap(componentName)(param.Key).ContainsKey(nestedState) Then
                                Dim nr = VisibilityToCodeValue(componentName, param.Key, nestedState)
                                If Not idCodeParts.ContainsKey(param.Key) Then
                                    idCodeParts(param.Key) = New List(Of String)
                                End If
                                idCodeParts(param.Key).Add(nr)
                            End If
                        Next
                    End If
                Next
            Next
        End If
        For Each Key In idCodeParts.Keys.ToList()
            Dim updatedList As New List(Of String)
            For Each part In idCodeParts(Key)
                If part IsNot Nothing AndAlso part.Contains("<outletCount - 1>") AndAlso Not part.Equals(componentName & "_Bypass_Bore") Then
                    Dim replaced = part.Replace("<outletCount - 1>", (Convert.ToInt32(ParentVisibilityState) - 1).ToString())
                    updatedList.Add(replaced)
                Else
                    updatedList.Add(part)
                End If
            Next
            idCodeParts(Key) = updatedList
        Next
        If Not String.IsNullOrEmpty(ParentVisibilityState) Then
            idCodeParts("<outletCount>") = New List(Of String) From {ParentVisibilityState}
            idCodeParts("<outletCount - plugg>") = New List(Of String) From {(Convert.ToInt32(ParentVisibilityState) - No_of_plugg).ToString()}
        End If

        Dim idCode As String = formatRule
        For Each kvp In idCodeParts
            Dim key = kvp.Key
            Dim valueList = kvp.Value

            Dim joinedValue = String.Join("", valueList)
            idCode = idCode.Replace("{" & key & "}", joinedValue)
        Next
        idCode = System.Text.RegularExpressions.Regex.Replace(idCode, "\{[^\}]+\}", "")
        Return idCode
    End Function
    Public Shared Function GetOutletPosition(index As Integer, totalOutlets As Integer, panel As Panel) As Point
        ' Let's assume outlets are evenly distributed
        Dim outletsPerSide As Integer = totalOutlets \ 2

        Dim spacingY As Integer = panel.Height \ (outletsPerSide)
        Dim xLeft As Integer = panel.Left - 160 ' offset left of rectangle
        Dim xRight As Integer = panel.Right + 10

        If index Mod 2 = 1 Then
            ' Odd numbers (1,3,5...) – right side, from bottom up
            Dim posIndex As Integer = (index - 1) \ 2
            Return New Point(xRight, panel.Bottom - spacingY * (posIndex + 1))
        Else
            ' Even numbers (2,4,6...) – left side, from bottom up
            Dim posIndex As Integer = (index - 2) \ 2
            Return New Point(xLeft, panel.Bottom - spacingY * (posIndex + 1))
        End If
    End Function

    Public Shared Function GetSidePositions(cfg As DeviceConfig, total As Integer, panel As Panel) _
    As Dictionary(Of Integer, Point)

        Dim result As New Dictionary(Of Integer, Point)
        Dim nSides As Integer = cfg.NumberOfSides
        Dim perSide = total / nSides
        ' Height of each control
        Const ctrlHeight As Integer = 30
        ' Compute spacing so that pos=0 → bottom, pos=perSide-1 → top
        Dim spacingY As Double = 0
        If perSide > 1 Then
            spacingY = (panel.Height - ctrlHeight) / CDbl(perSide - 1)
        End If
        Dim xLeft = panel.Left - 195
        Dim xRight = panel.Right + 5
        ' Determine the horizontal starting side based on the new config property
        Dim startCornerLower = cfg.StartAt.Trim().ToLower()
        Dim startOnLeft As Boolean = startCornerLower.Contains("left") ' True if "left" is in the corner name
        Dim startOnRight As Boolean = startCornerLower.Contains("right") ' True if "right" is in the corner name

        For i As Integer = 1 To total
            Dim sideIdx As Integer
            Dim posInSide As Integer ' Index within the vertical stack on its side

            If nSides = 2 Then
                posInSide = (i - 1) \ 2 ' Simplified: (1-1)\2=0, (2-1)\2=0, (3-1)\2=1, (4-1)\2=1 etc.
                If startOnLeft Then
                    ' If starting on Left (e.g., bottom-left, top-left):                   
                    sideIdx = If(i Mod 2 = 1, xLeft, xRight)
                ElseIf startOnRight Then
                    ' If starting on Right (e.g., bottom-right, top-right):
                    sideIdx = If(i Mod 2 = 1, xRight, xLeft) ' This is the original logic
                Else
                    ' Default behavior if StartingCorner for 2 sides is invalid or missing
                    sideIdx = If(i Mod 2 = 1, xRight, xLeft)
                End If
            Else ' NumberOfSides is 1
                posInSide = i - 1

                ' The side is determined solely by the Sides array for nSides=1
                sideIdx = If(cfg.Sides(0).Trim().ToLower() = "right", xRight, xLeft)
            End If


            Dim yPos As Integer
            Dim startVerticalLower = cfg.StartAt.Trim().ToLower()
            Dim startbottom = startVerticalLower.Contains("bottom")
            Dim starttop = startVerticalLower.Contains("top")
            If startbottom Then
                ' Start at panel.Bottom - ctrlHeight, spacing up
                yPos = CInt(Math.Round(panel.Bottom - ctrlHeight - spacingY * posInSide))
            ElseIf starttop Then
                ' Start at panel.Top, spacing down
                yPos = CInt(Math.Round(panel.Top + spacingY * posInSide))
            Else
                ' Default to bottom if StartAt is missing or invalid
                yPos = CInt(Math.Round(panel.Bottom - ctrlHeight - spacingY * posInSide))
            End If
            result(i) = New Point(sideIdx, yPos)
        Next

        Return result
    End Function

    Public Shared Function EvaluateExpression(expr As String, total As Integer) As Integer
        'only handles "TotalOutletCount / N"
        Dim parts = expr.Split("/"c)
        If parts.Length = 2 AndAlso parts(0).Trim() = "TotalOutletCount" Then
            Return (total \ Integer.Parse(parts(1).Trim())) - 1
        End If
        Return Integer.Parse(expr)
    End Function

    Public Shared Function PlaceFlipCheckBox(ParameterValueID As Integer?, FlipCheckBoxes As Dictionary(Of ComboBox, CheckBox), cb As ComboBox, Fliplabels As Dictionary(Of ComboBox, Label))
        Dim fliplabel As New Label
        Dim flipCB As New CheckBox
        fliplabel.Text = "Flip: "
        flipCB.Text = "Flip"

        If DatabaseFunctions.GetFlipParameter(ParameterValueID) Then
            ' Create checkbox if not already created
            If Not FlipCheckBoxes.ContainsKey(cb) Then
                Dim childname As String = cb.Tag
                If childname.Contains("Outlet") Then
                    fliplabel.Location = New Point(cb.Location.X, cb.Location.Y + cb.Height + 26)
                    fliplabel.AutoSize = True
                    flipCB.Location = New Point(cb.Location.X + fliplabel.Width + 10, cb.Location.Y + cb.Height + 23)
                Else

                    fliplabel.Location = New Point(cb.Location.X, cb.Location.Y + cb.Height + 5)
                    fliplabel.AutoSize = True

                    flipCB.Location = New Point(cb.Location.X + fliplabel.Width + 10, cb.Location.Y + cb.Height + 2)
                End If
                cb.Parent.Controls.Add(fliplabel)
                cb.Parent.Controls.Add(flipCB)
                FlipCheckBoxes(cb) = flipCB
                Fliplabels(cb) = fliplabel
            Else
                FlipCheckBoxes(cb).Visible = True
                Fliplabels(cb).Visible = True
            End If
        Else
            ' Hide if it exists but not needed
            If FlipCheckBoxes.ContainsKey(cb) Then
                FlipCheckBoxes(cb).Visible = False
                Fliplabels(cb).Visible = False
            End If
        End If
    End Function

    Public Shared Function GetNestedBlockReferences(trans As Transaction, parentBr As BlockReference) As List(Of BlockReference)
        Dim nestedRefs As New List(Of BlockReference)

        Dim explodedObjects As New DBObjectCollection()
        parentBr.Explode(explodedObjects)

        For Each obj As DBObject In explodedObjects
            If TypeOf obj Is BlockReference Then
                nestedRefs.Add(DirectCast(obj, BlockReference))
            End If
        Next

        Return nestedRefs
    End Function

    Public Shared Function GetRuntimeNestedBlocks(trans As Transaction, parentBr As BlockReference) As Dictionary(Of BlockReference, String)
        Dim result As New Dictionary(Of BlockReference, String)
        Dim btr As BlockTableRecord = CType(trans.GetObject(parentBr.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)

        For Each id As ObjectId In btr
            Dim ent As Entity = TryCast(trans.GetObject(id, OpenMode.ForRead), Entity)
            If TypeOf ent Is BlockReference Then

                Dim nestedBr As BlockReference = CType(ent, BlockReference)

                If nestedBr.IsDynamicBlock Then
                    nestedBr.UpgradeOpen()
                    Dim dynprops As DynamicBlockReferencePropertyCollection = nestedBr.DynamicBlockReferencePropertyCollection
                    Dim dynDef As BlockTableRecord = trans.GetObject(nestedBr.DynamicBlockTableRecord, OpenMode.ForRead)
                    Dim name As String = dynDef.Name

                    ' Retrieve visibility parameters
                    For Each prop As DynamicBlockReferenceProperty In dynprops
                        If prop.PropertyName = "Visibility" Then
                            result.Add(nestedBr, name)
                        End If
                    Next
                End If
            End If
        Next

        Return result
    End Function


    Public Shared Function GetNestedAttribute(trans As Transaction, blkref As BlockReference)
        Dim Attribute As String
        Dim attRefs As AttributeCollection = blkref.AttributeCollection
        For Each attId As ObjectId In attRefs
            Dim attRef As AttributeReference = TryCast(trans.GetObject(attId, OpenMode.ForRead), AttributeReference)
            If attRef IsNot Nothing AndAlso attRef.Tag.ToUpper() = "PANEL" Then
                Attribute = attRef.TextString
                Exit For
            End If
        Next

        Return Attribute
    End Function

    Public Shared Function GetAttribute(trans As Transaction, blkref As BlockReference, attributetag As String)
        Dim attribute As String
        Dim attrefs As AttributeCollection = blkref.AttributeCollection
        For Each attId As ObjectId In attrefs
            Dim attref As AttributeReference = TryCast(trans.GetObject(attId, OpenMode.ForRead), AttributeReference)
            If attref IsNot Nothing AndAlso attref.Tag.ToString() = attributetag Then
                attribute = attref.TextString
                Exit For
            End If
        Next

        Return attribute
    End Function

    Public Shared Function AppendAttribute(actrans, br)
        If Not br.IsWriteEnabled Then
            br.UpgradeOpen()
        End If
        Dim btr As BlockTableRecord = DirectCast(actrans.GetObject(br.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)

        ' Check if the block definition contains attribute definitions.
        If btr.HasAttributeDefinitions Then
            For Each objId As ObjectId In btr
                Dim attDef As AttributeDefinition = TryCast(actrans.GetObject(objId, OpenMode.ForRead), AttributeDefinition)
                If attDef IsNot Nothing AndAlso Not attDef.Constant Then
                    ' Create a new attribute reference based on the attribute definition.
                    Dim attRef As New AttributeReference()
                    attRef.SetAttributeFromBlock(attDef, br.BlockTransform)

                    ' Optional: you might initialize the attribute value here if needed.
                    ' attRef.TextString = "Your Default Value Here"

                    ' Append the attribute reference to the block reference's attribute collection.
                    br.AttributeCollection.AppendAttribute(attRef)
                    actrans.AddNewlyCreatedDBObject(attRef, True)
                End If
            Next
        End If

    End Function

    ''' Gets the current visibility state of a block reference.
    Public Shared Function GetCurrentVisibilityState(blkRef As BlockReference) As String
        Dim dynProps As DynamicBlockReferencePropertyCollection = blkRef.DynamicBlockReferencePropertyCollection
        For Each prop As DynamicBlockReferenceProperty In dynProps
            If prop.PropertyName.Contains("Visibility") Then
                Return prop.Value.ToString()
            End If
        Next
        Return String.Empty
    End Function

    Public Shared Function GetMText(trans As Transaction, br As BlockReference)
        Dim blockdef As BlockTableRecord = TryCast(trans.GetObject(br.BlockTableRecord, OpenMode.ForWrite), BlockTableRecord)
        Dim mtextlist As New List(Of MText)

        For Each objId As ObjectId In blockdef
            Dim ent As Entity = TryCast(trans.GetObject(objId, OpenMode.ForWrite), Entity)
            If TypeOf ent Is MText Then
                If ent.Visible Then
                    Dim mtext As MText = CType(ent, MText)
                    mtextlist.Add(mtext)
                End If
            End If
        Next

        Return mtextlist
    End Function

    Public Shared Function GetParameterNames(br As BlockReference) As List(Of String)
        Dim paramNames As New List(Of String)()
        ' First check if the block reference is a dynamic block
        If br.IsDynamicBlock Then
            Dim dynProps As DynamicBlockReferencePropertyCollection = br.DynamicBlockReferencePropertyCollection
            ' Loop through each property in the collection
            For Each prop As DynamicBlockReferenceProperty In dynProps
                If prop.PropertyName = "Visibility" Then
                    paramNames.Add(prop.PropertyName)
                End If
            Next
        End If
        Return paramNames
    End Function

    Public Shared Function UpdateComboBoxParameter(blkref As BlockReference, parent As Control, ComboBoxTag As String)
        blkref.UpgradeOpen()
        Dim dynProps As DynamicBlockReferencePropertyCollection = blkref.DynamicBlockReferencePropertyCollection
        Dim activevalue As String
        For Each prop As DynamicBlockReferenceProperty In dynProps
            If prop.PropertyName = "Visibility" Then
                activevalue = prop.Value
            End If
        Next
        Dim CB As ComboBox = Functions.FindComboBoxByTag(parent, ComboBoxTag)
        If CB IsNot Nothing Then
            CB.SelectedItem = activevalue
        Else
            ' Handle case where ComboBox is not found
            MessageBox.Show("ComboBox with tag 'MainVisibility' not found.")
        End If

    End Function

    Public Shared Function FindComboBoxByTag(Parent As Control, tagname As String) As ComboBox
        For Each ctrl As Control In Parent.Controls
            If TypeOf ctrl Is ComboBox Then
                Dim cb As ComboBox = DirectCast(ctrl, ComboBox)
                If cb.Tag IsNot Nothing AndAlso cb.Tag.ToString().Equals(tagname, StringComparison.OrdinalIgnoreCase) Then
                    Return cb
                End If
            End If
        Next
        Return Nothing
    End Function

    Public Shared Function FindTextBoxByTag(Parent As Control, tagname As String) As TextBox
        For Each ctrl As Control In Parent.Controls
            If TypeOf ctrl Is TextBox Then
                Dim tb As TextBox = DirectCast(ctrl, TextBox)
                If tb.Tag IsNot Nothing AndAlso tb.Tag.ToString().Equals(tagname, StringComparison.OrdinalIgnoreCase) Then
                    Return tb
                End If
            End If
        Next
        Return Nothing
    End Function

    'Gather all the "same" type of components. For example Pump elements with different definitions for different views. In P203 there are 2 side views called P203_Pump_Element and one front view calle P203_Pump_Element_F
    Public Shared Function NormalizeBlockName(BaseName As String) As String
        Dim parts = BaseName.Split("_"c)
        If parts.Length <= 1 Then Return BaseName
        ' Rebuild all except last part if it's a direction label
        Dim suffixes = {"F", "Front", "Side", "Back"}
        If suffixes.Contains(parts.Last(), StringComparer.OrdinalIgnoreCase) Then
            Return String.Join("_", parts.Take(parts.Length - 1))
        End If
        Return BaseName
    End Function

    'Create almost all UI
    Public Shared Function CreateComboBoxes(Parent As Control, parentvisibilitystates As List(Of String), parentblockname As String, basepointx As Integer, basepointy As Integer, UIElementsTier2 As List(Of Object), OutletCBVisibilityHandler As EventHandler, FlipHandler As EventHandler) As ComboBox
        Dim ParentCB As New ComboBox
        Dim xOffset As Integer = 220
        Dim yOffset As Integer = 80
        Dim maxRows As Integer = 6
        Dim currentRow As Integer = 0
        Dim currentColumn As Integer = 0

        ' Create
        ' ComboBox if visibility states exist.
        If parentvisibilitystates IsNot Nothing AndAlso parentvisibilitystates.Count > 0 Then
            Dim ParentLabel As New Label
            With ParentLabel
                .Location = New Point(basepointx, basepointy)
                .Size = New Size(160, 20)
                .Text = "Select " & parentblockname
            End With
            Parent.Controls.Add(ParentLabel)
            UIElementsTier2.Add(ParentLabel)
            With ParentCB
                .Location = New Point(basepointx, basepointy + 20)
                .Size = New Size(180, 30)
                ParentCB.Text = "Select " & parentblockname
                ParentCB.Tag = "ParentVisibility"
            End With
            Parent.Controls.Add(ParentCB)
            UIElementsTier2.Add(ParentCB)
            For Each state As String In parentvisibilitystates
                ParentCB.Items.Add(state)
            Next
            currentRow = 1 ' move layout one row below the Outlet combo box.
            Dim ComponentType As String = DatabaseFunctions.GetParentAttribute(parentblockname, "COMPONENTTYPE")
            If ComponentType.ToString().ToUpper().Contains("METERING DEVICE") Then
                AddHandler ParentCB.SelectedIndexChanged, OutletCBVisibilityHandler
            End If
        End If
        ' Retrieve child block names from the DB.
        Dim childBlockNames As List(Of String) = DatabaseFunctions.GetChildBlockDefinitions(parentblockname)
        'Check if it's double nested
        Dim doublenested As String = DatabaseFunctions.GetParentAttribute(parentblockname, "DOUBLENESTED")

        Dim groupedBlocks As New Dictionary(Of String, List(Of String)) ' baseName -> list of full block names
        Dim groupedAmounts As New Dictionary(Of String, Integer)        ' baseName -> summed amount
        '
        For Each childBlock In childBlockNames
            Dim baseName = NormalizeBlockName(childBlock)

            If Not groupedBlocks.ContainsKey(baseName) Then
                groupedBlocks(baseName) = New List(Of String)
                groupedAmounts(baseName) = 0
            End If

            groupedBlocks(baseName).Add(childBlock)

            ' Get and add AMOUNTINPARENTBLOCK
            Dim amountRaw As String = DatabaseFunctions.GetChildAttribute(childBlock, "AMOUNTINPARENTBLOCK")
            Dim parsedAmount As Integer = 1 ' Default if missing
            If Not String.IsNullOrWhiteSpace(amountRaw) Then
                Dim tryParseAmount As Integer
                If Integer.TryParse(amountRaw, tryParseAmount) AndAlso tryParseAmount > 0 Then
                    parsedAmount = tryParseAmount
                End If
            End If
            groupedAmounts(baseName) += parsedAmount
        Next

        ' Create ComboBoxes for child blocks flagged as SINGLE or with no "PANEL" attribute.
        'For Each childBlock In childBlockNames
        For Each childblock In groupedBlocks.Keys
            Dim totalCombos As Integer = groupedAmounts(childblock)
            Dim blockDefs = groupedBlocks(childblock)

            Dim childGroup As String = DatabaseFunctions.GetChildAttribute(childblock, "PANEL")
            Dim amountRaw As String = DatabaseFunctions.GetChildAttribute(childblock, "AMOUNTINPARENTBLOCK")
            Dim doubleChildBlocks As List(Of String) = DatabaseFunctions.GetDoubleChildBlocks(childblock)

            If String.IsNullOrEmpty(childGroup) OrElse childGroup.Trim().ToUpper() = "SINGLE" Then
                For i As Integer = 1 To totalCombos
                    If currentRow >= maxRows Then
                        currentRow = 0
                        currentColumn += 1
                    End If

                    ' First remove prefix (everything up to the first underscore)
                    Dim cleanChildBlockName As String = childblock
                    Dim underscoreIndex As Integer = childblock.IndexOf("_")
                    If underscoreIndex >= 0 AndAlso underscoreIndex < childblock.Length - 1 Then
                        cleanChildBlockName = childblock.Substring(underscoreIndex + 1)
                    End If

                    ' Then remove ALL underscores from the remaining name
                    cleanChildBlockName = cleanChildBlockName.Replace("_", " ")

                    ' Add label for the first ComboBox only
                    If i = 1 Then
                        Dim headerLabel As New Label With {
                     .Size = New Drawing.Size(160, 20),
                    .Text = cleanChildBlockName,
                    .Location = New Drawing.Point(basepointx + (currentColumn * xOffset), basepointy + (currentRow * yOffset))
                }
                        Parent.Controls.Add(headerLabel)
                        UIElementsTier2.Add(headerLabel)
                        'currentRow += 1
                    End If

                    ' Create ComboBox
                    Dim childVisibilityStates As List(Of String) = DatabaseFunctions.GetChildVisibilityStatesFromDatabase(childblock)
                    Dim childCB As New ComboBox With {
                .Location = New Point(basepointx + currentColumn * xOffset, basepointy + (currentRow * yOffset) + 20),
                .Size = New Size(215, 30),
                .Text = "Select " & cleanChildBlockName & If(totalCombos > 1, $" #{i}", ""),
                .Tag = childblock & "_Visibility_" & i
            }

                    For Each state As String In childVisibilityStates
                        childCB.Items.Add(state)
                    Next

                    Parent.Controls.Add(childCB)
                    UIElementsTier2.Add(childCB)
                    'comboboxes2.Add(childCB)
                    AddHandler childCB.SelectedIndexChanged, FlipHandler
                    currentRow += 1


                    If doubleChildBlocks IsNot Nothing AndAlso doubleChildBlocks.Count > 0 Then
                        ' Just use the first one to extract parameter values for now
                        Dim firstdoublechild As String = doubleChildBlocks(0)
                        Dim visibilityValues As List(Of String) = DatabaseFunctions.GetDoubleChildVisibilityStatesFromDatabase(firstdoublechild)

                        Dim cleandoubleChildBlockName As String = firstdoublechild
                        Dim underscoreIndexdouble As Integer = firstdoublechild.IndexOf("_")
                        If underscoreIndexdouble >= 0 AndAlso underscoreIndexdouble < firstdoublechild.Length - 1 Then
                            cleandoubleChildBlockName = firstdoublechild.Substring(underscoreIndexdouble + 1)
                        End If

                        ' Remove duplicates if needed
                        visibilityValues = visibilityValues.Distinct().ToList()

                        ' Create ComboBox for double child block
                        Dim doubleChildCB As New ComboBox With {
                            .Location = New Point(basepointx + currentColumn * xOffset, basepointy + (currentRow * yOffset) + 20),
                            .Size = New Size(200, 30),
                            .Text = "Select " & cleandoubleChildBlockName,
                            .Tag = cleandoubleChildBlockName & "_Visibility_" & i
                        }

                        For Each val As String In visibilityValues
                            doubleChildCB.Items.Add(val)
                        Next

                        Parent.Controls.Add(doubleChildCB)
                        UIElementsTier2.Add(doubleChildCB)
                        Dim doubleBinding As New DoubleNestedBinding With {
                        .DoubleCombo = doubleChildCB,
                        .DoubleBlockCandidates = doubleChildBlocks ' this is your full "2L_Reservoir_Lid"
                          }
                        InsertComponent.ChildToDoubleBindingMap(childCB) = doubleBinding

                        currentRow += 1
                    End If
                Next
            End If
        Next

        Return ParentCB
    End Function

    Public Shared Sub MoveHatchToBack(actrans As Transaction, insertedblockref As BlockReference)
        Dim childblockrefs As Dictionary(Of BlockReference, String) = Functions.GetRuntimeNestedBlocks(actrans, insertedblockref)
        For Each kvp In childblockrefs
            Dim OutletChildBlockRefs = Functions.GetNestedBlockReferences(actrans, kvp.Key)
            For Each item In OutletChildBlockRefs
                Dim nestedblkref As BlockReference = item
                Dim dynBlockId As ObjectId = nestedblkref.DynamicBlockTableRecord
                Dim dynBlockDef As BlockTableRecord = DirectCast(actrans.GetObject(dynBlockId, OpenMode.ForRead), BlockTableRecord)
                Dim nestedBlockName As String = dynBlockDef.Name

                Dim draworder As DrawOrderTable = CType(actrans.GetObject(dynBlockDef.DrawOrderTableId, OpenMode.ForWrite), DrawOrderTable)
                Dim moveToBack As New ObjectIdCollection()
                For Each id In dynBlockDef
                    Dim ent As Entity = TryCast(actrans.GetObject(id, OpenMode.ForRead), Entity)
                    If TypeOf ent Is Hatch Then
                        Dim hatchEnt As Hatch = CType(ent, Hatch)
                        If ent IsNot Nothing AndAlso ent.Layer.ToLower() = "mask_hatch" Then
                            moveToBack.Add(id)
                        End If
                    End If
                Next
                If moveToBack.Count > 0 Then
                    draworder.MoveToBottom(moveToBack)
                End If
            Next
        Next
    End Sub

    Public Shared Function MoveBROrder(actrans As Transaction, insertedblockref As BlockReference, selectedcomponentblock As String)
        Dim childblocks = Functions.GetRuntimeNestedBlocks(actrans, insertedblockref)
        For Each kvp In childblocks
            Dim childbr As BlockReference = kvp.Key
            Dim dynBlockId As ObjectId = insertedblockref.DynamicBlockTableRecord
            Dim dynBlockId2 As ObjectId = childbr.DynamicBlockTableRecord
            Dim dynBlockDef As BlockTableRecord = DirectCast(actrans.GetObject(dynBlockId, OpenMode.ForRead), BlockTableRecord)
            Dim dynBlockDef2 As BlockTableRecord = DirectCast(actrans.GetObject(dynBlockId2, OpenMode.ForRead), BlockTableRecord)
            Dim ChildBlockName As String = dynBlockDef2.Name
            Dim moveback As String
            moveback = DatabaseFunctions.GetChildAttribute(ChildBlockName, "MOVETOBACK")
            If moveback IsNot Nothing AndAlso moveback.Trim().ToUpper() = "TRUE" Then
                Dim draworder As DrawOrderTable = CType(actrans.GetObject(dynBlockDef.DrawOrderTableId, OpenMode.ForWrite), DrawOrderTable)
                Dim moveToBack As New ObjectIdCollection()
                moveToBack.Add(childbr.ObjectId)
                If moveToBack.Count > 0 Then
                    draworder.MoveToBottom(moveToBack)
                End If
            End If
            Dim movefront = DatabaseFunctions.GetChildAttribute(ChildBlockName, "MOVETOFRONT")
            If movefront IsNot Nothing AndAlso movefront.Trim().ToUpper() = "TRUE" Then
                Dim draworder As DrawOrderTable = CType(actrans.GetObject(dynBlockDef.DrawOrderTableId, OpenMode.ForWrite), DrawOrderTable)
                Dim moveToFront As New ObjectIdCollection()
                moveToFront.Add(childbr.ObjectId)
                If moveToFront.Count > 0 Then
                    draworder.MoveToTop(moveToFront)
                End If
            End If
        Next
    End Function

    Public Shared Sub UpdateMText(Parent As Control, MTextList As List(Of MText))
        If MTextList.Count > 5 Then
            Dim cbMains As ComboBox = Functions.FindComboBoxByTag(Parent, "MainVisibility")
            Dim numToUpdate As Integer = CInt(cbMains.SelectedItem) \ 2      ' For example, if SelectedItem is 14, then numToUpdate = 7.
            Dim startIndex As Integer = MTextList.Count - numToUpdate        ' For an 11-item list, startIndex = 11 - numToUpdate.
            Dim tbcounter As Integer = 1

            'Change Mtext
            For i As Integer = startIndex + 1 To MTextList.Count      ' tbcounter here is 1-based (e.g., 5 to 11 when numToUpdate is 7)
                Dim mtext As MText = MTextList(i - 1)                ' Convert tbcounter to 0-based index.
                Dim tbTagName As String = "MText" & tbcounter                   ' Example: "MText5", "MText6", etc.
                Dim tb As TextBox = FindTextBoxByTag(Parent, tbTagName)

                If tb IsNot Nothing AndAlso IsNumeric(tb.Text) Then
                    mtext.Contents = tb.Text
                End If
                tbcounter += 1
            Next
        End If
    End Sub

    Public Shared Sub UpdateParentVisibility(parent As Control, blkref As BlockReference)
        If blkref.IsDynamicBlock Then
            ' Get the dynamic properties for the main block.
            Dim dynProps As DynamicBlockReferencePropertyCollection = blkref.DynamicBlockReferencePropertyCollection
            ' Loop through each dynamic property and update if a combobox exists for it.
            For Each prop As DynamicBlockReferenceProperty In dynProps
                ' Check if the property is one of the ones controlling visibility.
                If prop.PropertyName.ToUpper().Contains("VISIBILITY") Then
                    ' Look for a combobox whose Tag matches this dynamic property name.
                    Dim cbOutlet As ComboBox = Functions.FindComboBoxByTag(parent, "ParentVisibility")
                    If cbOutlet IsNot Nothing AndAlso cbOutlet.SelectedItem IsNot Nothing Then
                        prop.Value = cbOutlet.SelectedItem
                    End If
                End If
            Next
        End If
    End Sub

    Public Shared Function GetVisibleChildren(blkref As BlockReference)
        Dim VisibleChildren As List(Of BlockReference)
        If blkref.Visible Then
            VisibleChildren.Add(blkref)
        End If
    End Function

    Public Structure ParameterDef
        Public ParameterID As Integer
        Public ParameterName As String
        Public ParentBlockID As Integer?
        Public ChildBlockID As Integer?
    End Structure

    Public Shared Function InsertCloneBR(folderpath As String, selectedcomponent As String, selectedcomponentblock As String, accurdb As Database, inspoint As Point3d)
        Dim br As BlockReference
        Dim blockid As ObjectId = ObjectId.Null

        ' Find the next available clone name in the current drawing
        Dim cloneIndex As Integer = 1
        Dim NewBlockName As String = selectedcomponentblock & "Clone" & cloneIndex

        Using acTrans As Transaction = accurdb.TransactionManager.StartTransaction()
            Dim bt As BlockTable = acTrans.GetObject(accurdb.BlockTableId, OpenMode.ForRead)
            Do While bt.Has(NewBlockName)
                cloneIndex += 1
                NewBlockName = selectedcomponentblock & "Clone" & cloneIndex
            Loop
            acTrans.Commit()
        End Using

        Using sourceDb As New Database(False, True)
            Dim sourceFile = Directory.GetFiles(folderpath,
                                                selectedcomponent & ".dwg",
                                                SearchOption.AllDirectories).FirstOrDefault()
            sourceDb.ReadDwgFile(sourceFile, FileOpenMode.OpenForReadAndWriteNoShare, False, "")
            sourceDb.CloseInput(True)  ' free file lock

            'Rename Source block name
            Using trSrc As Transaction = sourceDb.TransactionManager.StartTransaction()
                Dim srcBt As BlockTable = trSrc.GetObject(sourceDb.BlockTableId, OpenMode.ForWrite)
                If srcBt.Has(selectedcomponentblock) Then
                    Dim srcBlock As BlockTableRecord = trSrc.GetObject(srcBt(selectedcomponentblock), OpenMode.ForWrite)
                    srcBlock.Name = NewBlockName  ' e.g., "SSV_BlockClone1"
                Else
                    Throw New Exception($"Source block {selectedcomponent} not found in SSV.dwg")
                End If


                ' Clone the renamed block from source to destination, replacing existing definition
                Dim srcBt2 As BlockTable = trSrc.GetObject(sourceDb.BlockTableId, OpenMode.ForRead)
                Dim srcCloneDefId As ObjectId = srcBt2(NewBlockName)
                Dim map As New IdMapping()
                accurdb.WblockCloneObjects(New ObjectIdCollection({srcCloneDefId}), accurdb.BlockTableId, map, DuplicateRecordCloning.Replace, False)


                ' Restore source block name for next iteration
                Dim srcBt3 As BlockTable = trSrc.GetObject(sourceDb.BlockTableId, OpenMode.ForWrite)
                Dim clonedBlock As BlockTableRecord = trSrc.GetObject(srcBt3(NewBlockName), OpenMode.ForWrite)
                clonedBlock.Name = selectedcomponentblock
                Dim tempBlockId As ObjectId = srcBt3(selectedcomponentblock)

                blockid = map(tempBlockId).Value

                ' Create and return a new block reference at the origin.
                br = New BlockReference(inspoint, blockid)
                trSrc.Commit()
            End Using
        End Using
        Return br
    End Function

    'Insert component with ID Code
    Public Shared Sub InsertBlockWithParameters(folderpath As String, filename As String, blockName As String, paramValues As List(Of Tuple(Of String, Object)), poscheckbox As CheckBox, cfg As DeviceConfig)
        Dim paramDefs = DatabaseFunctions.GetAllParameterDefinitions()

        Dim acDoc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim acCurDb As Autodesk.AutoCAD.DatabaseServices.Database = acDoc.Database
        Dim ed As Editor = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor
        Dim filepath As String = FindDWGFilePath(folderpath, filename)

        Using docLock As DocumentLock = acDoc.LockDocument()
            Using acTrans As Transaction = acCurDb.TransactionManager.StartTransaction()
                Dim bt As BlockTable = acTrans.GetObject(acCurDb.BlockTableId, OpenMode.ForRead)
                Dim inspoint As Point3d = New Point3d(0, 0, 0)
                ' If checkbox is checked, ask user for insert point
                If poscheckbox.Checked Then
                    Dim ppo As New PromptPointOptions(vbLf & "Specify insertion point:")
                    Dim ppr As PromptPointResult = ed.GetPoint(ppo)
                    If ppr.Status = PromptStatus.OK Then
                        inspoint = ppr.Value
                    Else
                        ed.WriteMessage(vbLf & "Insertion point not selected. Command canceled.")
                        Return
                    End If
                End If

                Dim br As BlockReference = Functions.InsertCloneBR(folderpath, filename, blockName, acCurDb, inspoint)
                Dim ms As BlockTableRecord = acTrans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                Functions.MoveHatchToBack(acTrans, br)
                Functions.MoveBROrder(acTrans, br, blockName)
                ms.AppendEntity(br)
                acTrans.AddNewlyCreatedDBObject(br, True)
                Functions.AppendAttribute(acTrans, br)

                ed.WriteMessage(vbLf & "Block '" & blockName & "' inserted successfully at " & inspoint.ToString())
                Dim outletcount As Integer
                'Update ParentBlock Visibility
                If br.IsDynamicBlock Then
                    Dim dynProps As DynamicBlockReferencePropertyCollection = br.DynamicBlockReferencePropertyCollection
                    For Each kv In paramValues.Where(Function(p) String.Equals(p.Item1, blockName, StringComparison.OrdinalIgnoreCase))

                        Dim val = kv.Item2
                        outletcount = kv.Item2
                        For Each prop As DynamicBlockReferenceProperty In dynProps
                            If prop.PropertyName.ToUpper().Contains("VISIBILITY") Then
                                If val IsNot Nothing Then
                                    prop.Value = val
                                    Exit For
                                End If
                            End If
                        Next
                    Next
                End If
                Dim nestedblocks As New Dictionary(Of BlockReference, String)
                nestedblocks = Functions.GetRuntimeNestedBlocks(acTrans, br)

                'Update Children Visibility
                For Each pair In nestedblocks
                    Dim childRef As BlockReference = pair.Key
                    Dim childName As String = pair.Value

                    ' 1) Try to read a Position attribute (if it exists)
                    Dim lookupKey As String = childName
                    Try
                        Dim posObj = Functions.GetAttribute(acTrans, childRef, "POSITION")
                        Dim pos As Integer
                        If posObj IsNot Nothing AndAlso
           Integer.TryParse(posObj.ToString(), pos) Then
                            If cfg IsNot Nothing AndAlso cfg.StartAt IsNot Nothing AndAlso cfg.StartAt.Trim().ToLower().Contains("bottom") Then
                                Dim adjustedPosition As Integer = pos - (cfg.TotalOutletCount - outletcount)
                                lookupKey = childName & adjustedPosition.ToString()
                            Else
                                ' Only if we successfully parse an integer do we append it
                                lookupKey &= pos.ToString()
                            End If
                        End If
                    Catch ex As KeyNotFoundException
                        ' GetAttribute might throw if the attribute isn't there;
                        ' in that case we just leave lookupKey = featureName
                    End Try

                    ' 2) Find the matching tuple
                    Dim match = paramValues.FirstOrDefault(Function(p) String.Equals(p.Item1, lookupKey, StringComparison.OrdinalIgnoreCase))
                    If match Is Nothing OrElse match.Item2 Is Nothing Then
                        Continue For
                    End If
                    Dim childVal = match.Item2

                    If childVal Is Nothing Then
                        Continue For
                    End If
                    ' only update dynamic children:
                    If childRef.IsDynamicBlock Then
                        For Each prop As DynamicBlockReferenceProperty In childRef.DynamicBlockReferencePropertyCollection
                            If prop.PropertyName.ToUpper().Contains("VISIBILITY") Then
                                prop.Value = childVal
                                Exit For
                            End If
                        Next
                    End If
                Next
                Functions.AssignIdCode(acTrans, br, blockName)
                acDoc.Editor.Regen()
                acTrans.Commit()
            End Using
        End Using
    End Sub

    'Matches tagname of combobox to correct block reference in autocad (name)
    Public Shared Function GetTagName(ComboBox As ComboBox, Name As String, actrans As Transaction, childblkref As BlockReference, cfg As DeviceConfig, childblockname As String)
        'In the case of a single combobox 
        Dim tagname As String = Name & "_Visibility_" & "1"
        'For Outlet Fittings, Couple combobox to correct block with this calculation using Position Attribute
        Dim PositionAtt As String = Functions.GetAttribute(actrans, childblkref, "POSITION")
        Dim Position As Integer
        If Integer.TryParse(PositionAtt, Position) Then
            If ComboBox.SelectedItem IsNot Nothing Then
                Dim OutletCount As Integer
                Integer.TryParse(ComboBox.SelectedItem.ToString(), OutletCount)
                Dim Mtext As String = DatabaseFunctions.GetChildAttribute(Name, "MTEXT")
                'Looks at cases where there are the same amount of outlets as outlet fittings instances 
                If cfg.StartAt IsNot Nothing AndAlso cfg.StartAt.Trim().ToLower().Contains("bottom") AndAlso Not String.Equals(Mtext, "True", StringComparison.OrdinalIgnoreCase) Then
                    Dim adjustedPosition As Integer = Position - (cfg.TotalOutletCount - OutletCount)
                    tagname = Name & "_Visibility_" & adjustedPosition.ToString()
                Else
                    tagname = Name & "_Visibility_" & Position
                End If
            Else
                'Works for Pumps and Non Metering Devices with Position Attribute 
                tagname = Name & "_Visibility_" & PositionAtt
            End If
        End If
        Return tagname
    End Function
    Public Shared Function FindDWGFilePath(folderPath As String, filename As String) As String
        If String.IsNullOrWhiteSpace(folderPath) OrElse String.IsNullOrWhiteSpace(filename) Then
            Return Nothing
        End If

        If Not Directory.Exists(folderPath) Then
            MessageBox.Show("Folder not found: " & folderPath, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return Nothing
        End If

        ' Normalize filename to ensure it ends with .dwg
        If Not filename.ToLower().EndsWith(".dwg") Then
            filename &= ".dwg"
        End If

        ' Recursively search for the file
        Dim matches = Directory.GetFiles(folderPath, filename, SearchOption.AllDirectories)

        ' Return the first match if found
        If matches.Length > 0 Then
            Return matches(0)
        End If

        ' Not found
        Return Nothing
    End Function
    Public Shared Sub AssignIdCode(tr As Transaction, NewBlockRef As BlockReference, selectedcomponentBlock As String)

        Dim idCode = GenerateIdCode(tr, NewBlockRef, selectedcomponentBlock)

        If NewBlockRef.AttributeCollection IsNot Nothing AndAlso NewBlockRef.AttributeCollection.Count > 0 Then
            Console.WriteLine($"AttributeCollection Count: {NewBlockRef.AttributeCollection.Count}")
            For Each attId As ObjectId In NewBlockRef.AttributeCollection
                Dim attRef As AttributeReference = TryCast(tr.GetObject(attId, OpenMode.ForWrite), AttributeReference)
                If attRef IsNot Nothing AndAlso attRef.Tag.ToUpper() = "IDENTIFICATION_CODE" Then
                    attRef.TextString = idCode
                End If
            Next
        Else
            Console.WriteLine("No attributes found in AttributeCollection.")
        End If

    End Sub

    Public Shared Sub UpdateParameters(Childbr As BlockReference, parent As Control, tagname As String, flipcheckboxes As Dictionary(Of ComboBox, CheckBox))
        Dim nesteddynprops As DynamicBlockReferencePropertyCollection = Childbr.DynamicBlockReferencePropertyCollection
        For Each prop As DynamicBlockReferenceProperty In nesteddynprops
            ' all nested blocks have a property named "visibility".
            If prop.PropertyName.ToUpper().Contains("VISIBILITY") Then
                Dim cbchild As ComboBox = Functions.FindComboBoxByTag(parent, tagname)
                If cbchild IsNot Nothing AndAlso cbchild.SelectedItem IsNot Nothing Then
                    Dim SelectedItem = cbchild.SelectedItem.ToString()
                    prop.Value = SelectedItem
                End If
            ElseIf prop.PropertyName.ToUpper().Contains("FLIP") Then
                Functions.FlipBlockReference(parent, tagname, prop, flipcheckboxes)
            End If
        Next
    End Sub

    Public Shared Function FlipBlockReference(parent As Control, tagname As String, prop As DynamicBlockReferenceProperty, FlipCheckboxes As Dictionary(Of ComboBox, CheckBox))
        ' Find the correct flip checkbox linked to this nested block's ComboBox
        Dim cbnested As ComboBox = Functions.FindComboBoxByTag(parent, tagname)
        Dim allowedValues = prop.GetAllowedValues()
        If allowedValues IsNot Nothing AndAlso allowedValues.Count > 0 Then
            For Each vals In allowedValues
                Console.Write(vbLf & "Allowed flip value: " & vals.ToString())
            Next
        End If

        If cbnested IsNot Nothing AndAlso FlipCheckboxes.ContainsKey(cbnested) Then
            Dim flipcheckbox As CheckBox = FlipCheckboxes(cbnested)

            If flipcheckbox.Checked Then
                prop.Value = allowedValues(1)
            Else
                prop.Value = allowedValues(0)
            End If
        End If
    End Function

    Public Shared Sub SafeDisposeAndRemoveFromControls(item As Object, parent As Control)
        Dim ctrl As Control = TryCast(item, Control)
        If ctrl IsNot Nothing Then
            If parent.Controls.Contains(ctrl) Then
                parent.Controls.Remove(ctrl)
            End If
            ' Check if it's one of the class-level fields that might be disposed and re-created
            ' This check is more relevant if they weren't always re-created with New.
            ' With the current approach of always using New in BuildMeteringDeviceControls,
            ' simply disposing is fine.
            If Not ctrl.IsDisposed Then
                ctrl.Dispose()
            End If
        Else
            ' Handle other IDisposable objects if they are not controls
            Dim disp As IDisposable = TryCast(item, IDisposable)
            If disp IsNot Nothing Then
                disp.Dispose()
            End If
        End If
    End Sub

    ' Helper to clear a list of UI elements
    Public Shared Sub ClearUIElementList(elementList As List(Of Object), parent As Control)
        Dim itemsToClear As New List(Of Object)(elementList) ' Iterate over a copy
        elementList.Clear() ' Clear original list immediately

        For Each item In itemsToClear
            SafeDisposeAndRemoveFromControls(item, parent)
        Next
    End Sub

    Public Shared Function BuildMeteringDeviceControls(Parent As Control, OutletCB As ComboBox, cfg As DeviceConfig, UIElementstier3 As List(Of Object), parentkey As String, ChildCB_VisChange As EventHandler, CheckBoxChecked As EventHandler, outletcount As Integer)
        Dim meteringdevicepanel As Panel
        Dim Checkboxall As CheckBox

        ' 1. configure panel size & position
        ' Dispose of the previous panel instance, if any, and create a new one.
        If meteringdevicepanel IsNot Nothing Then
            If Parent.Controls.Contains(meteringdevicepanel) Then
                Parent.Controls.Remove(meteringdevicepanel)
            End If
            If Not meteringdevicepanel.IsDisposed Then
                meteringdevicepanel.Dispose()
            End If
        End If
        meteringdevicepanel = New Panel() ' *** Instantiate a new panel ***


        meteringdevicepanel.Location = New Point(OutletCB.Location.X + 450, OutletCB.Location.Y)
        ' e.g. height = rows * spacing
        Dim outletsPerSide As Integer = outletcount / cfg.NumberOfSides
        Dim spacingY As Integer = 65
        meteringdevicepanel.Size = New Size(300, outletsPerSide * spacingY)
        meteringdevicepanel.BorderStyle = BorderStyle.FixedSingle
        Parent.Controls.Add(meteringdevicepanel)
        UIElementstier3.Add(meteringdevicepanel)

        ' 2. outlet comboboxes
        Dim ChildBlockName As String = cfg.ChildBlockName
        Dim allowedValues = DatabaseFunctions.GetChildVisibilityStatesFromDatabase(ChildBlockName)
        Dim sidePositions = Functions.GetSidePositions(cfg, outletcount, meteringdevicepanel)

        For i As Integer = 1 To outletcount
            Dim pos = sidePositions(i)
            Dim cb As New ComboBox With {
            .Location = pos,
            .Size = New Size(190, 30),
            .Text = $"Select Outlet Fitting {i}",
            .Tag = $"{ChildBlockName}_Visibility_{i}"
        }
            allowedValues.ForEach(Sub(v) cb.Items.Add(v))
            Parent.Controls.Add(cb)
            UIElementstier3.Add(cb)
            AddHandler cb.SelectedIndexChanged, ChildCB_VisChange

            If cfg.SubElements IsNot Nothing Then
                Dim verticalOffset As Integer = 22  ' spacing between main and subelement CB
                For Each subEl In cfg.SubElements
                    Dim subCb As New ComboBox With {
                        .Location = New Point(pos.X, pos.Y + verticalOffset),
                        .Size = New Size(190, 30),
                        .Text = subEl.PromptText,
                        .Tag = $"{subEl.ChildBlockName}_Visibility_{i}"
                    }

                    ' Populate with visibility states from sub-element block
                    Dim subVals = DatabaseFunctions.GetChildVisibilityStatesFromDatabase(subEl.ChildBlockName)
                    subVals.ForEach(Sub(v) subCb.Items.Add(v))

                    Parent.Controls.Add(subCb)
                    UIElementstier3.Add(subCb)
                    verticalOffset += 35 ' stack more subelements if needed
                Next
            End If
        Next

        ' 3. any middle elements (e.g. metering screws)
        Dim midpointX = meteringdevicepanel.Left + meteringdevicepanel.Width \ 2
        For Each meCfg In cfg.MiddleElements
            ' check parent attribute
            If DatabaseFunctions.GetChildAttribute(meCfg.ChildBlockName, meCfg.AttributeName) = "True" Then
                Dim count = Functions.EvaluateExpression(meCfg.CountExpression, outletcount)
                Dim yOffset As Double = (meteringdevicepanel.Height - 30) \ count
                For j As Integer = 1 To count + 1
                    Dim cb As New ComboBox With {
                    .Size = New Size(190, 30),
                    .Location = New Point(midpointX - 95, meteringdevicepanel.Top + (yOffset * (j - 1))),
                    .Text = meCfg.PromptText,
                    .Tag = $"{meCfg.ChildBlockName}_Visibility_{j}"
                }
                    Dim vals = DatabaseFunctions.GetChildVisibilityStatesFromDatabase(meCfg.ChildBlockName)
                    vals.ForEach(Sub(v) cb.Items.Add(v))
                    Parent.Controls.Add(cb)
                    cb.BringToFront()
                    UIElementstier3.Add(cb)
                Next
            End If
        Next

        '4 all controls
        Dim alllbl As New Label With {
        .TextAlign = ContentAlignment.MiddleCenter,
        .Size = New Drawing.Size(190, 20),
        .Location = New Drawing.Point(midpointX - 90, meteringdevicepanel.Bottom + 10),
        .Text = "Change All Outlet Fittings"
        }

        Parent.Controls.Add(alllbl)
        UIElementstier3.Add(alllbl)
        Dim allCb As New ComboBox With {
        .Size = New Size(190, 30),
        .Location = New Point(midpointX - 95, meteringdevicepanel.Bottom + 35),
        .Text = "Select All Outlet Fittings",
        .Tag = $"{ChildBlockName}_All_Visibility"
    }
        allowedValues.ForEach(Sub(v) allCb.Items.Add(v))
        Parent.Controls.Add(allCb)
        UIElementstier3.Add(allCb)
        AddHandler allCb.SelectedIndexChanged, CheckBoxChecked
        'Return Checkboxall
    End Function

    Public Shared Function AddDoubleChildToBRList(selectedcomponentblock As String, nestedblocks As Dictionary(Of BlockReference, String), actrans As Transaction)
        Dim DoubleNestedAttr As String = DatabaseFunctions.GetParentAttribute(selectedcomponentblock, "DOUBLENESTED")
        ' Check if we need to retrieve double nested visibility states.
        If DoubleNestedAttr = "True" Then
            ' Iterate over a copy of the current keys so that we can safely add new ones.
            For Each kvp In nestedblocks.ToList()
                ' Get the nested visibility states for this inner block reference.
                Dim innerblocks As Dictionary(Of BlockReference, String) = Functions.GetRuntimeNestedBlocks(actrans, kvp.Key)
                For Each innerKvp As KeyValuePair(Of BlockReference, String) In innerblocks
                    ' Only add if it is not already in the dictionary.
                    If Not nestedblocks.ContainsKey(innerKvp.Key) Then
                        nestedblocks.Add(innerKvp.Key, innerKvp.Value)
                    End If
                Next
            Next
        End If
    End Function

    Public Shared Sub CheckBoxChecked(sender As Object, UIElementsTier3 As List(Of Object))
        Dim sourceComboBox As ComboBox = TryCast(sender, ComboBox)
        If sourceComboBox Is Nothing OrElse sourceComboBox.SelectedItem Is Nothing Then
            ' Sender is not the expected ComboBox or no item is selected in it.
            Exit Sub
        End If

        Dim selectedValueToApply As Object = sourceComboBox.SelectedItem
        Dim selectedIndexToApply As Integer = sourceComboBox.SelectedIndex ' Useful for fallback or if items are simple strings

        ' Determine the base ChildBlockName from the sourceComboBox's Tag.
        Dim sourceTagString As String = If(sourceComboBox.Tag IsNot Nothing, sourceComboBox.Tag.ToString(), String.Empty)
        If String.IsNullOrEmpty(sourceTagString) OrElse Not sourceTagString.EndsWith("_All_Visibility") Then
            ' Source ComboBox's Tag is not in the expected format. Cannot reliably identify target ComboBoxes.
            Exit Sub
        End If

        Dim childBlockName As String = sourceTagString.Substring(0, sourceTagString.Length - "_All_Visibility".Length)

        Dim targetTagPrefix As String = childBlockName & "_Visibility_"

        ' Iterate through UIElementsTier3 to find and update the individual outlet fitting ComboBoxes.
        For Each controlElement As Control In UIElementsTier3
            ' Check if the current element is a ComboBox
            If TypeOf controlElement Is ComboBox Then
                Dim targetComboBox As ComboBox = DirectCast(controlElement, ComboBox)

                ' Check if this ComboBox is an individual outlet fitting ComboBox:
                ' - It must not be the sourceComboBox itself.           
                If targetComboBox IsNot sourceComboBox AndAlso
               targetComboBox.Tag IsNot Nothing Then

                    Dim targetTagString As String = targetComboBox.Tag.ToString()
                    If targetTagString.StartsWith(targetTagPrefix) AndAlso targetTagString.Length > targetTagPrefix.Length Then
                        ' This ComboBox is identified as an individual outlet fitting ComboBox.
                        ' Set its SelectedItem to match the sourceComboBox's SelectedItem.                         
                        If targetComboBox.Items.Contains(selectedValueToApply) Then
                            targetComboBox.SelectedItem = selectedValueToApply
                        Else
                            ' Fallback
                            If selectedIndexToApply >= 0 AndAlso selectedIndexToApply < targetComboBox.Items.Count Then
                                targetComboBox.SelectedIndex = selectedIndexToApply
                            Else
                            End If
                        End If
                    End If
                End If
            End If
        Next
    End Sub

End Class

Public Class DualMapEntry
    Public Property article As String
    Public Property codeValue As String
End Class