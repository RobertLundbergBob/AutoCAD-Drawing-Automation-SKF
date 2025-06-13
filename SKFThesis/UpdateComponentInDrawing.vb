Imports Autodesk.AutoCAD.ApplicationServices
Imports System.IO
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.Runtime
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.TrackBar
Imports System.Runtime.InteropServices.ComTypes
Imports Autodesk.AutoCAD.Geometry

Public Class UpdateComponentInDrawing

    Public Shared Sub UpdateClonedBlocksFromExternal(sourcelibraryfolder)
        Dim acadDoc As Document = Application.DocumentManager.MdiActiveDocument
        Dim db As Database = acadDoc.Database
        acadDoc.Editor.WriteMessage(vbLf & "Updating cloned block definitions..." & vbLf)
        ' Lock the document to prevent changes by the user during update
        Using acadDoc.LockDocument()
            ' Identify all clone block names
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Dim CurrentStateDict As New Dictionary(Of BlockReference, String)
                Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim prefixes As New List(Of String)
                For Each btrId As ObjectId In bt
                    Dim btr As BlockTableRecord = CType(btrId.GetObject(OpenMode.ForRead), BlockTableRecord)
                    Dim cloneName = btr.Name
                    If Not cloneName.Contains("Clone") Then Continue For

                    ' Open the source DWG containing the master block definitions
                    Using sourceDb As New Database(False, True)
                        Dim baseName = cloneName.Substring(0, cloneName.IndexOf("Clone", StringComparison.Ordinal))
                        Dim baseType = baseName.Split("_"c)(0)
                        Dim prefix = baseType & "_"
                        prefixes.Add(prefix)
                        Dim sourceFile = Directory.GetFiles(sourcelibraryfolder,
                                                        baseType & ".dwg",
                                                        SearchOption.AllDirectories).FirstOrDefault()
                        sourceDb.ReadDwgFile(sourceFile, FileOpenMode.OpenForReadAndWriteNoShare, False, "")
                        sourceDb.CloseInput(True)  ' free file lock
                        ' Save all current reference states (dynamic props & attributes) for this block
                        Dim refStates As New List(Of BlockRefState)
                        Using trans As Transaction = db.TransactionManager.StartTransaction()
                            Dim btDest As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)
                            Dim cloneDef As BlockTableRecord = trans.GetObject(btDest(cloneName), OpenMode.ForRead)
                            Dim bt2 As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)
                            Dim ms3 As BlockTableRecord = DirectCast(trans.GetObject(bt2(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)
                            'Dim CloneOBJColl As ObjectIdCollection = cloneDef.GetBlockReferenceIds(True, False)
                            For Each id As ObjectId In ms3
                                Dim ent As Entity = TryCast(trans.GetObject(id, OpenMode.ForRead), Entity)
                                If TypeOf ent Is BlockReference Then
                                    Dim br As BlockReference = DirectCast(ent, BlockReference)
                                    Dim brbtr As BlockTableRecord = DirectCast(trans.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                                    Dim Name As String = brbtr.Name
                                    If String.Equals(Name, cloneName, StringComparison.OrdinalIgnoreCase) Then
                                        refStates.Add(New BlockRefState(br, trans))
                                        br.UpgradeOpen()
                                        br.Erase()
                                    End If
                                End If
                            Next
                            trans.Commit()
                        End Using

                        ' Rename source block to match the clone name
                        Using trSrc As Transaction = sourceDb.TransactionManager.StartTransaction()
                            Dim srcBt As BlockTable = trSrc.GetObject(sourceDb.BlockTableId, OpenMode.ForWrite)
                            If srcBt.Has(baseName) Then
                                Dim srcBlock As BlockTableRecord = trSrc.GetObject(srcBt(baseName), OpenMode.ForWrite)
                                srcBlock.Name = cloneName  ' e.g., "SSV_BlockClone1"
                            Else
                                Throw New Exception($"Source block {baseName} not found in SSV.dwg")
                            End If
                            trSrc.Commit()
                        End Using

                        ' Clone the renamed block from source to destination, replacing existing definition
                        Using trClone As Transaction = sourceDb.TransactionManager.StartTransaction()
                            Dim srcBt2 As BlockTable = trClone.GetObject(sourceDb.BlockTableId, OpenMode.ForRead)
                            Dim srcCloneDefId As ObjectId = srcBt2(cloneName)
                            trClone.Commit()
                            Dim map As New IdMapping()
                            db.WblockCloneObjects(New ObjectIdCollection({srcCloneDefId}), db.BlockTableId, map, DuplicateRecordCloning.Replace, False)
                        End Using

                        ' Restore source block name for next iteration
                        Using trSrcReset As Transaction = sourceDb.TransactionManager.StartTransaction()
                            Dim srcBt3 As BlockTable = trSrcReset.GetObject(sourceDb.BlockTableId, OpenMode.ForWrite)
                            Dim clonedBlock As BlockTableRecord = trSrcReset.GetObject(srcBt3(cloneName), OpenMode.ForWrite)
                            clonedBlock.Name = baseName
                            trSrcReset.Commit()
                        End Using

                        Using trReplace As Transaction = db.TransactionManager.StartTransaction()
                            For Each state As BlockRefState In refStates
                                Dim bt2 As BlockTable = trReplace.GetObject(db.BlockTableId, OpenMode.ForRead)
                                Dim btrId2 As ObjectId = bt2(state.BlockName)
                                Dim ms2 As BlockTableRecord = trReplace.GetObject(bt2(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                                Dim newbr As BlockReference = ReplaceBlockReference(db, trReplace, state)
                                Functions.MoveHatchToBack(trReplace, newbr)
                                'Functions.MoveBRToBack(trReplace, newbr)
                                Dim parentDef As BlockTableRecord = CType(trReplace.GetObject(newbr.BlockTableRecord, OpenMode.ForWrite), BlockTableRecord)
                                Dim drawOrder As DrawOrderTable = CType(trReplace.GetObject(parentDef.DrawOrderTableId, OpenMode.ForWrite), DrawOrderTable)

                                Dim moveToBack As New ObjectIdCollection()

                                For Each id In parentDef
                                    Dim ent = TryCast(trReplace.GetObject(id, OpenMode.ForRead), Entity)
                                    If TypeOf ent Is BlockReference Then
                                        Dim br As BlockReference = CType(ent, BlockReference)
                                        Dim childDef As BlockTableRecord = CType(trReplace.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                                        Dim Childblockname As String = childDef.Name
                                        Dim blockname As String = "indicator_pin"
                                        If ChildBlockName.IndexOf(Blockname, StringComparison.OrdinalIgnoreCase) >= 0 Then
                                            moveToBack.Add(br.ObjectId)
                                        End If
                                    End If
                                Next
                                If moveToBack.Count > 0 Then
                                    drawOrder.MoveToBottom(moveToBack)
                                End If
                            Next
                            trReplace.Commit()
                        End Using
                    End Using
                    acadDoc.Editor.WriteMessage($"Updated block definition: {cloneName}" & vbLf)
                Next  ' next cloneName
                Application.DocumentManager.MdiActiveDocument.Editor.Regen()
                tr.Commit()
            End Using

            Using tr3 As Transaction = db.TransactionManager.StartTransaction()
                Dim ChildBR As New Dictionary(Of BlockReference, String)
                Dim usedblocks As New HashSet(Of ObjectId)
                Dim bt As BlockTable = DirectCast(tr3.GetObject(db.BlockTableId, OpenMode.ForRead), BlockTable)
                Dim ms As BlockTableRecord = DirectCast(tr3.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)
                For Each entid As ObjectId In ms
                    Dim ent As Entity = TryCast(tr3.GetObject(entid, OpenMode.ForRead), Entity)
                    If TypeOf ent Is BlockReference Then
                        Dim blkRef As BlockReference = DirectCast(ent, BlockReference)
                        Dim btrid As ObjectId = blkRef.DynamicBlockTableRecord
                        usedblocks.Add(btrid)
                        ChildBR = Functions.GetRuntimeNestedBlocks(tr3, blkRef)
                        For Each kvp In ChildBR
                            Dim br As BlockReference = kvp.Key
                            Dim nesterbtrid As ObjectId = br.DynamicBlockTableRecord
                            usedblocks.Add(nesterbtrid)
                        Next
                    End If
                Next
                For Each objid As ObjectId In bt
                    If Not usedblocks.Contains(objid) Then
                        Dim btr As BlockTableRecord = tr3.GetObject(objid, OpenMode.ForWrite)
                        btr.Erase(True) ' True = erase dependent objects (should be none)
                    End If
                Next
            End Using

        End Using
        acadDoc.Editor.WriteMessage("Update complete." & vbLf)
    End Sub

    Private Shared Function ReplaceBlockReference(db As Database, tr As Transaction, savedState As BlockRefState) As BlockReference

        Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
        Dim btrId As ObjectId = bt(savedState.BlockName)
        Dim ms As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

        Dim newBr As New BlockReference(savedState.Position, btrId)
        newBr.Rotation = savedState.Rotation
        newBr.ScaleFactors = savedState.ScaleFactors
        ms.AppendEntity(newBr)
        tr.AddNewlyCreatedDBObject(newBr, True)

        ' Reapply dynamic props
        If newBr.IsDynamicBlock Then
            Dim propCol = newBr.DynamicBlockReferencePropertyCollection
            For Each prop In propCol
                If savedState.DynamicProps.ContainsKey(prop.PropertyName) Then
                    Try
                        prop.Value = savedState.DynamicProps(prop.PropertyName)
                    Catch
                        ' Skip invalid values
                    End Try
                End If
            Next
        End If

        ' Reapply attributes
        Dim defBtr = CType(tr.GetObject(newBr.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
        If defBtr.HasAttributeDefinitions Then
            For Each objId In defBtr
                Dim ent = TryCast(tr.GetObject(objId, OpenMode.ForRead), Entity)
                If TypeOf ent Is AttributeDefinition Then
                    Dim attDef = CType(ent, AttributeDefinition)
                    Dim attRef = New AttributeReference()
                    attRef.SetAttributeFromBlock(attDef, newBr.BlockTransform)
                    If savedState.Attributes.ContainsKey(attDef.Tag) Then
                        attRef.TextString = savedState.Attributes(attDef.Tag)
                    End If
                    newBr.AttributeCollection.AppendAttribute(attRef)
                    tr.AddNewlyCreatedDBObject(attRef, True)
                End If
            Next
        End If

        Dim btr As BlockTableRecord = CType(tr.GetObject(newBr.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
        For Each id As ObjectId In btr
            Dim ent As Entity = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)
            If TypeOf ent Is BlockReference Then
                Dim childBr As BlockReference = CType(ent, BlockReference)
                If childBr.IsDynamicBlock Then
                    Dim childBtr As BlockTableRecord = CType(tr.GetObject(childBr.DynamicBlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                    If savedState.NestedChildStates.ContainsKey(childBtr.Name) Then
                        Dim visState = savedState.NestedChildStates(childBtr.Name)
                        Dim propCol = childBr.DynamicBlockReferencePropertyCollection
                        For Each prop In propCol
                            If prop.PropertyName.ToUpper().Contains("VISIBILITY") Then
                                prop.Value = visState
                            End If
                        Next
                    End If
                End If
            End If
        Next

        Return newBr
    End Function



End Class

''' <summary>
''' Stores dynamic properties and attributes of a block reference for restoration.
''' </summary>
Public Class BlockRefState
    Public Property ReferenceId As ObjectId
    Public Property Position As Point3d
    Public Property Rotation As Double
    Public Property ScaleFactors As Scale3d
    Public Property BlockName As String
    Public Property DynamicProps As Dictionary(Of String, Object)
    Public Property Attributes As Dictionary(Of String, String)
    Public Property NestedChildStates As Dictionary(Of String, String) ' block name → vis state


    Public Sub New(br As BlockReference, tr As Transaction)
        ReferenceId = br.ObjectId
        Position = br.Position
        Rotation = br.Rotation
        ScaleFactors = br.ScaleFactors

        Dim btr As BlockTableRecord = CType(tr.GetObject(br.DynamicBlockTableRecord, OpenMode.ForRead), BlockTableRecord)
        BlockName = btr.Name

        DynamicProps = New Dictionary(Of String, Object)
        Attributes = New Dictionary(Of String, String)

        If br.IsDynamicBlock Then
            For Each prop As DynamicBlockReferenceProperty In br.DynamicBlockReferencePropertyCollection
                If prop.PropertyName.ToUpper().Contains("VISIBILITY") Then
                    DynamicProps(prop.PropertyName) = prop.Value
                End If
            Next
        End If

        For Each attId As ObjectId In br.AttributeCollection
            Dim attRef As AttributeReference = CType(tr.GetObject(attId, OpenMode.ForRead), AttributeReference)
            Attributes(attRef.Tag) = attRef.TextString
        Next

        ' Save nested child visibility states
        NestedChildStates = New Dictionary(Of String, String)

        Dim btr2 As BlockTableRecord = CType(tr.GetObject(br.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
        For Each id In btr2
            Dim ent As Entity = TryCast(tr.GetObject(id, OpenMode.ForRead), Entity)
            If TypeOf ent Is BlockReference Then
                Dim childBr As BlockReference = CType(ent, BlockReference)
                If childBr.IsDynamicBlock Then
                    Dim childBtr As BlockTableRecord = CType(tr.GetObject(childBr.DynamicBlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                    Dim visState = Functions.GetCurrentVisibilityState(childBr)
                    If Not String.IsNullOrEmpty(visState) Then
                        NestedChildStates(childBtr.Name) = visState
                    End If
                End If
            End If
        Next
    End Sub
End Class