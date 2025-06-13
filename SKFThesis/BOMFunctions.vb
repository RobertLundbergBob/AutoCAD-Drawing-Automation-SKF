Imports System.IO
Imports System.Reflection.Emit
Imports System.Windows.Forms
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.Runtime
Imports Microsoft.SqlServer.Server
Imports Microsoft.VisualBasic.Devices
Imports Autodesk.AutoCAD.ApplicationServices.Core.Application
Imports Newtonsoft.Json
Imports System.CodeDom.Compiler
Imports Microsoft.Office.Interop.Excel

Public Class BOMFunctions
    Private Shared ComponentMap As New Dictionary(Of ObjectId, String)()
    Private Shared ComponentCount As New Dictionary(Of String, Integer)()
    'Private Shared NestedChildMap As New Dictionary(Of ObjectId, Dictionary(Of ObjectId, String))()
    Private Shared outletLeaderCreated As New HashSet(Of String)()
    Private Shared ComponentTypeMap As New Dictionary(Of ObjectId, String)
    Private Shared ComponentLeaderMap As New Dictionary(Of ObjectId, ObjectId)()
    Private Shared ComponentPositions As New Dictionary(Of ObjectId, Point3d)()
    ' Store state per layout name
    Private Shared LayoutComponentMapStore As New Dictionary(Of String, Dictionary(Of ObjectId, String))()
    Private Shared LayoutComponentCountStore As New Dictionary(Of String, Dictionary(Of String, Integer))()


    <CommandMethod("BOM")>
    Public Sub Form()
        Dim form As New MainForm
        form.Show()
    End Sub

    Public Shared Sub BOMButton(layoutName As String)
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Autodesk.AutoCAD.DatabaseServices.Database = doc.Database
        Dim ed As Editor = doc.Editor

        Using docLock As DocumentLock = doc.LockDocument()
            Using tr As Transaction = db.TransactionManager.StartTransaction()
                Dim bt As BlockTable = tr.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim ms As BlockTableRecord = tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForWrite)
                Dim layoutDict As DBDictionary = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead)
                Dim layoutId As ObjectId = layoutDict.GetAt(layoutName)
                Dim layout As Layout = tr.GetObject(layoutId, OpenMode.ForRead)
                Dim ps As BlockTableRecord = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite)

                ' Find existing BOM table
                Dim bomTable As Table = FindBOMTable(ps, tr)

                ' Always get current modelspace state
                Dim currentData = GetComponentData(ms, tr)
                Dim currentComponentMap = currentData.componentMap
                Dim currentComponentCount = currentData.componentCount
                Dim currentComponentPositions = currentData.componentPositions
                Dim currentComponentTypeMap = currentData.componentTypeMap


                Dim previousComponentMap As Dictionary(Of ObjectId, String) =
                If(LayoutComponentMapStore.ContainsKey(layoutName),
                   New Dictionary(Of ObjectId, String)(LayoutComponentMapStore(layoutName)),
                   New Dictionary(Of ObjectId, String)())

                Dim previousComponentCount As Dictionary(Of String, Integer) =
                If(LayoutComponentCountStore.ContainsKey(layoutName),
                   New Dictionary(Of String, Integer)(LayoutComponentCountStore(layoutName)),
                   New Dictionary(Of String, Integer)())
                bomTable = UpdateOrCreateBOMAndLeaders(layoutName, ms, tr, db, ed, bomTable,
                               previousComponentMap, previousComponentCount,
                               currentComponentMap, currentComponentCount,
                               currentComponentPositions, currentComponentTypeMap)

                ' Save current state
                ComponentMap = currentComponentMap
                ComponentCount = currentComponentCount
                ComponentPositions = currentComponentPositions
                ComponentTypeMap = currentComponentTypeMap

                LayoutComponentMapStore(layoutName) = New Dictionary(Of ObjectId, String)(ComponentMap)
                LayoutComponentCountStore(layoutName) = New Dictionary(Of String, Integer)(ComponentCount)
                tr.Commit()
                If bomTable IsNot Nothing AndAlso MainForm.ExportToExcel Then
                    ExportTableToExcel(bomTable)
                End If
            End Using
        End Using
    End Sub

    Public Shared Function UpdateOrCreateBOMAndLeaders(
    layoutName As String,
    ms As BlockTableRecord,
    tr As Transaction,
    db As Autodesk.AutoCAD.DatabaseServices.Database,
    ed As Editor,
    bomTable As Table,
    previousComponentMap As Dictionary(Of ObjectId, String),
    previousComponentCount As Dictionary(Of String, Integer),
    currentComponentMap As Dictionary(Of ObjectId, String),
    currentComponentCount As Dictionary(Of String, Integer),
    currentComponentPositions As Dictionary(Of ObjectId, Point3d),
    currentComponentTypeMap As Dictionary(Of ObjectId, String)
) As Table
        Dim isNewTable As Boolean = (bomTable Is Nothing)
        If isNewTable Then
            ' Create new table if missing
            Dim layoutDict As DBDictionary = tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead)
            Dim layoutId As ObjectId = layoutDict.GetAt(layoutName)
            Dim layout As Layout = tr.GetObject(layoutId, OpenMode.ForRead)
            Dim ps As BlockTableRecord = tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite)
            bomTable = CreateTable(tr, db, ps, ed, currentComponentCount, layoutName)
        Else
            If Not bomTable.IsWriteEnabled Then bomTable.UpgradeOpen()
            Dim Lookup = CreateBOMLookUp(bomTable)
            For Each idCode As String In currentComponentCount.Keys
                Dim newQty = currentComponentCount(idCode)
                If Not isNewTable AndAlso previousComponentCount.ContainsKey(idCode) Then
                    Dim oldQty = previousComponentCount(idCode)
                    If newQty <> oldQty Then
                        Dim rowIndex = Integer.Parse(Lookup(idCode))
                        bomTable.Cells(rowIndex, 1).TextString = newQty.ToString()
                    End If
                ElseIf isNewTable OrElse Not Lookup.ContainsKey(idCode) Then
                    AddNewBOMRow(bomTable, idCode, currentComponentCount)
                End If
            Next
            RemoveBOMRows(bomTable, currentComponentCount, Lookup)
        End If

        ' Remove leaders for components that became invisible
        Dim becameInvisible = previousComponentMap.Keys.Except(currentComponentMap.Keys)
        For Each objId In becameInvisible
            If ComponentLeaderMap.ContainsKey(objId) Then
                Dim leaderId = ComponentLeaderMap(objId)
                Dim leaderObj = tr.GetObject(leaderId, OpenMode.ForWrite, True)
                If leaderObj IsNot Nothing Then leaderObj.Erase()
                ComponentLeaderMap.Remove(objId)
                Dim idCode = previousComponentMap(objId)
                outletLeaderCreated.Remove(idCode)
            End If
        Next

        ' Refresh leader tag info
        EraseUnusedLeaders(tr, currentComponentMap)

        ' Updated BOMLookup
        Dim bomLookup = CreateBOMLookUp(bomTable)

        For Each kvp In currentComponentMap
            Dim objId = kvp.Key
            Dim idCode = kvp.Value

            Dim newPos As Point3d = currentComponentPositions(objId)
            If ComponentPositions.ContainsKey(objId) Then
                Dim oldPos As Point3d = ComponentPositions(objId)
                If Not newPos.Equals(oldPos) Then
                    UpdateLeaderPosition(tr, objId, oldPos, newPos)
                End If
            End If
            ComponentPositions(objId) = newPos

            Dim tag = GetTagNumber(bomLookup, idCode)
            If String.IsNullOrEmpty(tag) Then
                ed.WriteMessage($"No matching TagNumber found for IdCode: {idCode}{vbLf}")
                Continue For
            End If

            If ComponentLeaderMap.ContainsKey(objId) Then
                Dim leaderId = ComponentLeaderMap(objId)
                Dim leaderObj = tr.GetObject(leaderId, OpenMode.ForWrite, True)
                If leaderObj Is Nothing OrElse leaderObj.IsErased Then
                    Dim newLeaderId = CreateNewLeader(db, tr, ms, newPos, tag)
                    ComponentLeaderMap(objId) = newLeaderId
                    ed.WriteMessage($"Leader for component {objId} was missing. A new leader has been created.")
                Else
                    UpdateLeaderTagNumber(tr, CType(leaderObj, MLeader), tag)
                End If
            Else
                If currentComponentTypeMap.ContainsKey(objId) AndAlso currentComponentTypeMap(objId) = "Nested" Then
                    If outletLeaderCreated.Contains(idCode) Then
                        If ComponentLeaderMap.ContainsKey(objId) Then
                            Dim leader As ObjectId = ComponentLeaderMap(objId)
                            Dim leaderObj = tr.GetObject(leader, OpenMode.ForWrite, True)
                            If leaderObj Is Nothing OrElse leaderObj.IsErased Then
                                Dim newleaderId As ObjectId = CreateNewLeader(db, tr, ms, newPos, tag)
                                ComponentLeaderMap(objId) = newleaderId
                            End If
                        End If
                        ed.WriteMessage($"Leader for article number {idCode} already exists.{vbLf}")
                        Continue For
                    End If

                    Dim leaderId = CreateNewLeader(db, tr, ms, newPos, tag)
                    ComponentLeaderMap(objId) = leaderId
                    outletLeaderCreated.Add(idCode)
                    ed.WriteMessage($"Created leader for article number {idCode}.{vbLf}")
                Else
                    Dim leaderId = CreateNewLeader(db, tr, ms, newPos, tag)
                    ComponentLeaderMap(objId) = leaderId
                End If
            End If
        Next
        Return bomTable
    End Function


    Private Shared Function CreateBOMLookUp(bomTable As Table) As Dictionary(Of String, String)
        Dim bomLookup As New Dictionary(Of String, String) ' (IdCode -> PosNumber)
        For i As Integer = 1 To bomTable.Rows.Count - 1
            Dim idCode As String = bomTable.Cells(i, 2).TextString
            Dim posNumber As String = bomTable.Cells(i, 0).TextString
            If Not bomLookup.ContainsKey(idCode) Then
                bomLookup.Add(idCode, posNumber) ' posNumber stored as value
            End If
        Next
        Return bomLookup
    End Function
    Private Shared Function GetTagNumber(bomLookup As Dictionary(Of String, String), IdCode As String) As String
        Dim TagNumber As String = If(bomLookup.ContainsKey(IdCode), bomLookup(IdCode), "")
        If String.IsNullOrEmpty(TagNumber) Then
            Return ""
        End If
        Return TagNumber
    End Function

    Private Shared Function EraseUnusedLeaders(tr As Transaction, components As Dictionary(Of ObjectId, String))
        Dim invalidComponents As New List(Of ObjectId)
        If ComponentLeaderMap.Count = 0 Then
            Return "No leaders to erase."
        End If

        For Each kvp In ComponentLeaderMap
            Dim objId As ObjectId = kvp.Key
            ' If the component is no longer in ComponentMap, mark for removal
            If Not components.ContainsKey(objId) Then
                invalidComponents.Add(objId)
            End If
        Next
        ' Remove invalid keys from ComponentLeaderMap
        For Each key In invalidComponents
            Dim leaderToErase As Entity = tr.GetObject(ComponentLeaderMap(key), OpenMode.ForWrite, True)
            If leaderToErase IsNot Nothing AndAlso Not leaderToErase.IsErased Then
                leaderToErase.Erase()
            End If
            ComponentLeaderMap.Remove(key)
        Next
        Return $"Unused leaders erased: {invalidComponents.Count}"
    End Function
    Private Shared Function CreateTable(
    tr As Transaction,
    db As Autodesk.AutoCAD.DatabaseServices.Database,
    ps As BlockTableRecord,
    ed As Editor,
    componentCount As Dictionary(Of String, Integer),
    layoutName As String
) As Table
        ' Create Table
        Dim tbl As New Table With {
            .TableStyle = db.Tablestyle ' Default drawing tablestyle
            }
        If layoutName = "A0" Then
            tbl.Position = New Point3d(1169 - 180, 20 + 30, 0) ' Adjust table position
        ElseIf layoutName = "A1" Then
            tbl.Position = New Point3d(821 - 180, 20 + 30, 0)
        ElseIf layoutName = "A2" Then
            tbl.Position = New Point3d(584 - 180, 10 + 30, 0)
        ElseIf layoutName = "A3" Then
            tbl.Position = New Point3d(410 - 180, 10 + 30, 0)
        ElseIf layoutName = "A3L" Then
            tbl.Position = New Point3d(831 - 180, 10 + 30, 0)
        ElseIf layoutName = "A4" Then
            tbl.Position = New Point3d(200 - 180, 10 + 30, 0)
        ElseIf layoutName = "A4SV" Then
            tbl.Position = New Point3d(200 - 180, 10 + 30, 0)
        ElseIf layoutName = "A3SV" Then
            tbl.Position = New Point3d(410 - 180, 10 + 30, 0)
        Else
            ed.WriteMessage(vbLf & "Selected layout not found in drawing")
        End If

        Dim numRows As Integer = Math.Max(componentCount.Count + 1, 2) ' Ensures at least 2 rows
        Dim numCols As Integer = 4
        tbl.SetSize(numRows, numCols)

        tbl.SetRowHeight(6) 'tbl.Rows(0).Height = 6
        tbl.Columns(0).Width = 10
        tbl.Columns(1).Width = 10
        tbl.Columns(2).Width = 38
        tbl.Columns(3).Width = 54

        tbl.Rows(0).Style = "Header"
        Dim headers() As String = {"Pos", "Qty", "Article Number", "Description"}
        For col As Integer = 0 To numCols - 1
            tbl.Cells(0, col).TextString = headers(col)
            tbl.Cells(0, col).Alignment = CellAlignment.MiddleLeft
        Next

        Dim rowIndex As Integer = 1
        Dim posNumber As Integer = 1
        If componentCount.Count = 0 Then
            For col As Integer = 0 To numCols - 1
                tbl.Cells(1, col).TextString = "-"
                tbl.Cells(1, col).Alignment = CellAlignment.MiddleCenter
            Next
            ed.WriteMessage($"No components found in Model Space")
        Else
            ' Fill Table
            For Each kvp In componentCount
                If rowIndex >= numRows Then Exit For
                Dim idCode As String = kvp.Key
                Dim quantity As Integer = kvp.Value

                tbl.Rows(rowIndex).Style = "Data"
                tbl.Cells(rowIndex, 0).TextString = posNumber.ToString()
                tbl.Cells(rowIndex, 1).TextString = quantity.ToString()
                tbl.Cells(rowIndex, 2).TextString = idCode
                tbl.Cells(rowIndex, 3).TextString = "Name"
                For col As Integer = 0 To numCols - 1
                    tbl.Cells(rowIndex, col).Alignment = CellAlignment.MiddleCenter
                Next

                rowIndex += 1
                posNumber += 1
            Next
        End If

        ' Add XData to identify the table as BOM
        Dim regAppTable As RegAppTable = tr.GetObject(db.RegAppTableId, OpenMode.ForWrite)
        If Not regAppTable.Has("BOM_TABLE") Then
            Dim regAppTableRecord As New RegAppTableRecord With {
                .Name = "BOM_TABLE"
            }
            regAppTable.UpgradeOpen()
            regAppTable.Add(regAppTableRecord)
            tr.AddNewlyCreatedDBObject(regAppTableRecord, True)
        End If

        Dim xdata As New ResultBuffer From {
            New TypedValue(DxfCode.ExtendedDataRegAppName, "BOM_TABLE"),
            New TypedValue(DxfCode.ExtendedDataAsciiString, "BillOfMaterials")
        }

        tbl.XData = xdata
        ' Add Table to Model Space
        ps.AppendEntity(tbl)
        tr.AddNewlyCreatedDBObject(tbl, True)
        Return tbl
    End Function

    Public Shared Sub ExportTableToExcel(tbl As Table)
        Dim doc As Document = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument
        Dim db As Database = doc.Database
        Dim ed As Editor = doc.Editor

        Dim excelApp As New Microsoft.Office.Interop.Excel.Application
        Dim workbook As Workbook = excelApp.Workbooks.Add()
        Dim sheet As Worksheet = workbook.Sheets(1)

        Using tr As Transaction = db.TransactionManager.StartTransaction()

            For r As Integer = 0 To tbl.Rows.Count - 1
                For c As Integer = 0 To tbl.Columns.Count - 1
                    sheet.Cells(r + 1, c + 1).Value = tbl.Cells(r, c).TextString
                Next
            Next
            tr.Commit()
        End Using

        Dim folderpath As String = Functions.GetRad1Path() & "\"
        Dim filename As String = MainForm.ExcelFileName & ".xlsx"
        Dim counter As Integer = 1
        While System.IO.File.Exists(folderpath & filename)
            filename = MainForm.ExcelFileName & "_" & counter & ".xlsx"
            counter += 1
        End While
        workbook.SaveAs(folderpath & filename)
        workbook.Close()
        excelApp.Quit()

        'Release COM objects
        System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(workbook)
        System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp)
    End Sub

    Public Shared Function GetComponentData(ms As BlockTableRecord, tr As Transaction
) As (componentCount As Dictionary(Of String, Integer),
componentMap As Dictionary(Of ObjectId, String),
componentPositions As Dictionary(Of ObjectId, Point3d),
componentTypeMap As Dictionary(Of ObjectId, String),
nestedChildMap As Dictionary(Of ObjectId, Dictionary(Of ObjectId, String)),
nestedPositionMap As Dictionary(Of ObjectId, Dictionary(Of ObjectId, Point3d)))

        Dim componentCount As New Dictionary(Of String, Integer) ' {ID Code → Quantity}
        Dim componentMap As New Dictionary(Of ObjectId, String) ' ComponentId -> IdCode
        Dim componentPositions As New Dictionary(Of ObjectId, Point3d) ' ComponentId -> position
        Dim componentTypeMap As New Dictionary(Of ObjectId, String) ' toplevel or nested
        Dim nestedChildMap As New Dictionary(Of ObjectId, Dictionary(Of ObjectId, String))
        Dim nestedPositionMap As New Dictionary(Of ObjectId, Dictionary(Of ObjectId, Point3d))

        For Each objId As ObjectId In ms
            Dim ent As Entity = TryCast(tr.GetObject(objId, OpenMode.ForRead), Entity)
            If TypeOf ent Is BlockReference Then
                Dim br As BlockReference = CType(ent, BlockReference)
                Dim idCode As String = GetIdentificationCode(br, tr)

                If Not String.IsNullOrEmpty(idCode) Then
                    If Not componentMap.ContainsKey(objId) Then
                        componentMap(objId) = idCode
                        componentPositions(objId) = br.Position
                        componentTypeMap(objId) = "Parent"
                    End If

                    If componentCount.ContainsKey(idCode) Then
                        componentCount(idCode) += 1
                    Else
                        componentCount(idCode) = 1
                    End If
                End If

                Dim nestedData = CollectNestedComponents(tr, br)
                If nestedData.NestedChildMap.ContainsKey(br.ObjectId) Then
                    nestedChildMap(br.ObjectId) = nestedData.NestedChildMap(br.ObjectId)
                    nestedPositionMap(br.ObjectId) = nestedData.NestedPositionMap(br.ObjectId)

                    For Each childId In nestedData.NestedChildMap(br.ObjectId).Keys
                        Dim articleNumber As String = nestedData.NestedChildMap(br.ObjectId)(childId)
                        Dim position As Point3d = nestedData.NestedPositionMap(br.ObjectId)(childId)

                        componentCount(articleNumber) = If(componentCount.ContainsKey(articleNumber), componentCount(articleNumber) + 1, 1)
                        componentMap(childId) = articleNumber
                        componentPositions(childId) = position
                        componentTypeMap(childId) = "Nested"
                    Next
                End If
            End If
        Next
        Return (componentCount, componentMap, componentPositions, componentTypeMap, nestedChildMap, nestedPositionMap)
    End Function
    Public Shared Function FindParameterForVisibility(componentName As String, visibilityState As String) As String
        If Not InsertComponent.visibilityMap.ContainsKey(componentName) Then Return Nothing

        For Each paramEntry In InsertComponent.visibilityMap(componentName)
            Dim parameterName = paramEntry.Key
            Dim visMap = paramEntry.Value

            If visMap.ContainsKey(visibilityState) Then
                Return parameterName
            End If
        Next

        Return Nothing
    End Function

    Public Shared Function CollectNestedComponents(
    tr As Transaction,
    parentBr As BlockReference
) As (NestedChildMap As Dictionary(Of ObjectId, Dictionary(Of ObjectId, String)), NestedPositionMap As Dictionary(Of ObjectId, Dictionary(Of ObjectId, Point3d)))

        Dim NestedChildMap As New Dictionary(Of ObjectId, Dictionary(Of ObjectId, String))
        Dim NestedPositionMap As New Dictionary(Of ObjectId, Dictionary(Of ObjectId, Point3d))
        Dim blockDef As BlockTableRecord
        If parentBr.DynamicBlockTableRecord <> ObjectId.Null Then
            blockDef = CType(tr.GetObject(parentBr.DynamicBlockTableRecord, OpenMode.ForRead), BlockTableRecord)
        Else
            blockDef = CType(tr.GetObject(parentBr.BlockTableRecord, OpenMode.ForRead), BlockTableRecord)
        End If
        Dim childblocks As Dictionary(Of BlockReference, String) = Functions.GetRuntimeNestedBlocks(tr, parentBr)
        Dim blockName As String = blockDef.Name.Split("_"c)(0)

        'For Each entId As ObjectId In blockDef
        Dim groupedNested = From kvp In childblocks
                            Group kvp By blockDefName = kvp.Value Into Group

        For Each group In groupedNested
            Dim instances = group.Group.ToList()
            For i As Integer = 1 To instances.Count
                'For Each entId As ObjectId In blockDef
                'Dim ent As Entity = TryCast(tr.GetObject(entId, OpenMode.ForRead), Entity)
                'If TypeOf ent Is BlockReference Then
                Dim nestedBr As BlockReference = instances(i - 1).Key 'CType(ent, BlockReference) 'instances(i - 1).Key
                Dim nestedblockdef = CType(tr.GetObject(nestedBr.DynamicBlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                ' Only include nested blocks visible in the drawing
                If nestedBr.Visible Then
                    ' Combined transform = parent × child
                    Dim nestedname As String = nestedblockdef.Name
                    Dim nestedTransform As Matrix3d = nestedBr.BlockTransform.PreMultiplyBy(parentBr.BlockTransform)
                    Dim worldPosition As Point3d = Point3d.Origin.TransformBy(nestedTransform)

                    Dim dynProps As DynamicBlockReferencePropertyCollection = nestedBr.DynamicBlockReferencePropertyCollection
                    For Each prop As DynamicBlockReferenceProperty In dynProps
                        If prop.PropertyName = "Visibility" Then
                            Dim visState As String = prop.Value.ToString()
                            Dim paramName As String = FindParameterForVisibility(blockName, visState)
                            Dim article As String = Functions.VisibilityToArticleNumber(blockName, paramName, visState)

                            ' Leader info
                            If nestedBr.AttributeCollection.Count <> 0 Then
                                For Each attId As ObjectId In nestedBr.AttributeCollection
                                    Dim attRef As AttributeReference = TryCast(tr.GetObject(attId, OpenMode.ForRead), AttributeReference)
                                    If attRef IsNot Nothing AndAlso attRef.Tag.ToUpper() = "NEEDLEADER" Then
                                        If Not String.IsNullOrEmpty(article) Then
                                            If Not NestedChildMap.ContainsKey(parentBr.ObjectId) Then
                                                NestedChildMap(parentBr.ObjectId) = New Dictionary(Of ObjectId, String)
                                                NestedPositionMap(parentBr.ObjectId) = New Dictionary(Of ObjectId, Point3d)
                                            End If
                                            ' Combined transform = parent × child
                                            NestedChildMap(parentBr.ObjectId)(nestedBr.ObjectId) = article
                                            NestedPositionMap(parentBr.ObjectId)(nestedBr.ObjectId) = worldPosition
                                        End If
                                    End If
                                Next
                            End If

                        End If
                    Next
                End If
                'End If
            Next
        Next

        Return (NestedChildMap, NestedPositionMap)
    End Function

    Private Shared Function GetIdentificationCode(br As BlockReference, tr As Transaction) As String
        If br.AttributeCollection.Count = 0 Then Return ""

        For Each attId As ObjectId In br.AttributeCollection
            Dim attRef As AttributeReference = TryCast(tr.GetObject(attId, OpenMode.ForRead), AttributeReference)
            If attRef IsNot Nothing AndAlso attRef.Tag.ToUpper() = "IDENTIFICATION_CODE" Then
                Return attRef.TextString
            End If
        Next
        Return ""
    End Function
    Public Shared Function FindBOMTable(ps As BlockTableRecord, tr As Transaction) As Table
        For Each objId As ObjectId In ps
            Dim ent As Entity = TryCast(tr.GetObject(objId, OpenMode.ForRead), Entity)
            If TypeOf ent Is Table Then
                Dim tbl As Table = CType(ent, Table)
                If tbl.XData IsNot Nothing Then
                    Dim xdata As TypedValue() = tbl.XData.AsArray()
                    If xdata.Length > 1 AndAlso xdata(1).Value.ToString() = "BillOfMaterials" Then
                        Return tbl
                    End If
                End If
            End If
        Next
        Return Nothing
    End Function
    Private Shared Function CreateNewLeader(db As Autodesk.AutoCAD.DatabaseServices.Database, tr As Transaction, ms As BlockTableRecord, position As Point3d, idCode As String) As ObjectId
        Dim mleaderStyleName As String = "SKF_Bubbla"
        Dim mlStyleId As ObjectId = ObjectId.Null
        Dim mld As DBDictionary = tr.GetObject(db.MLeaderStyleDictionaryId, OpenMode.ForRead)
        If mld.Contains(mleaderStyleName) Then
            mlStyleId = mld(mleaderStyleName)
        Else
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "MultiLeaderStyle not found!")

        End If

        Dim mleader As New MLeader()
        mleader.SetDatabaseDefaults()
        mleader.MLeaderStyle = mlStyleId ' Apply MultiLeaderStyle
        mleader.ContentType = ContentType.BlockContent
        mleader.BlockRotation = 0
        mleader.EnableAnnotationScale = True

        Dim leaderIndex As Integer = mleader.AddLeader()
        Dim lineIndex As Integer = mleader.AddLeaderLine(leaderIndex)
        Dim x_offset As Integer = 60
        Dim y_offset As Integer = 30
        Dim leaderStart As New Point3d(position.X, position.Y, 0)
        Dim landingPoint As New Point3d(position.X + x_offset, position.Y + y_offset, 0)
        mleader.AddFirstVertex(lineIndex, leaderStart) ' Start leader at component
        mleader.AddLastVertex(lineIndex, landingPoint) ' Ensure the leader starts at the correct point

        Dim mleaderStyle As MLeaderStyle = tr.GetObject(mlStyleId, OpenMode.ForWrite)
        If Not mleaderStyle.Annotative Then
            mleader.BlockContentId = mleaderStyle.BlockId ' Ensures the correct predefined block is used
        End If
        Dim blockId As ObjectId = mleaderStyle.BlockId

        mleader.TextAlignmentType = TextAlignmentType.LeftAlignment
        mleader.EnableLanding = True

        ' Add XData to identify MultiLeader as BOM-related
        Dim regAppTable As RegAppTable = tr.GetObject(db.RegAppTableId, OpenMode.ForWrite)
        If Not regAppTable.Has("BOM_LEADER") Then
            Dim regAppTableRecord As New RegAppTableRecord With {
                                .Name = "BOM_LEADER"
                            }
            regAppTable.UpgradeOpen()
            regAppTable.Add(regAppTableRecord)
            tr.AddNewlyCreatedDBObject(regAppTableRecord, True)
        End If

        Dim xdata As New ResultBuffer From {
            New TypedValue(DxfCode.ExtendedDataRegAppName, "BOM_LEADER"),
            New TypedValue(DxfCode.ExtendedDataAsciiString, idCode) ' Stores the component IdCode as xdata to connect the leader with the component
            }
        mleader.XData = xdata
        UpdateLeaderTagNumber(tr, mleader, idCode)

        ' Add leader to Model Space
        ms.AppendEntity(mleader)
        tr.AddNewlyCreatedDBObject(mleader, True)

        Return mleader.ObjectId
    End Function
    Private Shared Sub UpdateLeaderTagNumber(tr As Transaction, mleader As MLeader, TagNumber As String)
        Dim blockId As ObjectId = mleader.BlockContentId
        Dim blockdef As BlockTableRecord = TryCast(tr.GetObject(blockId, OpenMode.ForWrite), BlockTableRecord)
        If blockdef Is Nothing Then
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Block definition not found!")
            Return
        End If
        ' Find the attribute definition in the block
        Dim attDef As AttributeDefinition = Nothing
        For Each entId As ObjectId In blockdef
            Dim attEntity As Entity = TryCast(tr.GetObject(entId, OpenMode.ForWrite), Entity)
            If TypeOf attEntity Is AttributeDefinition Then
                Dim attrDefTemp As AttributeDefinition = DirectCast(attEntity, AttributeDefinition)
                If attrDefTemp.Tag.ToUpper() = "TAGNUMBER" Then
                    attDef = attrDefTemp
                    Exit For
                End If
            End If
        Next

        If attDef Is Nothing Then
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "TAGNUMBER attribute not found in block definition.")
            Return
        End If
        ' Apply the attribute change to the Multileader
        Try
            Dim attRef As AttributeReference = mleader.GetBlockAttribute(attDef.ObjectId)
            If attRef IsNot Nothing Then
                attRef.TextString = TagNumber
                mleader.SetBlockAttribute(attDef.ObjectId, attRef) ' Apply update
            End If
        Catch ex As Autodesk.AutoCAD.Runtime.Exception
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Error updating attribute: " & ex.Message)
        End Try
    End Sub

    Private Shared Sub UpdateLeaderPosition(tr As Transaction, componentId As ObjectId, oldPosition As Point3d, newPosition As Point3d)
        If Not ComponentLeaderMap.ContainsKey(componentId) Then Exit Sub

        If Not ComponentLeaderMap.ContainsKey(componentId) Then Return
        Dim leaderId As ObjectId = ComponentLeaderMap(componentId)
        If leaderId.IsErased OrElse leaderId.IsNull Then Return
        Dim leader As MLeader = TryCast(tr.GetObject(leaderId, OpenMode.ForWrite, True), MLeader)
        If leader Is Nothing OrElse leader.IsErased Then Return

        ' Calculate the offset
        Dim offset As Vector3d = oldPosition.GetVectorTo(newPosition)

        ' Apply the offset to the leader
        Dim trans As Matrix3d = Matrix3d.Displacement(offset)
        leader.TransformBy(trans)
    End Sub
    Private Shared Sub AddNewBOMRow(
    bomTable As Table,
    idCode As String,
    componentCount As Dictionary(Of String, Integer)
)
        ' Ensure BOM Table is write-enabled
        If Not bomTable.IsWriteEnabled Then bomTable.UpgradeOpen()
        ' Insert the new row into the BOM table
        bomTable.InsertRows(bomTable.Rows.Count, bomTable.Rows(0).Height, 1)
        Dim rowIndex As Integer = bomTable.Rows.Count - 1
        bomTable.Cells(rowIndex, 0).TextString = rowIndex.ToString() ' Position number
        bomTable.Cells(rowIndex, 1).TextString = componentCount(idCode).ToString() ' Quantity
        bomTable.Cells(rowIndex, 2).TextString = idCode ' ID Code
        bomTable.Cells(rowIndex, 3).TextString = "Name" ' Placeholder for description

        For col As Integer = 0 To bomTable.Columns.Count - 1
            bomTable.Cells(rowIndex, col).Alignment = CellAlignment.MiddleCenter
        Next
    End Sub
    Private Shared Sub RemoveBOMRows(
    bomTable As Table,
    componentCount As Dictionary(Of String, Integer),
    bomLookup As Dictionary(Of String, String)
)
        Dim rowsToRemove As New List(Of Integer)
        For Each idCode In bomLookup.Keys
            If Not componentCount.ContainsKey(idCode) Then
                Dim rowIndex As Integer
                If Integer.TryParse(bomLookup(idCode), rowIndex) Then
                    rowsToRemove.Add(rowIndex)
                Else ' when BomTable is empty: "-"
                    rowIndex = 1
                    rowsToRemove.Add(rowIndex)
                End If
            End If
        Next

        If bomTable.IsWriteEnabled = False Then
            bomTable.UpgradeOpen()
        End If
        ' Delete rows from the BOM
        rowsToRemove.Sort()
        For i = rowsToRemove.Count - 1 To 0 Step -1
            bomTable.DeleteRows(rowsToRemove(i), 1)
        Next

        ' Reassign PosNumbers for all remaining rows (in reverse order to prevent index shifting)
        For i As Integer = 1 To bomTable.Rows.Count - 1
            bomTable.Cells(i, 0).TextString = i.ToString()
        Next
    End Sub

End Class