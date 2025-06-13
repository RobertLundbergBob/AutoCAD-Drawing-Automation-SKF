Imports System.Data.SQLite
Imports System.Windows.Forms
Imports Autodesk.AutoCAD.DatabaseServices
Imports System.Configuration
Imports System.IO
Imports System.Security.Cryptography
Imports System.Text.RegularExpressions
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry
Imports Autodesk.AutoCAD.GraphicsInterface
Imports Autodesk.AutoCAD.PlottingServices
Imports Autodesk.AutoCAD.Runtime
Imports Microsoft.SqlServer.Server
Imports Microsoft.VisualBasic.Devices
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application
Public Class UploadToDatabase

    Public Shared baseUserPath As String = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile)
    Public Shared folderPath As String = Functions.GetRad1Path & "\"
    Public Shared connectionstring As String = "Data Source=" & folderPath & "YourDatabaseFilePath.sqlite"

    Public Shared Sub UploadBlockLibraryData()
        'Browse File Explorer for the DWG File of the Parent Block
        Dim filepath As String
        Dim DrawingName As String
        Dim openFile As New OpenFileDialog()
        openFile.Title = "Select an AutoCAD Drawing"
        ' Filter so that only DWG (or any relevant format) files are displayed.
        openFile.Filter = "AutoCAD Files (*.dwg)|*.dwg|All Files (*.*)|*.*"

        If openFile.ShowDialog() = DialogResult.OK Then
            ' Get the selected file path.
            filepath = openFile.FileName
            ' Use Path.GetFileNameWithoutExtension to extract the drawing name from the file name.
            DrawingName = Path.GetFileNameWithoutExtension(filepath)
        End If

        'Dim filepath As String = "C:\Users\XK8940\OneDrive - SKF\Shared Documents - O365-AE Team Sweden\Examensarbete\RAD1\Pumps\P203\P203.dwg"
        Dim parentBlockName As String = DrawingName & "_Block"

        Dim parameternames As New List(Of String)
        UploadParentBlockDefinition(parentBlockName)

        Using db As New Autodesk.AutoCAD.DatabaseServices.Database(False, True)
            db.ReadDwgFile(filepath, FileShare.Read, True, Nothing) 'Get DWG File
            db.CloseInput(True)
            ' Start a transaction
            Using trans As Transaction = db.TransactionManager.StartTransaction()
                Dim bt As BlockTable = trans.GetObject(db.BlockTableId, OpenMode.ForRead)
                Dim ms As BlockTableRecord = DirectCast(trans.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)

                Dim parentblkref As BlockReference = Functions.GetBlockReference(trans, bt, parentBlockName)
                Dim ParentParameterNames As List(Of String) = Functions.GetParameterNames(parentblkref)
                'Get ParentBlockID
                Dim ParentBlockID As Integer? = GetParentBlockID(parentBlockName)
                'Insert into parameterdefinition the ParentBlockID, Parameternames
                For Each paramName As String In ParentParameterNames
                    DeleteParameterDefinition(ParentBlockID, Nothing, paramName)
                    InsertIntoParameterDefinition(ParentBlockID, Nothing, Nothing, paramName)
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
        vbLf & "Parameter '" & paramName & "' inserted (replacing existing if any) for ParentBlockID: " & ParentBlockID.Value.ToString())
                Next

                'Add Parent Parameter Values to the ParameterValueLibrary
                Dim visibilitystates As List(Of String) = Functions.GetVisibilityStates(trans, bt, parentblkref)
                Dim ParameterID As Integer? = GetParameterID(Nothing, ParentBlockID, Nothing, "Visibility")
                If ParameterID.HasValue Then
                    ' Delete all previous values for this ParameterID
                    DeleteParameterValues(ParameterID)
                    For Each value In visibilitystates
                        ' Always insert new visibility state
                        InsertIntoParameterValueLibrary(ParameterID, value)
                        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "Value '" & value & "' inserted for ParameterID: " & ParameterID.ToString())
                    Next
                End If

                'ADD PARENT TO ATTRIBUTEDEFINITION
                If parentblkref.AttributeCollection IsNot Nothing Then
                    For Each attRefId As ObjectId In parentblkref.AttributeCollection
                        Dim attRef As AttributeReference = TryCast(trans.GetObject(attRefId, OpenMode.ForRead), AttributeReference)
                        If attRef IsNot Nothing Then
                            Dim attributeName As String = attRef.Tag ' Tag = Attribute name
                            Dim attributeValue As String = attRef.TextString ' Actual Value

                            ' Only insert if it doesn't already exist
                            If Not AttributeDefinitionExists(ParentBlockID, Nothing, attributeName) Then
                                InsertIntoAttributeDefinition(ParentBlockID, Nothing, attributeName)
                                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                                        vbLf & "Inserted attribute '" & attributeName & "' for ParentBlockID: " & ParentBlockID.ToString())
                            Else
                                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                                        vbLf & "Attribute '" & attributeName & "' already exists for ParentBlockID: " & ParentBlockID.ToString())
                            End If

                            ' Get AttributeID for this (freshly inserted or existing)
                            Dim attributeID As Integer? = GetAttributeID(Nothing, ParentBlockID, attributeName)

                            If attributeID.HasValue Then
                                ' Insert AttributeValue
                                InsertIntoAttributeValueLibrary(attributeID, attributeValue)
                                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                                        vbLf & "Inserted attribute value '" & attributeValue & "' for AttributeID: " & attributeID.Value.ToString())
                            End If
                        End If
                    Next
                End If

                Dim ChildBlockID As Integer?
                Dim oldChildBlockIDs As List(Of Integer) = GetChildBlockIDsByParent(ParentBlockID)
                DeleteChildBlockDefinitions(ParentBlockID)
                DeleteDoubleChildBlockDefinitions(oldChildBlockIDs)
                DeleteOrphanedParameterDefinitions()
                DeleteOrphanedParameterValues()
                DeleteOrphanedAttributes()
                DeleteOrphanedAttributeValues()
                DeleteOrphanedFlipVisibilityStates()
                'Get Nested Block References
                Dim childblocks As Dictionary(Of BlockReference, String) = Functions.GetRuntimeNestedBlocks(trans, parentblkref)
                'Add ChildBlockDefinitions and their parameters with values
                For Each kvp In childblocks
                    Dim childBr As BlockReference = kvp.Key
                    Dim childblockname As String = kvp.Value
                    ' Only upload if this childblockname isn't already inserted
                    If Not ChildBlockDefinitionExists(ParentBlockID, childblockname) Then
                        UploadChildBlockDefinition(ParentBlockID, childblockname)
                    Else
                        Continue For
                        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "ChildBlockName '" & childblockname & "' already exists for ParentBlockID: " & ParentBlockID.ToString())
                    End If
                    ' Retrieve the parameter names from the child block.
                    Dim childParamNames As List(Of String) = Functions.GetParameterNames(childBr)
                    'Get ChildBlockID'
                    ChildBlockID = GetChildBlockID(childblockname)
                    'Insert into ParameterDefinition the ChildBlockID and ParameterNames
                    For Each childParam As String In childParamNames
                        If Not ParameterDefinitionExists(Nothing, ChildBlockID, Nothing, childParam) Then
                            InsertIntoParameterDefinition(Nothing, ChildBlockID, Nothing, childParam)
                        Else
                            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "ChildBlockName '" & childblockname & "' already exists for ParentBlockID: " & ParentBlockID.ToString())
                        End If
                    Next

                    'Add Child Parameter Values to the ParameterValueLibrary
                    Dim ChildVisibilityStates As List(Of String) = Functions.GetChildVisibilityStates(trans, childBr)
                    'Get ParameterID
                    Dim ParameterID2 As Integer? = GetParameterID(ChildBlockID, Nothing, Nothing, "Visibility")
                    For Each state In ChildVisibilityStates
                        InsertIntoParameterValueLibrary(ParameterID2, state)
                        Dim dynProps As DynamicBlockReferencePropertyCollection = childBr.DynamicBlockReferencePropertyCollection
                        Dim visProp = dynProps.Cast(Of DynamicBlockReferenceProperty)().
              FirstOrDefault(Function(p) p.PropertyName = "Visibility")
                        ' Set the visibility state temporarily
                        visProp.Value = state
                        ' Regenerate to reflect the new state
                        Application.DocumentManager.MdiActiveDocument.Editor.Regen()
                        ' Now check for Flip parameter
                        Dim flipVisible As Boolean = IsFlipParameterVisible(childBr)
                        If flipVisible = True Then
                            Dim ParameterValueID As Integer? = GetParameterValueID(ParameterID2, state)
                            InsertIntoFlipVisibilityStatesLibrary(ParameterValueID)
                        End If
                    Next

                    '----------ATTRIBUTES----------
                    If childBr.AttributeCollection IsNot Nothing Then
                        For Each attRefId As ObjectId In childBr.AttributeCollection
                            Dim attRef As AttributeReference = TryCast(trans.GetObject(attRefId, OpenMode.ForRead), AttributeReference)
                            If attRef IsNot Nothing Then
                                Dim attributeName As String = attRef.Tag ' Tag = Attribute name
                                Dim attributeValue As String = attRef.TextString ' Actual Value

                                ' Only insert if it doesn't already exist
                                If Not AttributeDefinitionExists(Nothing, ChildBlockID, attributeName) Then
                                    InsertIntoAttributeDefinition(Nothing, ChildBlockID, attributeName)
                                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                                        vbLf & "Inserted attribute '" & attributeName & "' for ChildBlockID: " & ChildBlockID.ToString())
                                Else
                                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                                        vbLf & "Attribute '" & attributeName & "' already exists for ChildBlockID: " & ChildBlockID.ToString())
                                End If

                                ' Get AttributeID for this (freshly inserted or existing)
                                Dim attributeID As Integer? = GetAttributeID(ChildBlockID, Nothing, attributeName)

                                If attributeID.HasValue Then
                                    ' Insert AttributeValue
                                    InsertIntoAttributeValueLibrary(attributeID, attributeValue)
                                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                                        vbLf & "Inserted attribute value '" & attributeValue & "' for AttributeID: " & attributeID.Value.ToString())
                                End If
                            End If
                        Next
                    End If
                Next

                Dim DoubleNestedAttr As String = Functions.GetAttribute(trans, parentblkref, "DOUBLENESTED")
                ' Check if we need to retrieve double nested visibility states.
                If DoubleNestedAttr = "True" Then
                    ' Iterate over a copy of the current keys so that we can safely add new ones.
                    For Each kvp In childblocks.ToList()
                        ' Get the nested visibility states for this inner block reference.
                        Dim innerblocks As Dictionary(Of BlockReference, String) = Functions.GetRuntimeNestedBlocks(trans, kvp.Key)
                        For Each kvps In innerblocks
                            Dim childname As String = kvp.Value
                            ChildBlockID = GetChildBlockID(childname)
                            Dim doublechild As BlockReference = kvps.Key
                            Dim doublechildname As String = kvps.Value
                            If Not DoubleChildBlockDefinitionExists(ChildBlockID, doublechildname) Then
                                UploadDoubleChildBlockDefinition(ChildBlockID, doublechildname)
                            Else
                                Continue For
                                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "ChildBlockName '" & doublechildname & "' already exists for ParentBlockID: " & ParentBlockID.ToString())
                            End If
                            ' Retrieve the parameter names from the child block. Insert Into 
                            Dim childParamNames As List(Of String) = Functions.GetParameterNames(doublechild)
                            Dim DoubleChildBlockID As Integer? = GetDoubleChildBlockID(doublechildname)
                            For Each childParam As String In childParamNames
                                If Not ParameterDefinitionExists(Nothing, Nothing, DoubleChildBlockID, childParam) Then
                                    InsertIntoParameterDefinition(Nothing, Nothing, DoubleChildBlockID, childParam)
                                Else
                                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(vbLf & "ChildBlockName '" & doublechildname & "' already exists for ParentBlockID: " & ParentBlockID.ToString())
                                End If
                            Next
                            'Add Child Parameter Values to the ParameterValueLibrary
                            Dim ChildVisibilityStates As List(Of String) = Functions.GetChildVisibilityStates(trans, doublechild)
                            'Get ParameterID
                            Dim ParameterID2 As Integer? = GetParameterID(Nothing, Nothing, DoubleChildBlockID, "Visibility")
                            For Each state In ChildVisibilityStates
                                InsertIntoParameterValueLibrary(ParameterID2, state)
                            Next
                        Next
                    Next
                End If
                trans.Commit()


            End Using

        End Using
    End Sub

    'Find potential flip parameter
    Private Shared Function IsFlipParameterVisible(blockRef As BlockReference) As Boolean
        ' The Flip grip is actually a BlockReference’s DynamicBlockReferenceProperty with a type Flip
        Dim props = blockRef.DynamicBlockReferencePropertyCollection
        For Each prop As DynamicBlockReferenceProperty In props
            If prop.PropertyName.ToLower().Contains("flip") Then
                Try
                    If prop.VisibleInCurrentVisibilityState = True Then
                        Return True
                    Else
                        ' If it is read-only, it's not active in this visibility state
                        Return False
                    End If
                Catch ex As Exception
                    ' If it's inaccessible in this state, assume it's not visible
                    Return False
                End Try
            End If
        Next
        Return False
    End Function

    'Upload to ParentBlockDefinition 
    Public Shared Function UploadParentBlockDefinition(parentBlockName As String) As Boolean
        Using conn As New SQLiteConnection(UploadToDatabase.connectionstring)
            conn.Open()

            ' Check if it already exists
            Dim checkSql As String = "SELECT 1 FROM ParentBlockDefinition WHERE ParentBlockName = @ParentBlockName LIMIT 1"
            Using checkCmd As New SQLiteCommand(checkSql, conn)
                checkCmd.Parameters.AddWithValue("@ParentBlockName", parentBlockName)
                Dim exists = checkCmd.ExecuteScalar()
                If exists IsNot Nothing Then
                    MessageBox.Show($"Parent block '{parentBlockName}' is already uploaded.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information)
                    Return False
                End If
            End Using

            ' If not exists, insert
            Dim insertSql As String = "INSERT INTO ParentBlockDefinition (ParentBlockName) VALUES (@ParentBlockName)"
            Using insertCmd As New SQLiteCommand(insertSql, conn)
                insertCmd.Parameters.AddWithValue("@ParentBlockName", parentBlockName)
                insertCmd.ExecuteNonQuery()
            End Using
        End Using

        Return True ' Successfully inserted
    End Function

    'Delete ChildBlockDefinitions
    Public Shared Sub DeleteChildBlockDefinitions(parentBlockID As Integer?)
        If Not parentBlockID.HasValue Then Exit Sub

        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()
            Dim sql As String = "DELETE FROM ChildBlockDefinition WHERE ParentBlockID = @ParentBlockID"
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@ParentBlockID", parentBlockID)
                cmd.ExecuteNonQuery()
            End Using
            conn.Close()
        End Using
    End Sub

    'Check childblock
    Public Shared Function ChildBlockDefinitionExists(parentBlockID As Integer, childBlockName As String) As Boolean
        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()
            Dim sql As String = "
            SELECT COUNT(*) 
            FROM ChildBlockDefinition 
            WHERE ParentBlockID = @ParentBlockID AND ChildBlockName = @ChildBlockName
        "
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@ParentBlockID", parentBlockID)
                cmd.Parameters.AddWithValue("@ChildBlockName", childBlockName)
                Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                Return (count > 0)
            End Using
        End Using
    End Function

    'Upload ChildBlockDefinition
    Public Shared Function UploadChildBlockDefinition(parentBlockID As Integer, childBlockName As String) As Integer
        Using conn As New SQLiteConnection(UploadToDatabase.connectionstring)
            conn.Open()

            Dim insertSql As String = "
            INSERT INTO ChildBlockDefinition (ParentBlockID, ChildBlockName)
            VALUES (@ParentBlockID, @ChildBlockName);
            SELECT last_insert_rowid();
        "
            Using insertCmd As New SQLiteCommand(insertSql, conn)
                insertCmd.Parameters.AddWithValue("@ParentBlockID", parentBlockID)
                insertCmd.Parameters.AddWithValue("@ChildBlockName", childBlockName)
                Dim newId = insertCmd.ExecuteScalar()
                Return Convert.ToInt32(newId)
            End Using
        End Using
    End Function

    Public Shared Function DoubleChildBlockDefinitionExists(ChildBlockID As Integer, doublechildBlockName As String) As Boolean
        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()
            Dim sql As String = "
            SELECT COUNT(*) 
            FROM DoubleChildBlockDefinition 
            WHERE ChildBlockID = @ChildBlockID AND DoubleChildBlockName = @DoubleChildBlockName
        "
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@ChildBlockID", ChildBlockID)
                cmd.Parameters.AddWithValue("@DoubleChildBlockName", doublechildBlockName)
                Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                Return (count > 0)
            End Using
        End Using
    End Function

    'UploadDoubleChildBlock
    Public Shared Function UploadDoubleChildBlockDefinition(ChildBlockID As Integer, doublechildBlockName As String) As Integer
        Using conn As New SQLiteConnection(UploadToDatabase.connectionstring)
            conn.Open()

            Dim insertSql As String = "
            INSERT INTO DoubleChildBlockDefinition (ChildBlockID, DoubleChildBlockName)
            VALUES (@ChildBlockID, @DoubleChildBlockName);
            SELECT last_insert_rowid();
        "
            Using insertCmd As New SQLiteCommand(insertSql, conn)
                insertCmd.Parameters.AddWithValue("@ChildBlockID", ChildBlockID)
                insertCmd.Parameters.AddWithValue("@DoubleChildBlockName", doublechildBlockName)
                Dim newId = insertCmd.ExecuteScalar()
                Return Convert.ToInt32(newId)
            End Using
        End Using
    End Function
    'Delete DoubleChildBlockDefinitions
    Public Shared Sub DeleteDoubleChildBlockDefinitions(ChildBlockIDs As List(Of Integer))
        If ChildBlockIDs Is Nothing OrElse ChildBlockIDs.Count = 0 Then Exit Sub

        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()
            For Each childBlockID In ChildBlockIDs
                Dim sql As String = "DELETE FROM DoubleChildBlockDefinition WHERE ChildBlockID = @ChildBlockID"
                Using cmd As New SQLiteCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@ChildBlockID", childBlockID)
                    cmd.ExecuteNonQuery()
                End Using
            Next
            conn.Close()
        End Using
    End Sub

    Public Shared Function GetDoubleChildBlockID(doublechildblockname As String) As Integer?
        Dim ChildBlockID As Integer? = Nothing

        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()
            ' The SQL assumes your table has columns "ParentBlockID" and "ParentBlockName"
            Dim sql As String = "
            SELECT DoubleChildBlockID
            FROM DoubleChildBlockDefinition 
            WHERE DoubleChildBlockName = @DoubleChildBlockName"
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@DoubleChildBlockName", doublechildblockname)
                Dim result As Object = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    ChildBlockID = Convert.ToInt32(result)
                End If
            End Using
            conn.Close()
        End Using
        Return ChildBlockID
    End Function

    Public Shared Function GetChildBlockIDsByParent(parentBlockID As Integer) As List(Of Integer)
        Dim childBlockIDs As New List(Of Integer)

        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()
            Dim sql As String = "SELECT ChildBlockID FROM ChildBlockDefinition WHERE ParentBlockID = @ParentBlockID"
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@ParentBlockID", parentBlockID)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        childBlockIDs.Add(Convert.ToInt32(reader("ChildBlockID")))
                    End While
                End Using
            End Using
            conn.Close()
        End Using

        Return childBlockIDs
    End Function

    Public Shared Function GetParentBlockID(parentblockname As String) As Integer?
        Dim ParentBlockID As Integer? = Nothing

        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()
            ' The SQL assumes your table has columns "ParentBlockID" and "ParentBlockName"
            Dim sql As String = "
            SELECT ParentBlockID
            FROM ParentBlockDefinition 
            WHERE ParentBlockName = @ParentBlockName"
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@ParentBlockName", parentblockname)
                Dim result As Object = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    ParentBlockID = Convert.ToInt32(result)
                End If
            End Using
            conn.Close()
        End Using
        Return ParentBlockID
    End Function

    Public Shared Function GetChildBlockID(childblockname As String) As Integer?
        Dim ChildBlockID As Integer? = Nothing

        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()
            ' The SQL assumes your table has columns "ParentBlockID" and "ParentBlockName"
            Dim sql As String = "
            SELECT ChildBlockID
            FROM ChildBlockDefinition 
            WHERE ChildBlockName = @ChildBlockName"
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@ChildBlockName", childblockname)
                Dim result As Object = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    ChildBlockID = Convert.ToInt32(result)
                End If
            End Using
            conn.Close()
        End Using
        Return ChildBlockID
    End Function

    Public Shared Function GetAttributeID(childBlockID As Integer?, parentBlockID As Integer?, attributeName As String) As Integer?
        Using conn As New SQLiteConnection(connectionString)
            conn.Open()

            Dim sql As String = "
            SELECT AttributeID 
            FROM AttributeDefinition
            WHERE (ParentBlockID = @ParentBlockID AND AttributeName = @AttributeName)
           OR (ChildBlockID = @ChildBlockID AND AttributeName = @AttributeName)
        "

            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@AttributeName", attributeName)
                If parentBlockID.HasValue Then
                    cmd.Parameters.AddWithValue("@ParentBlockID", parentBlockID.Value)
                Else
                    cmd.Parameters.AddWithValue("@ParentBlockID", DBNull.Value)
                End If
                If childBlockID.HasValue Then
                    cmd.Parameters.AddWithValue("@ChildBlockID", childBlockID.Value)
                Else
                    cmd.Parameters.AddWithValue("@ChildBlockID", DBNull.Value)
                End If
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing Then
                    Return Convert.ToInt32(result)
                Else
                    Return Nothing
                End If
            End Using

            conn.Close()
        End Using
    End Function

    Public Shared Sub DeleteParameterDefinitionsByChildBlockIDs(childBlockIDs As List(Of Integer))
        If childBlockIDs Is Nothing OrElse childBlockIDs.Count = 0 Then Exit Sub

        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()
            For Each childBlockID In childBlockIDs
                Dim sql As String = "DELETE FROM ParameterDefinition WHERE ChildBlockID = @ChildBlockID"
                Using cmd As New SQLiteCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@ChildBlockID", childBlockID)
                    cmd.ExecuteNonQuery()
                End Using
            Next
            conn.Close()
        End Using
    End Sub

    Public Shared Sub DeleteOrphanedParameterDefinitions()
        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()

            Dim sql As String = "
            DELETE FROM ParameterDefinition
            WHERE 
                (ParentBlockID IS NOT NULL AND ParentBlockID NOT IN (SELECT ParentBlockID FROM ParentBlockDefinition))
                OR
                (ChildBlockID IS NOT NULL AND ChildBlockID NOT IN (SELECT ChildBlockID FROM ChildBlockDefinition))
                OR
                (DoubleChildBlockID IS NOT NULL AND DoubleChildBlockID NOT IN (SELECT DoubleChildBlockID FROM DoubleChildBlockDefinition))
        "

            Using cmd As New SQLiteCommand(sql, conn)
                Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                vbLf & "Deleted " & rowsAffected.ToString() & " orphaned ParameterDefinition records.")
            End Using

            conn.Close()
        End Using
    End Sub


    Public Shared Sub DeleteParameterDefinition(parentBlockID As Integer?, childBlockID As Integer?, parameterName As String)
        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()
            Dim sql As String = "DELETE FROM ParameterDefinition WHERE ParameterName = @ParameterName"

            If parentBlockID.HasValue Then
                sql &= " AND ParentBlockID = @ParentBlockID"
            Else
                sql &= " AND ParentBlockID IS NULL"
            End If

            If childBlockID.HasValue Then
                sql &= " AND ChildBlockID = @ChildBlockID"
            Else
                sql &= " AND ChildBlockID IS NULL"
            End If

            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@ParameterName", parameterName)

                If parentBlockID.HasValue Then
                    cmd.Parameters.AddWithValue("@ParentBlockID", parentBlockID.Value)
                End If
                If childBlockID.HasValue Then
                    cmd.Parameters.AddWithValue("@ChildBlockID", childBlockID.Value)
                End If

                cmd.ExecuteNonQuery()
            End Using

            conn.Close()
        End Using
    End Sub

    Public Shared Function ParameterDefinitionExists(parentBlockID As Integer?, childBlockID As Integer?, doublechildblockid As Integer?, parameterName As String) As Boolean
        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()
            Dim sql As String = "
            SELECT COUNT(*) FROM ParameterDefinition 
            WHERE ParameterName = @ParameterName"

            If parentBlockID.HasValue Then
                sql &= " AND ParentBlockID = @ParentBlockID"
            Else
                sql &= " AND ParentBlockID IS NULL"
            End If

            If childBlockID.HasValue Then
                sql &= " AND ChildBlockID = @ChildBlockID"
            Else
                sql &= " AND ChildBlockID IS NULL"
            End If

            If doublechildblockid.HasValue Then
                sql &= " AND DoubleChildBlockID = @DoubleChildBlockID"
            Else
                sql &= " AND DoubleChildBlockID IS NULL"
            End If

            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@ParameterName", parameterName)
                If parentBlockID.HasValue Then cmd.Parameters.AddWithValue("@ParentBlockID", parentBlockID.Value)
                If childBlockID.HasValue Then cmd.Parameters.AddWithValue("@ChildBlockID", childBlockID.Value)
                If doublechildblockid.HasValue Then cmd.Parameters.AddWithValue("@DoubleChildBlockID", doublechildblockid.Value)
                Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                Return (count > 0)
            End Using
        End Using
    End Function

    Public Shared Function InsertIntoParameterDefinition(ParentBlockID As Integer?, ChildBlockID As Integer?, DoubleChildBlockID As Integer?, ParamName As String)

        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()
            ' The SQL command below assumes your table has columns named ParentBlockID and ParentBlockName.
            Dim sql As String = "INSERT INTO ParameterDefinition (ParentBlockID, ChildBlockID, DoubleChildBlockID, ParameterName) 
            VALUES (@ParentBlockID, @ChildBlockID, @DoubleChildBlockID, @ParameterName)            
            "
            Using cmd As New SQLiteCommand(sql, conn)
                If ParentBlockID Is Nothing Then
                    cmd.Parameters.AddWithValue("@ParentBlockID", DBNull.Value)
                Else
                    cmd.Parameters.AddWithValue("@ParentBlockID", ParentBlockID)
                End If

                If ChildBlockID Is Nothing Then
                    cmd.Parameters.AddWithValue("@ChildBlockID", DBNull.Value)
                Else
                    cmd.Parameters.AddWithValue("@ChildBlockID", ChildBlockID)
                End If

                If DoubleChildBlockID Is Nothing Then
                    cmd.Parameters.AddWithValue("@DoubleChildBlockID", DBNull.Value)
                Else
                    cmd.Parameters.AddWithValue("@DoubleChildBlockID", DoubleChildBlockID)
                End If

                cmd.Parameters.AddWithValue("@ParameterName", ParamName)
                cmd.ExecuteNonQuery()
            End Using
            conn.Close()
        End Using
    End Function

    Public Shared Function GetParameterID(ChildBlockID As Integer?, ParentBlockID As Integer?, DoubleChildBlockID As Integer?, ParameterName As String)
        Dim ParameterID As Integer? = Nothing

        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()
            Dim sql = "
            Select ParameterID
            FROM ParameterDefinition
            WHERE ChildBlockID = @ChildBlockID OR ParentBlockID = @ParentBlockID OR DoubleChildBlockID = @DoubleChildBlockID AND ParameterName = @ParameterName
            "
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@ChildBlockID", ChildBlockID)
                cmd.Parameters.AddWithValue("@ParentBlockID", ParentBlockID)
                cmd.Parameters.AddWithValue("@ParameterName", ParameterName)
                cmd.Parameters.AddWithValue("@DoubleChildBlockID", DoubleChildBlockID)
                Dim result As Object = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    ParameterID = Convert.ToInt32(result)
                End If
            End Using
            conn.Close()
        End Using
        Return ParameterID
    End Function

    Public Shared Sub DeleteOrphanedParameterValues()
        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()

            Dim sql As String = "
            DELETE FROM ParameterValueLibrary
            WHERE ParameterID NOT IN (
                SELECT ParameterID FROM ParameterDefinition
            )
        "

            Using cmd As New SQLiteCommand(sql, conn)
                Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                vbLf & "Deleted " & rowsAffected.ToString() & " orphaned parameter values.")
            End Using

            conn.Close()
        End Using
    End Sub

    Public Shared Sub DeleteOrphanedAttributes()
        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()

            ' Delete attributes whose ParentBlockID doesn't exist anymore
            Dim sql1 As String = "
            DELETE FROM AttributeDefinition
            WHERE ParentBlockID IS NOT NULL
              AND ParentBlockID NOT IN (SELECT ParentBlockID FROM ParentBlockDefinition)
        "

            ' Delete attributes whose ChildBlockID doesn't exist anymore
            Dim sql2 As String = "
            DELETE FROM AttributeDefinition
            WHERE ChildBlockID IS NOT NULL
              AND ChildBlockID NOT IN (SELECT ChildBlockID FROM ChildBlockDefinition)
        "

            Using cmd1 As New SQLiteCommand(sql1, conn)
                Dim rowsAffected1 As Integer = cmd1.ExecuteNonQuery()
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                vbLf & "Deleted " & rowsAffected1.ToString() & " orphaned attributes linked to missing ParentBlockID.")
            End Using

            Using cmd2 As New SQLiteCommand(sql2, conn)
                Dim rowsAffected2 As Integer = cmd2.ExecuteNonQuery()
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                vbLf & "Deleted " & rowsAffected2.ToString() & " orphaned attributes linked to missing ChildBlockID.")
            End Using

            conn.Close()
        End Using
    End Sub

    Public Shared Sub InsertIntoAttributeValueLibrary(attributeID As Integer, attributeValue As String)
        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()

            Dim sql As String = "
            INSERT INTO AttributeValueLibrary (AttributeID, AttributeValue)
            VALUES (@AttributeID, @AttributeValue)
        "

            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@AttributeID", attributeID)
                cmd.Parameters.AddWithValue("@AttributeValue", attributeValue)
                cmd.ExecuteNonQuery()
            End Using

            conn.Close()
        End Using
    End Sub

    Public Shared Function InsertIntoParameterValueLibrary(ParameterID As Integer?, value As String)
        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()
            ' The SQL command below assumes your table has columns named ParentBlockID and ParentBlockName.
            Dim sql As String = "INSERT INTO ParameterValueLibrary (ParameterID, Value) 
            VALUES (@ParameterID, @Value)            
            "
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@ParameterID", ParameterID)
                cmd.Parameters.AddWithValue("@Value", value)
                cmd.ExecuteNonQuery()
            End Using
            conn.Close()
        End Using
    End Function

    Public Shared Function AttributeDefinitionExists(parentBlockID As Integer?, childBlockID As Integer?, attributeName As String) As Boolean
        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()
            Dim sql As String = "SELECT COUNT(*) FROM AttributeDefinition WHERE AttributeName = @AttributeName"

            If parentBlockID.HasValue Then
                sql &= " AND ParentBlockID = @ParentBlockID"
            Else
                sql &= " AND ParentBlockID IS NULL"
            End If

            If childBlockID.HasValue Then
                sql &= " AND ChildBlockID = @ChildBlockID"
            Else
                sql &= " AND ChildBlockID IS NULL"
            End If

            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@AttributeName", attributeName)
                If parentBlockID.HasValue Then cmd.Parameters.AddWithValue("@ParentBlockID", parentBlockID.Value)
                If childBlockID.HasValue Then cmd.Parameters.AddWithValue("@ChildBlockID", childBlockID.Value)
                Dim count As Integer = Convert.ToInt32(cmd.ExecuteScalar())
                Return (count > 0)
            End Using
        End Using
    End Function

    Public Shared Sub InsertIntoAttributeDefinition(parentBlockID As Integer?, childBlockID As Integer?, attributeName As String)
        Using conn As New SQLiteConnection(connectionString)
            conn.Open()

            Dim sql As String = "
            INSERT INTO AttributeDefinition (ParentBlockID, ChildBlockID, AttributeName)
            VALUES (@ParentBlockID, @ChildBlockID, @AttributeName)
        "

            Using cmd As New SQLiteCommand(sql, conn)
                If parentBlockID.HasValue Then
                    cmd.Parameters.AddWithValue("@ParentBlockID", parentBlockID.Value)
                Else
                    cmd.Parameters.AddWithValue("@ParentBlockID", DBNull.Value)
                End If

                If childBlockID.HasValue Then
                    cmd.Parameters.AddWithValue("@ChildBlockID", childBlockID.Value)
                Else
                    cmd.Parameters.AddWithValue("@ChildBlockID", DBNull.Value)
                End If

                cmd.Parameters.AddWithValue("@AttributeName", attributeName)
                cmd.ExecuteNonQuery()
            End Using

            conn.Close()
        End Using
    End Sub

    Public Shared Sub DeleteParameterValues(ByVal parameterID As Integer?)
        If Not parameterID.HasValue Then Exit Sub

        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()
            Dim sql As String = "DELETE FROM ParameterValueLibrary WHERE ParameterID = @ParameterID"
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@ParameterID", parameterID.Value)
                cmd.ExecuteNonQuery()
            End Using
            conn.Close()
        End Using
    End Sub

    Public Shared Sub DeleteOrphanedAttributeValues()
        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()

            Dim sql As String = "
            DELETE FROM AttributeValueLibrary
            WHERE AttributeID NOT IN (
                SELECT AttributeID FROM AttributeDefinition
            )
        "

            Using cmd As New SQLiteCommand(sql, conn)
                Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                vbLf & "Deleted " & rowsAffected.ToString() & " orphaned attribute values from AttributeValueLibrary.")
            End Using

            conn.Close()
        End Using
    End Sub

    Public Shared Function GetParameterValueID(parameterID As Integer, Value As String) As Integer?
        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()

            Dim sql As String = "
            SELECT ParameterValueID 
            FROM ParameterValueLibrary
            WHERE ParameterID = @ParameterID AND Value = @Value
        "

            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@ParameterID", parameterID)
                cmd.Parameters.AddWithValue("@Value", Value)

                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing Then
                    Return Convert.ToInt32(result)
                Else
                    Return Nothing
                End If
            End Using

            conn.Close()
        End Using
    End Function

    Public Shared Sub InsertIntoFlipVisibilityStatesLibrary(Value As Integer?)
        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()

            Dim sql As String = "
            INSERT INTO FlipVisibilityStatesLibrary (ParameterValueID)
            VALUES (@ParameterValueID)
        "

            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@ParameterValueID", Value)
                cmd.ExecuteNonQuery()
            End Using

            conn.Close()
        End Using
    End Sub

    Public Shared Sub DeleteOrphanedFlipVisibilityStates()
        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()

            Dim sql As String = "
            DELETE FROM FlipVisibilityStatesLibrary
            WHERE ParameterValueID NOT IN (
                SELECT ParameterValueID FROM ParameterValueLibrary
            )
        "

            Using cmd As New SQLiteCommand(sql, conn)
                Dim rowsAffected As Integer = cmd.ExecuteNonQuery()
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                vbLf & "Deleted " & rowsAffected.ToString() & " orphaned entries from FlipVisibilityStatesLibrary.")
            End Using

            conn.Close()
        End Using
    End Sub


    Private Sub UploadToParameterDefinition_Click(sender As Object, e As EventArgs)
        UploadBlockLibraryData()
    End Sub


    '--------POTENTIAL TO UPLOAD DRAWING DATA TO THE DATABASE. BUT THIS INFORMATION MIGHT JUST BE THE SAME AS FROM SAP. FOR EXAMPLE TRACKING HOW OFTEN DIFFERENT P203 Pumps ARE USED--------

    'Upload DWG Files to Database
    Private Function UploadDrawingToDatabase_Click(sender As Object, e As EventArgs)
        Dim filepath As String
        Dim openFile As New OpenFileDialog()
        openFile.Title = "Select an AutoCAD Drawing"
        ' Filter so that only DWG (or any relevant format) files are displayed.
        openFile.Filter = "AutoCAD Files (*.dwg)|*.dwg|All Files (*.*)|*.*"

        If openFile.ShowDialog() = DialogResult.OK Then
            ' Get the selected file path.
            filepath = openFile.FileName
            ' Use Path.GetFileNameWithoutExtension to extract the drawing name from the file name.
            Dim drawingName As String = Path.GetFileNameWithoutExtension(filepath)
            ' Retrieve the file creation date from the file system.
            Dim dateCreated As DateTime = File.GetCreationTime(filepath)

            ' Call the method to insert the drawing info into your database.
            InsertDrawingRecord(drawingName, filepath, dateCreated)
        End If
        Return filepath
    End Function

    ' This method handles the database connection and the insertion of the drawing record.
    Private Sub InsertDrawingRecord(drawingName As String, filePath As String, dateCreated As DateTime)
        ' Connection string to the SQLite database.
        Using conn As New SQLiteConnection(connectionstring)
            Try
                conn.Open()
                ' Use a parameterized query to safely insert data.
                Dim query As String = "INSERT INTO Drawings (DrawingName, FilePath, DateCreated) VALUES (@DrawingName, @FilePath, @DateCreated)"
                Using cmd As New SQLiteCommand(query, conn)
                    cmd.Parameters.AddWithValue("@DrawingName", drawingName)
                    cmd.Parameters.AddWithValue("@FilePath", filePath)
                    cmd.Parameters.AddWithValue("@DateCreated", dateCreated)
                    cmd.ExecuteNonQuery()
                End Using
                MessageBox.Show("The drawing record has been successfully uploaded.", "Upload Success", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Catch ex As Exception
                ' You can log the error or show a message box if the insertion fails.
                MessageBox.Show("Error inserting record: " & ex.Message, "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Finally
                conn.Close()
            End Try
        End Using
    End Sub

    Private Sub UploadParentBlockReferenceBtn_Click(sender As Object, e As EventArgs)
        LoadBlockReferences()
    End Sub

    Private Sub LoadBlockReferences()
        Dim dwgfilepath As String = DatabaseFunctions.GetFilePath("TestDatabase")
        Dim DrawingID As Integer? = GetDrawingID("TestDatabase")
        ClearChildBlocksForDrawing(DrawingID)
        ClearParentBlockReferencesForDrawing(DrawingID)
        ClearOrphanedParameterRows()
        Dim newDb As New Autodesk.AutoCAD.DatabaseServices.Database(False, True)
        newDb.ReadDwgFile(dwgfilepath, FileShare.Read, True, "")

        Using tr As Transaction = newDb.TransactionManager.StartTransaction()
            Dim bt As BlockTable = DirectCast(tr.GetObject(newDb.BlockTableId, OpenMode.ForRead), BlockTable)
            Dim ms As BlockTableRecord = DirectCast(tr.GetObject(bt(BlockTableRecord.ModelSpace), OpenMode.ForRead), BlockTableRecord)
            Dim i As Integer = 1
            ' Iterate through Model Space to find block references
            For Each entId As ObjectId In ms
                Dim ent As Entity = TryCast(tr.GetObject(entId, OpenMode.ForRead), Entity)
                If TypeOf ent Is BlockReference Then
                    Dim blkRef As BlockReference = DirectCast(ent, BlockReference)
                    Dim blockDef As BlockTableRecord = DirectCast(tr.GetObject(blkRef.DynamicBlockTableRecord, OpenMode.ForRead), BlockTableRecord)
                    Dim blkName As String = blockDef.Name ' Use the block definition name

                    ' Remove any "Clone" suffix with numbers.
                    Dim cleanName As String = RemoveCloneSuffix(blkName)
                    ' You can now use cleanName (for example, "SSV_Block") for further processing.
                    ' For demonstration, you might want to output the result to the console or a log.
                    System.Diagnostics.Debug.WriteLine("Original: " & blkName & " - Cleaned: " & cleanName)
                    Dim PositionX As Integer = blkRef.Position.X
                    Dim PositionY As Integer = blkRef.Position.Y
                    Dim Rotation As Integer = blkRef.Rotation
                    Dim Scale As Integer = blkRef.ScaleFactors.X
                    Dim parenthandle As String = blkRef.Handle.ToString()
                    Dim parentdynprops As DynamicBlockReferencePropertyCollection = blkRef.DynamicBlockReferencePropertyCollection
                    Dim ParentParameterNameValue As New List(Of Tuple(Of String, String))()
                    For Each props As DynamicBlockReferenceProperty In parentdynprops
                        If props.PropertyName.ToUpper().Contains("VISIBILITY") Then
                            Dim propname As String = props.PropertyName
                            Dim Currentvalue As String = props.Value.ToString()
                            ParentParameterNameValue.Add(Tuple.Create(propname, Currentvalue))
                        End If
                    Next

                    'Add/update parent blocks 
                    UploadToParentBlockReference(cleanName, PositionX, PositionY, Rotation, Scale, DrawingID, parenthandle) ' replaces the parent
                    Dim newParentId As Integer? = GetParentBlockReferenceID(parenthandle)
                    SyncBlockReferenceParameters(Nothing, newParentId, ParentParameterNameValue)

                    ' Collect visible child block refs
                    Dim visibleChildren As New List(Of Tuple(Of String, String))()
                    Dim childParamMap As New Dictionary(Of String, List(Of Tuple(Of String, String)))()


                    Dim ChildBlockReferences = Functions.GetRuntimeNestedBlocks(tr, blkRef)
                    For Each kvp In ChildBlockReferences
                        Dim childbr As BlockReference = kvp.Key
                        If childbr.Visible Then
                            Dim dynBlockId As ObjectId = childbr.DynamicBlockTableRecord
                            Dim dynBlockDef As BlockTableRecord = DirectCast(tr.GetObject(dynBlockId, OpenMode.ForRead), BlockTableRecord)
                            Dim childBlockName As String = dynBlockDef.Name
                            Dim childHandle As String = childbr.Handle.ToString()
                            visibleChildren.Add(Tuple.Create(childBlockName, childHandle))

                            'Parameters
                            Dim paramList As New List(Of Tuple(Of String, String))()
                            Dim dynProps As DynamicBlockReferencePropertyCollection = childbr.DynamicBlockReferencePropertyCollection
                            For Each prop As DynamicBlockReferenceProperty In dynProps
                                If prop.PropertyName.ToUpper().Contains("VISIBILITY") Then
                                    paramList.Add(Tuple.Create(prop.PropertyName, prop.Value.ToString()))
                                End If
                            Next
                            childParamMap(childHandle) = paramList ' Map handle to its specific parameters

                        End If
                    Next

                    SyncChildBlockReferences(newParentId, visibleChildren)

                    For Each kvp In childParamMap
                        Dim handle As String = kvp.Key
                        Dim paramList As List(Of Tuple(Of String, String)) = kvp.Value
                        Dim childID As Integer = GetChildBlockReferenceID(handle)

                        SyncBlockReferenceParameters(childID, Nothing, paramList)
                    Next

                    i += 1
                End If
            Next

            tr.Commit()
        End Using
        newDb.Dispose()
    End Sub

    Public Shared Function GetDrawingID(DrawingName As String)
        Dim DrawingID As Integer? = Nothing
        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()
            Dim sql = "
            Select DrawingID
            FROM Drawings
            WHERE DrawingName = @DrawingName
            "
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@DrawingName", DrawingName)
                Dim result As Object = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    DrawingID = Convert.ToInt32(result)
                End If
            End Using
            conn.Close()
        End Using
        Return DrawingID
    End Function

    ' Helper function to remove a "Clone" suffix plus trailing digits.
    ' For example, "SSV_BlockClone5" or "SSV_BlockClone10" will become "SSV_Block".
    Private Function RemoveCloneSuffix(ByVal blockName As String) As String
        ' Pattern matches "Clone" followed by one or more digits at the end of the string (case-insensitive).
        Dim pattern As String = "Clone\d+$"
        Dim regex As New Regex(pattern, RegexOptions.IgnoreCase)
        Dim newName As String = regex.Replace(blockName, "")
        Return newName
    End Function

    Public Shared Sub ClearParentBlockReferencesForDrawing(drawingID As Integer)
        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()
            Dim deleteSql As String = "DELETE FROM ParentBlockReferencesInDrawing WHERE DrawingID = @DrawingID"
            Using cmd As New SQLiteCommand(deleteSql, conn)
                cmd.Parameters.AddWithValue("@DrawingID", drawingID)
                cmd.ExecuteNonQuery()
            End Using
            conn.Close()
        End Using
    End Sub

    Public Shared Function UploadToParentBlockReference(BlockName, PositionX, PositionY, Rotation, Scale, DrawingID, handle)
        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()
            ' The SQL command below assumes your table has columns named ParentBlockID and ParentBlockName.
            Dim sql As String = "
            INSERT INTO ParentBlockReferencesInDrawing 
            (DrawingID, BlockName, InsertionX, InsertionY, Rotation, Scale, Handle)
            VALUES (@DrawingID, @BlockName, @InsertionX, @InsertionY, @Rotation, @Scale, @Handle)
            ON CONFLICT(Handle)
            DO UPDATE SET
            InsertionX = excluded.InsertionX,
            InsertionY = excluded.InsertionY,
            Rotation = excluded.Rotation,
            Scale = excluded.Scale
"
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@DrawingID", DrawingID)
                cmd.Parameters.AddWithValue("@BlockName", BlockName)
                cmd.Parameters.AddWithValue("@InsertionX", PositionX)
                cmd.Parameters.AddWithValue("@InsertionY", PositionY)
                cmd.Parameters.AddWithValue("@Rotation", Rotation)
                cmd.Parameters.AddWithValue("@Scale", Scale)
                cmd.Parameters.AddWithValue("@Handle", handle)
                cmd.ExecuteNonQuery()
            End Using
            conn.Close()
        End Using
    End Function

    Public Shared Function GetParentBlockReferenceID(Handle As String)
        Dim ParentBlockReferenceID As Integer? = Nothing
        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()
            Dim sql = "
            Select ParentBlockReferenceID
            FROM ParentBlockReferencesInDrawing
            WHERE Handle = @Handle
            "
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@Handle", Handle)
                Dim result As Object = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    ParentBlockReferenceID = Convert.ToInt32(result)
                End If
            End Using
            conn.Close()
        End Using
        Return ParentBlockReferenceID
    End Function

    'Clear child blocks connected to a parent block which is connected to the drawingid
    Public Shared Sub ClearChildBlocksForDrawing(drawingID As Integer)
        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()

            ' Delete all child blocks for all parent blocks in this drawing
            Dim sql As String = "
            DELETE FROM ChildBlockReferencesInDrawing
            WHERE ParentBlockReferenceID IN (
                SELECT ParentBlockReferenceID
                FROM ParentBlockReferencesInDrawing
                WHERE DrawingID = @DrawingID
            )"
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@DrawingID", drawingID)
                cmd.ExecuteNonQuery()
            End Using

            conn.Close()
        End Using
    End Sub

    Public Shared Sub SyncChildBlockReferences(ParentBlockReferenceID As Integer, visibleChildren As List(Of Tuple(Of String, String)))
        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()

            '' Step 1: Delete all existing child blocks for this parent
            'Dim deleteSql As String = "DELETE FROM ChildBlockReferencesInDrawing WHERE ParentBlockReferenceID = @ParentBlockReferenceID"
            'Using deleteCmd As New SQLiteCommand(deleteSql, conn)
            '    deleteCmd.Parameters.AddWithValue("@ParentBlockReferenceID", ParentBlockReferenceID)
            '    deleteCmd.ExecuteNonQuery()
            'End Using

            ' Step 2: Insert current visible child blocks
            For Each child In visibleChildren
                Dim sql As String = "
            INSERT INTO ChildBlockReferencesInDrawing 
            (ParentBlockReferenceID, BlockName, Handle)
            VALUES (@ParentBlockReferenceID, @BlockName, @Handle)          
"
                Using cmd As New SQLiteCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@BlockName", child.Item1)
                    cmd.Parameters.AddWithValue("@ParentBlockReferenceID", ParentBlockReferenceID)
                    cmd.Parameters.AddWithValue("@Handle", child.Item2)
                    cmd.ExecuteNonQuery()
                End Using
            Next
            conn.Close()
        End Using
    End Sub

    Public Shared Function GetChildBlockReferenceID(Handle As String)
        Dim ChildBlockReferenceID As Integer? = Nothing
        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()
            Dim sql = "
            Select ChildBlockReferenceID
            FROM ChildBlockReferencesInDrawing
            WHERE Handle = @Handle
            "
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@Handle", Handle)
                Dim result As Object = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Not IsDBNull(result) Then
                    ChildBlockReferenceID = Convert.ToInt32(result)
                End If
            End Using
            conn.Close()
        End Using
        Return ChildBlockReferenceID
    End Function

    Public Shared Function SyncBlockReferenceParameters(childID As Integer?, ParentID As Integer?, parameternamevalue As List(Of Tuple(Of String, String)))
        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()

            ' Step 1: Delete existing parameters for this block (either parent or child)
            Dim deleteSql As String = ""
            If childID.HasValue Then
                deleteSql = "DELETE FROM BlockReferenceParameters WHERE ChildBlockReferenceID = @ChildBlockReferenceID"
            ElseIf ParentID.HasValue Then
                deleteSql = "DELETE FROM BlockReferenceParameters WHERE ParentBlockReferenceID = @ParentBlockReferenceID"
            End If

            If Not String.IsNullOrEmpty(deleteSql) Then
                Using deleteCmd As New SQLiteCommand(deleteSql, conn)
                    If childID.HasValue Then
                        deleteCmd.Parameters.AddWithValue("@ChildBlockReferenceID", childID)
                    ElseIf ParentID.HasValue Then
                        deleteCmd.Parameters.AddWithValue("@ParentBlockReferenceID", ParentID)
                    End If
                    deleteCmd.ExecuteNonQuery()
                End Using
            End If

            ' Step 2: Insert new parameters
            For Each param In parameternamevalue
                Dim insertSql As String = "
                INSERT INTO BlockReferenceParameters 
                (ParentBlockReferenceID, ChildBlockReferenceID, ParameterName, ParameterValue) 
                VALUES (@ParentBlockReferenceID, @ChildBlockReferenceID, @ParamName, @ParamValue)
            "

                Using insertCmd As New SQLiteCommand(insertSql, conn)
                    insertCmd.Parameters.AddWithValue("@ParamName", param.Item1)
                    insertCmd.Parameters.AddWithValue("@ParamValue", param.Item2)

                    If childID.HasValue Then
                        insertCmd.Parameters.AddWithValue("@ChildBlockReferenceID", childID)
                    Else
                        insertCmd.Parameters.AddWithValue("@ChildBlockReferenceID", DBNull.Value)
                    End If

                    If ParentID.HasValue Then
                        insertCmd.Parameters.AddWithValue("@ParentBlockReferenceID", ParentID)
                    Else
                        insertCmd.Parameters.AddWithValue("@ParentBlockReferenceID", DBNull.Value)
                    End If

                    insertCmd.ExecuteNonQuery()
                End Using
            Next

            conn.Close()
        End Using

        Return True


    End Function

    Public Shared Sub ClearOrphanedParameterRows()
        Using conn As New SQLiteConnection(connectionstring)
            conn.Open()

            ' Delete parameter rows with child IDs that no longer exist
            Dim deleteOrphanedChildrenSql As String = "
            DELETE FROM BlockReferenceParameters
            WHERE ChildBlockReferenceID IS NOT NULL
              AND ChildBlockReferenceID NOT IN (
                SELECT ChildBlockReferenceID FROM ChildBlockReferencesInDrawing
              )
        "
            Using cmd1 As New SQLiteCommand(deleteOrphanedChildrenSql, conn)
                cmd1.ExecuteNonQuery()
            End Using

            ' Delete parameter rows with parent IDs that no longer exist
            Dim deleteOrphanedParentsSql As String = "
            DELETE FROM BlockReferenceParameters
            WHERE ParentBlockReferenceID IS NOT NULL
              AND ParentBlockReferenceID NOT IN (
                SELECT ParentBlockReferenceID FROM ParentBlockReferencesInDrawing
              )
        "
            Using cmd2 As New SQLiteCommand(deleteOrphanedParentsSql, conn)
                cmd2.ExecuteNonQuery()
            End Using

            conn.Close()
        End Using
    End Sub

End Class