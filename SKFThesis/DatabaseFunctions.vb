Imports System.Configuration
Imports System.Data.Entity.Core.Mapping
Imports System.Data.SQLite
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
Imports Microsoft.SqlServer.Server
Imports Microsoft.VisualBasic.Devices
Imports Rad1.Functions
Imports Application = Autodesk.AutoCAD.ApplicationServices.Application

Public Class DatabaseFunctions

    'Works
    Public Shared Function GetVisibilityStatesFromDatabase(blockName As String) As List(Of String)
        Dim visibilityStates As New List(Of String)

        Using conn As New SQLiteConnection(UploadToDatabase.connectionstring)
            conn.Open()

            ' The query links the three tables:
            ' 1. ParentBlockDefinition to check the block name (e.g. "SSV_Block")
            ' 2. ParameterDefinition to get the ParameterID associated with that block via ParentBlockID.
            ' 3. ParameterValueLibrary to get the allowed values.
            Dim sql As String = "
        SELECT pvl.Value 
        FROM ParameterValueLibrary pvl 
        INNER JOIN ParameterDefinition pd ON pvl.ParameterID = pd.ParameterID 
        INNER JOIN ParentBlockDefinition pbd ON pd.ParentBlockID = pbd.ParentBlockID 
        WHERE pbd.ParentBlockName = @ParentBlockName"

            Using command As New SQLiteCommand(sql, conn)
                command.Parameters.AddWithValue("@ParentBlockName", blockName)
                Using reader As SQLiteDataReader = command.ExecuteReader()
                    While reader.Read()
                        visibilityStates.Add(reader("Value").ToString())
                    End While
                End Using
            End Using
        End Using

        Return visibilityStates
    End Function

    Public Shared Function GetChildVisibilityStatesFromDatabase(blockname As String) As List(Of String)
        Dim childvisibilityStates As New List(Of String)

        Using conn As New SQLiteConnection(UploadToDatabase.connectionstring)
            conn.Open()

            ' The query links the three tables:
            ' 1. ParentBlockDefinition to check the block name (e.g. "SSV_Block")
            ' 2. ParameterDefinition to get the ParameterID associated with that block via ParentBlockID.
            ' 3. ParameterValueLibrary to get the allowed values.
            Dim sql As String = "
        SELECT pvl.Value 
        FROM ParameterValueLibrary pvl 
        INNER JOIN ParameterDefinition pd ON pvl.ParameterID = pd.ParameterID 
        INNER JOIN ChildBlockDefinition cbd ON pd.ChildBlockID = cbd.ChildBlockID
        INNER JOIN ParentBlockDefinition pbd ON cbd.ParentBlockID = pbd.ParentBlockID 
        WHERE pd.ParameterName = @Visibility AND cbd.ChildBlockName = @ChildBlockName"

            Using command As New SQLiteCommand(sql, conn)
                command.Parameters.AddWithValue("@ChildBlockName", blockname)
                command.Parameters.AddWithValue("@Visibility", "Visibility")
                Using reader As SQLiteDataReader = command.ExecuteReader()
                    While reader.Read()
                        childvisibilityStates.Add(reader("Value").ToString())
                    End While
                End Using
            End Using
        End Using

        Return childvisibilityStates
    End Function

    Public Shared Function GetDoubleChildVisibilityStatesFromDatabase(blockname As String) As List(Of String)
        Dim doublechildvisibilityStates As New List(Of String)

        Using conn As New SQLiteConnection(UploadToDatabase.connectionstring)
            conn.Open()

            ' The query links the three tables:
            ' 1. ParentBlockDefinition to check the block name (e.g. "SSV_Block")
            ' 2. ParameterDefinition to get the ParameterID associated with that block via ParentBlockID.
            ' 3. ParameterValueLibrary to get the allowed values.
            Dim sql As String = "
        SELECT pvl.Value 
        FROM ParameterValueLibrary pvl 
        INNER JOIN ParameterDefinition pd ON pvl.ParameterID = pd.ParameterID 
        INNER JOIN DoubleChildBlockDefinition dcbd ON pd.DoubleChildBlockID = dcbd.DoubleChildBlockID
        WHERE pd.ParameterName = @Visibility AND dcbd.DoubleChildBlockName = @DoubleChildBlockName
"

            Using command As New SQLiteCommand(sql, conn)
                command.Parameters.AddWithValue("@DoubleChildBlockName", blockname)
                command.Parameters.AddWithValue("@Visibility", "Visibility")
                Using reader As SQLiteDataReader = command.ExecuteReader()
                    While reader.Read()
                        doublechildvisibilityStates.Add(reader("Value").ToString())
                    End While
                End Using
            End Using
        End Using

        Return doublechildvisibilityStates
    End Function

    Public Shared Function GetDoubleChildBlocks(childBlockName As String) As List(Of String)
        Dim doubleChildBlockNames As New List(Of String)

        Using conn As New SQLiteConnection(UploadToDatabase.connectionstring)
            conn.Open()

            ' First, get the ChildBlockID for the given childBlockName
            Dim getChildIDSql As String = "SELECT ChildBlockID FROM ChildBlockDefinition WHERE ChildBlockName = @ChildBlockName"
            Dim childBlockID As Integer = -1

            Using cmd As New SQLiteCommand(getChildIDSql, conn)
                cmd.Parameters.AddWithValue("@ChildBlockName", childBlockName)
                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing AndAlso Integer.TryParse(result.ToString(), childBlockID) Then
                    ' ChildBlockID found
                Else
                    Return doubleChildBlockNames ' Return empty list if not found
                End If
            End Using

            ' Now get DoubleChildBlockNames from DoubleChildBlockDefinition
            Dim sql As String = "SELECT DoubleChildBlockName FROM DoubleChildBlockDefinition WHERE ChildBlockID = @ChildBlockID"
            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@ChildBlockID", childBlockID)
                Using reader As SQLiteDataReader = cmd.ExecuteReader()
                    While reader.Read()
                        doubleChildBlockNames.Add(reader("DoubleChildBlockName").ToString())
                    End While
                End Using
            End Using
        End Using

        Return doubleChildBlockNames
    End Function

    Public Shared Function GetChildBlockDefinitions(parentblockname As String)
        Dim ChildBlockNames As New List(Of String)

        Using conn As New SQLiteConnection(UploadToDatabase.connectionstring)
            conn.Open()
            Dim sql As String = "
            SELECT cbd.ChildBlockName
            FROM ChildBlockDefinition cbd
            INNER JOIN ParentBlockDefinition pbd ON cbd.ParentBlockID = pbd.ParentBlockID
            WHERE pbd.ParentBlockName = @ParentBlockName
            "
            Using command As New SQLiteCommand(sql, conn)
                command.Parameters.AddWithValue("@ParentBlockName", parentblockname)
                Using reader As SQLiteDataReader = command.ExecuteReader()
                    While reader.Read()
                        ChildBlockNames.Add(reader("ChildBlockName").ToString())
                    End While
                End Using
            End Using
        End Using

        Return ChildBlockNames
    End Function

    Public Shared Function GetChildAttribute(childblockname As String, attributename As String)
        Dim ChildAttribute As String

        Using conn As New SQLiteConnection(UploadToDatabase.connectionstring)
            conn.Open()
            Dim sql As String = "
            SELECT avl.AttributeValue
            FROM AttributeValueLibrary avl
            INNER JOIN AttributeDefinition ad ON avl.AttributeID = ad.AttributeID
            INNER JOIN ChildBlockDefinition cbd ON ad.ChildBlockID = cbd.ChildBlockID
            WHERE cbd.ChildBlockName = @ChildBlockName AND ad.AttributeName = @AttributeName
            "
            Using command As New SQLiteCommand(sql, conn)
                command.Parameters.AddWithValue("@ChildBlockName", childblockname)
                command.Parameters.AddWithValue("@AttributeName", attributename)
                Using reader As SQLiteDataReader = command.ExecuteReader()
                    While reader.Read()
                        ChildAttribute = reader("AttributeValue").ToString()
                    End While
                End Using
            End Using
        End Using

        Return ChildAttribute

    End Function

    Public Shared Function GetParentAttribute(parentblockname As String, attributename As String)
        Dim ParentAttribute As String
        Using conn As New SQLiteConnection(UploadToDatabase.connectionstring)
            conn.Open()
            Dim sql As String = "
            SELECT avl.AttributeValue
            FROM AttributeValueLibrary avl
            INNER JOIN AttributeDefinition ad ON avl.AttributeID = ad.AttributeID
            INNER JOIN ParentBlockDefinition pbd ON ad.ParentBlockID = pbd.ParentBlockID
            WHERE pbd.ParentBlockName = @ParentBlockName AND ad.AttributeName = @AttributeName
            "
            Using command As New SQLiteCommand(sql, conn)
                command.Parameters.AddWithValue("@ParentBlockName", parentblockname)
                command.Parameters.AddWithValue("@AttributeName", attributename)
                Using reader As SQLiteDataReader = command.ExecuteReader()
                    While reader.Read()
                        ParentAttribute = reader("AttributeValue").ToString()
                    End While
                End Using
            End Using
        End Using

        Return ParentAttribute
    End Function

    Public Shared Function UploadParameterDefinition()

        Dim sql As String = "INSERT INTO ParameterDefinition (ParameterName, ParentID) " &
                           "VALUES (@BlockName, @ParentID); SELECT last_insert_rowid();"
    End Function

    Public Shared Function GetFilePath(drawingname As String)
        Dim Filepath As String

        Using conn As New SQLiteConnection(UploadToDatabase.connectionstring)
            conn.Open()
            Dim sql As String = "
            SELECT d.FilePath
            FROM Drawings d
            WHERE d.DrawingName = @DrawingName
            "
            Using command As New SQLiteCommand(sql, conn)
                command.Parameters.AddWithValue("@DrawingName", drawingname)
                Using reader As SQLiteDataReader = command.ExecuteReader()
                    While reader.Read()
                        Filepath = reader("FilePath").ToString()
                    End While
                End Using
            End Using
        End Using
        Return Filepath
    End Function

    Public Shared Function GetFlipParameter(ParameterValueID As Integer?) As Boolean
        Using conn As New SQLiteConnection(UploadToDatabase.connectionstring)
            conn.Open()
            Dim sql As String = "
            SELECT 1 
            FROM FlipVisibilityStatesLibrary
            WHERE ParameterValueID = @ParameterValueID
            LIMIT 1
        "

            Using command As New SQLiteCommand(sql, conn)
                command.Parameters.AddWithValue("@ParameterValueID", ParameterValueID)
                Dim result = command.ExecuteScalar()
                Return result IsNot Nothing
            End Using
        End Using
    End Function

    Public Shared Function GetParameterValueID(parameterID As Integer, parameterValue As String) As Integer?
        Using conn As New SQLiteConnection(UploadToDatabase.connectionstring)
            conn.Open()

            Dim sql As String = "
            SELECT ParameterValueID
            FROM ParameterValueLibrary
            WHERE ParameterID = @ParameterID AND Value = @Value
            LIMIT 1
        "

            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@ParameterID", parameterID)
                cmd.Parameters.AddWithValue("@Value", parameterValue)

                Dim result = cmd.ExecuteScalar()
                If result IsNot Nothing Then
                    Return Convert.ToInt32(result)
                Else
                    Return Nothing
                End If
            End Using
        End Using
    End Function

    Public Shared Function GetMatchingParentBlock(code As String) As String
        Dim bestMatch As String = Nothing
        Dim longestMatchLength As Integer = 0

        Using conn As New SQLiteConnection(UploadToDatabase.connectionstring)
            conn.Open()
            Dim sql = "SELECT ParentBlockName FROM ParentBlockDefinition"

            Using cmd As New SQLiteCommand(sql, conn)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim blockName = reader("ParentBlockName").ToString()
                        Dim prefix = blockName.Replace("_Block", "").ToUpper()

                        If code.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then
                            If prefix.Length > longestMatchLength Then
                                longestMatchLength = prefix.Length
                                bestMatch = blockName
                            End If
                        End If
                    End While
                End Using
            End Using
        End Using

        Return bestMatch
    End Function

    Public Shared Function ExtractRemainingIDCode(afterPrefixFrom As String, parentBlockName As String) As String
        ' Example: parentBlockName = "SSV_Block" → prefix = "SSV"
        Dim prefix As String = parentBlockName.Replace("_Block", "").ToUpper()

        If afterPrefixFrom.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) Then
            Return afterPrefixFrom.Substring(prefix.Length)
        End If

        Return afterPrefixFrom ' fallback
    End Function

    Public Shared Function MatchIDCodeValuesInString(parentBlockID As Integer, remainingID As String) As List(Of (ParameterID As Integer, ParameterValue As String))
        Dim matches As New List(Of (ParameterID As Integer, ParameterValue As String, IDCodeValue As String))

        Using conn As New SQLiteConnection(UploadToDatabase.connectionstring)
            conn.Open()

            Dim sql = "
            SELECT pd.ParameterID, pvl.IDCodeValue, pvl.Value
            FROM ParameterDefinition pd
            LEFT JOIN ParameterValueLibrary pvl ON pd.ParameterID = pvl.ParameterID
            LEFT JOIN ChildBlockDefinition cbd ON pd.ChildBlockID = cbd.ChildBlockID
            WHERE pd.ParentBlockID = @ParentBlockID OR cbd.ParentBlockID = @ParentBlockID
        "

            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@ParentBlockID", parentBlockID)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim paramID = Convert.ToInt32(reader("ParameterID"))
                        Dim idCodeValue = reader("IDCodeValue").ToString()
                        Dim value = reader("Value").ToString()

                        ' Check if the remaining string contains the IDCodeValue
                        If Not String.IsNullOrWhiteSpace(idCodeValue) AndAlso
                       remainingID.IndexOf(idCodeValue, StringComparison.OrdinalIgnoreCase) >= 0 Then

                            matches.Add((paramID, value, idCodeValue))
                        End If
                    End While
                End Using
            End Using
        End Using

        ' ✅ Group by ParameterID, keep the longest matching IDCodeValue
        Dim finalResults As New List(Of (Integer, String))

        Dim grouped = matches.GroupBy(Function(m) m.ParameterID)

        For Each group In grouped
            ' Pick the longest IDCodeValue match
            Dim bestMatch = group.OrderByDescending(Function(m) m.IDCodeValue.Length).First()
            finalResults.Add((bestMatch.ParameterID, bestMatch.ParameterValue))
        Next

        Return finalResults
    End Function

    Public Shared Function GetParameterValuesForTokens(parentBlockID As Integer, tokens As List(Of String)) As List(Of (Integer, String))
        Dim results As New List(Of (Integer, String))

        Using conn As New SQLiteConnection(UploadToDatabase.connectionstring)
            conn.Open()

            ' Get all ParameterIDs for parent and its children
            Dim paramIDs As New List(Of Integer)
            Dim sql = "
            SELECT pd.ParameterID, pvl.IDCodeValue, pvl.Value
            FROM ParameterDefinition pd
            LEFT JOIN ParameterValueLibrary pvl ON pd.ParameterID = pvl.ParameterID
            LEFT JOIN ChildBlockDefinition cbd ON pd.ChildBlockID = cbd.ChildBlockID
            WHERE pd.ParentBlockID = @ParentBlockID OR cbd.ParentBlockID = @ParentBlockID
        "

            Using cmd As New SQLiteCommand(sql, conn)
                cmd.Parameters.AddWithValue("@ParentBlockID", parentBlockID)

                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim token = reader("IDCodeValue").ToString()
                        If tokens.Contains(token) Then
                            Dim paramID As Integer = Convert.ToInt32(reader("ParameterID"))
                            Dim paramValue As String = reader("Value").ToString()
                            results.Add((paramID, paramValue))
                        End If
                    End While
                End Using
            End Using
        End Using

        Return results
    End Function

    Public Shared Function GetAllParameterDefinitions() As List(Of ParameterDef)
        Dim list As New List(Of ParameterDef)

        Using conn As New SQLiteConnection(UploadToDatabase.connectionstring)
            conn.Open()
            Dim sql = "SELECT ParameterID, ParameterName, ParentBlockID, ChildBlockID FROM ParameterDefinition"

            Using cmd As New SQLiteCommand(sql, conn)
                Using reader = cmd.ExecuteReader()
                    While reader.Read()
                        Dim paramID As Integer = Convert.ToInt32(reader("ParameterID"))
                        Dim paramName As String = reader("ParameterName").ToString()

                        ' Handle NULLs properly using TryCast or nullable Convert
                        Dim parentID As Integer? = Nothing
                        If Not IsDBNull(reader("ParentBlockID")) Then
                            parentID = Convert.ToInt32(reader("ParentBlockID"))
                        End If

                        Dim childID As Integer? = Nothing
                        If Not IsDBNull(reader("ChildBlockID")) Then
                            childID = Convert.ToInt32(reader("ChildBlockID"))
                        End If

                        list.Add(New ParameterDef With {
                        .ParameterID = paramID,
                        .ParameterName = paramName,
                        .ParentBlockID = parentID,
                        .ChildBlockID = childID
                    })
                    End While
                End Using
            End Using
        End Using

        Return list
    End Function
End Class
