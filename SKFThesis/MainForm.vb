Imports System.Windows.Forms
Imports Autodesk.AutoCAD.ApplicationServices

Public Class MainForm
    Inherits Form

    Private btnOpenBlockSelector As New Button() With {.Text = "Open Block Selector"}
    Public Shared ExportToExcel As Boolean = False
    Public Shared ExcelFileName As String = "BOMExport"

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles ModifyBtn.Click
        Dim blockSelectorForm As New ModifyComponent()
        blockSelectorForm.Show()

        'Close mainform
        Me.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles InsertBtn.Click
        Dim form As New InsertComponent()
        form.Show()

        'Close mainform
        Me.Close()
    End Sub

    Private Sub UpdateBtn_Click(sender As Object, e As EventArgs) Handles UpdateBtn.Click
        If LayoutComboBox.SelectedItem = Nothing Then
            MessageBox.Show("Please enter layout paper size")
        Else
            Dim layoutName As String = LayoutComboBox.SelectedItem.ToString()
            BOMFunctions.BOMButton(layoutName)
        End If
    End Sub

    Private Sub UpdateBlockDef_Click(sender As Object, e As EventArgs) Handles UpdateBlockDef.Click
        Dim sourceLibPath = Functions.GetRad1Path() & "\" ' or wherever you store the library
        UpdateComponentInDrawing.UpdateClonedBlocksFromExternal(sourceLibPath)
        Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.Regen()
    End Sub

    Private Sub ExportExcelCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles ExportExcelCheckBox.CheckedChanged
        ExportToExcel = ExportExcelCheckBox.Checked
    End Sub

    Private Sub ExcelFileTextBox_TextChanged(sender As Object, e As EventArgs) Handles ExcelFileTextBox.TextChanged
        If Not String.IsNullOrEmpty(ExcelFileTextBox.Text) Then
            ExcelFileName = ExcelFileTextBox.Text
        End If
    End Sub
End Class