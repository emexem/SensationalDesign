Imports System.Data.SqlServerCe
Imports System.Configuration
Imports System.Text.RegularExpressions

Public Class Form1

    Dim searchTxt, db, c, pk As String
    Public Property strSelectedValue As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load
        db = "PROJECT"
        c = "PROJECTNAME"
        pk = "TEAMID"
        Database.SearchRecords(db, DataGridView1)
    End Sub
    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        Close()
    End Sub

    Private Sub btnSearch_Click(sender As Object, e As EventArgs) Handles btnSearch.Click
        searchTxt = txtSearch.Text.ToString()
        If searchTxt = "" Then
            Database.SearchRecords(db, DataGridView1)
        Else
            Database.SearchRecords(searchTxt, c, db, DataGridView1)
        End If
        strSelectedValue = DataGridView1.Rows(0).Selected
    End Sub

    Private Sub BtnRemove_Click(sender As Object, e As EventArgs) Handles BtnRemove.Click
        Database.DeleteRecords(strSelectedValue, db, DataGridView1, pk)
    End Sub
End Class
