
Public Class Form1

    Private Sub BtnDataAndRun_Click(sender As Object, e As EventArgs) Handles BtnDataAndRun.Click

        Call Main()

    End Sub

    Private Sub BtnClearAllSheets_Click(sender As Object, e As EventArgs) Handles BtnClearAllSheets.Click

        Clear_All_Sheets()

    End Sub

    Private Sub BtnExit_Click(sender As Object, e As EventArgs) Handles BtnExit.Click

        Close()

    End Sub

End Class
