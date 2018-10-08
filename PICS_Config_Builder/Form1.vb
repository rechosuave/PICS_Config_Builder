
Public Class Form1

    Private Sub BtnDataAndRun_Click(sender As Object, e As EventArgs) Handles BtnDataAndRun.Click

        Call Main()

        PleaseWaitForm.Visible = False
        ' Set cursor as hourglass
        Cursor.Current = Cursors.Default

    End Sub

    Private Sub BtnExit_Click(sender As Object, e As EventArgs) Handles BtnExit.Click

        Close()

    End Sub

    Private Sub PleaseWaitForm_Click(sender As Object, e As EventArgs) Handles PleaseWaitForm.Click

    End Sub
End Class
