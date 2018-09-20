
Imports Office = Microsoft.Office.Interop
Imports Excel = Microsoft.Office.Interop.Excel

Public Class Form1

    ' Declare variables, start Excel and get Application object
    Dim XLApp As New Excel.Application
    Dim XLWrkBook As Excel.Workbook = XLApp.Workbooks
    Dim XLWrkSheet As Excel.Worksheet = CType(XLWrkBook.ActiveSheet, Worksheet)

    Private Sub BtnDataAndRun_Click(sender As Object, e As EventArgs) Handles BtnDataAndRun.Click

        XLWrkBook.Visible = True
        XLApp.UserControl = True

        Call Button_Data_And_Run(XLWrkBook)
        MsgBox("Outputs Generated")

    End Sub

    Private Sub BtnClearAllSheets_Click(sender As Object, e As EventArgs) Handles BtnClearAllSheets.Click

        Clear_All_Sheets(XLWrkBook)
        MsgBox("All Sheets Cleared")

    End Sub

    Private Sub BtnExit_Click(sender As Object, e As EventArgs) Handles BtnExit.Click

        Close()

    End Sub
End Class
