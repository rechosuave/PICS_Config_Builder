
Imports Office = Microsoft.Office.Interop
Imports Excel = Microsoft.Office.Interop.Excel

Module ImportData

    Sub Button_Data_And_Run(ByRef wb As Workbook)

        Call Button_Import_Data(wb)

        wb.Application.ScreenUpdating = False

        Call Generate_Sim_Data(wb)
        Call Generate_Memory_Data(wb)
        Call Generate_Wire_Data(wb)

        Dim outFolder As String
        outFolder = Create_Output_Folder(wb)

        Call Export_CSV(outFolder, "SimData", "OPC_Tags.csv")
        Call Export_CSV(outFolder, "MemoryData", "GLOBAL_Tags.csv")
        Call Export_Wire_Data(wb, outFolder)

        wb.Application.ScreenUpdating = False

    End Sub

    Sub Button_Import_Data(ByRef wb As Workbook)

        Dim projectfN As String
        Dim picsBuilder As String
        Dim projectBuilder As String
        Dim cpuImport As String

        wb.Application.ScreenUpdating = False

        wb.Application.DisplayAlerts = False 'Turn safety alerts OFF

        Call Unhide_All_Sheets(wb)

        Dim ws As Worksheet = wb.Sheets("IO Sheets").Select
        ws.Range("A2:AA9999").Clear()

        picsBuilder = wb.Name

        projectfN = GetProjectFN()
        If projectfN = Nothing Then Exit Sub

        Dim xlApp As New Excel.Application
        Dim xlProjectWorkBook As Workbook = xlApp.Workbooks.Open(projectfN)

        projectBuilder = xlProjectWorkBook.Name

        cpuImport = xlProjectWorkBook.Sheets("Instructions").Range("C3").Value
        xlProjectWorkBook.Sheets("IO Sheets").UsedRange.Copy

        ' Paste entire IO sheet
        wb.Activate()
        ws = wb.Sheets("IO Sheets").Select
        ws.Range("A1").PasteSpecial.xlPasteValues

        ' Remove any white space at the top
        Do While ws.Range("A1").Value <> "PLCBaseTag"
            ws.Range("A1").EntireRow.Delete()
        Loop

        ' Fix all selections to look nice
        If wb.Sheets("Instructions").Range("CPU_PREFIX").Value = "" Then
            wb.Sheets("Instructions").Range("CPU_PREFIX").Value = cpuImport
        End If

        Reset_Sheet(wb, "Instructions")
        Reset_Sheet(wb, "IO Sheets")
        ws = wb.Sheets("Instructions").Select

        Call Hide_Sheets(wb)

        xlProjectWorkBook.Close(SaveChanges:=False)
        wb.Application.DisplayAlerts = True 'Turn safety alerts ON

        wb.Application.ScreenUpdating = True

    End Sub

    Sub Button_Clear_All_Sheets(ByRef wrkBook As Workbook)
        '
        'WARNING!!! This will clear all data AND delete all Wire sheets
        If MsgBox("WARNING! This will clear all data from this workbook and delete existing Wire data sheets.", vbOKCancel) = vbCancel Then Exit Sub

        wrkBook.Application.ScreenUpdating = False

        Call Button_Unhide_All_Sheets(wrkBook)

        Clear_All_Sheets(wrkBook)

        Call Delete_Wire_Sheets(wrkBook, "Wire_AIn Template")
        Call Delete_Wire_Sheets(wrkBook, "Wire_DIn Template")
        Call Delete_Wire_Sheets(wrkBook, "Wire_ValveC Template")
        Call Delete_Wire_Sheets(wrkBook, "Wire_ValveMO Template")
        Call Delete_Wire_Sheets(wrkBook, "Wire_ValveSO Template")
        Call Delete_Wire_Sheets(wrkBook, "Wire_Motor Template")
        Call Delete_Wire_Sheets(wrkBook, "Wire_VSD Template")

        Call Button_Hide_Sheets(wrkBook)
        wrkBook.Sheets("Instructions").Range("CPU_PREFIX").ClearContents

        wrkBook.Application.ScreenUpdating = True

    End Sub

    Public Function GetProjectFN() As String
        ' OpenFile method used to quickly open a file from the dialog box. 
        ' The file Is opened In read-only mode For security purposes. 
        ' To open a file In read/write mode, you must use another method, such as FileStream.

        Dim title, fnXtnFilter As String

        title = "Select Project Config File"
        fnXtnFilter = "Excel files (*.xls;*.xlsm)"

        Dim openFileDialog1 = New OpenFileDialog()

        openFileDialog1.Title = title
        openFileDialog1.InitialDirectory = "c:\\"
        openFileDialog1.Filter = fnXtnFilter
        openFileDialog1.FilterIndex = 2
        openFileDialog1.RestoreDirectory = True

        If (openFileDialog1.ShowDialog() = DialogResult.OK) Then
            GetProjectFN = openFileDialog1.FileName
        Else
            GetProjectFN = Nothing

        End If

    End Function
End Module