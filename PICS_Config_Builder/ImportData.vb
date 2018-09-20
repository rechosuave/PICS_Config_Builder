
Imports Excel = Microsoft.Office.Interop.Excel

Module ImportData

    Sub Button_Data_And_Run()

        'Select and open the Excel project file (PLC IO Mapping) that will be used to create PICS simulation files
        Dim XLProjectWB = OpenXLProjectFN()

        'Create or update the Excel PICS Config file that will organize all data
        Call Button_Import_Data(XLProjectWB)

        'XLWrkBook.Application.ScreenUpdating = False

        'Call Generate_Sim_Data(XLWrkBook)
        'Call Generate_Memory_Data(XLWrkBook)
        'Call Generate_Wire_Data(XLWrkBook)

        'Dim outFolder As String
        'outFolder = Create_Output_Folder(XLWrkBook)

        'Call Export_CSV(outFolder, "SimData", "OPC_Tags.csv")
        'Call Export_CSV(outFolder, "MemoryData", "GLOBAL_Tags.csv")
        'Call Export_Wire_Data(XLWrkBook, outFolder)

        'XLWrkBook.Application.ScreenUpdating = False
        'XLApp.Quit()

    End Sub

    Sub Button_Import_Data(ByRef xlProjectWB As Workbook)

        'Create an Excel PICS Config file
        'Dim projectfN, picsBuilder, projectBuilder, cpuImport As String

        'wb.Application.ScreenUpdating = False
        'wb.Application.DisplayAlerts = False 'Turn safety alerts OFF

        'Call Unhide_All_Sheets(wb)

        'Dim ws As Worksheet = wb.Sheets("IO Sheets").Select
        'ws.Range("A2:AA9999").Clear()

        'picsBuilder = wb.Name

        'Dim xlApp As New Excel.Application
        'Dim xlProjectWorkBook As Workbook = xlApp.Workbooks.Open(projectfN)

        'projectBuilder = xlProjectWorkBook.Name

        'cpuImport = xlProjectWorkBook.Sheets("Instructions").Range("C3").Value
        'xlProjectWorkBook.Sheets("IO Sheets").UsedRange.Copy

        '' Paste entire IO sheet
        'wb.Activate()
        'ws = wb.Sheets("IO Sheets").Select
        'ws.Range("A1").PasteSpecial.xlPasteValues

        '' Remove any white space at the top
        'Do While ws.Range("A1").Value <> "PLCBaseTag"
        '    ws.Range("A1").EntireRow.Delete()
        'Loop

        '' Fix all selections to look nice
        'If wb.Sheets("Instructions").Range("CPU_PREFIX").Value = "" Then
        '    wb.Sheets("Instructions").Range("CPU_PREFIX").Value = cpuImport
        'End If

        'Reset_Sheet(wb, "Instructions")
        'Reset_Sheet(wb, "IO Sheets")
        'ws = wb.Sheets("Instructions").Select

        'Call Hide_Sheets(wb)

        'xlProjectWorkBook.Close(SaveChanges:=False)
        'wb.Application.DisplayAlerts = True 'Turn safety alerts ON

        'wb.Application.ScreenUpdating = True

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
            Return openFileDialog1.FileName
        Else
            Return Nothing

        End If

    End Function

    Function OpenXLProjectFN()

        'Select and open the Excel project file (PLC IO Mapping) that will be used to create PICS simulation files
        Dim XLApp As Excel.Application
        Dim XLWrkBook As Excel.Workbook
        Dim XLWrkSheet As Excel.Worksheet
        Dim FileName As String
        Dim title = "Open - Select Project Config File"
        Dim fnXtnFilter = "Excel Files (*.xls;*.xlsm),*.xls;*xlsm"

        XLApp = CType(CreateObject("Excel.Application"), Excel.Application)
        FileName = XLApp.GetOpenFilename(FileFilter:=fnXtnFilter, FilterIndex:=2, Title:=title)
        XLWrkBook = XLApp.Workbooks.Open(FileName)
        XLWrkSheet = XLWrkBook.ActiveSheet

        XLWrkSheet.Visible = True
        '        XLWrkBook.UserControl = True
        Return XLWrkBook

    End Function

End Module