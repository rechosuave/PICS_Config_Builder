
Imports Excel = Microsoft.Office.Interop.Excel

Module ImportData

    Sub Main()

        'Select and open the Excel project file (PLC IO Mapping) that will be used to create PICS simulation files
        Dim XLProjectWB As Workbook
        XLProjectWB = OpenXLProjectFN()
        If XLProjectWB Is Nothing Then 'No project file selected
            Exit Sub
        End If

        'Create or update the Excel PICS Config file that will organize data
        Dim pathname As New IO.FileInfo(XLProjectWB.Name)
        Dim CurDir As String = pathname.Directory.Name
        Dim XLpicsWB As Workbook
        Dim response As MsgBoxResult
        Dim XLpicsFN As String

        response = MsgBox("Create A New PICS Config Builder file?", vbYesNo)

        If response = vbYes Then
            XLpicsFN = InputBox("Enter New PICS Config Builder File Name:", "New File Name", "PICS_Config_Builder",,)
            XLpicsWB = OpenXLpicsFN(CurDir + XLpicsFN)

        Else
            XLpicsFN = ""
            XLpicsWB = OpenXLpicsFN(XLpicsFN)

        End If

        Call Button_Import_Data(XLProjectWB)

        XLpicsWB.Application.ScreenUpdating = False

        Call Generate_Sim_Data(XLpicsWB)
        Call Generate_Memory_Data(XLpicsWB)
        Call Generate_Wire_Data(XLpicsWB)

        Dim outFolder As String
        outFolder = Create_Output_Folder(XLpicsWB)

        Call Export_CSV(outFolder, "SimData", "OPC_Tags.csv")
        Call Export_CSV(outFolder, "MemoryData", "GLOBAL_Tags.csv")
        Call Export_Wire_Data(XLpicsWB, outFolder)

        XLpicsWB.Application.ScreenUpdating = False
        XLpicsWB.Close()

        If XLProjectWB.Name.Contains(".xlsm") Then  'Re-enable Excel application macros security settings prior to closing file
            XLProjectWB.Application.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityLow
            XLProjectWB.Close()
        Else
            XLProjectWB.Close()

        End If

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

        Dim ws As Worksheet
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
        ws = CType(wrkBook.Sheets("Instructions"), Worksheet)
        ws.Range("CPU_PREFIX").ClearContents()

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

    Function OpenXLProjectFN() As Workbook

        'Select and open the Excel project file (PLC IO Mapping) that will be used to create PICS simulation files
        Dim XLApp As Excel.Application
        Dim XLWrkBook As Excel.Workbook
        Dim XLWrkSheet As Excel.Worksheet
        Dim sFileN As String
        Dim title = "Open - Select Project Config File"
        Dim fnXtnFilter = "Excel Files (*.xls;*.xlsm),*.xls;*xlsm"

        XLApp = CType(CreateObject("Excel.Application"), Excel.Application)
        sFileN = CType(XLApp.GetOpenFilename(FileFilter:=fnXtnFilter, FilterIndex:=2, Title:=title), String)

        If sFileN = False Then 'operator cancelled operation to open the project file
            XLApp.Quit()
            Return Nothing
        End If

        If sFileN.Contains(".xlsm") Then  'Disable Excel application macros prior to opening file
            XLApp.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable
        End If

        XLWrkBook = XLApp.Workbooks.Open(Filename:=sFileN, ReadOnly:=True)
        XLWrkSheet = CType(XLWrkBook.ActiveSheet, Worksheet)

        XLWrkSheet.Visible = CType(True, Excel.XlSheetVisibility)

        Return CType(XLWrkBook, Workbook)

    End Function

    Function OpenXLpicsFN(ByVal FileName As String) As Workbook

        'Select or create PICS Excel file that will be used to create PICS simulation files
        Dim XLApp As Excel.Application
        Dim XLWrkBook As Excel.Workbook
        Dim XLWrkSheet As Excel.Worksheet
        Dim title = "Open - Select PICS Config File"
        Dim fnXtnFilter = "Excel Files (*.xls;*.xlsm),*.xls;*xlsm"

        XLApp = CType(CreateObject("Excel.Application"), Excel.Application)

        If FileName <> "" Then 'open existing or create a new PICS Config Excel File
            XLWrkBook = XLApp.Workbooks.Open(FileName)
            XLWrkSheet = CType(XLWrkBook.ActiveSheet, Excel.Worksheet)
        Else
            FileName = CType(XLApp.GetOpenFilename(FileFilter:=fnXtnFilter, FilterIndex:=2, Title:=title), String)
            XLWrkBook = XLApp.Workbooks.Open(FileName)
            XLWrkSheet = CType(XLWrkBook.ActiveSheet, Excel.Worksheet)

        End If

        XLWrkSheet.Visible = CType(True, Excel.XlSheetVisibility)

        Return CType(XLWrkBook, Workbook)

    End Function
End Module