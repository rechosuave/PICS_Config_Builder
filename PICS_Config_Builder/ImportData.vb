
Imports Microsoft.Office.Interop.Excel

Module ImportData

    Public XLApp As New Application
    Public XLProjectWB, XLpicsWB As Workbook
    Public DirectoryName, xlProjectFN, XLpicsFN, CPU_Name As String      'user selected directory for PICS Config File
    Public picsBuilder, projectBuilder As String    'Workbook names
    Const fnXtnFilter = "Excel Files (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*xlsm"

    Sub Main()

        'Open Excel Application sub-process
        Call Open_Excel()

        Call OpenXLProjectFN()      ' Select and open the Excel project file (PLC IO Mapping) 

        If XLProjectWB Is Nothing Then ' Project file not selected
            Exit Sub
        End If

        Call OpenXLpicsFN()     ' Open or Create the Excel PICS Config Builder file that is used to organize data

        If XLpicsWB Is Nothing Then 'No PICS Config Builder file selected
            Exit Sub
        End If

        Call Validate_PICS_WB()    ' Validate or add required worksheets to PICS config builder file

        ' If Not IsNothing(XLTemplateWB) Then ' no wire template file selected

        Call Import_Data()         ' Import "IO Sheets" worksheet from Project file into "IO Sheets" in PICS file

            Call Generate_Sim_Data()    ' Build OPC tags
            Call Generate_Memory_Data() ' Build Global tags
            Call Generate_Wire_Data()   ' Build Wire data files for PICS simulator

            Dim outFolder As String
            outFolder = Create_Output_Folder(XLpicsWB)

            Call Export_CSV(outFolder, "SimData", "OPC_Tags.csv")
            Call Export_CSV(outFolder, "MemoryData", "GLOBAL_Tags.csv")
            Call Export_Wire_Data(XLpicsWB, outFolder)

        '    End If

        If IsFileOpen(xlProjectFN) Then        ' is the project file workbook still open
            XLProjectWB.Close(SaveChanges:=False)
        End If

        If IsFileOpen(XLpicsFN) Then        ' is the PICS Config file workbook still open
            If XLpicsFN.Contains(".xlsm") Then     'Re-enable Excel application macros security settings prior to closing file
                Dim ws As Worksheet = XLpicsWB.Sheets(2)    ' return to "IO Sheets" worksheet
                ws.Activate()
            Else
                Dim ws As Worksheet = XLpicsWB.Sheets(1)
                ws.Activate()
            End If
            XLpicsWB.Close(SaveChanges:=True)
        End If

        XLApp.Quit()

    End Sub

    Sub Import_Data()

        'Populate Excel PICS Config file with data from Project file "IO Sheets" worksheet
        Dim shtName As String = "IO Sheets"
        Dim picsCellVal As String = "PLCBaseTag"
        Dim ws As Worksheet = XLpicsWB.Sheets(shtName)

        picsBuilder = XLpicsWB.Name
        projectBuilder = XLProjectWB.Name

        XLProjectWB.Sheets(shtName).UsedRange.Copy          ' copy from Project worksheet to clipboard
        ws.Range("A1").PasteSpecial(XlPasteType.xlPasteValues)  ' paste from clipboard to PICS worksheet

        Do While ws.Range("A1").Value <> picsCellVal   ' Remove white space from top row
            ws.Range("A1").EntireRow.Delete()
        Loop

        ' Add CPU_PREFIX to form title
        If Form1.CPU_PREFIX.Text = "" Then
            CPU_Name = XLProjectWB.Sheets("Instructions").Range("C3").Value
            Form1.CPU_PREFIX.Text = CPU_Name
        Else
            CPU_Name = Form1.CPU_PREFIX.Text
        End If

        If XLProjectWB.Name.Contains(".xlsm") Then  'Re-enable workbook macros security settings prior to closing Project file
            XLProjectWB.Application.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityLow
            XLProjectWB.Close(SaveChanges:=False)
        Else
            XLProjectWB.Close(SaveChanges:=False)

        End If

    End Sub

    Sub OpenXLProjectFN()

        'Select and open the Excel project file (PLC IO Mapping) that will be used to create PICS simulation files
        Dim title = "Open - Select Project Config File"

        xlProjectFN = XLApp.GetOpenFilename(FileFilter:=fnXtnFilter, FilterIndex:=2, Title:=title)
        If xlProjectFN Is Nothing Or xlProjectFN = "False" Then 'operator cancelled operation to open the project file
            XLApp.Quit()
            Exit Sub
        End If

        DirectoryName = IO.Path.GetDirectoryName(xlProjectFN)
        XLApp.DefaultFilePath = DirectoryName
        FileIO.FileSystem.CurrentDirectory = DirectoryName     ' set directory path

        If xlProjectFN.Contains(".xlsm") Then    'Disable Excel application macros security settings when opening file
            XLApp.Application.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable
        End If

        XLProjectWB = XLApp.Workbooks.Open(Filename:=xlProjectFN, ReadOnly:=True)

    End Sub

    Sub OpenXLpicsFN()

        'Select or create PICS Excel file(workbook) that will be used to create PICS simulation files
        Dim title = "Open - Select PICS Config File"
        Dim response As MsgBoxResult
        Dim fn As String

        response = MsgBox("Create A New PICS Config Builder file?", vbYesNo)

        If response = vbYes Then
            Do
                fn = InputBox("Enter New PICS Config Builder File Name:", "New File Name", "PICS_Config_Builder")
                If fn = "" Then     ' Cancel button pressed
                    response = MsgBox("Cancel Operation?", vbYesNo)
                    If response = vbYes Then Exit Sub
                End If
            Loop Until Not IsNothing(fn)
            XLpicsFN = DirectoryName & "\" & fn & ".xlsx"

        Else
            ' if response is no  - assume that you want to open an existing PICS file
            response = MsgBox("Open an existing PICS Config Builder file?", vbYesNo)
            If response = vbYes Then
                XLpicsFN = XLApp.GetOpenFilename(FileFilter:=fnXtnFilter, FilterIndex:=2, Title:=title)
            Else
                XLApp.Quit()  ' Since no PICS file was selected or created - close Excel app and exit to main screen form
                Exit Sub
            End If

        End If

        If IO.File.Exists(XLpicsFN) Then     'open existing PICS Config File
            If XLpicsFN.Contains(".xlsm") Then  'Disable Excel application macros prior to opening file
                XLApp.Application.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable
            End If
            XLpicsWB = XLApp.Workbooks.Open(XLpicsFN)
        Else
            XLpicsWB = XLApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet)  ' create new workbook for PICS Config file
            XLpicsWB.SaveAs(XLpicsFN)

        End If

    End Sub

    Sub Open_Excel()

        XLApp = CreateObject("Excel.Application")
        XLApp.Application.ScreenUpdating = False
        XLApp.Application.DisplayAlerts = False 'Turn safety alerts OFF

    End Sub

    Function IsNothing(ByRef objValue) As Boolean

        ' Check for required user response - returns True if Empty or NULL
        If VarType(objValue) = vbEmpty Or VarType(objValue) = vbNull Then
            Return True
        ElseIf VarType(objValue) = vbString Then
            If objValue = "" Or objValue = "False" Then
                Return True
            End If
        ElseIf VarType(objValue) = vbObject Then
            If objValue Is Nothing Then
                Return True
            End If
        End If

        Return False

    End Function

End Module