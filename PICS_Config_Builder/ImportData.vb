
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Module ImportData

    Public XLApp As New Application
    Public XLProjectWB, XLpicsWB As Workbook
    Public DirectoryName, CPU_Name As String      'user selected directory for PICS Config File
    Public picsBuilder, projectBuilder As String    'Workbook names
    Const fnXtnFilter = "Excel Files (*.xls;*.xlsx;*.xlsm),*.xls;*.xlsx;*xlsm"

    Sub Main()

        'Open Excel Application sub-process
        Call Open_Excel()

        Call OpenXLProjectFN()      ' Select and open the Excel project file (PLC IO Mapping) 

        If XLProjectWB Is Nothing Then ' Project file not selected
            Exit Sub
        End If

        Call OpenXLpicsFN()     ' Create or update the Excel PICS Config file that will organize data

        If XLpicsWB Is Nothing Then 'No PICS Config file selected
            Exit Sub
        End If

        Call Import_Data()

        Call Generate_Sim_Data()
        Call Generate_Memory_Data()
        Call Generate_Wire_Data()

        'Dim outFolder As String
        'outFolder = Create_Output_Folder(XLpicsWB)

        'Call Export_CSV(outFolder, "SimData", "OPC_Tags.csv")
        'Call Export_CSV(outFolder, "MemoryData", "GLOBAL_Tags.csv")
        'Call Export_Wire_Data(XLpicsWB, outFolder)

        If XLpicsWB.Name.Contains(".xlsm") Then     'Re-enable Excel application macros security settings prior to closing file
            XLpicsWB.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityLow
        End If

        XLpicsWB.Application.ScreenUpdating = True
        XLpicsWB.Application.DisplayAlerts = True 'Turn safety alerts back On
        XLpicsWB.Close()
        XLApp.Quit()

    End Sub

    Sub Import_Data()

        'Populate Excel PICS Config file with data from Project file
        Dim shtName As String = "IO Sheets"
        Dim shtFound As Boolean = False
        Dim ws As Worksheet
        Dim shtCount As Integer = XLpicsWB.Sheets.Count

        For Each ws In XLpicsWB.Sheets      ' does worksheet exist?
            If ws.Name.Equals(shtName) Then shtFound = True
        Next

        If Not shtFound Then        ' create worksheet
            If shtCount = 1 Then        ' rename sheet 1
                XLpicsWB.Sheets(shtCount).Name = shtName
                ws = XLpicsWB.Sheets(shtName)
                ws.Range("A2:AA9999").Clear()
            Else
                XLpicsWB.Worksheets.Add().Name = shtName
                ws = XLpicsWB.Sheets(shtName)
            End If
        Else
            ws = XLpicsWB.Sheets(shtName)

        End If

        picsBuilder = XLpicsWB.Name
        projectBuilder = XLProjectWB.Name

        CPU_Name = XLProjectWB.Sheets("Instructions").Range("C3").Value
        Dim values = XLProjectWB.Sheets("IO Sheets").UsedRange.Copy        ' copy from Project worksheet and paste to new PICS worksheet
        ws.Range("A1").PasteSpecial(XlPasteType.xlPasteValues)

        Do While ws.Range("A1").Value <> "PLCBaseTag"   ' Remove white space from top row
            ws.Range("A1").EntireRow.Delete()
        Loop

        ' Add CPU_PREFIX to form title
        If Form1.CPU_PREFIX.Text = "" Then
            Form1.CPU_PREFIX.Text = CPU_Name
        End If

        If XLProjectWB.Name.Contains(".xlsm") Then  'Re-enable Excel application macros security settings prior to closing Project file
            XLProjectWB.Application.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityLow
            XLProjectWB.Close(SaveChanges:=False)
        Else
            XLProjectWB.Close(SaveChanges:=False)

        End If

    End Sub

    Sub OpenXLProjectFN()

        'Select and open the Excel project file (PLC IO Mapping) that will be used to create PICS simulation files
        Dim sFileN As String
        Dim title = "Open - Select Project Config File"

        sFileN = XLApp.GetOpenFilename(FileFilter:=fnXtnFilter, FilterIndex:=2, Title:=title)
        DirectoryName = IO.Path.GetDirectoryName(sFileN)
        XLApp.DefaultFilePath = DirectoryName
        FileIO.FileSystem.CurrentDirectory = DirectoryName     ' set directory path

        If sFileN Is Nothing Then 'operator cancelled operation to open the project file
            XLApp.Quit()
        End If


        If sFileN.Contains(".xlsm") Then    'Disable Excel application macros security settings when opening file
            XLApp.Application.AutomationSecurity = Microsoft.Office.Core.MsoAutomationSecurity.msoAutomationSecurityForceDisable
        End If

        XLProjectWB = XLApp.Workbooks.Open(Filename:=sFileN, ReadOnly:=True)

    End Sub

    Sub OpenXLpicsFN()

        'Select or create PICS Excel file that will be used to create PICS simulation files
        Dim title = "Open - Select PICS Config File"
        Dim response As MsgBoxResult
        Dim fn, XLpicsFN As String

        response = MsgBox("Create A New PICS Config Builder file?", vbYesNo)

        If response = vbYes Then
            Do
                fn = InputBox("Enter New PICS Config Builder File Name:", "New File Name", "PICS_Config_Builder")
                If fn = "" Then     ' Cancel button pressed
                    response = MsgBox("Cancel Operation?", vbYesNo)
                    If response = vbYes Then Exit Sub
                End If
            Loop Until Not IsBlank(fn)
            XLpicsFN = DirectoryName & "\" & fn & ".xlsx"

        Else
            ' if response is no  - assume that you want to open an existing PICS file
            response = MsgBox("Open an existing PICS Config Builder file?", vbYesNo)
            If response = vbYes Then
                XLpicsFN = XLApp.GetOpenFilename(FileFilter:=fnXtnFilter, FilterIndex:=2, Title:=title)
            Else
                XLApp.Quit()  ' Since no PICS file was selected or created - close Excel and exit sub
                Exit Sub
            End If

        End If

        If IO.File.Exists(XLpicsFN) Then     'open existing workbook for PICS Config File
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

    Function IsBlank(ByVal Value)

        'returns True if Empty or NULL
        If VarType(Value) = vbEmpty Or VarType(Value) = vbNull Then
            Return True
        ElseIf VarType(Value) = vbString Then
            If Value = "" Then
                Return True
            End If
        ElseIf VarType(Value) = vbObject Then
            If Value Is Nothing Then
                Return True
            End If
        End If

        Return False

    End Function

End Module