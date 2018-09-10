Sub Button_Import_Data(ByRef x As Integer)
    Dim CopiedSheet As Worksheet
    Dim x As Integer
    Dim fName As String
    Dim picsBuilder As String
    Dim projectBuilder As String
    Dim cpuImport As String

    Application.ScreenUpdating = False

    Application.DisplayAlerts = False 'Turn safety alerts OFF

    Call Unhide_All_Sheets()

    Sheets("IO Sheets").Select
    Range("A2:AA9999").Clear

    picsBuilder = ActiveWorkbook.Name

    fName = Application.GetOpenFilename("Excel files(*.xls; *.xlsm), *.xls;*.xlsm", 2, "Select a Project Config file.")
    If fName = "" Or fName = "False" Then Exit Sub

    Workbooks.Open(fName)
    projectBuilder = ActiveWorkbook.Name

    cpuImport = Sheets("Instructions").Range("C3").Value
    Sheets("IO Sheets").UsedRange.Copy

    ' Paste entire IO sheet
    Workbooks(picsBuilder).Activate
    Sheets("IO Sheets").Select
    Range("A1").PasteSpecial(Paste:=xlValues)

    ' Remove any white space at the top
    Do While Range("A1") <> "PLCBaseTag"
        Range("A1").EntireRow.Delete
    Loop

    ' Fix all selections to look nice
    If Sheets("Instructions").Range("CPU_PREFIX").Value = "" Then
        Sheets("Instructions").Range("CPU_PREFIX").Value = cpuImport
    End If

    Reset_Sheet("Instructions")
    Reset_Sheet("IO Sheets")
    Sheets("Instructions").Select

    Call Hide_Sheets()

    Workbooks(projectBuilder).Close(SaveChanges:=False)
    Application.DisplayAlerts = True 'Turn safety alerts ON

    Application.ScreenUpdating = True

End Sub

Sub Button_Data_And_Run(ByRef x As Integer)

    Call Button_Import_Data()

    Application.ScreenUpdating = False

    Call Generate_Sim_Data()
    Call Generate_Memory_Data()
    Call Generate_Wire_Data()

    Dim outFolder As String
    outFolder = Create_Output_Folder()

    Call Export_CSV(outFolder, "SimData", "OPC_Tags.csv")
    Call Export_CSV(outFolder, "MemoryData", "GLOBAL_Tags.csv")
    Call Export_Wire_Data(outFolder)

    Application.ScreenUpdating = False

End Sub

Sub Button_Clear_All_Sheets(ByRef x As Integer)
    '
    '
    '
    'WARNING! This will clear all the data delete all Wire sheets
    If MsgBox("WARNING! This will clear all the data from this workbook and delete existing Wire data sheets.", vbOKCancel) = vbCancel Then Exit Sub

    Application.ScreenUpdating = False

    Call Button_Unhide_All_Sheets()

    Clear_All_Sheets

    Call Delete_Wire_Sheets("Wire_AIn Template")
    Call Delete_Wire_Sheets("Wire_DIn Template")
    Call Delete_Wire_Sheets("Wire_ValveC Template")
    Call Delete_Wire_Sheets("Wire_ValveMO Template")
    Call Delete_Wire_Sheets("Wire_ValveSO Template")
    Call Delete_Wire_Sheets("Wire_Motor Template")
    Call Delete_Wire_Sheets("Wire_VSD Template")

    Call Button_Hide_Sheets()
    Sheets("Instructions").Range("CPU_PREFIX").ClearContents

    Application.ScreenUpdating = True

End Sub

