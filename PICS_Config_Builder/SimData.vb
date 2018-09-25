Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Module SimData

    Sub Generate_Sim_Data()

        Dim ws As Worksheet = XLpicsWB.Sheets("IO Sheets")

        If ws.FilterMode Then ws.ShowAllData()

        Clear_Sheet_Type("SimData")
        Clear_Sheet_Type("IOTags")
        Clear_Sheet_Type("MinMax")

        Call Make_Sim_Tags("IO Sheets", "SimData")

        Call CheckMinMaxData("MinMax - AIn")

        Call Rem_Spaces("SimData", "E")
        Call Rem_Spaces("IOTags - AIn", "E")
        Call Rem_Spaces("IOTags - DIn", "E")
        Call Rem_Spaces("IOTags - ValveMO", "E")
        Call Rem_Spaces("IOTags - ValveSO", "E")
        Call Rem_Spaces("IOTags - ValveC", "E")
        Call Rem_Spaces("IOTags - Motor", "E")
        Call Rem_Spaces("IOTags - VSD", "E")

        Call SortByColumn("IOTags - ValveC", "E")

        ws = XLpicsWB.Sheets("IOTags - AIn")
        ws.Range("A2").Select()

        ws = XLpicsWB.Sheets("IOTags - DIn")
        ws.Range("A2").Select()

        ws = XLpicsWB.Sheets("SimData")
        ws.Range("A8").Select()

        ws = XLpicsWB.Sheets("Instructions")

    End Sub

    Sub Button_Hide_Sheets(ByRef wrkBook As Workbook)

        wrkBook.Application.ScreenUpdating = False

        Hide_Sheets(wrkBook)

        wrkBook.Application.ScreenUpdating = True

    End Sub

    Sub Button_Unhide_All_Sheets(ByRef wrkBook As Workbook)

        wrkBook.Application.ScreenUpdating = False

        Unhide_All_Sheets(wrkBook)

        wrkBook.Application.ScreenUpdating = True

    End Sub

    Sub ShowStatusBar(ByRef wrkBook As Workbook, Message As String)
        '
        '
        wrkBook.Application.StatusBar = Message
        wrkBook.Application.OnTime(Now() + TimeSerial(0, 0, 5), "hideStatusBar")

    End Sub
    Sub HideStatusBar(ByRef wrkBook As Workbook)
        '
        '
        wrkBook.Application.StatusBar = False

    End Sub

    Sub Make_Sim_Tags(ByRef sourceSheet As String, ByRef DataSheet As String)
        '
        Dim SimName, SimType, SimDefVal, SimIOAddr, SimDesc As String
        Dim Prefix, PLCBaseTag, DataType, IOVariable, IOAddress, IOType, DesignTag, Description As String
        Dim IOPrefix, MinMaxPrefix As String
        Dim InputMin, InputMax, OutputMin, OutputMax As Integer

        Prefix = ImportData.CPU_Name
        IOPrefix = "IOTags - "
        MinMaxPrefix = "MinMax - "

        'Source data is in sourceSheet, DataSheet is the destination
        Dim ws As Worksheet = XLpicsWB.Sheets(sourceSheet)
        Dim SourceRowCount As Integer = ws.Range("A").End(XlDirection.xlUp).Row
        Dim PLCBaseTag_Col As Integer = Find_Header_Column(sourceSheet, "PLCBaseTag")
        Dim DataType_Col As Integer = Find_Header_Column(sourceSheet, "Data Type")
        Dim IOVariable_Col As Integer = Find_Header_Column(sourceSheet, "Variable")
        Dim IOAddress_Col As Integer = Find_Header_Column(sourceSheet, "IOAddress")
        Dim IOType_Col As Integer = Find_Header_Column(sourceSheet, "IOType")
        Dim DesignTag_Col As Integer = Find_Header_Column(sourceSheet, "DesignTag")
        Dim Description_Col As Integer = Find_Header_Column(sourceSheet, "Description")
        Dim InputMin_Col As Integer = Find_Header_Column(sourceSheet, "InputMin")
        Dim InputMax_Col As Integer = Find_Header_Column(sourceSheet, "InputMax")
        Dim OutputMin_Col As Integer = Find_Header_Column(sourceSheet, "OutputMin")
        Dim OutputMax_Col As Integer = Find_Header_Column(sourceSheet, "OutputMax")

        For i = 2 To SourceRowCount

            PLCBaseTag = CType(ws.Range(i, PLCBaseTag_Col).Value, String)
            DataType = CType(ws.Range(i, DataType_Col).Value, String)
            IOVariable = CType(ws.Range(i, IOVariable_Col).Value, String)
            IOAddress = CType(ws.Range(i, IOAddress_Col).Value, String)
            IOType = CType(ws.Range(i, IOType_Col).Value, String)
            DesignTag = CType(ws.Range(i, DesignTag_Col).Value, String)
            Description = CType(ws.Range(i, Description_Col).Value, String)
            InputMin = CType(ws.Range(i, InputMin_Col).Value, Integer)
            InputMax = CType(ws.Range(i, InputMax_Col).Value, Integer)
            OutputMin = CType(ws.Range(i, OutputMin_Col).Value, Integer)
            OutputMax = CType(ws.Range(i, OutputMax_Col).Value, Integer)

            ' Since these are all the same in PICS functionally, make them all AIn
            DataType = Replace(DataType, "AInAdv", "AIn")
            DataType = Replace(DataType, "AInHART", "AIn")

            'Ignores spares, and types that have no use here
            If UCase(DesignTag) <> "SPARE" And
            UCase(DataType) <> "SPARE" And
            UCase(PLCBaseTag) <> "SPARE" And
            UCase(PLCBaseTag) <> "" Then

                ' Strip of the P_ or PC_ to get the 'base' type
                ' There should be a corresponding sheet
                '   - If not error and say so.
                Dim stripType As String
                Dim stripSheet As String
                Dim stripMinMax As String
                stripType = Mid(DataType, InStr(DataType, "_") + 1, Len(DataType))
                stripSheet = IOPrefix & stripType
                stripMinMax = MinMaxPrefix & stripType

                DesignTag = Replace(DesignTag, "-", "_")    'Change dashes to underscores
                IOVariable = Replace(IOVariable, ".", "_")  'Change dot to underscore

                If InStr(IOType, "DI") > 0 Then
                    'Paste First Row
                    SimName = IOVariable
                    SimType = "B R/W"
                    SimDefVal = "0"
                    SimIOAddr = "[" & Prefix & "_Sim]" & IOAddress
                    SimDesc = Description

                    ' Write data to IO tag sheet
                    ws = XLpicsWB.Sheets(DataSheet)
                    Dim RowCount As Integer = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row
                    ws.Range("A" & RowCount + 1).Select()
                    ws.Range("A" & RowCount + 1).Cells.Value = SimName
                    ws.Range("B" & RowCount + 1).Cells.Value = SimType
                    ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    ' Write data to IO tag sheet
                    ws = XLpicsWB.Sheets(stripSheet)
                    RowCount = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row
                    ws.Range("A" & RowCount + 1).Select()
                    ws.Range("A" & RowCount + 1).Cells.Value = SimName
                    ws.Range("B" & RowCount + 1).Cells.Value = SimType
                    ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    'Paste Second (Fault) Row
                    SimName = SimName & "_Flt"
                    SimIOAddr = Replace(SimIOAddr, "Data", "Fault")
                    SimDesc = Description & " CH_FLT"

                    ws = XLpicsWB.Sheets(DataSheet)
                    RowCount = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row
                    ws.Range("A" & RowCount + 1).Select()
                    ws.Range("A" & RowCount + 1).Cells.Value = SimName
                    ws.Range("B" & RowCount + 1).Cells.Value = SimType
                    ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    ' Write channel fault item to IO tag sheet
                    ws = XLpicsWB.Sheets(stripSheet)
                    RowCount = ws.Cells(ws.Rows.Count, "A").EndxlUp.Row
                    ws.Range("A" & RowCount + 1).Select()
                    ws.Range("A" & RowCount + 1).Cells.Value = SimName
                    ws.Range("B" & RowCount + 1).Cells.Value = SimType
                    ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                ElseIf InStr(IOType, "DO") > 0 Then
                    'Paste Row
                    SimName = IOVariable
                    SimType = "B R"
                    SimDefVal = ""
                    SimIOAddr = "[" & Prefix & "_Sim]" & IOAddress
                    SimDesc = Description

                    ws = XLpicsWB.Sheets(DataSheet)
                    Dim RowCount As Integer = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row
                    ws.Range("A" & RowCount + 1).Select()
                    ws.Range("A" & RowCount + 1).Cells.Value = SimName
                    ws.Range("B" & RowCount + 1).Cells.Value = SimType
                    ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    ws = XLpicsWB.Sheets(stripSheet)
                    RowCount = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row
                    ws.Range("A" & RowCount + 1).Select()
                    ws.Range("A" & RowCount + 1).Cells.Value = SimName
                    ws.Range("B" & RowCount + 1).Cells.Value = SimType
                    ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                ElseIf InStr(IOType, "AI") > 0 Then
                    'Paste First Row
                    SimName = IOVariable
                    SimType = "F R/W"
                    SimDefVal = "0"
                    SimIOAddr = "[" & Prefix & "_Sim]" & IOAddress
                    SimDesc = Description

                    ws = XLpicsWB.Sheets(DataSheet)
                    Dim RowCount As Integer = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row
                    ws.Range("A" & RowCount + 1).Select()
                    ws.Range("A" & RowCount + 1).Cells.Value = SimName
                    ws.Range("B" & RowCount + 1).Cells.Value = SimType
                    ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    ' Write data to IO tag sheets
                    ws = XLpicsWB.Sheets(stripSheet)
                    RowCount = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row
                    ws.Range("A" & RowCount + 1).Select()
                    ws.Range("A" & RowCount + 1).Cells.Value = SimName
                    ws.Range("B" & RowCount + 1).Cells.Value = SimType
                    ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    ws = XLpicsWB.Sheets(stripMinMax)
                    RowCount = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row
                    ws.Range("A" & RowCount + 1).Select()
                    ws.Range("A" & RowCount + 1).Cells.Value = SimName
                    ws.Range("B" & RowCount + 1).Cells.Value = InputMin
                    ws.Range("C" & RowCount + 1).Cells.Value = InputMax
                    ws.Range("D" & RowCount + 1).Cells.Value = OutputMin
                    ws.Range("E" & RowCount + 1).Cells.Value = OutputMax

                    'Paste Second (Fault) Row
                    SimName = SimName & "_Flt"
                    SimType = "B R/W"

                    ' Handle HART scenario
                    If InStr(IOType, "H") > 0 Then
                        SimIOAddr = Replace(SimIOAddr, ".Data", "Fault")
                    Else
                        SimIOAddr = Replace(SimIOAddr, "Data", "Fault")
                    End If

                    SimDesc = Description & " CH_FLT"

                    ws = XLpicsWB.Sheets(DataSheet)
                    RowCount = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row
                    ws.Range("A" & RowCount + 1).Select()
                    ws.Range("A" & RowCount + 1).Cells.Value = SimName
                    ws.Range("B" & RowCount + 1).Cells.Value = SimType
                    ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    ' Add faults to IO tag sheet
                    ws = XLpicsWB.Sheets(stripSheet)
                    RowCount = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row
                    ws.Range("A" & RowCount + 1).Select()
                    ws.Range("A" & RowCount + 1).Cells.Value = SimName
                    ws.Range("B" & RowCount + 1).Cells.Value = SimType
                    ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    ws = XLpicsWB.Sheets(stripMinMax)
                    RowCount = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row
                    ws.Range("A" & RowCount + 1).Select()
                    ws.Range("A" & RowCount + 1).Cells.Value = SimName
                    ws.Range("B" & RowCount + 1).Cells.Value = InputMin
                    ws.Range("C" & RowCount + 1).Cells.Value = InputMax
                    ws.Range("D" & RowCount + 1).Cells.Value = OutputMin
                    ws.Range("E" & RowCount + 1).Cells.Value = OutputMax

                ElseIf InStr(IOType, "AO") > 0 Then
                    'Paste Row
                    SimName = IOVariable
                    SimType = "F R"
                    SimDefVal = ""
                    SimIOAddr = "[" & Prefix & "_Sim]" & IOAddress
                    SimDesc = Description

                    ws = XLpicsWB.Sheets(DataSheet)
                    Dim RowCount As Integer = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row
                    ws.Range("A" & RowCount + 1).Select()
                    ws.Range("A" & RowCount + 1).Cells.Value = SimName
                    ws.Range("B" & RowCount + 1).Cells.Value = SimType
                    ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    ' Write data to IO tag sheet
                    ws = XLpicsWB.Sheets(stripSheet)
                    RowCount = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row
                    ws.Range("A" & RowCount + 1).Select()
                    ws.Range("A" & RowCount + 1).Cells.Value = SimName
                    ws.Range("B" & RowCount + 1).Cells.Value = SimType
                    ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    ws = XLpicsWB.Sheets(stripMinMax)
                    RowCount = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row
                    ws.Range("A" & RowCount + 1).Select()
                    ws.Range("A" & RowCount + 1).Cells.Value = SimName
                    ws.Range("B" & RowCount + 1).Cells.Value = InputMin
                    ws.Range("C" & RowCount + 1).Cells.Value = InputMax
                    ws.Range("D" & RowCount + 1).Cells.Value = OutputMin
                    ws.Range("E" & RowCount + 1).Cells.Value = OutputMax

                ElseIf InStr(IOType, "RTD") > 0 Then
                    'Paste First Row
                    SimName = IOVariable
                    SimType = "F R/W"
                    SimDefVal = "0"
                    SimIOAddr = "[" & Prefix & "_Sim]" & IOAddress
                    SimDesc = Description

                    ws = XLpicsWB.Sheets(DataSheet)
                    Dim RowCount As Integer = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row
                    ws.Range("A" & RowCount + 1).Select()
                    ws.Range("A" & RowCount + 1).Cells.Value = SimName
                    ws.Range("B" & RowCount + 1).Cells.Value = SimType
                    ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    'Write data to IO tag sheet
                    ws = XLpicsWB.Sheets(stripSheet)
                    RowCount = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row
                    ws.Range("A" & RowCount + 1).Select()
                    ws.Range("A" & RowCount + 1).Cells.Value = SimName
                    ws.Range("B" & RowCount + 1).Cells.Value = SimType
                    ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    ws = XLpicsWB.Sheets(stripMinMax)
                    RowCount = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row
                    ws.Range("A" & RowCount + 1).Select()
                    ws.Range("A" & RowCount + 1).Cells.Value = SimName
                    ws.Range("B" & RowCount + 1).Cells.Value = InputMin
                    ws.Range("C" & RowCount + 1).Cells.Value = InputMax
                    ws.Range("D" & RowCount + 1).Cells.Value = OutputMin
                    ws.Range("E" & RowCount + 1).Cells.Value = OutputMax

                    'Paste Second (Fault) Row
                    SimName = SimName & "_Flt"
                    SimType = "B R/W"
                    SimIOAddr = Replace(SimIOAddr, "Data", "Fault")
                    SimDesc = Description & " CH_FLT"

                    ws = XLpicsWB.Sheets(DataSheet)
                    RowCount = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row
                    ws.Range("A" & RowCount + 1).Select()
                    ws.Range("A" & RowCount + 1).Cells.Value = SimName
                    ws.Range("B" & RowCount + 1).Cells.Value = SimType
                    ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    ws = XLpicsWB.Sheets(stripSheet)
                    RowCount = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row
                    ws.Range("A" & RowCount + 1).Select()
                    ws.Range("A" & RowCount + 1).Cells.Value = SimName
                    ws.Range("B" & RowCount + 1).Cells.Value = SimType
                    ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    ws = XLpicsWB.Sheets(stripMinMax)
                    RowCount = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row
                    ws.Range("A" & RowCount + 1).Select()
                    ws.Range("A" & RowCount + 1).Cells.Value = SimName
                    ws.Range("B" & RowCount + 1).Cells.Value = InputMin
                    ws.Range("C" & RowCount + 1).Cells.Value = InputMax
                    ws.Range("D" & RowCount + 1).Cells.Value = OutputMin
                    ws.Range("E" & RowCount + 1).Cells.Value = OutputMax

                End If
            End If
        Next i

    End Sub
    Sub CheckMinMaxData(minMaxSheet As String)
        '
        '   Checks to make sure the Min Max data is numeric.
        Dim ws As Worksheet = XLpicsWB.Sheets("minMaxSheet").Select
        Dim RowCount As Integer = ws.Cells(ws.Rows.Count, "A").End.xlUp.Row

        For i = 2 To RowCount
            If Not IsNumeric(XLpicsWB.Sheets("minMaxSheet").Range("B" & i).Cells.Value) Then
                ws = XLpicsWB.Sheets("IO Sheets").Select
                ws.Columns("L").Select
                MsgBox("InputMin must be numeric values.")
                Exit For
            End If
        Next i

        For i = 2 To RowCount
            If Not IsNumeric(XLpicsWB.Sheets(minMaxSheet).Range("C" & i).Cells.Value) Then
                ws = XLpicsWB.Sheets("IO Sheets").Select
                ws.Columns("M").Select
                MsgBox("InputMax must be numeric values.")
                Exit For
            End If
        Next i

        For i = 2 To RowCount
            If Not IsNumeric(XLpicsWB.Sheets(minMaxSheet).Range("D" & i).Cells.Value) Then
                ws = XLpicsWB.Sheets("IO Sheets").Select
                ws.Columns("N").Select
                MsgBox("OutputMin must be numeric values.")
                Exit For
            End If
        Next i

        For i = 2 To RowCount
            If Not IsNumeric(XLpicsWB.Sheets(minMaxSheet).Range("E" & i).Cells.Value) Then
                ws = XLpicsWB.Sheets("IO Sheets").Select
                ws.Columns("O").Select
                MsgBox("OutputMax must be numeric values.")
                Exit For
            End If
        Next i

    End Sub

    Function Is_Cell_Blank(ByRef wrkBook As Workbook, DataSheet As String) As Boolean

        Dim ws As Worksheet = wrkBook.Sheets(DataSheet).Select
        If IsNothing(ws.Range("A8")) Then
            'MsgBox "Empty"
            Is_Cell_Blank = True
        ElseIf ws.Range("A8").Value = "" Then
            'MsgBox "Empty Text"
            If ws.Range("A8").HasFormula Then
                'MsgBox "Empty Text is the result of a formula"
            End If
            Is_Cell_Blank = False
        Else
            'MsgBox "Contains data"
            Is_Cell_Blank = False
        End If

    End Function
    Sub Rem_Spaces(ByRef destSheet As String, ByRef DestCol As String)
        '
        Dim ws As Worksheet = XLpicsWB.Sheets(destSheet).Select

        ws.Columns(DestCol).Replace(What:="  ",
                        Replacement:=" ",
                        LookAt:=XlLookAt.xlPart,
                        SearchOrder:=XlSearchOrder.xlByRows,
                        MatchCase:=False,
                        SearchFormat:=False,
                        ReplaceFormat:=False)
        ws.Range("A1").Select()

    End Sub

    Sub Remove_From_Desc(ByRef ws As Worksheet, ByRef x As Integer)

        Dim DelWord, OldDesc, NewDesc As String
        Dim RowCount As Integer = ws.Cells(ws.Rows.Count, "D").End(XlDirection.xlUp).Row
        DelWord = InputBox("Please enter the word you wish to delete:", "Delete From Descriptions")

        'Range("H1").FormulaR1C1 = DelWord
        If DelWord <> "" Then
            For i = 8 To RowCount
                ws.Range("D" & i).Select()
                OldDesc = ws.Range("D" & i).Value
                NewDesc = Replace(OldDesc, DelWord, "")
                ws.Range("D" & i).Value = NewDesc
            Next
        End If

        ws.Range("D8").Select()

    End Sub

    Sub SortByColumn(ByRef sheetName As String, ByRef SortCol As String)
        '
        '
        Dim ws As Worksheet = XLpicsWB.Sheets(sheetName).Select
        Dim RowCount As Integer = ws.Cells(ws.Rows.Count, "D").End(XlDirection.xlUp).Row

        ws.Range("A2:E" & RowCount).Sort(Key1:=ws.Range(SortCol & 2), Order1:=XlSortOrder.xlAscending)

    End Sub

End Module
