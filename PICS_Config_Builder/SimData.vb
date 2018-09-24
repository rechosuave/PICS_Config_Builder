
Module SimData

    Sub Generate_Sim_Data(ByRef wrkBook As Workbook)

        Call Unhide_All_Sheets(wrkBook)

        'Unfilter IO Sheet if someone decided to filter it
        Dim wrkSheet As Worksheet = CType(wrkBook.Sheets("IO Sheets"), Worksheet)

        If wrkSheet.FilterMode Then wrkSheet.ShowAllData()

        Clear_Sheet_Type(wrkBook, "SimData")
        Clear_Sheet_Type(wrkBook, "IOTags")
        Clear_Sheet_Type(wrkBook, "MinMax")

        Call Make_Sim_Tags(wrkBook, "IO Sheets", "SimData")

        Call CheckMinMaxData(wrkBook, "MinMax - AIn")

        Call Rem_Spaces(wrkBook, "SimData", "E")
        Call Rem_Spaces(wrkBook, "IOTags - AIn", "E")
        Call Rem_Spaces(wrkBook, "IOTags - DIn", "E")
        Call Rem_Spaces(wrkBook, "IOTags - ValveMO", "E")
        Call Rem_Spaces(wrkBook, "IOTags - ValveSO", "E")
        Call Rem_Spaces(wrkBook, "IOTags - ValveC", "E")
        Call Rem_Spaces(wrkBook, "IOTags - Motor", "E")
        Call Rem_Spaces(wrkBook, "IOTags - VSD", "E")

        Call SortByColumn(wrkBook, "IOTags - ValveC", "E")

        wrkSheet = CType(wrkBook.Sheets("IOTags - AIn"), Worksheet)
        wrkSheet.Range("A2").Select()

        wrkSheet = CType(wrkBook.Sheets("IOTags - DIn"), Worksheet)
        wrkSheet.Range("A2").Select()

        wrkSheet = CType(wrkBook.Sheets("SimData"), Worksheet)
        wrkSheet.Range("A8").Select()

        wrkSheet = CType(wrkBook.Sheets("Instructions"), Worksheet)

        Call Hide_Sheets(wrkBook)

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
        '
        wrkBook.Application.StatusBar = Message
        wrkBook.Application.OnTime(Now() + TimeSerial(0, 0, 5), "hideStatusBar")

    End Sub
    Sub HideStatusBar(ByRef wrkBook As Workbook)
        '
        '
        wrkBook.Application.StatusBar = False

    End Sub

    Sub Make_Sim_Tags(ByRef wrkBook As Workbook, sourceSheet As String, DataSheet As String)
        '
        Dim SimName, SimType, SimDefVal, SimIOAddr, SimDesc As String
        Dim Prefix, PLCBaseTag, DataType, IOVariable, IOAddress, IOType, DesignTag, Description As String
        Dim IOPrefix, MinMaxPrefix As String
        Dim InputMin, InputMax, OutputMin, OutputMax As Integer

        Prefix = Get_CPU_Name(wrkBook)
        IOPrefix = "IOTags - "
        MinMaxPrefix = "MinMax - "

        'Source data is in sourceSheet, DataSheet is the destination
        Dim wrkSheet As Worksheet = CType(wrkBook.Sheets(sourceSheet), Worksheet)
        Dim SourceRowCount As Integer = wrkSheet.Range("A").End(Microsoft.Office.Interop.Excel.XlDirection.xlUp).Row

        Dim PLCBaseTag_Col As Integer = Find_Header_Column(wrkBook, sourceSheet, "PLCBaseTag")
        Dim DataType_Col As Integer = Find_Header_Column(wrkBook, sourceSheet, "Data Type")
        Dim IOVariable_Col As Integer = Find_Header_Column(wrkBook, sourceSheet, "Variable")
        Dim IOAddress_Col As Integer = Find_Header_Column(wrkBook, sourceSheet, "IOAddress")
        Dim IOType_Col As Integer = Find_Header_Column(wrkBook, sourceSheet, "IOType")
        Dim DesignTag_Col As Integer = Find_Header_Column(wrkBook, sourceSheet, "DesignTag")
        Dim Description_Col As Integer = Find_Header_Column(wrkBook, sourceSheet, "Description")
        Dim InputMin_Col As Integer = Find_Header_Column(wrkBook, sourceSheet, "InputMin")
        Dim InputMax_Col As Integer = Find_Header_Column(wrkBook, sourceSheet, "InputMax")
        Dim OutputMin_Col As Integer = Find_Header_Column(wrkBook, sourceSheet, "OutputMin")
        Dim OutputMax_Col As Integer = Find_Header_Column(wrkBook, sourceSheet, "OutputMax")

        For i = 2 To SourceRowCount

            PLCBaseTag = CType(wrkSheet.Range(i, PLCBaseTag_Col).Value, String)
            DataType = CType(wrkSheet.Range(i, DataType_Col).Value, String)
            IOVariable = CType(wrkSheet.Range(i, IOVariable_Col).Value, String)
            IOAddress = CType(wrkSheet.Range(i, IOAddress_Col).Value, String)
            IOType = CType(wrkSheet.Range(i, IOType_Col).Value, String)
            DesignTag = CType(wrkSheet.Range(i, DesignTag_Col).Value, String)
            Description = CType(wrkSheet.Range(i, Description_Col).Value, String)
            InputMin = CType(wrkSheet.Range(i, InputMin_Col).Value, Integer)
            InputMax = CType(wrkSheet.Range(i, InputMax_Col).Value, Integer)
            OutputMin = CType(wrkSheet.Range(i, OutputMin_Col).Value, Integer)
            OutputMax = CType(wrkSheet.Range(i, OutputMax_Col).Value, Integer)

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

                    wrkSheet = CType(wrkBook.Sheets(DataSheet), Worksheet)

                    Dim RowCount As Integer = wrkSheet.Cells(wrkSheet.Rows.Count, "A").End.xlUp.Row
                    wrkSheet.Range("A" & RowCount + 1).Select()
                    wrkSheet.Range("A" & RowCount + 1).Cells.Value = SimName
                    wrkSheet.Range("B" & RowCount + 1).Cells.Value = SimType
                    wrkSheet.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    wrkSheet.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    wrkSheet.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    ' Write data to IO tag sheet
                    wrkSheet = CType(wrkBook.Sheets(stripSheet), Worksheet)
                    RowCount = wrkSheet.Cells(wrkSheet.Rows.Count, "A").End.xlUp.Row
                    wrkSheet.Range("A" & RowCount + 1).Select()
                    wrkSheet.Range("A" & RowCount + 1).Cells.Value = SimName
                    wrkSheet.Range("B" & RowCount + 1).Cells.Value = SimType
                    wrkSheet.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    wrkSheet.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    wrkSheet.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    'Paste Second (Fault) Row
                    SimName = SimName & "_Flt"
                    SimIOAddr = Replace(SimIOAddr, "Data", "Fault")
                    SimDesc = Description & " CH_FLT"

                    wrkSheet = CType(wrkBook.Sheets(DataSheet), Worksheet)
                    RowCount = wrkSheet.Cells(wrkSheet.Rows.Count, "A").End.xlUp.Row
                    wrkSheet.Range("A" & RowCount + 1).Select()
                    wrkSheet.Range("A" & RowCount + 1).Cells.Value = SimName
                    wrkSheet.Range("B" & RowCount + 1).Cells.Value = SimType
                    wrkSheet.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    wrkSheet.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    wrkSheet.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    ' Write channel fault item to IO tag sheet
                    wrkSheet = CType(wrkBook.Sheets(stripSheet), Worksheet)
                    RowCount = wrkSheet.Cells(wrkSheet.Rows.Count, "A").EndxlUp.Row
                    wrkSheet.Range("A" & RowCount + 1).Select()
                    wrkSheet.Range("A" & RowCount + 1).Cells.Value = SimName
                    wrkSheet.Range("B" & RowCount + 1).Cells.Value = SimType
                    wrkSheet.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    wrkSheet.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    wrkSheet.Range("E" & RowCount + 1).Cells.Value = SimDesc

                ElseIf InStr(IOType, "DO") > 0 Then
                    'Paste Row
                    SimName = IOVariable
                    SimType = "B R"
                    SimDefVal = ""
                    SimIOAddr = "[" & Prefix & "_Sim]" & IOAddress
                    SimDesc = Description

                    wrkSheet = CType(wrkBook.Sheets(DataSheet), Worksheet)
                    Dim RowCount As Integer = wrkSheet.Cells(wrkSheet.Rows.Count, "A").End.xlUp.Row
                    wrkSheet.Range("A" & RowCount + 1).Select()
                    wrkSheet.Range("A" & RowCount + 1).Cells.Value = SimName
                    wrkSheet.Range("B" & RowCount + 1).Cells.Value = SimType
                    wrkSheet.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    wrkSheet.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    wrkSheet.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    wrkSheet = CType(wrkBook.Sheets(stripSheet), Worksheet)
                    RowCount = wrkSheet.Cells(wrkSheet.Rows.Count, "A").End.xlUp.Row
                    wrkSheet.Range("A" & RowCount + 1).Select()
                    wrkSheet.Range("A" & RowCount + 1).Cells.Value = SimName
                    wrkSheet.Range("B" & RowCount + 1).Cells.Value = SimType
                    wrkSheet.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    wrkSheet.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    wrkSheet.Range("E" & RowCount + 1).Cells.Value = SimDesc

                ElseIf InStr(IOType, "AI") > 0 Then
                    'Paste First Row
                    SimName = IOVariable
                    SimType = "F R/W"
                    SimDefVal = "0"
                    SimIOAddr = "[" & Prefix & "_Sim]" & IOAddress
                    SimDesc = Description

                    wrkSheet = CType(wrkBook.Sheets(DataSheet), Worksheet)
                    Dim RowCount As Integer = wrkSheet.Cells(wrkSheet.Rows.Count, "A").End.xlUp.Row
                    wrkSheet.Range("A" & RowCount + 1).Select()
                    wrkSheet.Range("A" & RowCount + 1).Cells.Value = SimName
                    wrkSheet.Range("B" & RowCount + 1).Cells.Value = SimType
                    wrkSheet.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    wrkSheet.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    wrkSheet.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    ' Write data to IO tag sheets
                    wrkSheet = CType(wrkBook.Sheets(stripSheet), Worksheet)
                    RowCount = wrkSheet.Cells(wrkSheet.Rows.Count, "A").End.xlUp.Row
                    wrkSheet.Range("A" & RowCount + 1).Select()
                    wrkSheet.Range("A" & RowCount + 1).Cells.Value = SimName
                    wrkSheet.Range("B" & RowCount + 1).Cells.Value = SimType
                    wrkSheet.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    wrkSheet.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    wrkSheet.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    wrkSheet = CType(wrkBook.Sheets(stripMinMax), Worksheet)
                    RowCount = wrkSheet.Cells(wrkSheet.Rows.Count, "A").End.xlUp.Row
                    wrkSheet.Range("A" & RowCount + 1).Select()
                    wrkSheet.Range("A" & RowCount + 1).Cells.Value = SimName
                    wrkSheet.Range("B" & RowCount + 1).Cells.Value = InputMin
                    wrkSheet.Range("C" & RowCount + 1).Cells.Value = InputMax
                    wrkSheet.Range("D" & RowCount + 1).Cells.Value = OutputMin
                    wrkSheet.Range("E" & RowCount + 1).Cells.Value = OutputMax

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

                    wrkSheet = CType(wrkBook.Sheets(DataSheet), Worksheet)
                    RowCount = wrkSheet.Cells(wrkSheet.Rows.Count, "A").End.xlUp.Row
                    wrkSheet.Range("A" & RowCount + 1).Select()
                    wrkSheet.Range("A" & RowCount + 1).Cells.Value = SimName
                    wrkSheet.Range("B" & RowCount + 1).Cells.Value = SimType
                    wrkSheet.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    wrkSheet.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    wrkSheet.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    ' Add faults to IO tag sheet
                    wrkSheet = CType(wrkBook.Sheets(stripSheet), Worksheet)
                    RowCount = wrkSheet.Cells(wrkSheet.Rows.Count, "A").End.xlUp.Row
                    wrkSheet.Range("A" & RowCount + 1).Select()
                    wrkSheet.Range("A" & RowCount + 1).Cells.Value = SimName
                    wrkSheet.Range("B" & RowCount + 1).Cells.Value = SimType
                    wrkSheet.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    wrkSheet.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    wrkSheet.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    wrkSheet = CType(wrkBook.Sheets(stripMinMax), Worksheet)
                    RowCount = wrkSheet.Cells(wrkSheet.Rows.Count, "A").End.xlUp.Row
                    wrkSheet.Range("A" & RowCount + 1).Select()
                    wrkSheet.Range("A" & RowCount + 1).Cells.Value = SimName
                    wrkSheet.Range("B" & RowCount + 1).Cells.Value = InputMin
                    wrkSheet.Range("C" & RowCount + 1).Cells.Value = InputMax
                    wrkSheet.Range("D" & RowCount + 1).Cells.Value = OutputMin
                    wrkSheet.Range("E" & RowCount + 1).Cells.Value = OutputMax

                ElseIf InStr(IOType, "AO") > 0 Then
                    'Paste Row
                    SimName = IOVariable
                    SimType = "F R"
                    SimDefVal = ""
                    SimIOAddr = "[" & Prefix & "_Sim]" & IOAddress
                    SimDesc = Description

                    wrkSheet = CType(wrkBook.Sheets(DataSheet), Worksheet)
                    Dim RowCount As Integer = wrkSheet.Cells(wrkSheet.Rows.Count, "A").End.xlUp.Row
                    wrkSheet.Range("A" & RowCount + 1).Select()
                    wrkSheet.Range("A" & RowCount + 1).Cells.Value = SimName
                    wrkSheet.Range("B" & RowCount + 1).Cells.Value = SimType
                    wrkSheet.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    wrkSheet.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    wrkSheet.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    ' Write data to IO tag sheet
                    wrkSheet = CType(wrkBook.Sheets(stripSheet), Worksheet)
                    RowCount = wrkSheet.Cells(wrkSheet.Rows.Count, "A").End.xlUp.Row
                    wrkSheet.Range("A" & RowCount + 1).Select()
                    wrkSheet.Range("A" & RowCount + 1).Cells.Value = SimName
                    wrkSheet.Range("B" & RowCount + 1).Cells.Value = SimType
                    wrkSheet.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    wrkSheet.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    wrkSheet.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    wrkSheet = CType(wrkBook.Sheets(stripMinMax), Worksheet)
                    RowCount = wrkSheet.Cells(wrkSheet.Rows.Count, "A").End.xlUp.Row
                    wrkSheet.Range("A" & RowCount + 1).Select()
                    wrkSheet.Range("A" & RowCount + 1).Cells.Value = SimName
                    wrkSheet.Range("B" & RowCount + 1).Cells.Value = InputMin
                    wrkSheet.Range("C" & RowCount + 1).Cells.Value = InputMax
                    wrkSheet.Range("D" & RowCount + 1).Cells.Value = OutputMin
                    wrkSheet.Range("E" & RowCount + 1).Cells.Value = OutputMax

                ElseIf InStr(IOType, "RTD") > 0 Then
                    'Paste First Row
                    SimName = IOVariable
                    SimType = "F R/W"
                    SimDefVal = "0"
                    SimIOAddr = "[" & Prefix & "_Sim]" & IOAddress
                    SimDesc = Description

                    wrkSheet = CType(wrkBook.Sheets(DataSheet), Worksheet)
                    Dim RowCount As Integer = wrkSheet.Cells(wrkSheet.Rows.Count, "A").End.xlUp.Row
                    wrkSheet.Range("A" & RowCount + 1).Select()
                    wrkSheet.Range("A" & RowCount + 1).Cells.Value = SimName
                    wrkSheet.Range("B" & RowCount + 1).Cells.Value = SimType
                    wrkSheet.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    wrkSheet.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    wrkSheet.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    'Write data to IO tag sheet
                    wrkSheet = CType(wrkBook.Sheets(stripSheet), Worksheet)
                    RowCount = wrkSheet.Cells(wrkSheet.Rows.Count, "A").End.xlUp.Row
                    wrkSheet.Range("A" & RowCount + 1).Select()
                    wrkSheet.Range("A" & RowCount + 1).Cells.Value = SimName
                    wrkSheet.Range("B" & RowCount + 1).Cells.Value = SimType
                    wrkSheet.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    wrkSheet.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    wrkSheet.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    wrkSheet = CType(wrkBook.Sheets(stripMinMax), Worksheet)
                    RowCount = wrkSheet.Cells(wrkSheet.Rows.Count, "A").End.xlUp.Row
                    wrkSheet.Range("A" & RowCount + 1).Select()
                    wrkSheet.Range("A" & RowCount + 1).Cells.Value = SimName
                    wrkSheet.Range("B" & RowCount + 1).Cells.Value = InputMin
                    wrkSheet.Range("C" & RowCount + 1).Cells.Value = InputMax
                    wrkSheet.Range("D" & RowCount + 1).Cells.Value = OutputMin
                    wrkSheet.Range("E" & RowCount + 1).Cells.Value = OutputMax

                    'Paste Second (Fault) Row
                    SimName = SimName & "_Flt"
                    SimType = "B R/W"
                    SimIOAddr = Replace(SimIOAddr, "Data", "Fault")
                    SimDesc = Description & " CH_FLT"

                    wrkSheet = CType(wrkBook.Sheets(DataSheet), Worksheet)
                    RowCount = wrkSheet.Cells(wrkSheet.Rows.Count, "A").End.xlUp.Row
                    wrkSheet.Range("A" & RowCount + 1).Select()
                    wrkSheet.Range("A" & RowCount + 1).Cells.Value = SimName
                    wrkSheet.Range("B" & RowCount + 1).Cells.Value = SimType
                    wrkSheet.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    wrkSheet.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    wrkSheet.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    wrkSheet = CType(wrkBook.Sheets(stripSheet), Worksheet)
                    RowCount = wrkSheet.Cells(wrkSheet.Rows.Count, "A").End.xlUp.Row
                    wrkSheet.Range("A" & RowCount + 1).Select()
                    wrkSheet.Range("A" & RowCount + 1).Cells.Value = SimName
                    wrkSheet.Range("B" & RowCount + 1).Cells.Value = SimType
                    wrkSheet.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                    wrkSheet.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                    wrkSheet.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    wrkSheet = CType(wrkBook.Sheets(stripMinMax), Worksheet)
                    RowCount = wrkSheet.Cells(wrkSheet.Rows.Count, "A").End.xlUp.Row
                    wrkSheet.Range("A" & RowCount + 1).Select()
                    wrkSheet.Range("A" & RowCount + 1).Cells.Value = SimName
                    wrkSheet.Range("B" & RowCount + 1).Cells.Value = InputMin
                    wrkSheet.Range("C" & RowCount + 1).Cells.Value = InputMax
                    wrkSheet.Range("D" & RowCount + 1).Cells.Value = OutputMin
                    wrkSheet.Range("E" & RowCount + 1).Cells.Value = OutputMax

                End If
            End If
        Next i

    End Sub
    Sub CheckMinMaxData(ByRef wrkBook As Workbook, minMaxSheet As String)
        '
        '   Checks to make sure the Min Max data is numeric.
        Dim wrkSheet As Worksheet = wrkBook.Sheets("minMaxSheet").Select
        Dim RowCount As Integer = wrkSheet.Cells(wrkSheet.Rows.Count, "A").End.xlUp.Row

        For i = 2 To RowCount
            If Not IsNumeric(wrkBook.Sheets("minMaxSheet").Range("B" & i).Cells.Value) Then
                wrkSheet = wrkBook.Sheets("IO Sheets").Select
                wrkSheet.Columns("L").Select
                MsgBox("InputMin must be numeric values.")
                Exit For
            End If
        Next i

        For i = 2 To RowCount
            If Not IsNumeric(wrkBook.Sheets(minMaxSheet).Range("C" & i).Cells.Value) Then
                wrkSheet = wrkBook.Sheets("IO Sheets").Select
                wrkSheet.Columns("M").Select
                MsgBox("InputMax must be numeric values.")
                Exit For
            End If
        Next i

        For i = 2 To RowCount
            If Not IsNumeric(wrkBook.Sheets(minMaxSheet).Range("D" & i).Cells.Value) Then
                wrkSheet = wrkBook.Sheets("IO Sheets").Select
                wrkSheet.Columns("N").Select
                MsgBox("OutputMin must be numeric values.")
                Exit For
            End If
        Next i

        For i = 2 To RowCount
            If Not IsNumeric(wrkBook.Sheets(minMaxSheet).Range("E" & i).Cells.Value) Then
                wrkSheet = wrkBook.Sheets("IO Sheets").Select
                wrkSheet.Columns("O").Select
                MsgBox("OutputMax must be numeric values.")
                Exit For
            End If
        Next i

    End Sub

    Function Is_Cell_Blank(ByRef wrkBook As Workbook, DataSheet As String) As Boolean

        Dim wrkSheet As Worksheet = wrkBook.Sheets(DataSheet).Select
        If IsNothing(wrkSheet.Range("A8")) Then
            'MsgBox "Empty"
            Is_Cell_Blank = True
        ElseIf wrkSheet.Range("A8").Value = "" Then
            'MsgBox "Empty Text"
            If wrkSheet.Range("A8").HasFormula Then
                'MsgBox "Empty Text is the result of a formula"
            End If
            Is_Cell_Blank = False
        Else
            'MsgBox "Contains data"
            Is_Cell_Blank = False
        End If

    End Function
    Sub Rem_Spaces(ByRef wrkBook As Workbook, destSheet As String, DestCol As String)
        '
        Dim wrkSheet As Worksheet = wrkBook.Sheets(destSheet).Select

        wrkSheet.Columns(DestCol).Replace(What:="  ",
                        Replacement:=" ",
                        LookAt:="xlPart",
                        SearchOrder:="xlByRows",
                        MatchCase:=False,
                        SearchFormat:=False,
                        ReplaceFormat:=False)
        wrkSheet.Range("A1").Select()

    End Sub

    Sub Remove_From_Desc(ByRef wrkSheet As Worksheet, ByRef x As Integer)

        Dim DelWord, OldDesc, NewDesc As String
        Dim RowCount As Integer = wrkSheet.Cells(wrkSheet.Rows.Count, "D").End.xlUp.Row
        DelWord = InputBox("Please enter the word you wish to delete:", "Delete From Descriptions")

        'Range("H1").FormulaR1C1 = DelWord
        If DelWord <> "" Then
            For i = 8 To RowCount
                wrkSheet.Range("D" & i).Select()
                OldDesc = wrkSheet.Range("D" & i).Value
                NewDesc = Replace(OldDesc, DelWord, "")
                wrkSheet.Range("D" & i).Value = NewDesc
            Next
        End If

        wrkSheet.Range("D8").Select()

    End Sub
    Sub SortByColumn(ByRef wrkBook As Workbook, sheetName As String, SortCol As String)
        '
        '
        Dim wrkSheet As Worksheet = wrkBook.Sheets(sheetName).Select
        Dim RowCount As Integer = wrkSheet.Cells(wrkSheet.Rows.Count, "D").End.xlUp.Row

        wrkSheet.Range("A2:E" & RowCount).Sort(Key1:=wrkSheet.Range(SortCol & 2), Order1:="xlAscending")

    End Sub

End Module
