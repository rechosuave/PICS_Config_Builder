Imports Microsoft.Office.Interop.Excel

Module SimData

    Sub Generate_Sim_Data()

        Call Make_Sim_Tags("IO Sheets", "SimData")

        Call CheckMinMaxData("MinMax - AIn")

        Call Remove_Spaces("SimData", "E")
        Call Remove_Spaces("IOTags - AIn", "E")
        Call Remove_Spaces("IOTags - DIn", "E")
        Call Remove_Spaces("IOTags - ValveMO", "E")
        Call Remove_Spaces("IOTags - ValveSO", "E")
        Call Remove_Spaces("IOTags - ValveC", "E")
        Call Remove_Spaces("IOTags - Motor", "E")
        Call Remove_Spaces("IOTags - VSD", "E")

        SortByColumn("IOTags - ValveC", "E")

    End Sub

    Sub Make_Sim_Tags(ByVal sourceSheet As String, ByVal DataSheet As String)
        '
        Dim SimName, SimType, SimDefVal, SimIOAddr, SimDesc As String
        Dim Prefix, PLCBaseTag, DataType, IOVariable, IOAddress, IOType, DesignTag, Description As String
        Dim IOPrefix, MinMaxPrefix As String
        Dim RowCount, InputMin, InputMax, OutputMin, OutputMax As Integer

        Prefix = ImportData.CPU_Name
        IOPrefix = "IOTags - "
        MinMaxPrefix = "MinMax - "

        'Source data is in sourceSheet, DataSheet is the destination
        Dim ws As Worksheet = XLpicsWB.Sheets(sourceSheet)
        Dim SourceRowCount As Integer = ws.Cells(ws.Cells.Rows.Count, "A").End(XlDirection.xlUp).Row
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

            ws = XLpicsWB.Sheets(sourceSheet)       ' select next row of data to copy
            PLCBaseTag = ws.Cells(i, PLCBaseTag_Col).Value
            DataType = ws.Cells(i, DataType_Col).Value
            IOVariable = ws.Cells(i, IOVariable_Col).Value
            IOAddress = ws.Cells(i, IOAddress_Col).Value
            IOType = ws.Cells(i, IOType_Col).Value
            DesignTag = ws.Cells(i, DesignTag_Col).Value
            Description = ws.Cells(i, Description_Col).Value
            InputMin = ws.Cells(i, InputMin_Col).Value
            InputMax = ws.Cells(i, InputMax_Col).Value
            OutputMin = ws.Cells(i, OutputMin_Col).Value
            OutputMax = ws.Cells(i, OutputMax_Col).Value

            ' Since these are all the same in PICS functionally, make them all AIn
            DataType = Replace(DataType, "AInAdv", "AIn")
            DataType = Replace(DataType, "AInHART", "AIn")

            'Ignores spares, and types that have no use here
            If UCase(DesignTag) <> "SPARE" And UCase(DataType) <> "SPARE" And
            UCase(PLCBaseTag) <> "SPARE" And UCase(PLCBaseTag) <> "" Then

                ' Strip out "P_" or "PC_" to get the 'base' type
                ' There should be a corresponding sheet - If not error and say so.
                Dim stripType As String
                Dim stripSheet As String
                Dim stripMinMax As String
                stripType = Mid(DataType, InStr(DataType, "_") + 1, Len(DataType))
                stripSheet = IOPrefix & stripType
                stripMinMax = MinMaxPrefix & stripType

                DesignTag = Replace(DesignTag, "-", "_")    'Change dashes to underscores
                IOVariable = Replace(IOVariable, ".", "_")  'Change dot to underscore

                Select Case IOType          ' 5 IO types - AI, AO, DI, DO, RTD
                    Case "AI"                           'Paste First Row
                        SimName = IOVariable
                        SimType = "F R/W"
                        SimDefVal = "0"
                        SimIOAddr = "[" & Prefix & "_Sim]" & IOAddress
                        SimDesc = Description

                        ws = XLpicsWB.Sheets(DataSheet)
                        RowCount = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row
                        ws.Range("A" & RowCount + 1).Cells.Value = SimName
                        ws.Range("B" & RowCount + 1).Cells.Value = SimType
                        ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                        ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                        ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                        ws = SelectWS(stripSheet)
                        RowCount = ws.Cells(ws.Cells.Rows.Count, "A").End(XlDirection.xlUp).Row
                        ws.Range("A" & RowCount + 1).Cells.Value = SimName
                        ws.Range("B" & RowCount + 1).Cells.Value = SimType
                        ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                        ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                        ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                        ws = XLpicsWB.Sheets(stripMinMax)
                        RowCount = ws.Cells(ws.Cells.Rows.Count, "A").End(XlDirection.xlUp).Row
                        ws.Range("A" & RowCount + 1).Cells.Value = SimName
                        ws.Range("B" & RowCount + 1).Cells.Value = InputMin
                        ws.Range("C" & RowCount + 1).Cells.Value = InputMax
                        ws.Range("D" & RowCount + 1).Cells.Value = OutputMin
                        ws.Range("E" & RowCount + 1).Cells.Value = OutputMax

                        SimName = SimName & "_Flt"      'Paste Second (Fault) Row
                        SimType = "B R/W"

                        If InStr(IOType, "H") > 0 Then      ' Handle HART scenario
                            SimIOAddr = Replace(SimIOAddr, ".Data", "Fault")
                        Else
                            SimIOAddr = Replace(SimIOAddr, "Data", "Fault")
                        End If

                        SimDesc = Description & " CH_FLT"

                        ws = XLpicsWB.Sheets(DataSheet)
                        RowCount = ws.Cells(ws.Cells.Rows.Count, "A").End(XlDirection.xlUp).Row
                        ws.Range("A" & RowCount + 1).Cells.Value = SimName
                        ws.Range("B" & RowCount + 1).Cells.Value = SimType
                        ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                        ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                        ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                        ws = XLpicsWB.Sheets(stripSheet)        ' Add faults to IO tag sheet
                        RowCount = ws.Cells(ws.Cells.Rows.Count, "A").End(XlDirection.xlUp).Row
                        ws.Range("A" & RowCount + 1).Cells.Value = SimName
                        ws.Range("B" & RowCount + 1).Cells.Value = SimType
                        ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                        ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                        ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                        ws = XLpicsWB.Sheets(stripMinMax)
                        RowCount = ws.Cells(ws.Cells.Rows.Count, "A").End(XlDirection.xlUp).Row
                        ws.Range("A" & RowCount + 1).Cells.Value = SimName
                        ws.Range("B" & RowCount + 1).Cells.Value = InputMin
                        ws.Range("C" & RowCount + 1).Cells.Value = InputMax
                        ws.Range("D" & RowCount + 1).Cells.Value = OutputMin
                        ws.Range("E" & RowCount + 1).Cells.Value = OutputMax

                    Case "AO"                           'Paste Row
                        SimName = IOVariable
                        SimType = "F R"
                        SimDefVal = ""
                        SimIOAddr = "[" & Prefix & "_Sim]" & IOAddress
                        SimDesc = Description

                        ws = XLpicsWB.Sheets(DataSheet)
                        RowCount = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row
                        ws.Range("A" & RowCount + 1).Cells.Value = SimName
                        ws.Range("B" & RowCount + 1).Cells.Value = SimType
                        ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                        ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                        ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                        ws = XLpicsWB.Sheets(stripSheet)
                        RowCount = ws.Cells(ws.Cells.Rows.Count, "A").End(XlDirection.xlUp).Row
                        ws.Range("A" & RowCount + 1).Cells.Value = SimName
                        ws.Range("B" & RowCount + 1).Cells.Value = SimType
                        ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                        ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                        ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                        ws = XLpicsWB.Sheets(stripMinMax)
                        RowCount = ws.Cells(ws.Cells.Rows.Count, "A").End(XlDirection.xlUp).Row
                        ws.Range("A" & RowCount + 1).Cells.Value = SimName
                        ws.Range("B" & RowCount + 1).Cells.Value = InputMin
                        ws.Range("C" & RowCount + 1).Cells.Value = InputMax
                        ws.Range("D" & RowCount + 1).Cells.Value = OutputMin
                        ws.Range("E" & RowCount + 1).Cells.Value = OutputMax

                    Case "DI"                           'Paste First Row
                        SimName = IOVariable
                        SimType = "B R/W"
                        SimDefVal = "0"
                        SimIOAddr = "[" & Prefix & "_Sim]" & IOAddress
                        SimDesc = Description

                        ws = XLpicsWB.Sheets(DataSheet)
                        RowCount = ws.Cells(ws.Cells.Rows.Count, "A").End(XlDirection.xlUp).Row
                        ws.Range("A" & RowCount + 1).Cells.Value = SimName
                        ws.Range("B" & RowCount + 1).Cells.Value = SimType
                        ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                        ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                        ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                        ws = XLpicsWB.Sheets(stripSheet)        ' Write data to IO tag sheet
                        RowCount = ws.Cells(ws.Cells.Rows.Count, "A").End(XlDirection.xlUp).Row
                        ws.Range("A" & RowCount + 1).Cells.Value = SimName
                        ws.Range("B" & RowCount + 1).Cells.Value = SimType
                        ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                        ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                        ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                        SimName = SimName & "_Flt"      'Paste Second (Fault) Row
                        SimIOAddr = Replace(SimIOAddr, "Data", "Fault")
                        SimDesc = Description & " CH_FLT"

                        ws = XLpicsWB.Sheets(DataSheet)
                        RowCount = ws.Cells(ws.Cells.Rows.Count, "A").End(XlDirection.xlUp).Row
                        ws.Range("A" & RowCount + 1).Cells.Value = SimName
                        ws.Range("B" & RowCount + 1).Cells.Value = SimType
                        ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                        ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                        ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                        ' Write channel fault item to IO tag sheet
                        ws = XLpicsWB.Sheets(stripSheet)
                        RowCount = ws.Cells(ws.Cells.Rows.Count, "A").End(XlDirection.xlUp).Row
                        ws.Range("A" & RowCount + 1).Cells.Value = SimName
                        ws.Range("B" & RowCount + 1).Cells.Value = SimType
                        ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                        ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                        ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                    Case "DO"                           'Paste Row
                        SimName = IOVariable
                        SimType = "B R"
                        SimDefVal = ""
                        SimIOAddr = "[" & Prefix & "_Sim]" & IOAddress
                        SimDesc = Description

                        ws = XLpicsWB.Sheets(DataSheet)
                        RowCount = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row
                        ws.Range("A" & RowCount + 1).Cells.Value = SimName
                        ws.Range("B" & RowCount + 1).Cells.Value = SimType
                        ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                        ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                        ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                        ws = XLpicsWB.Sheets(stripSheet)
                        RowCount = ws.Cells(ws.Cells.Rows.Count, "A").End(XlDirection.xlUp).Row
                        ws.Range("A" & RowCount + 1).Cells.Value = SimName
                        ws.Range("B" & RowCount + 1).Cells.Value = SimType
                        ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                        ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                        ws.Range("E" & RowCount + 1).Cells.Value = SimDesc


                    Case "RTD"                       'Paste First Row
                        SimName = IOVariable
                        SimType = "F R/W"
                        SimDefVal = "0"
                        SimIOAddr = "[" & Prefix & "_Sim]" & IOAddress
                        SimDesc = Description

                        ws = XLpicsWB.Sheets(DataSheet)
                        RowCount = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row
                        ws.Range("A" & RowCount + 1).Cells.Value = SimName
                        ws.Range("B" & RowCount + 1).Cells.Value = SimType
                        ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                        ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                        ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                        ws = XLpicsWB.Sheets(stripSheet)        'Write data to IO tag sheet
                        RowCount = ws.Cells(ws.Cells.Rows.Count, "A").End(XlDirection.xlUp).Row
                        ws.Range("A" & RowCount + 1).Cells.Value = SimName
                        ws.Range("B" & RowCount + 1).Cells.Value = SimType
                        ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                        ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                        ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                        ws = XLpicsWB.Sheets(stripMinMax)
                        RowCount = ws.Cells(ws.Cells.Rows.Count, "A").End(XlDirection.xlUp).Row
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
                        RowCount = ws.Cells(ws.Cells.Rows.Count, "A").End(XlDirection.xlUp).Row
                        ws.Range("A" & RowCount + 1).Cells.Value = SimName
                        ws.Range("B" & RowCount + 1).Cells.Value = SimType
                        ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                        ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                        ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                        ws = XLpicsWB.Sheets(stripSheet)
                        RowCount = ws.Cells(ws.Cells.Rows.Count, "A").End(XlDirection.xlUp).Row
                        ws.Range("A" & RowCount + 1).Cells.Value = SimName
                        ws.Range("B" & RowCount + 1).Cells.Value = SimType
                        ws.Range("C" & RowCount + 1).Cells.Value = SimDefVal
                        ws.Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                        ws.Range("E" & RowCount + 1).Cells.Value = SimDesc

                        ws = XLpicsWB.Sheets(stripMinMax)
                        RowCount = ws.Cells(ws.Cells.Rows.Count, "A").End(XlDirection.xlUp).Row
                        ws.Range("A" & RowCount + 1).Cells.Value = SimName
                        ws.Range("B" & RowCount + 1).Cells.Value = InputMin
                        ws.Range("C" & RowCount + 1).Cells.Value = InputMax
                        ws.Range("D" & RowCount + 1).Cells.Value = OutputMin
                        ws.Range("E" & RowCount + 1).Cells.Value = OutputMax
                End Select

            End If

        Next

    End Sub

    Sub CheckMinMaxData(ByVal minMaxSheet As String)

        '   Checks to make sure the Min Max data is numeric.
        Dim ws As Worksheet
        Dim sourcesheet As String = "IO Sheets"
        Dim wsSource As Worksheet = XLpicsWB.Sheets(sourcesheet)
        Dim wsDestination As Worksheet = XLpicsWB.Sheets(minMaxSheet)
        Dim RowCount As Integer = wsDestination.Cells(wsDestination.Rows.Count, "A").End(XlDirection.xlUp).Row

        For i = 2 To RowCount
            If Not IsNumeric(wsDestination.Range("B" & i).Cells.Value) Then
                ws = wsSource
                ws.Columns("L").Select
                MsgBox(minMaxSheet & " InputMin must be numeric values.")
                Exit For
            End If
        Next i

        For i = 2 To RowCount
            If Not IsNumeric(wsDestination.Range("C" & i).Cells.Value) Then
                ws = wsSource
                ws.Columns("M").Select
                MsgBox(minMaxSheet & " InputMax must be numeric values.")
                Exit For
            End If
        Next i

        For i = 2 To RowCount
            If Not IsNumeric(wsDestination.Range("D" & i).Cells.Value) Then
                ws = wsSource
                ws.Columns("N").Select
                MsgBox(minMaxSheet & " OutputMin must be numeric values.")
                Exit For
            End If
        Next i

        For i = 2 To RowCount
            If Not IsNumeric(wsDestination.Range("E" & i).Cells.Value) Then
                ws = wsSource
                ws.Columns("O").Select
                MsgBox(minMaxSheet & " OutputMax must be numeric values.")
                Exit For
            End If
        Next i

    End Sub

    Sub Remove_Spaces(ByVal destSheet As String, ByVal DestCol As String)
        '
        Dim ws As Worksheet = XLpicsWB.Sheets(destSheet)
        Dim RowCount As Integer = ws.Cells(ws.Rows.Count, DestCol).End(XlDirection.xlUp).Row
        Dim rng As Range = ws.Range(DestCol & "1", ws.Range(DestCol & "1").End(XlDirection.xlDown))

        ' following commented statement no longer works in Excel for VB
        ' ws.Columns(DestCol).Replace(What:="  ", Replacement:=" ", LookAt:=XlLookAt.xlPart, SearchOrder:=XlSearchOrder.xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False)

        For i = 1 To RowCount
            rng(i).Value = Replace(rng(i).Value, "  ", " ")
        Next i

    End Sub

    Sub SortByColumn(ByVal shtName As String, ByVal SortCol As String)
        '
        Dim ws As Worksheet = XLpicsWB.Sheets(shtName)
        Dim RowCount As Integer = ws.Cells(ws.Rows.Count, "D").End(XlDirection.xlUp).Row

        ws.Range("A2:E" & RowCount).Sort(Key1:=ws.Range(SortCol & 2), Order1:=XlSortOrder.xlAscending)

    End Sub

    Function Is_Cell_Blank(ByRef wrkBook As Workbook, DataSheet As String) As Boolean

        Dim ws As Worksheet = wrkBook.Sheets(DataSheet).Select
        If Information.IsNothing(ws.Range("A8")) Then
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

    Function SelectWS(ByVal sheet As String) As Worksheet
        ' 
        Dim shtFound As Boolean = False
        Dim ws As Worksheet

        For Each ws In XLpicsWB.Sheets      ' does worksheet exist?
            If ws.Name.Equals(sheet) Then shtFound = True
        Next

        If shtFound Then        ' was worksheet found?
            Return XLpicsWB.Sheets(sheet)
        Else
            XLpicsWB.Worksheets.Add().Name = sheet
            Return XLpicsWB.Sheets(sheet)
        End If

    End Function

    Function WS_Exists(ByVal sheet As String) As Boolean

        Dim shtFound As Boolean = False
        Dim ws As Worksheet

        For Each ws In XLpicsWB.Sheets      ' does worksheet exist?
            If ws.Name.Equals(sheet) Then shtFound = True
        Next

        Return shtFound

    End Function

End Module
