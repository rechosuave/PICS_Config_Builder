Sub Generate_Sim_Data(ByRef x As Integer)

    Call Unhide_All_Sheets

    'Unfilter IO Sheet if someone decided to filter it
    Sheets("IO Sheets").Select
    If ActiveSheet.FilterMode Then ActiveSheet.ShowAllData

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

    Sheets("IOTags - AIn").Select
    Range("A2").Select
    Sheets("IOTags - DIn").Select
    Range("A2").Select
    Sheets("SimData").Select
    Range("A8").Select

    Sheets("Instructions").Select

    Call Hide_Sheets

End Sub

Sub Button_Hide_Sheets(ByRef x As Integer)

    Application.ScreenUpdating = False

    Hide_Sheets

    Application.ScreenUpdating = True

End Sub

Sub Button_Unhide_All_Sheets(ByRef x As Integer)

    Application.ScreenUpdating = False

    Unhide_All_Sheets

    Application.ScreenUpdating = True

End Sub

Sub showStatusBar(Message As String)
'
'
'
    Application.StatusBar = Message
    Application.OnTime Now() + TimeSerial(0, 0, 5), "hideStatusBar"
        
End Sub
Sub hideStatusBar()
'
'
'
    Application.StatusBar = False
    
End Sub

Sub Make_Sim_Tags(sourceSheet As String, DataSheet As String)
'
'
'
    Dim SimName As String
    Dim SimType As String
    Dim SimDefVal As String
    Dim SimIOAddr As String
    Dim SimDesc As String
    
    Dim Prefix As String
    Dim PLCBaseTag As String
    Dim DataType As String
    Dim IOVariable As String
    Dim IOAddress As String
    Dim IOType As String
    Dim DesignTag As String
    Dim Description As String
    Dim Rack As String
    Dim Module As String
    Dim Channel As String
    
    Dim IOPrefix
    Dim AInSheet As String
    Dim ValveCSheet As String
    Dim DInSheet As String
    Dim ValveMOSheet As String
    Dim ValveSOSheet As String
    Dim MotorSheet As String
    Dim VSDSheet As String
    
    Prefix = Get_CPU_Name()
    IOPrefix = "IOTags - "
    MinMaxPrefix = "MinMax - "

    'Souce data is in SourceSheet, DataSheet is the destination
    Sheets(sourceSheet).Select
    SourceRowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
    
    PLCBaseTag_Col = Find_Header_Column(sourceSheet, "PLCBaseTag")
    DataType_Col = Find_Header_Column(sourceSheet, "Data Type")
    IOVariable_Col = Find_Header_Column(sourceSheet, "Variable")
    IOAddress_Col = Find_Header_Column(sourceSheet, "IOAddress")
    IOType_Col = Find_Header_Column(sourceSheet, "IOType")
    DesignTag_Col = Find_Header_Column(sourceSheet, "DesignTag")
    Description_Col = Find_Header_Column(sourceSheet, "Description")
    InputMin_Col = Find_Header_Column(sourceSheet, "InputMin")
    InputMax_Col = Find_Header_Column(sourceSheet, "InputMax")
    OutputMin_Col = Find_Header_Column(sourceSheet, "OutputMin")
    OutputMax_Col = Find_Header_Column(sourceSheet, "OutputMax")
    
    For i = 2 To SourceRowCount
        
        PLCBaseTag = Worksheets(sourceSheet).Cells(i, PLCBaseTag_Col).Value
        DataType = Worksheets(sourceSheet).Cells(i, DataType_Col).Value
        IOVariable = Worksheets(sourceSheet).Cells(i, IOVariable_Col).Value
        IOAddress = Worksheets(sourceSheet).Cells(i, IOAddress_Col).Value
        IOType = Worksheets(sourceSheet).Cells(i, IOType_Col).Value
        DesignTag = Worksheets(sourceSheet).Cells(i, DesignTag_Col).Value
        Description = Worksheets(sourceSheet).Cells(i, Description_Col).Value
        InputMin = Worksheets(sourceSheet).Cells(i, InputMin_Col).Value
        InputMax = Worksheets(sourceSheet).Cells(i, InputMax_Col).Value
        OutputMin = Worksheets(sourceSheet).Cells(i, OutputMin_Col).Value
        OutputMax = Worksheets(sourceSheet).Cells(i, OutputMax_Col).Value
        
        ' Since these are all the same in PICS functionally, make them all AIn
        DataType = Replace(DataType, "AInAdv", "AIn")
        DataType = Replace(DataType, "AInHART", "AIn")
        
        'Ignores spares, and types that have no use here
        If UCase(DesignTag) <> "SPARE" And _
            UCase(DataType) <> "SPARE" And _
            UCase(PLCBaseTag) <> "SPARE" And _
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

                Sheets(DataSheet).Select
                RowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
                Range("A" & RowCount + 1).Select
                Range("A" & RowCount + 1).Cells.Value = SimName
                Range("B" & RowCount + 1).Cells.Value = SimType
                Range("C" & RowCount + 1).Cells.Value = SimDefVal
                Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                Range("E" & RowCount + 1).Cells.Value = SimDesc
                
                ' Write data to IO tag sheet
                Sheets(stripSheet).Select
                RowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
                Range("A" & RowCount + 1).Select
                Range("A" & RowCount + 1).Cells.Value = SimName
                Range("B" & RowCount + 1).Cells.Value = SimType
                Range("C" & RowCount + 1).Cells.Value = SimDefVal
                Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                Range("E" & RowCount + 1).Cells.Value = SimDesc
                                                                                   
                'Paste Second (Fault) Row
                SimName = SimName & "_Flt"
                SimIOAddr = Replace(SimIOAddr, "Data", "Fault")
                SimDesc = Description & " CH_FLT"

                Sheets(DataSheet).Select
                RowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
                Range("A" & RowCount + 1).Select
                Range("A" & RowCount + 1).Cells.Value = SimName
                Range("B" & RowCount + 1).Cells.Value = SimType
                Range("C" & RowCount + 1).Cells.Value = SimDefVal
                Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                Range("E" & RowCount + 1).Cells.Value = SimDesc
                
                ' Write channel fault item to IO tag sheet
                Sheets(stripSheet).Select
                RowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
                Range("A" & RowCount + 1).Select
                Range("A" & RowCount + 1).Cells.Value = SimName
                Range("B" & RowCount + 1).Cells.Value = SimType
                Range("C" & RowCount + 1).Cells.Value = SimDefVal
                Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                Range("E" & RowCount + 1).Cells.Value = SimDesc
                
            ElseIf InStr(IOType, "DO") > 0 Then
                'Paste Row
                SimName = IOVariable
                SimType = "B R"
                SimDefVal = ""
                SimIOAddr = "[" & Prefix & "_Sim]" & IOAddress
                SimDesc = Description

                Sheets(DataSheet).Select
                RowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
                Range("A" & RowCount + 1).Select
                Range("A" & RowCount + 1).Cells.Value = SimName
                Range("B" & RowCount + 1).Cells.Value = SimType
                Range("C" & RowCount + 1).Cells.Value = SimDefVal
                Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                Range("E" & RowCount + 1).Cells.Value = SimDesc

                Sheets(stripSheet).Select
                RowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
                Range("A" & RowCount + 1).Select
                Range("A" & RowCount + 1).Cells.Value = SimName
                Range("B" & RowCount + 1).Cells.Value = SimType
                Range("C" & RowCount + 1).Cells.Value = SimDefVal
                Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                Range("E" & RowCount + 1).Cells.Value = SimDesc
            
            ElseIf InStr(IOType, "AI") > 0 Then
                'Paste First Row
                SimName = IOVariable
                SimType = "F R/W"
                SimDefVal = "0"
                SimIOAddr = "[" & Prefix & "_Sim]" & IOAddress
                SimDesc = Description

                Sheets(DataSheet).Select
                RowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
                Range("A" & RowCount + 1).Select
                Range("A" & RowCount + 1).Cells.Value = SimName
                Range("B" & RowCount + 1).Cells.Value = SimType
                Range("C" & RowCount + 1).Cells.Value = SimDefVal
                Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                Range("E" & RowCount + 1).Cells.Value = SimDesc
                
                ' Write data to IO tag sheets
                Sheets(stripSheet).Select
                RowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
                Range("A" & RowCount + 1).Select
                Range("A" & RowCount + 1).Cells.Value = SimName
                Range("B" & RowCount + 1).Cells.Value = SimType
                Range("C" & RowCount + 1).Cells.Value = SimDefVal
                Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                Range("E" & RowCount + 1).Cells.Value = SimDesc
                
                Sheets(stripMinMax).Select
                RowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
                Range("A" & RowCount + 1).Select
                Range("A" & RowCount + 1).Cells.Value = SimName
                Range("B" & RowCount + 1).Cells.Value = InputMin
                Range("C" & RowCount + 1).Cells.Value = InputMax
                Range("D" & RowCount + 1).Cells.Value = OutputMin
                Range("E" & RowCount + 1).Cells.Value = OutputMax
                
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

                Sheets(DataSheet).Select
                RowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
                Range("A" & RowCount + 1).Select
                Range("A" & RowCount + 1).Cells.Value = SimName
                Range("B" & RowCount + 1).Cells.Value = SimType
                Range("C" & RowCount + 1).Cells.Value = SimDefVal
                Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                Range("E" & RowCount + 1).Cells.Value = SimDesc
                
                ' Add faults to IO tag sheet
                Sheets(stripSheet).Select
                RowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
                Range("A" & RowCount + 1).Select
                Range("A" & RowCount + 1).Cells.Value = SimName
                Range("B" & RowCount + 1).Cells.Value = SimType
                Range("C" & RowCount + 1).Cells.Value = SimDefVal
                Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                Range("E" & RowCount + 1).Cells.Value = SimDesc
                
                Sheets(stripMinMax).Select
                RowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
                Range("A" & RowCount + 1).Select
                Range("A" & RowCount + 1).Cells.Value = SimName
                Range("B" & RowCount + 1).Cells.Value = InputMin
                Range("C" & RowCount + 1).Cells.Value = InputMax
                Range("D" & RowCount + 1).Cells.Value = OutputMin
                Range("E" & RowCount + 1).Cells.Value = OutputMax
            
            ElseIf InStr(IOType, "AO") > 0 Then
                'Paste Row
                SimName = IOVariable
                SimType = "F R"
                SimDefVal = ""
                SimIOAddr = "[" & Prefix & "_Sim]" & IOAddress
                SimDesc = Description

                Sheets(DataSheet).Select
                RowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
                Range("A" & RowCount + 1).Select
                Range("A" & RowCount + 1).Cells.Value = SimName
                Range("B" & RowCount + 1).Cells.Value = SimType
                Range("C" & RowCount + 1).Cells.Value = SimDefVal
                Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                Range("E" & RowCount + 1).Cells.Value = SimDesc
                
                ' Write data to IO tag sheet
                Sheets(stripSheet).Select
                RowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
                Range("A" & RowCount + 1).Select
                Range("A" & RowCount + 1).Cells.Value = SimName
                Range("B" & RowCount + 1).Cells.Value = SimType
                Range("C" & RowCount + 1).Cells.Value = SimDefVal
                Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                Range("E" & RowCount + 1).Cells.Value = SimDesc
            
                Sheets(stripMinMax).Select
                RowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
                Range("A" & RowCount + 1).Select
                Range("A" & RowCount + 1).Cells.Value = SimName
                Range("B" & RowCount + 1).Cells.Value = InputMin
                Range("C" & RowCount + 1).Cells.Value = InputMax
                Range("D" & RowCount + 1).Cells.Value = OutputMin
                Range("E" & RowCount + 1).Cells.Value = OutputMax
                                        
            ElseIf InStr(IOType, "RTD") > 0 Then
                'Paste First Row
                SimName = IOVariable
                SimType = "F R/W"
                SimDefVal = "0"
                SimIOAddr = "[" & Prefix & "_Sim]" & IOAddress
                SimDesc = Description

                Sheets(DataSheet).Select
                RowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
                Range("A" & RowCount + 1).Select
                Range("A" & RowCount + 1).Cells.Value = SimName
                Range("B" & RowCount + 1).Cells.Value = SimType
                Range("C" & RowCount + 1).Cells.Value = SimDefVal
                Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                Range("E" & RowCount + 1).Cells.Value = SimDesc
                
                'Write data to IO tag sheet
                Sheets(stripSheet).Select
                RowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
                Range("A" & RowCount + 1).Select
                Range("A" & RowCount + 1).Cells.Value = SimName
                Range("B" & RowCount + 1).Cells.Value = SimType
                Range("C" & RowCount + 1).Cells.Value = SimDefVal
                Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                Range("E" & RowCount + 1).Cells.Value = SimDesc
                
                Sheets(stripMinMax).Select
                RowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
                Range("A" & RowCount + 1).Select
                Range("A" & RowCount + 1).Cells.Value = SimName
                Range("B" & RowCount + 1).Cells.Value = InputMin
                Range("C" & RowCount + 1).Cells.Value = InputMax
                Range("D" & RowCount + 1).Cells.Value = OutputMin
                Range("E" & RowCount + 1).Cells.Value = OutputMax
                
                'Paste Second (Fault) Row
                SimName = SimName & "_Flt"
                SimType = "B R/W"
                SimIOAddr = Replace(SimIOAddr, "Data", "Fault")
                SimDesc = Description & " CH_FLT"

                Sheets(DataSheet).Select
                RowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
                Range("A" & RowCount + 1).Select
                Range("A" & RowCount + 1).Cells.Value = SimName
                Range("B" & RowCount + 1).Cells.Value = SimType
                Range("C" & RowCount + 1).Cells.Value = SimDefVal
                Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                Range("E" & RowCount + 1).Cells.Value = SimDesc
                
                Sheets(stripSheet).Select
                RowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
                Range("A" & RowCount + 1).Select
                Range("A" & RowCount + 1).Cells.Value = SimName
                Range("B" & RowCount + 1).Cells.Value = SimType
                Range("C" & RowCount + 1).Cells.Value = SimDefVal
                Range("D" & RowCount + 1).Cells.Value = SimIOAddr
                Range("E" & RowCount + 1).Cells.Value = SimDesc
                
                Sheets(stripMinMax).Select
                RowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
                Range("A" & RowCount + 1).Select
                Range("A" & RowCount + 1).Cells.Value = SimName
                Range("B" & RowCount + 1).Cells.Value = InputMin
                Range("C" & RowCount + 1).Cells.Value = InputMax
                Range("D" & RowCount + 1).Cells.Value = OutputMin
                Range("E" & RowCount + 1).Cells.Value = OutputMax
                
            End If
        End If
    Next i
    
End Sub
Sub CheckMinMaxData(minMaxSheet As String)
'
'   Checks to make sure the Min Max data is numeric.
'
    
    Sheets(minMaxSheet).Select
    RowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To RowCount
        If Not IsNumeric(Worksheets(minMaxSheet).Range("B" & i).Cells.Value) Then
            Sheets("IO Sheets").Select
            Columns("L").Select
            MsgBox "InputMin must be numeric values."
            Exit For
        End If
    Next i
        
    For i = 2 To RowCount
        If Not IsNumeric(Worksheets(minMaxSheet).Range("C" & i).Cells.Value) Then
            Sheets("IO Sheets").Select
            Columns("M").Select
            MsgBox "InputMax must be numeric values."
            Exit For
        End If
    Next i
    
    For i = 2 To RowCount
        If Not IsNumeric(Worksheets(minMaxSheet).Range("D" & i).Cells.Value) Then
            Sheets("IO Sheets").Select
            Columns("N").Select
            MsgBox "OutputMin must be numeric values."
            Exit For
        End If
    Next i
    
    For i = 2 To RowCount
        If Not IsNumeric(Worksheets(minMaxSheet).Range("E" & i).Cells.Value) Then
            Sheets("IO Sheets").Select
            Columns("O").Select
            MsgBox "OutputMax must be numeric values."
            Exit For
        End If
    Next i
        
End Sub

Function Is_Cell_Blank(DataSheet As String) As Boolean
'
'
'
    Sheets(DataSheet).Select
    If IsEmpty(Range("A8")) Then
        'MsgBox "Empty"
        IsCellBlank = True
    ElseIf Range("A8") = "" Then
        'MsgBox "Empty Text"
        If Range("A8").HasFormula Then
            'MsgBox "Empty Text is the result of a formula"
        End If
        IsCellBlank = False
    Else
        'MsgBox "Contains data"
        IsCellBlank = False
    End If
End Function
Sub Rem_Spaces(destSheet As String, DestCol As String)
'
'
'
    Sheets(destSheet).Select
    Columns(DestCol).Replace What:="  ", _
                        Replacement:=" ", _
                        LookAt:=xlPart, _
                        SearchOrder:=xlByRows, _
                        MatchCase:=False, _
                        SearchFormat:=False, _
                        ReplaceFormat:=False
    Range("A1").Select
End Sub

Sub Remove_From_Desc(ByRef x As Integer)
    Dim DelWord As String
    Dim OldDesc As String
    Dim NewDesc As String

    RowCount = Cells(Cells.Rows.Count, "D").End(xlUp).Row
    DelWord = InputBox("Please enter the word you wish to delete:", "Delete From Descriptions")

    'Range("H1").FormulaR1C1 = DelWord
    If DelWord <> "" Then
        For i = 8 To RowCount
            Range("D" & i).Select
            OldDesc = Range("D" & i).Value
            NewDesc = Replace(OldDesc, DelWord, "")
            Range("D" & i).Value = NewDesc
        Next
    End If

    Range("D8").Select

End Sub
Sub SortByColumn(sheetName As String, SortCol As String)
'
'
'
    Sheets(sheetName).Select
    RowCount = Cells(Cells.Rows.Count, "D").End(xlUp).Row
    Range("A2:E" & RowCount).Sort Key1:=Range(SortCol & 2), Order1:=xlAscending
    
End Sub
