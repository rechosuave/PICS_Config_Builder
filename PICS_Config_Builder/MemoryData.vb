Sub Generate_Memory_Data(ByRef x As Integer)

    Call Unhide_All_Sheets

    Call Clear_Sheet_Type("MemoryData")
    Call Clear_Sheet_Type("IOMem")

    Call Generate_AI_Memory("IOTags - AIn", "IOMem - AIn")
    Call Generate_DI_Memory("IOTags - DIn", "IOMem - DIn")
    Call Generate_ValvesC_Memory("IOTags - ValveC", "IOMem - ValveC")
    Call Generate_ValvesMO_Memory("IOTags - ValveMO", "IOMem - ValveMO")
    Call Generate_ValvesSO_Memory("IOTags - ValveSO", "IOMem - ValveSO")
    Call Generate_Motor_Memory("IOTags - Motor", "IOMem - Motor")
    Call Generate_VSD_Memory("IOTags - VSD", "IOMem - VSD")

    Call Remove_From_Descriptions()
    Call Rem_Spaces("IOMem - AIn", "F")
    Call Rem_Spaces("IOMem - DIn", "F")
    Call Rem_Spaces("IOMem - ValveC", "F")
    Call Rem_Spaces("IOMem - ValveMO", "F")
    Call Rem_Spaces("IOMem - ValveSO", "F")
    Call Rem_Spaces("IOMem - Motor", "F")
    Call Rem_Spaces("IOMem - VSD", "F")

    Call Copy_Memory_Data()

    Sheets("Instructions").Select

    Call Hide_Sheets

End Sub

Private Function Write_Memory(destSheet As String, Optional inNum As String, Optional inName As String, _
                            Optional inType As String, Optional inVal As String, Optional inDesc As String)

    Sheets(destSheet).Select
    RowCount = Cells(Cells.Rows.Count, "B").End(xlUp).Row
    Range("A" & RowCount + 1).Select
    
    Range("A" & RowCount + 1).Cells.Value = inNum
    Range("B" & RowCount + 1).Cells.Value = inName
    Range("C" & RowCount + 1).Cells.Value = inType
    Range("D" & RowCount + 1).Cells.Value = inVal
    Range("F" & RowCount + 1).Cells.Value = inDesc

End Function

Sub Generate_AI_Memory(sourceSheet As String, destSheet As String)
'
'
'   Generate AI Memory
    Dim IO_Number As String
    Dim IO_Name As String
    Dim IO_Type As String
    Dim IO_Val As String
    Dim IO_Addr As String
    Dim IO_Desc As String
    
    Sheets(sourceSheet).Select
    SourceRowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
    
    IO_Number = 0

    For i = 2 To SourceRowCount
        
        IO_Name = Worksheets(sourceSheet).Range("A" & i).Cells.Value
        IO_Type = Worksheets(sourceSheet).Range("B" & i).Cells.Value
        IO_Val = Worksheets(sourceSheet).Range("C" & i).Cells.Value
        IO_Desc = Worksheets(sourceSheet).Range("E" & i).Cells.Value
        
        IO_Name = Replace(IO_Name, "_Inp_PV", "")
        IO_Name = Replace(IO_Name, "_Inp_AV", "")
        
        If InStr(IO_Name, "Flt") = False Then
            IO_Number = IO_Number + 1
            ' Write lines
            Write_Memory destSheet, IO_Number, IO_Name, IO_Type, IO_Val, IO_Desc
            Write_Memory destSheet, "", IO_Name & "_Flt", "B R/W", "0", IO_Desc & " IO Fault"
            Write_Memory destSheet, "", IO_Name & "_OR", "F R/W", "0", IO_Desc & " Override Level"
            Write_Memory destSheet, "", IO_Name & "_OR_EN", "B R/W", "0", IO_Desc & " Override Enable"
            Write_Memory destSheet, "", IO_Name & "_PV_DB", "F R/W", "0.025", IO_Desc & " Noise Level"
            Write_Memory destSheet, "", IO_Name & "_PV_EN", "B R/W", "0", IO_Desc & " Noise Enable Bit"
            Write_Memory destSheet, "", IO_Name & "_String", "STR R/W", IO_Name, ""
        End If
    Next

End Sub
Sub Generate_DI_Memory(sourceSheet As String, destSheet As String)
'
'
'   Generate DI Memory
    Dim IO_Number As String
    Dim IO_Name As String
    Dim IO_Type As String
    Dim IO_Val As String
    Dim IO_Addr As String
    Dim IO_Desc As String
        
    Sheets(sourceSheet).Select
    SourceRowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
    
    IO_Number = 0

    For i = 2 To SourceRowCount
        IO_Name = Worksheets(sourceSheet).Range("A" & i).Cells.Value
        IO_Type = Worksheets(sourceSheet).Range("B" & i).Cells.Value
        IO_Val = Worksheets(sourceSheet).Range("C" & i).Cells.Value
        IO_Desc = Worksheets(sourceSheet).Range("E" & i).Cells.Value
        
        IO_Name = Replace(IO_Name, "_Inp_PV", "")
    
        If InStr(IO_Name, "Flt") = False Then
            IO_Number = IO_Number + 1
            
            ' Write Lines
            Write_Memory destSheet, IO_Number, IO_Name, IO_Type, IO_Val, IO_Desc
            Write_Memory destSheet, "", IO_Name & "_Flt", "B R/W", "0", IO_Desc & " IO Fault"
            Write_Memory destSheet, "", IO_Name & "_String", "STR R/W", IO_Name, ""
            
        End If
    Next

End Sub
Sub Generate_ValvesC_Memory(sourceSheet As String, destSheet As String)
'
'
'   Generate ValvesC Memory
    Dim IO_Number As String
    Dim IO_Name As String
    Dim IO_Type As String
    Dim IO_Val As String
    Dim IO_Addr As String
    Dim IO_Desc As String
        
    Sheets(sourceSheet).Select
    SourceRowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
    
    IO_Number = 0

    For i = 2 To SourceRowCount
        IO_Name = Worksheets(sourceSheet).Range("A" & i).Cells.Value
        IO_Type = Worksheets(sourceSheet).Range("B" & i).Cells.Value
        IO_Val = Worksheets(sourceSheet).Range("C" & i).Cells.Value
        IO_Desc = Worksheets(sourceSheet).Range("E" & i).Cells.Value
        
        IO_Name = Replace(IO_Name, "_Out_CV", "")
        
        If IO_Type = "F R" Then
            IO_Number = IO_Number + 1
            
            ' Write lines
            Write_Memory destSheet, IO_Number, IO_Name & "_Fbk_Flt", "B R/W", "0", IO_Desc & " Feedback Fault"
            Write_Memory destSheet, "", IO_Name & "_OR", "F R/W", "0", IO_Desc & " Override Level"
            Write_Memory destSheet, "", IO_Name & "_OR_EN", "B R/W", "0", IO_Desc & " Override Enable Bit"
            Write_Memory destSheet, "", IO_Name & "_String", "STR R/W", IO_Name, ""
            
        End If
        
        Next

End Sub
Sub Generate_ValvesMO_Memory(sourceSheet As String, destSheet As String)
'
'
'   Generate ValvesSO Memory
    Dim IO_Number As String
    Dim IO_Name As String
    Dim IO_Type As String
    Dim IO_Val As String
    Dim IO_Addr As String
    Dim IO_Desc As String
        
    Sheets(sourceSheet).Select
    SourceRowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
    
    IO_Number = 0

    For i = 2 To SourceRowCount
        IO_Name = Worksheets(sourceSheet).Range("A" & i).Cells.Value
        IO_Type = Worksheets(sourceSheet).Range("B" & i).Cells.Value
        IO_Val = Worksheets(sourceSheet).Range("C" & i).Cells.Value
        IO_Desc = Worksheets(sourceSheet).Range("E" & i).Cells.Value
    
        IO_Name = Replace(IO_Name, "_Out", "")
    
        If IO_Type = "B R" Then
            IO_Number = IO_Number + 1
            
            ' Write lines
            Write_Memory destSheet, IO_Number, IO_Name & "_FTC", "B R/W", "0", IO_Desc & " Fail to Close"
            Write_Memory destSheet, "", IO_Name & "_FTO", "B R/W", "0", IO_Desc & " Fail to Open"
            Write_Memory destSheet, "", IO_Name & "_Stuck", "B R/W", "0", IO_Desc & " Is Stuck"
            Write_Memory destSheet, "", IO_Name & "_Inp_ActuatorFault", "B R/W", "0", IO_Desc & " Act Fault"
            Write_Memory destSheet, "", IO_Name & "_Inp_Hand", "B R/W", "0", IO_Desc & " Input Hand"
            Write_Memory destSheet, "", IO_Name & "_String", "STR R/W", IO_Name, ""
            
        End If
        
    Next

End Sub

Sub Generate_ValvesSO_Memory(sourceSheet As String, destSheet As String)
'
'
'   Generate ValvesSO Memory
    Dim IO_Number As String
    Dim IO_Name As String
    Dim IO_Type As String
    Dim IO_Val As String
    Dim IO_Addr As String
    Dim IO_Desc As String
        
    Sheets(sourceSheet).Select
    SourceRowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
    
    IO_Number = 0

    For i = 2 To SourceRowCount
        IO_Name = Worksheets(sourceSheet).Range("A" & i).Cells.Value
        IO_Type = Worksheets(sourceSheet).Range("B" & i).Cells.Value
        IO_Val = Worksheets(sourceSheet).Range("C" & i).Cells.Value
        IO_Desc = Worksheets(sourceSheet).Range("E" & i).Cells.Value
    
        IO_Name = Replace(IO_Name, "_Out", "")
    
        If IO_Type = "B R" Then
            IO_Number = IO_Number + 1
            
            ' Write lines
            Write_Memory destSheet, IO_Number, IO_Name & "_FTC", "B R/W", "0", IO_Desc & " Fail to Close"
            Write_Memory destSheet, "", IO_Name & "_FTO", "B R/W", "0", IO_Desc & " Fail to Open"
            Write_Memory destSheet, "", IO_Name & "_Inp_Hand", "B R/W", "0", IO_Desc & " Input Hand"
            Write_Memory destSheet, "", IO_Name & "_Stuck", "B R/W", "0", IO_Desc & " Is Stuck"
            Write_Memory destSheet, "", IO_Name & "_String", "STR R/W", IO_Name, ""
            
        End If
        
    Next

End Sub

Sub Generate_Motor_Memory(sourceSheet As String, destSheet As String)
'
'
'   Generate Motor Memory
    Dim IO_Number As String
    Dim IO_Name As String
    Dim IO_Type As String
    Dim IO_Val As String
    Dim IO_Addr As String
    Dim IO_Desc As String
        
    Sheets(sourceSheet).Select
    SourceRowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
    
    IO_Number = 0

    For i = 2 To SourceRowCount
        IO_Name = Worksheets(sourceSheet).Range("A" & i).Cells.Value
        IO_Type = Worksheets(sourceSheet).Range("B" & i).Cells.Value
        IO_Val = Worksheets(sourceSheet).Range("C" & i).Cells.Value
        IO_Desc = Worksheets(sourceSheet).Range("E" & i).Cells.Value
    
        IO_Name = Replace(IO_Name, "_Out_Run", "")
    
        If IO_Type = "B R" Then
            IO_Number = IO_Number + 1
        
            ' Write lines
            Write_Memory destSheet, IO_Number, IO_Name & "_Inp_Faulted", "B R/W", "0", IO_Desc & " Faulted"
            Write_Memory destSheet, "", IO_Name & "_FTR", "B R/W", "0", IO_Desc & " Fail to Run"
            Write_Memory destSheet, "", IO_Name & "_FTS", "B R/W", "0", IO_Desc & " Fail to Stop"
            Write_Memory destSheet, "", IO_Name & "_Inp_Hand", "B R/W", "0", IO_Desc & " Input Hand"
            Write_Memory destSheet, "", IO_Name & "_OverLoad", "B R/W", "0", IO_Desc & " OverLoad"
            Write_Memory destSheet, "", IO_Name & "_String", "STR R/W", IO_Name, ""
            
        End If
    Next

End Sub
Sub Generate_VSD_Memory(sourceSheet As String, destSheet As String)
'
'
'   Generate Motor Memory
    Dim IO_Number As String
    Dim IO_Name As String
    Dim IO_Type As String
    Dim IO_Val As String
    Dim IO_Addr As String
    Dim IO_Desc As String
        
    Sheets(sourceSheet).Select
    SourceRowCount = Cells(Cells.Rows.Count, "A").End(xlUp).Row
    
    IO_Number = 0

    For i = 2 To SourceRowCount
        IO_Name = Worksheets(sourceSheet).Range("A" & i).Cells.Value
        IO_Type = Worksheets(sourceSheet).Range("B" & i).Cells.Value
        IO_Val = Worksheets(sourceSheet).Range("C" & i).Cells.Value
        IO_Desc = Worksheets(sourceSheet).Range("E" & i).Cells.Value
        
        IO_Name = Replace(IO_Name, "_Out_Run", "")
        
        If IO_Type = "B R" Then
            IO_Number = IO_Number + 1
            
            ' Write lines
            Write_Memory destSheet, IO_Number, IO_Name & "_Inp_Faulted", "B R/W", "0", IO_Desc & " Faulted"
            Write_Memory destSheet, "", IO_Name & "_FTR", "B R/W", "0", IO_Desc & " Fail to Run"
            Write_Memory destSheet, "", IO_Name & "_FTS", "B R/W", "0", IO_Desc & " Fail to Stop"
            Write_Memory destSheet, "", IO_Name & "_Inp_Hand", "B R/W", "0", IO_Desc & " Input Hand"
            Write_Memory destSheet, "", IO_Name & "_String", "STR R/W", IO_Name, ""
            
        End If
    Next

End Sub

Sub Copy_Memory_Data(ByRef x As Integer)
    '
    '
    '
    'Clear MemoryData sheet
    Sheets("MemoryData").Select
    Range("A2:F9999").Clear

    'Copy AIn Memory Data
    Sheets("IOMem - AIn").Select
    RowCount = Cells(Cells.Rows.Count, "B").End(xlUp).Row
    If RowCount > 1 Then
        Range("B2:F" & RowCount).Copy

        'Paste data into MemoryData sheet
        Sheets("MemoryData").Select
        Range("A2").PasteSpecial(xlPasteValues)
    End If

    'Copy DIn Memory Data
    Sheets("IOMem - DIn").Select
    RowCount = Cells(Cells.Rows.Count, "B").End(xlUp).Row
    If RowCount > 1 Then
        Range("B2:F" & RowCount).Copy

        'Paste data into MemoryData sheet
        Sheets("MemoryData").Select
        MemRowCount = Cells(Cells.Rows.Count, "B").End(xlUp).Row
        MemRow = MemRowCount + 1
        Range("A" & MemRow).PasteSpecial(xlPasteValues)
    End If

    'Copy ValveC Memory Data
    Sheets("IOMem - ValveC").Select
    RowCount = Cells(Cells.Rows.Count, "B").End(xlUp).Row
    If RowCount > 1 Then
        Range("B2:F" & RowCount).Copy

        'Paste data into MemoryData sheet
        Sheets("MemoryData").Select
        MemRowCount = Cells(Cells.Rows.Count, "B").End(xlUp).Row
        MemRow = MemRowCount + 1
        Range("A" & MemRow).PasteSpecial(xlPasteValues)
    End If

    'Copy ValveMO Memory Data
    Sheets("IOMem - ValveMO").Select
    RowCount = Cells(Cells.Rows.Count, "B").End(xlUp).Row
    If RowCount > 1 Then
        Range("B2:F" & RowCount).Copy

        'Paste data into MemoryData sheet
        Sheets("MemoryData").Select
        MemRowCount = Cells(Cells.Rows.Count, "B").End(xlUp).Row
        MemRow = MemRowCount + 1
        Range("A" & MemRow).PasteSpecial(xlPasteValues)
    End If

    'Copy ValveSO Memory Data
    Sheets("IOMem - ValveSO").Select
    RowCount = Cells(Cells.Rows.Count, "B").End(xlUp).Row
    If RowCount > 1 Then
        Range("B2:F" & RowCount).Copy

        'Paste data into MemoryData sheet
        Sheets("MemoryData").Select
        MemRowCount = Cells(Cells.Rows.Count, "B").End(xlUp).Row
        MemRow = MemRowCount + 1
        Range("A" & MemRow).PasteSpecial(xlPasteValues)
    End If

    'Copy Motor Memory Data
    Sheets("IOMem - Motor").Select
    RowCount = Cells(Cells.Rows.Count, "B").End(xlUp).Row
    If RowCount > 1 Then
        Range("B2:F" & RowCount).Copy

        'Paste data into MemoryData sheet
        Sheets("MemoryData").Select
        MemRowCount = Cells(Cells.Rows.Count, "B").End(xlUp).Row
        MemRow = MemRowCount + 1
        Range("A" & MemRow).PasteSpecial(xlPasteValues)
    End If

    'Copy VSD Memory Data
    Sheets("IOMem - VSD").Select
    RowCount = Cells(Cells.Rows.Count, "B").End(xlUp).Row
    If RowCount > 1 Then
        Range("B2:F" & RowCount).Copy
        Range("A1").Select

        Sheets("MemoryData").Select
        MemRowCount = Cells(Cells.Rows.Count, "B").End(xlUp).Row
        MemRow = MemRowCount + 1
        Range("A" & MemRow).PasteSpecial(xlPasteValues)
    End If

    Range("A1").Select
    Application.CutCopyMode = False

End Sub
Sub Remove_From_Descriptions(ByRef x As Integer)
    '
    '
    '
    Dim Keyword As String

    ' Remove keywords from ValveC Memory
    Sheets("IOMem - ValveC").Select
    For i = 10 To 17    'Data is in rows 10 to 17
        Keyword = Worksheets("Instructions").Range("C" & i).Cells.Value
        If Keyword <> "" Then
            Columns("F").Replace What:=" " & Keyword,
                        Replacement:="",
                        LookAt:=xlPart,
                        SearchOrder:=xlByRows,
                        MatchCase:=False,
                        SearchFormat:=False,
                        ReplaceFormat:=False
        End If
    Next

    ' Remove keywords from ValveSO_MO Memory
    Sheets("IOMem - ValveMO").Select
    For i = 10 To 17    'Data is in rows 10 to 17
        Keyword = Worksheets("Instructions").Range("D" & i).Cells.Value
        If Keyword <> "" Then
            Columns("F").Replace What:=" " & Keyword,
                        Replacement:="",
                        LookAt:=xlPart,
                        SearchOrder:=xlByRows,
                        MatchCase:=False,
                        SearchFormat:=False,
                        ReplaceFormat:=False
        End If
    Next

    ' Remove keywords from ValveSO_MO Memory
    Sheets("IOMem - ValveSO").Select
    For i = 10 To 17    'Data is in rows 10 to 17
        Keyword = Worksheets("Instructions").Range("D" & i).Cells.Value
        If Keyword <> "" Then
            Columns("F").Replace What:=" " & Keyword,
                        Replacement:="",
                        LookAt:=xlPart,
                        SearchOrder:=xlByRows,
                        MatchCase:=False,
                        SearchFormat:=False,
                        ReplaceFormat:=False
        End If
    Next

    ' Remove keywords from Motor Memory
    Sheets("IOMem - Motor").Select
    For i = 10 To 17    'Data is in rows 10 to 17
        Keyword = Worksheets("Instructions").Range("E" & i).Cells.Value
        If Keyword <> "" Then
            Columns("F").Replace What:=" " & Keyword,
                        Replacement:="",
                        LookAt:=xlPart,
                        SearchOrder:=xlByRows,
                        MatchCase:=False,
                        SearchFormat:=False,
                        ReplaceFormat:=False
        End If
    Next

    ' Remove keywords from VSD Memory
    Sheets("IOMem - VSD").Select
    For i = 10 To 17    'Data is in rows 10 to 17
        Keyword = Worksheets("Instructions").Range("F" & i).Cells.Value
        If Keyword <> "" Then
            Columns("F").Replace What:=" " & Keyword,
                        Replacement:="",
                        LookAt:=xlPart,
                        SearchOrder:=xlByRows,
                        MatchCase:=False,
                        SearchFormat:=False,
                        ReplaceFormat:=False
        End If
    Next


End Sub
