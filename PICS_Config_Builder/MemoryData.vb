
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Module MemoryData

    Sub Generate_Memory_Data()

        If WS_Exists("MemoryData") Then Call Clear_Sheet_Type("MemoryData")
        If WS_Exists("IOMem") Then Call Clear_Sheet_Type("IOMem")

        If WS_Exists("IOTags - AIn") Then Call Generate_AI_Memory("IOTags - AIn", "IOMem - AIn")
        If WS_Exists("IOTags - DIn") Then Call Generate_DI_Memory("IOTags - DIn", "IOMem - DIn")
        If WS_Exists("IOTags - ValveC") Then Call Generate_ValvesC_Memory("IOTags - ValveC", "IOMem - ValveC")
        If WS_Exists("IOTags - ValveMO") Then Call Generate_ValvesMO_Memory("IOTags - ValveMO", "IOMem - ValveMO")
        If WS_Exists("IOTags - Motor") Then Call Generate_Motor_Memory("IOTags - Motor", "IOMem - Motor")
        If WS_Exists("IOTags - VSD") Then Call Generate_VSD_Memory("IOTags - VSD", "IOMem - VSD")

        Call Remove_From_Descriptions()

        If WS_Exists("IOMem - AIn") Then Call Remove_Spaces("IOMem - AIn", "F")
        If WS_Exists("IOMem - DIn") Then Call Remove_Spaces("IOMem - DIn", "F")
        If WS_Exists("IOMem - ValveC") Then Call Remove_Spaces("IOMem - ValveC", "F")
        If WS_Exists("IOMem - ValveMO") Then Call Remove_Spaces("IOMem - ValveMO", "F")
        If WS_Exists("IOMem - ValveSO") Then Call Remove_Spaces("IOMem - ValveSO", "F")
        If WS_Exists("IOMem - Motor") Then Call Remove_Spaces("IOMem - Motor", "F")
        If WS_Exists("IOMem - VSD") Then Call Remove_Spaces("IOMem - VSD", "F")

        Call Copy_Memory_Data()

    End Sub

    Private Sub Write_Memory(ByVal destSheet As String, Optional ByVal inNum As String = "", Optional ByVal inName As String = "",
                             Optional ByVal inType As String = "", Optional ByVal inVal As String = "", Optional ByVal inDesc As String = "")

        Dim ws As Worksheet
        Dim RowCount As Integer

        If WS_Exists(destSheet) Then    ' Does worksheet exist 
            ws = XLpicsWB.Sheets(destSheet)
        Else    ' create new worksheet
            ws = XLpicsWB.Sheets.Add(,,, XlSheetType.xlWorksheet)
            ws.Name = destSheet
        End If

        RowCount = ws.Cells(ws.Rows.Count, "B").End(XlDirection.xlUp).Row
        ws.Range("A" & RowCount + 1).Select()

        ws.Range("A" & RowCount + 1).Cells.Value = inNum
        ws.Range("B" & RowCount + 1).Cells.Value = inName
        ws.Range("C" & RowCount + 1).Cells.Value = inType
        ws.Range("D" & RowCount + 1).Cells.Value = inVal
        ws.Range("F" & RowCount + 1).Cells.Value = inDesc

    End Sub

    Sub Generate_AI_Memory(ByVal sourceSheet As String, ByVal destSheet As String)
        '
        '   Generate AI Memory
        Dim IO_Number As String
        Dim IO_Name As String
        Dim IO_Type As String
        Dim IO_Val As String
        Dim IO_Desc As String
        Dim SourceRowCount As Integer
        Dim ws As Worksheet = XLpicsWB.Sheets(sourceSheet)

        SourceRowCount = ws.Cells(ws.Rows.Count, "A").End(XlDirection.xlUp).Row

        IO_Number = 0

        For i = 2 To SourceRowCount

            IO_Name = ws.Range("A" & i).Cells.Value
            IO_Type = ws.Range("B" & i).Cells.Value
            IO_Val = ws.Range("C" & i).Cells.Value
            IO_Desc = ws.Range("E" & i).Cells.Value

            IO_Name = Replace(IO_Name, "_Inp_PV", "")
            IO_Name = Replace(IO_Name, "_Inp_AV", "")

            If InStr(IO_Name, "Flt") = False Then
                IO_Number = IO_Number + 1

                ' Write lines
                Write_Memory(destSheet, IO_Number, IO_Name, IO_Type, IO_Val, IO_Desc)
                Write_Memory(destSheet, "", IO_Name & "_Flt", "B R/W", "0", IO_Desc & " IO Fault")
                Write_Memory(destSheet, "", IO_Name & "_OR", "F R/W", "0", IO_Desc & " Override Level")
                Write_Memory(destSheet, "", IO_Name & "_OR_EN", "B R/W", "0", IO_Desc & " Override Enable")
                Write_Memory(destSheet, "", IO_Name & "_PV_DB", "F R/W", "0.025", IO_Desc & " Noise Level")
                Write_Memory(destSheet, "", IO_Name & "_PV_EN", "B R/W", "0", IO_Desc & " Noise Enable Bit")
                Write_Memory(destSheet, "", IO_Name & "_String", "STR R/W", IO_Name, "")
            End If
        Next

    End Sub
    Sub Generate_DI_Memory(ByRef sourceSheet As String, ByRef destSheet As String)
        '
        '   Generate DI Memory
        Dim IO_Number As String
        Dim IO_Name As String
        Dim IO_Type As String
        Dim IO_Val As String
        Dim IO_Desc As String
        Dim SourceRowCount As Integer
        Dim ws As Worksheet = XLpicsWB.Sheets(sourceSheet).Select

        SourceRowCount = ws.Cells(ws.Rows.Count, "A").End.xlUp.Row

        IO_Number = 0

        For i = 2 To SourceRowCount
            IO_Name = ws.Range("A" & i).Cells.Value
            IO_Type = ws.Range("B" & i).Cells.Value
            IO_Val = ws.Range("C" & i).Cells.Value
            IO_Desc = ws.Range("E" & i).Cells.Value

            IO_Name = Replace(IO_Name, "_Inp_PV", "")

            If InStr(IO_Name, "Flt") = False Then
                IO_Number = IO_Number + 1

                ' Write Lines
                Write_Memory(destSheet, IO_Number, IO_Name, IO_Type, IO_Val, IO_Desc)
                Write_Memory(destSheet, "", IO_Name & "_Flt", "B R/W", "0", IO_Desc & " IO Fault")
                Write_Memory(destSheet, "", IO_Name & "_String", "STR R/W", IO_Name, "")

            End If
        Next

    End Sub
    Sub Generate_ValvesC_Memory(ByRef sourceSheet As String, ByRef destSheet As String)
        '
        '
        '   Generate ValvesC Memory
        Dim IO_Number As String
        Dim IO_Name As String
        Dim IO_Type As String
        Dim IO_Val As String
        Dim IO_Desc As String
        Dim SourceRowCount As Integer
        Dim ws As Worksheet = XLpicsWB.Sheets(sourceSheet).Select

        SourceRowCount = ws.Cells(ws.Rows.Count, "A").End.xlUp.Row

        IO_Number = 0

        For i = 2 To SourceRowCount
            IO_Name = ws.Range("A" & i).Cells.Value
            IO_Type = ws.Range("B" & i).Cells.Value
            IO_Val = ws.Range("C" & i).Cells.Value
            IO_Desc = ws.Range("E" & i).Cells.Value

            IO_Name = Replace(IO_Name, "_Out_CV", "")

            If IO_Type = "F R" Then
                IO_Number = IO_Number + 1

                ' Write lines
                Write_Memory(destSheet, IO_Number, IO_Name & "_Fbk_Flt", "B R/W", "0", IO_Desc & " Feedback Fault")
                Write_Memory(destSheet, "", IO_Name & "_OR", "F R/W", "0", IO_Desc & " Override Level")
                Write_Memory(destSheet, "", IO_Name & "_OR_EN", "B R/W", "0", IO_Desc & " Override Enable Bit")
                Write_Memory(destSheet, "", IO_Name & "_String", "STR R/W", IO_Name, "")

            End If

        Next

    End Sub
    Sub Generate_ValvesMO_Memory(ByRef sourceSheet As String, ByRef destSheet As String)
        '
        '
        '   Generate ValvesSO Memory
        Dim IO_Number As String
        Dim IO_Name As String
        Dim IO_Type As String
        Dim IO_Val As String
        Dim IO_Desc As String
        Dim SourceRowCount As Integer
        Dim ws As Worksheet = XLpicsWB.Sheets(sourceSheet).Select

        SourceRowCount = ws.Cells(ws.Rows.Count, "A").End.xlUp.Row

        IO_Number = 0

        For i = 2 To SourceRowCount
            IO_Name = ws.Range("A" & i).Cells.Value
            IO_Type = ws.Range("B" & i).Cells.Value
            IO_Val = ws.Range("C" & i).Cells.Value
            IO_Desc = ws.Range("E" & i).Cells.Value

            IO_Name = Replace(IO_Name, "_Out", "")

            If IO_Type = "B R" Then
                IO_Number = IO_Number + 1

                ' Write lines
                Write_Memory(destSheet, IO_Number, IO_Name & "_FTC", "B R/W", "0", IO_Desc & " Fail to Close")
                Write_Memory(destSheet, "", IO_Name & "_FTO", "B R/W", "0", IO_Desc & " Fail to Open")
                Write_Memory(destSheet, "", IO_Name & "_Stuck", "B R/W", "0", IO_Desc & " Is Stuck")
                Write_Memory(destSheet, "", IO_Name & "_Inp_ActuatorFault", "B R/W", "0", IO_Desc & " Act Fault")
                Write_Memory(destSheet, "", IO_Name & "_Inp_Hand", "B R/W", "0", IO_Desc & " Input Hand")
                Write_Memory(destSheet, "", IO_Name & "_String", "STR R/W", IO_Name, "")

            End If

        Next

    End Sub

    Sub Generate_ValvesSO_Memory(ByRef sourceSheet As String, ByRef destSheet As String)
        '
        '   Generate ValvesSO Memory
        Dim IO_Number As String
        Dim IO_Name As String
        Dim IO_Type As String
        Dim IO_Val As String
        Dim IO_Desc As String
        Dim SourceRowCount As Integer
        Dim ws As Worksheet = XLpicsWB.Sheets(sourceSheet).Select

        SourceRowCount = ws.Cells(ws.Rows.Count, "A").End.xlUp.Row

        IO_Number = 0

        For i = 2 To SourceRowCount
            IO_Name = ws.Range("A" & i).Cells.Value
            IO_Type = ws.Range("B" & i).Cells.Value
            IO_Val = ws.Range("C" & i).Cells.Value
            IO_Desc = ws.Range("E" & i).Cells.Value

            IO_Name = Replace(IO_Name, "_Out", "")

            If IO_Type = "B R" Then
                IO_Number = IO_Number + 1

                ' Write lines
                Write_Memory(destSheet, IO_Number, IO_Name & "_FTC", "B R/W", "0", IO_Desc & " Fail to Close")
                Write_Memory(destSheet, "", IO_Name & "_FTO", "B R/W", "0", IO_Desc & " Fail to Open")
                Write_Memory(destSheet, "", IO_Name & "_Inp_Hand", "B R/W", "0", IO_Desc & " Input Hand")
                Write_Memory(destSheet, "", IO_Name & "_Stuck", "B R/W", "0", IO_Desc & " Is Stuck")
                Write_Memory(destSheet, "", IO_Name & "_String", "STR R/W", IO_Name, "")

            End If

        Next

    End Sub

    Sub Generate_Motor_Memory(ByRef sourceSheet As String, ByRef destSheet As String)
        '
        '   Generate Motor Memory
        Dim IO_Number As String
        Dim IO_Name As String
        Dim IO_Type As String
        Dim IO_Val As String
        Dim IO_Desc As String
        Dim SourceRowCount As Integer
        Dim ws As Worksheet = XLpicsWB.Sheets(sourceSheet).Select

        SourceRowCount = ws.Cells(ws.Rows.Count, "A").End.xlUp.Row

        IO_Number = 0

        For i = 2 To SourceRowCount
            IO_Name = ws.Range("A" & i).Cells.Value
            IO_Type = ws.Range("B" & i).Cells.Value
            IO_Val = ws.Range("C" & i).Cells.Value
            IO_Desc = ws.Range("E" & i).Cells.Value

            IO_Name = Replace(IO_Name, "_Out_Run", "")

            If IO_Type = "B R" Then
                IO_Number = IO_Number + 1

                ' Write lines
                Write_Memory(destSheet, IO_Number, IO_Name & "_Inp_Faulted", "B R/W", "0", IO_Desc & " Faulted")
                Write_Memory(destSheet, "", IO_Name & "_FTR", "B R/W", "0", IO_Desc & " Fail to Run")
                Write_Memory(destSheet, "", IO_Name & "_FTS", "B R/W", "0", IO_Desc & " Fail to Stop")
                Write_Memory(destSheet, "", IO_Name & "_Inp_Hand", "B R/W", "0", IO_Desc & " Input Hand")
                Write_Memory(destSheet, "", IO_Name & "_OverLoad", "B R/W", "0", IO_Desc & " OverLoad")
                Write_Memory(destSheet, "", IO_Name & "_String", "STR R/W", IO_Name, "")

            End If
        Next

    End Sub
    Sub Generate_VSD_Memory(ByRef sourceSheet As String, ByRef destSheet As String)
        '
        '   Generate Motor Memory
        Dim IO_Number As String
        Dim IO_Name As String
        Dim IO_Type As String
        Dim IO_Val As String
        Dim IO_Desc As String
        Dim SourceRowCount As Integer
        Dim ws As Worksheet = XLpicsWB.Sheets(sourceSheet).Select

        SourceRowCount = ws.Cells(ws.Rows.Count, "A").End.xlUp.Row

        IO_Number = 0

        For i = 2 To SourceRowCount
            IO_Name = ws.Range("A" & i).Cells.Value
            IO_Type = ws.Range("B" & i).Cells.Value
            IO_Val = ws.Range("C" & i).Cells.Value
            IO_Desc = ws.Range("E" & i).Cells.Value

            IO_Name = Replace(IO_Name, "_Out_Run", "")

            If IO_Type = "B R" Then
                IO_Number = IO_Number + 1

                ' Write lines
                Write_Memory(destSheet, IO_Number, IO_Name & "_Inp_Faulted", "B R/W", "0", IO_Desc & " Faulted")
                Write_Memory(destSheet, "", IO_Name & "_FTR", "B R/W", "0", IO_Desc & " Fail to Run")
                Write_Memory(destSheet, "", IO_Name & "_FTS", "B R/W", "0", IO_Desc & " Fail to Stop")
                Write_Memory(destSheet, "", IO_Name & "_Inp_Hand", "B R/W", "0", IO_Desc & " Input Hand")
                Write_Memory(destSheet, "", IO_Name & "_String", "STR R/W", IO_Name, "")

            End If
        Next

    End Sub

    Sub Copy_Memory_Data()
        '
        'Clear MemoryData sheet

        Dim ws As Worksheet
        Dim RowCount As Integer

        If WS_Exists("MemoryData") Then
            ws = XLpicsWB.Sheets("MemoryData")
            ws.Range("A2:F9999").Clear()
        End If

        'Copy AIn Memory Data
        If WS_Exists("IOMem - AddIn") Then
            ws = XLpicsWB.Sheets("IOMem - AIn")
            RowCount = ws.Cells(ws.Rows.Count, "B").End(XlDirection.xlUp).Row
            If RowCount > 1 Then
                ws.Range("B2:F" & RowCount).Copy()

                'Paste data into MemoryData sheet
                ws = XLpicsWB.Sheets("MemoryData")
                ws.Range("A2").PasteSpecial(XlPasteType.xlPasteValues)
            End If
        End If

        'Copy DIn Memory Data
        If WS_Exists("IOMem - DIn") Then
            ws = XLpicsWB.Sheets("IOMem - DIn")
            RowCount = ws.Cells(ws.Rows.Count, "B").End(XlDirection.xlUp).Row
            If RowCount > 1 Then
                ws.Range("B2:F" & RowCount).Copy()

                'Paste data into MemoryData sheet
                ws = XLpicsWB.Sheets("MemoryData")
                Dim MemRowCount As Integer = ws.Cells(ws.Rows.Count, "B").End(XlDirection.xlUp).Row
                Dim MemRow As Integer = MemRowCount + 1
                ws.Range("A" & MemRow).PasteSpecial(XlPasteType.xlPasteValues)
            End If

        End If

        'Copy ValveC Memory Data
        If WS_Exists("IOMem - ValveC") Then
            ws = XLpicsWB.Sheets("IOMem - ValveC")
            RowCount = ws.Cells(ws.Rows.Count, "B").End(XlDirection.xlUp).Row
            If RowCount > 1 Then
                ws.Range("B2:F" & RowCount).Copy()

                'Paste data into MemoryData sheet
                ws = XLpicsWB.Sheets("MemoryData")
                Dim MemRowCount As Integer = ws.Cells(ws.Rows.Count, "B").End(XlDirection.xlUp).Row
                Dim MemRow As Integer = MemRowCount + 1
                ws.Range("A" & MemRow).PasteSpecial(XlPasteType.xlPasteValues)
            End If
        End If

        'Copy ValveMO Memory Data
        If WS_Exists("IOMem - ValveMO") Then
            ws = XLpicsWB.Sheets("IOMem - ValveMO")
            RowCount = ws.Cells(ws.Rows.Count, "B").End(XlDirection.xlUp).Row
            If RowCount > 1 Then
                ws.Range("B2:F" & RowCount).Copy()

                'Paste data into MemoryData sheet
                ws = XLpicsWB.Sheets("MemoryData")
                Dim MemRowCount As Integer = ws.Cells(ws.Rows.Count, "B").End(XlDirection.xlUp).Row
                Dim MemRow As Integer = MemRowCount + 1
                ws.Range("A" & MemRow).PasteSpecial(XlPasteType.xlPasteValues)
            End If
        End If

        'Copy ValveSO Memory Data
        If WS_Exists("IOMem - ValveSO") Then
            ws = XLpicsWB.Sheets("IOMem - ValveSO")
            RowCount = ws.Cells(ws.Rows.Count, "B").End(XlDirection.xlUp).Row
            If RowCount > 1 Then
                ws.Range("B2:F" & RowCount).Copy()

                'Paste data into MemoryData sheet
                ws = XLpicsWB.Sheets("MemoryData")
                Dim MemRowCount As Integer = ws.Cells(ws.Rows.Count, "B").End(XlDirection.xlUp).Row
                Dim MemRow As Integer = MemRowCount + 1
                ws.Range("A" & MemRow).PasteSpecial(XlPasteType.xlPasteValues)
            End If
        End If

        'Copy Motor Memory Data
        If WS_Exists("IOMem - Motor") Then
            ws = XLpicsWB.Sheets("IOMem - Motor")
            RowCount = ws.Cells(ws.Rows.Count, "B").End(XlDirection.xlUp).Row
            If RowCount > 1 Then
                ws.Range("B2:F" & RowCount).Copy()

                'Paste data into MemoryData sheet
                ws = XLpicsWB.Sheets("MemoryData")
                Dim MemRowCount As Integer = ws.Cells(ws.Rows.Count, "B").End(XlDirection.xlUp).Row
                Dim MemRow As Integer = MemRowCount + 1
                ws.Range("A" & MemRow).PasteSpecial(XlPasteType.xlPasteValues)
            End If
        End If

        'Copy VSD Memory Data
        If WS_Exists("IOMem - VSD") Then
            ws = XLpicsWB.Sheets("IOMem - VSD")
            RowCount = ws.Cells(ws.Rows.Count, "B").End(XlDirection.xlUp).Row
            If RowCount > 1 Then
                ws.Range("B2:F" & RowCount).Copy()
                ws.Range("A1").Select()

                'Paste data into MemoryData sheet
                ws = XLpicsWB.Sheets("MemoryData")
                Dim MemRowCount As Integer = ws.Cells(ws.Rows.Count, "B").End(XlDirection.xlUp).Row
                Dim MemRow As Integer = MemRowCount + 1
                ws.Range("A" & MemRow).PasteSpecialPaste(XlPasteType.xlPasteValues)
            End If
        End If

    End Sub

    Sub Remove_From_Descriptions()
        '
        Dim ws As Worksheet
        Dim Keyword As String

        ' Remove keywords from ValveC Memory
        If WS_Exists("IOMem - ValveC") Then
            ws = XLpicsWB.Sheets("IOMem - ValveC")
            For i = 10 To 17    'Data is in rows 10 to 17
                Keyword = XLpicsWB.Sheets("Instructions").Range("C" & i).Cells.Value
                If Keyword <> "" Then
                    ws.Columns("F").Replace(What:=" " & Keyword, Replacement:="",
                            LookAt:=XlLookAt.xlPart, SearchOrder:=XlSearchOrder.xlByRows,
                            MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False)
                End If
            Next
        End If

        ' Remove keywords from ValveSO_MO Memory
        If WS_Exists("IOMem - ValveMO") Then
            ws = XLpicsWB.Sheets("IOMem - ValveMO").Select
            For i = 10 To 17    'Data is in rows 10 to 17
                Keyword = XLpicsWB.Sheets("Instructions").Range("D" & i).Cells.Value
                If Keyword <> "" Then
                    ws.Columns("F").Replace(What:=" " & Keyword, Replacement:="",
                            LookAt:=XlLookAt.xlPart, SearchOrder:=XlSearchOrder.xlByRows,
                            MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False)
                End If
            Next
        End If

        ' Remove keywords from ValveSO_MO Memory
        If WS_Exists("IOMem - ValveSO") Then
            ws = XLpicsWB.Sheets("IOMem - ValveSO").Select
            For i = 10 To 17    'Data is in rows 10 to 17
                Keyword = XLpicsWB.Sheets("Instructions").Range("D" & i).Cells.Value
                If Keyword <> "" Then
                    ws.Columns("F").Replace(What:=" " & Keyword, Replacement:="",
                            LookAt:=XlLookAt.xlPart, SearchOrder:=XlSearchOrder.xlByRows,
                            MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False)
                End If
            Next
        End If

        ' Remove keywords from Motor Memory
        If WS_Exists("IOMem - Motor") Then
            ws = XLpicsWB.Sheets("IOMem - Motor").Select
            For i = 10 To 17    'Data is in rows 10 to 17
                Keyword = XLpicsWB.Sheets("Instructions").Range("E" & i).Cells.Value
                If Keyword <> "" Then
                    ws.Columns("F").Replace(What:=" " & Keyword, Replacement:="",
                            LookAt:=XlLookAt.xlPart, SearchOrder:=XlSearchOrder.xlByRows,
                            MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False)
                End If
            Next
        End If

        ' Remove keywords from VSD Memory
        If WS_Exists("IOMem - VSD") Then
            ws = XLpicsWB.Sheets("IOMem - VSD").Select
            For i = 10 To 17    'Data is in rows 10 to 17
                Keyword = XLpicsWB.Sheets("Instructions").Range("F" & i).Cells.Value
                If Keyword <> "" Then
                    ws.Columns("F").Replace(What:=" " & Keyword, Replacement:="",
                        LookAt:=XlLookAt.xlPart, SearchOrder:=XlSearchOrder.xlByRows,
                        MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False)
                End If
            Next
        End If

    End Sub

End Module
