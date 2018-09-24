
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Module MemoryData

    Const xlPasteValues As Integer = Microsoft.Office.Interop.Excel.XlPasteType.xlPasteValues

    Sub Generate_Memory_Data(ByRef wb As Workbook)

        Dim ws As Worksheet

        Call Unhide_All_Sheets(wb)

        Call Clear_Sheet_Type(wb, "MemoryData")
        Call Clear_Sheet_Type(wb, "IOMem")

        Call Generate_AI_Memory(wb, "IOTags - AIn", "IOMem - AIn")
        Call Generate_DI_Memory(wb, "IOTags - DIn", "IOMem - DIn")
        Call Generate_ValvesC_Memory(wb, "IOTags - ValveC", "IOMem - ValveC")
        Call Generate_ValvesMO_Memory(wb, "IOTags - ValveMO", "IOMem - ValveMO")
        Call Generate_Motor_Memory(wb, "IOTags - Motor", "IOMem - Motor")
        Call Generate_VSD_Memory(wb, "IOTags - VSD", "IOMem - VSD")

        Call Remove_From_Descriptions(wb)
        Call Rem_Spaces(wb, "IOMem - AIn", "F")
        Call Rem_Spaces(wb, "IOMem - DIn", "F")
        Call Rem_Spaces(wb, "IOMem - ValveC", "F")
        Call Rem_Spaces(wb, "IOMem - ValveMO", "F")
        Call Rem_Spaces(wb, "IOMem - ValveSO", "F")
        Call Rem_Spaces(wb, "IOMem - Motor", "F")
        Call Rem_Spaces(wb, "IOMem - VSD", "F")

        Call Copy_Memory_Data(wb)

        ws = CType(wb.Sheets("Instructions"), Worksheet)

        Call Hide_Sheets(wb)

    End Sub

    Private Sub Write_Memory(ByRef wrkBook As Workbook, ByVal destSheet As String, Optional ByVal inNum As String = "", Optional ByVal inName As String = "",
                             Optional ByVal inType As String = "", Optional ByVal inVal As String = "", Optional ByVal inDesc As String = "")

        Dim RowCount As Integer
        Dim XLWorkBook As Workbook = wrkBook
        Dim Wrksheet As Worksheet = XLWorkBook.Sheets(destSheet).Select

        RowCount = Wrksheet.Cells(Wrksheet.Rows.Count, "B").End.xlUp.Row
        Wrksheet.Range("A" & RowCount + 1).Select()

        Wrksheet.Range("A" & RowCount + 1).Cells.Value = inNum
        Wrksheet.Range("B" & RowCount + 1).Cells.Value = inName
        Wrksheet.Range("C" & RowCount + 1).Cells.Value = inType
        Wrksheet.Range("D" & RowCount + 1).Cells.Value = inVal
        Wrksheet.Range("F" & RowCount + 1).Cells.Value = inDesc

    End Sub

    Sub Generate_AI_Memory(ByRef wrkBook As Workbook, sourceSheet As String, destSheet As String)
        '
        '   Generate AI Memory
        Dim IO_Number As String
        Dim IO_Name As String
        Dim IO_Type As String
        Dim IO_Val As String
        Dim IO_Desc As String
        Dim SourceRowCount As Integer
        Dim XLWorkBook As Workbook = wrkBook
        Dim Wrksheet As Worksheet = XLWorkBook.Sheets(sourceSheet).Select

        SourceRowCount = Wrksheet.Cells(Wrksheet.Rows.Count, "A").End.xlUp.Row

        IO_Number = 0

        For i = 2 To SourceRowCount

            IO_Name = Wrksheet.Range("A" & i).Cells.Value
            IO_Type = Wrksheet.Range("B" & i).Cells.Value
            IO_Val = Wrksheet.Range("C" & i).Cells.Value
            IO_Desc = Wrksheet.Range("E" & i).Cells.Value

            IO_Name = Replace(IO_Name, "_Inp_PV", "")
            IO_Name = Replace(IO_Name, "_Inp_AV", "")

            If InStr(IO_Name, "Flt") = False Then
                IO_Number = IO_Number + 1
                ' Write lines
                Write_Memory(wrkBook, destSheet, IO_Number, IO_Name, IO_Type, IO_Val, IO_Desc)
                Write_Memory(wrkBook, destSheet, "", IO_Name & "_Flt", "B R/W", "0", IO_Desc & " IO Fault")
                Write_Memory(wrkBook, destSheet, "", IO_Name & "_OR", "F R/W", "0", IO_Desc & " Override Level")
                Write_Memory(wrkBook, destSheet, "", IO_Name & "_OR_EN", "B R/W", "0", IO_Desc & " Override Enable")
                Write_Memory(wrkBook, destSheet, "", IO_Name & "_PV_DB", "F R/W", "0.025", IO_Desc & " Noise Level")
                Write_Memory(wrkBook, destSheet, "", IO_Name & "_PV_EN", "B R/W", "0", IO_Desc & " Noise Enable Bit")
                Write_Memory(wrkBook, destSheet, "", IO_Name & "_String", "STR R/W", IO_Name, "")
            End If
        Next

    End Sub
    Sub Generate_DI_Memory(ByRef wrkBook As Workbook, sourceSheet As String, destSheet As String)
        '
        '   Generate DI Memory
        Dim IO_Number As String
        Dim IO_Name As String
        Dim IO_Type As String
        Dim IO_Val As String
        Dim IO_Desc As String
        Dim SourceRowCount As Integer
        Dim XLWorkBook As Workbook = wrkBook
        Dim Wrksheet As Worksheet = XLWorkBook.Sheets(sourceSheet).Select

        SourceRowCount = Wrksheet.Cells(Wrksheet.Rows.Count, "A").End.xlUp.Row

        IO_Number = 0

        For i = 2 To SourceRowCount
            IO_Name = Wrksheet.Range("A" & i).Cells.Value
            IO_Type = Wrksheet.Range("B" & i).Cells.Value
            IO_Val = Wrksheet.Range("C" & i).Cells.Value
            IO_Desc = Wrksheet.Range("E" & i).Cells.Value

            IO_Name = Replace(IO_Name, "_Inp_PV", "")

            If InStr(IO_Name, "Flt") = False Then
                IO_Number = IO_Number + 1

                ' Write Lines
                Write_Memory(wrkBook, destSheet, IO_Number, IO_Name, IO_Type, IO_Val, IO_Desc)
                Write_Memory(wrkBook, destSheet, "", IO_Name & "_Flt", "B R/W", "0", IO_Desc & " IO Fault")
                Write_Memory(wrkBook, destSheet, "", IO_Name & "_String", "STR R/W", IO_Name, "")

            End If
        Next

    End Sub
    Sub Generate_ValvesC_Memory(ByRef wrkBook As Workbook, sourceSheet As String, destSheet As String)
        '
        '
        '   Generate ValvesC Memory
        Dim IO_Number As String
        Dim IO_Name As String
        Dim IO_Type As String
        Dim IO_Val As String
        Dim IO_Desc As String
        Dim SourceRowCount As Integer
        Dim XLWorkBook As Workbook = wrkBook
        Dim Wrksheet As Worksheet = XLWorkBook.Sheets(sourceSheet).Select

        SourceRowCount = Wrksheet.Cells(Wrksheet.Rows.Count, "A").End.xlUp.Row

        IO_Number = 0

        For i = 2 To SourceRowCount
            IO_Name = Wrksheet.Range("A" & i).Cells.Value
            IO_Type = Wrksheet.Range("B" & i).Cells.Value
            IO_Val = Wrksheet.Range("C" & i).Cells.Value
            IO_Desc = Wrksheet.Range("E" & i).Cells.Value

            IO_Name = Replace(IO_Name, "_Out_CV", "")

            If IO_Type = "F R" Then
                IO_Number = IO_Number + 1

                ' Write lines
                Write_Memory(wrkBook, destSheet, IO_Number, IO_Name & "_Fbk_Flt", "B R/W", "0", IO_Desc & " Feedback Fault")
                Write_Memory(wrkBook, destSheet, "", IO_Name & "_OR", "F R/W", "0", IO_Desc & " Override Level")
                Write_Memory(wrkBook, destSheet, "", IO_Name & "_OR_EN", "B R/W", "0", IO_Desc & " Override Enable Bit")
                Write_Memory(wrkBook, destSheet, "", IO_Name & "_String", "STR R/W", IO_Name, "")

            End If

        Next

    End Sub
    Sub Generate_ValvesMO_Memory(ByRef wrkBook As Workbook, sourceSheet As String, destSheet As String)
        '
        '
        '   Generate ValvesSO Memory
        Dim IO_Number As String
        Dim IO_Name As String
        Dim IO_Type As String
        Dim IO_Val As String
        Dim IO_Desc As String
        Dim SourceRowCount As Integer
        Dim XLWorkBook As Workbook = wrkBook
        Dim Wrksheet As Worksheet = XLWorkBook.Sheets(sourceSheet).Select

        SourceRowCount = Wrksheet.Cells(Wrksheet.Rows.Count, "A").End.xlUp.Row

        IO_Number = 0

        For i = 2 To SourceRowCount
            IO_Name = Wrksheet.Range("A" & i).Cells.Value
            IO_Type = Wrksheet.Range("B" & i).Cells.Value
            IO_Val = Wrksheet.Range("C" & i).Cells.Value
            IO_Desc = Wrksheet.Range("E" & i).Cells.Value

            IO_Name = Replace(IO_Name, "_Out", "")

            If IO_Type = "B R" Then
                IO_Number = IO_Number + 1

                ' Write lines
                Write_Memory(wrkBook, destSheet, IO_Number, IO_Name & "_FTC", "B R/W", "0", IO_Desc & " Fail to Close")
                Write_Memory(Wrksheet, destSheet, "", IO_Name & "_FTO", "B R/W", "0", IO_Desc & " Fail to Open")
                Write_Memory(Wrksheet, destSheet, "", IO_Name & "_Stuck", "B R/W", "0", IO_Desc & " Is Stuck")
                Write_Memory(Wrksheet, destSheet, "", IO_Name & "_Inp_ActuatorFault", "B R/W", "0", IO_Desc & " Act Fault")
                Write_Memory(Wrksheet, destSheet, "", IO_Name & "_Inp_Hand", "B R/W", "0", IO_Desc & " Input Hand")
                Write_Memory(Wrksheet, destSheet, "", IO_Name & "_String", "STR R/W", IO_Name, "")

            End If

        Next

    End Sub

    Sub Generate_ValvesSO_Memory(ByRef wrkBook As Workbook, sourceSheet As String, destSheet As String)
        '
        '   Generate ValvesSO Memory
        Dim IO_Number As String
        Dim IO_Name As String
        Dim IO_Type As String
        Dim IO_Val As String
        Dim IO_Desc As String
        Dim SourceRowCount As Integer
        Dim XLWorkBook As Workbook = wrkBook
        Dim Wrksheet As Worksheet = XLWorkBook.Sheets(sourceSheet).Select

        SourceRowCount = Wrksheet.Cells(Wrksheet.Rows.Count, "A").End.xlUp.Row

        IO_Number = 0

        For i = 2 To SourceRowCount
            IO_Name = Wrksheet.Range("A" & i).Cells.Value
            IO_Type = Wrksheet.Range("B" & i).Cells.Value
            IO_Val = Wrksheet.Range("C" & i).Cells.Value
            IO_Desc = Wrksheet.Range("E" & i).Cells.Value

            IO_Name = Replace(IO_Name, "_Out", "")

            If IO_Type = "B R" Then
                IO_Number = IO_Number + 1

                ' Write lines
                Write_Memory(wrkBook, destSheet, IO_Number, IO_Name & "_FTC", "B R/W", "0", IO_Desc & " Fail to Close")
                Write_Memory(wrkBook, destSheet, "", IO_Name & "_FTO", "B R/W", "0", IO_Desc & " Fail to Open")
                Write_Memory(wrkBook, destSheet, "", IO_Name & "_Inp_Hand", "B R/W", "0", IO_Desc & " Input Hand")
                Write_Memory(wrkBook, destSheet, "", IO_Name & "_Stuck", "B R/W", "0", IO_Desc & " Is Stuck")
                Write_Memory(wrkBook, destSheet, "", IO_Name & "_String", "STR R/W", IO_Name, "")

            End If

        Next

    End Sub

    Sub Generate_Motor_Memory(ByRef wrkBook As Workbook, sourceSheet As String, destSheet As String)
        '
        '   Generate Motor Memory
        Dim IO_Number As String
        Dim IO_Name As String
        Dim IO_Type As String
        Dim IO_Val As String
        Dim IO_Desc As String
        Dim SourceRowCount As Integer
        Dim XLWorkBook As Workbook = wrkBook
        Dim Wrksheet As Worksheet = XLWorkBook.Sheets(sourceSheet).Select

        SourceRowCount = Wrksheet.Cells(Wrksheet.Rows.Count, "A").End.xlUp.Row

        IO_Number = 0

        For i = 2 To SourceRowCount
            IO_Name = Wrksheet.Range("A" & i).Cells.Value
            IO_Type = Wrksheet.Range("B" & i).Cells.Value
            IO_Val = Wrksheet.Range("C" & i).Cells.Value
            IO_Desc = Wrksheet.Range("E" & i).Cells.Value

            IO_Name = Replace(IO_Name, "_Out_Run", "")

            If IO_Type = "B R" Then
                IO_Number = IO_Number + 1

                ' Write lines
                Write_Memory(wrkBook, destSheet, IO_Number, IO_Name & "_Inp_Faulted", "B R/W", "0", IO_Desc & " Faulted")
                Write_Memory(wrkBook, destSheet, "", IO_Name & "_FTR", "B R/W", "0", IO_Desc & " Fail to Run")
                Write_Memory(wrkBook, destSheet, "", IO_Name & "_FTS", "B R/W", "0", IO_Desc & " Fail to Stop")
                Write_Memory(wrkBook, destSheet, "", IO_Name & "_Inp_Hand", "B R/W", "0", IO_Desc & " Input Hand")
                Write_Memory(wrkBook, destSheet, "", IO_Name & "_OverLoad", "B R/W", "0", IO_Desc & " OverLoad")
                Write_Memory(wrkBook, destSheet, "", IO_Name & "_String", "STR R/W", IO_Name, "")

            End If
        Next

    End Sub
    Sub Generate_VSD_Memory(ByRef wrkBook As Workbook, sourceSheet As String, destSheet As String)
        '
        '   Generate Motor Memory
        Dim IO_Number As String
        Dim IO_Name As String
        Dim IO_Type As String
        Dim IO_Val As String
        Dim IO_Desc As String
        Dim SourceRowCount As Integer
        Dim XLWorkBook As Workbook = wrkBook
        Dim Wrksheet As Worksheet = XLWorkBook.Sheets(sourceSheet).Select

        SourceRowCount = Wrksheet.Cells(Wrksheet.Rows.Count, "A").End.xlUp.Row

        IO_Number = 0

        For i = 2 To SourceRowCount
            IO_Name = Wrksheet.Range("A" & i).Cells.Value
            IO_Type = Wrksheet.Range("B" & i).Cells.Value
            IO_Val = Wrksheet.Range("C" & i).Cells.Value
            IO_Desc = Wrksheet.Range("E" & i).Cells.Value

            IO_Name = Replace(IO_Name, "_Out_Run", "")

            If IO_Type = "B R" Then
                IO_Number = IO_Number + 1

                ' Write lines
                Write_Memory(wrkBook, destSheet, IO_Number, IO_Name & "_Inp_Faulted", "B R/W", "0", IO_Desc & " Faulted")
                Write_Memory(wrkBook, destSheet, "", IO_Name & "_FTR", "B R/W", "0", IO_Desc & " Fail to Run")
                Write_Memory(wrkBook, destSheet, "", IO_Name & "_FTS", "B R/W", "0", IO_Desc & " Fail to Stop")
                Write_Memory(wrkBook, destSheet, "", IO_Name & "_Inp_Hand", "B R/W", "0", IO_Desc & " Input Hand")
                Write_Memory(wrkBook, destSheet, "", IO_Name & "_String", "STR R/W", IO_Name, "")

            End If
        Next

    End Sub

    Sub Copy_Memory_Data(ByRef wrkBook As Workbook)
        '
        'Clear MemoryData sheet

        Dim XLWorkBook As Workbook = wrkBook
        Dim Wrksheet As Worksheet = XLWorkBook.Sheets("MemoryData").Select
        Wrksheet.Range("A2:F9999").Clear()

        'Copy AIn Memory Data
        Wrksheet = XLWorkBook.Sheets("IOMem - AIn").Select
        Dim RowCount As Integer = Wrksheet.Cells(Wrksheet.Rows.Count, "B").End.xlUp.Row

        If RowCount > 1 Then
            Wrksheet.Range("B2:F" & RowCount).Copy()

            'Paste data into MemoryData sheet
            Wrksheet = XLWorkBook.Sheets("MemoryData").Select
            Wrksheet.Range("A2").PasteSpecial(Paste:=xlPasteValues)
        End If

        'Copy DIn Memory Data
        Wrksheet = XLWorkBook.Sheets("IOMem - DIn").Select
        RowCount = Wrksheet.Cells(Wrksheet.Rows.Count, "B").End.xlUp.Row

        If RowCount > 1 Then
            Wrksheet.Range("B2:F" & RowCount).Copy()

            'Paste data into MemoryData sheet
            Wrksheet = XLWorkBook.Sheets("MemoryData").Select
            Dim MemRowCount As Integer = Wrksheet.Cells(Wrksheet.Rows.Count, "B").End.xlUp.Row
            Dim MemRow As Integer = MemRowCount + 1
            Wrksheet.Range("A" & MemRow).PasteSpecial(Paste:=xlPasteValues)
        End If

        'Copy ValveC Memory Data
        Wrksheet = XLWorkBook.Sheets("IOMem - ValveC").Select
        RowCount = Wrksheet.Cells(Wrksheet.Rows.Count, "B").End.xlUp.Row
        If RowCount > 1 Then
            Wrksheet.Range("B2:F" & RowCount).Copy()

            'Paste data into MemoryData sheet
            Wrksheet = XLWorkBook.Sheets("MemoryData").Select
            Dim MemRowCount As Integer = Wrksheet.Cells(Wrksheet.Rows.Count, "B").End.xlUp.Row
            Dim MemRow As Integer = MemRowCount + 1
            Wrksheet.Range("A" & MemRow).PasteSpecial(Paste:=xlPasteValues)
        End If

        'Copy ValveMO Memory Data
        Wrksheet = XLWorkBook.Sheets("IOMem - ValveMO").Select
        RowCount = Wrksheet.Cells(Wrksheet.Rows.Count, "B").End.xlUp.Row
        If RowCount > 1 Then
            Wrksheet.Range("B2:F" & RowCount).Copy()

            'Paste data into MemoryData sheet
            Wrksheet = XLWorkBook.Sheets("MemoryData").Select
            Dim MemRowCount As Integer = Wrksheet.Cells(Wrksheet.Rows.Count, "B").End.xlUp.Row
            Dim MemRow As Integer = MemRowCount + 1
            Wrksheet.Range("A" & MemRow).PasteSpecial(Paste:=xlPasteValues)
        End If

        'Copy ValveSO Memory Data
        Wrksheet = XLWorkBook.Sheets("IOMem - ValveSO").Select
        RowCount = Wrksheet.Cells(Wrksheet.Rows.Count, "B").End.xlUp.Row
        If RowCount > 1 Then
            Wrksheet.Range("B2:F" & RowCount).Copy()

            'Paste data into MemoryData sheet
            Wrksheet = XLWorkBook.Sheets("MemoryData").Select
            Dim MemRowCount As Integer = Wrksheet.Cells(Wrksheet.Rows.Count, "B").End.xlUp.Row
            Dim MemRow As Integer = MemRowCount + 1
            Wrksheet.Range("A" & MemRow).PasteSpecial(Paste:=xlPasteValues)
        End If

        'Copy Motor Memory Data
        Wrksheet = XLWorkBook.Sheets("IOMem - Motor").Select
        RowCount = Wrksheet.Cells(Wrksheet.Rows.Count, "B").End.xlUp.Row
        If RowCount > 1 Then
            Wrksheet.Range("B2:F" & RowCount).Copy()

            'Paste data into MemoryData sheet
            Wrksheet = XLWorkBook.Sheets("MemoryData").Select
            Dim MemRowCount As Integer = Wrksheet.Cells(Wrksheet.Rows.Count, "B").End.xlUp.Row
            Dim MemRow As Integer = MemRowCount + 1
            Wrksheet.Range("A" & MemRow).PasteSpecial(Paste:=xlPasteValues)
        End If

        'Copy VSD Memory Data
        Wrksheet = XLWorkBook.Sheets("IOMem - VSD").Select
        RowCount = Wrksheet.Cells(Wrksheet.Rows.Count, "B").End.xlUp.Row
        If RowCount > 1 Then
            Wrksheet.Range("B2:F" & RowCount).Copy()
            Wrksheet.Range("A1").Select()

            Wrksheet = XLWorkBook.Sheets("MemoryData").Select
            Dim MemRowCount As Integer = Wrksheet.Cells(Wrksheet.Rows.Count, "B").End.xlUp.Row
            Dim MemRow As Integer = MemRowCount + 1
            Wrksheet.Range("A" & MemRow).PasteSpecialPaste(Paste:=xlPasteValues)
        End If

        Wrksheet.Range("A1").Select()
        XLWorkBook.Application.CutCopyMode = False

    End Sub

    Sub Remove_From_Descriptions(ByRef XLWorkBook As Workbook)
        '
        Dim WrkBook As Workbook = XLWorkBook
        Dim WrkSheet As Worksheet
        Dim Keyword As String

        ' Remove keywords from ValveC Memory
        WrkSheet = WrkBook.Sheets("IOMem - ValveC").Select
        For i = 10 To 17    'Data is in rows 10 to 17
            Keyword = WrkBook.Sheets("Instructions").Range("C" & i).Cells.Value
            If Keyword <> "" Then
                WrkSheet.Columns("F").Replace(What:=" " & Keyword,
                        Replacement:="",
                        LookAt:="XlPart",
                        SearchOrder:="XlByRows",
                        MatchCase:=False,
                        SearchFormat:=False,
                        ReplaceFormat:=False)
            End If
        Next

        ' Remove keywords from ValveSO_MO Memory
        WrkSheet = WrkBook.Sheets("IOMem - ValveMO").Select
        For i = 10 To 17    'Data is in rows 10 to 17
            Keyword = WrkBook.Sheets("Instructions").Range("D" & i).Cells.Value
            If Keyword <> "" Then
                WrkSheet.Columns("F").Replace(What:=" " & Keyword,
                        Replacement:="",
                        LookAt:="xlPart",
                        SearchOrder:="xlByRows",
                        MatchCase:=False,
                        SearchFormat:=False,
                        ReplaceFormat:=False)
            End If
        Next

        ' Remove keywords from ValveSO_MO Memory
        WrkSheet = WrkBook.Sheets("IOMem - ValveSO").Select
        For i = 10 To 17    'Data is in rows 10 to 17
            Keyword = WrkBook.Sheets("Instructions").Range("D" & i).Cells.Value
            If Keyword <> "" Then
                WrkSheet.Columns("F").Replace(What:=" " & Keyword,
                        Replacement:="",
                        LookAt:="xlPart",
                        SearchOrder:="xlByRows",
                        MatchCase:=False,
                        SearchFormat:=False,
                        ReplaceFormat:=False)
            End If
        Next

        ' Remove keywords from Motor Memory
        WrkSheet = WrkBook.Sheets("IOMem - Motor").Select
        For i = 10 To 17    'Data is in rows 10 to 17
            Keyword = WrkBook.Sheets("Instructions").Range("E" & i).Cells.Value
            If Keyword <> "" Then
                WrkSheet.Columns("F").Replace(What:=" " & Keyword,
                        Replacement:="",
                        LookAt:="xlPart",
                        SearchOrder:="xlByRows",
                        MatchCase:=False,
                        SearchFormat:=False,
                        ReplaceFormat:=False)
            End If
        Next

        ' Remove keywords from VSD Memory
        WrkSheet = WrkBook.Sheets("IOMem - VSD").Select
        For i = 10 To 17    'Data is in rows 10 to 17
            Keyword = WrkBook.Sheets("Instructions").Range("F" & i).Cells.Value
            If Keyword <> "" Then
                WrkSheet.Columns("F").Replace(What:=" " & Keyword,
                        Replacement:="",
                        LookAt:="xlPart",
                        SearchOrder:="xlByRows",
                        MatchCase:=False,
                        SearchFormat:=False,
                        ReplaceFormat:=False)
            End If
        Next


    End Sub

End Module
