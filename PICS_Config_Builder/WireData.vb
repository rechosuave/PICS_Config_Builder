Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Module WireData

    Const xlPasteValues As Integer = XlPasteType.xlPasteValues
    Const xlWhole As Integer = XlLookAt.xlWhole
    Const xlPart As Integer = XlLookAt.xlPart
    Const xlByRows As Integer = XlSearchOrder.xlByRows
    Const xlNext As Integer = XlSearchDirection.xlNext
    Const xlTextWindows As Integer = XlFileFormat.xlTextWindows


    Public CPU_Name As String
    Public Template_Name As String

    Sub Generate_Wire_Data(ByRef wrkBook As Workbook)

        Call Unhide_All_Sheets(wrkBook)

        CPU_Name = Get_CPU_Name(wrkBook)
        Template_Name = "Object"

        'Call Create_Wire_AIn_Sheets
        Call Create_Basic_Wire_Sheets(wrkBook, "AIn", "*_Inp_?V")
        Call Create_Basic_Wire_Sheets(wrkBook, "DIn", "*_Inp_PV")
        Call Create_Basic_Wire_Sheets(wrkBook, "Motor", "*_Out_Run")
        Call Create_Basic_Wire_Sheets(wrkBook, "ValveC", "*_Out_CV")
        Call Create_Basic_Wire_Sheets(wrkBook, "ValveMO", "*_Out_Open")
        Call Create_Basic_Wire_Sheets(wrkBook, "ValveSO", "*_Out")
        Call Create_Basic_Wire_Sheets(wrkBook, "VSD", "*_Out_SpeedRef")

        Call Color_Wire_Tabs(wrkBook)

        wrkBook.Sheets("Instructions").Select

        Call Hide_Sheets(wrkBook)

    End Sub

    Private Function GetRowGap(ByRef wrkBook As Workbook, sheet As String) As Integer

        Dim Count As Integer
        Count = 0

        Dim wrkSht As Worksheet
        wrkSht = wrkBook.Sheets(sheet)

        Dim itemRng As Excel.Range
        itemRng = wrkSht.Range("B1")

        Do While ExtractNumber(itemRng.Value) = 1
            itemRng = itemRng.Offset(1, 0)
            Count = Count + 1
        Loop

        GetRowGap = Count

    End Function

    Private Function GetMaxItems(ByRef wrkBook As Workbook, sheet As String) As Integer

        Dim toSub As String
        Dim lastRng As Excel.Range
        lastRng = wrkBook.Sheets(sheet).Range("B1").End.xlDown
        toSub = lastRng.Value

        GetMaxItems = ExtractNumber(toSub)

    End Function

    Private Function ExtractNumber(str As String) As Integer

        Dim length As Integer
        length = 1

        Do While IsNumeric(Right(str, length))
            length = length + 1
        Loop

        ExtractNumber = CInt(Right(str, length - 1))

    End Function

    Private Function Find_Column(ByRef wrkBook As Workbook, sheet As String, str As String) As Integer

        Find_Column = wrkBook.Sheets(sheet).Range("A:ZZ").Find(str, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext).Column

    End Function

    Sub ValidateOPC1(ByRef wrkBook As Workbook, sheetStr As String)
        '
        ' Pre-Pass OPC1 tags for existance
        ' Checks against SimData sheet for existance of OPC1 tag that it is looking to use

        Dim searchRng As Excel.Range, firstMatch As Excel.Range = Nothing

        Dim wrkSht As Worksheet

        wrkSht = wrkBook.Sheets(sheetStr)

        searchRng = wrkSht.Range("A1").Value

        Do While Not searchRng Is Nothing

            searchRng = wrkSht.Range("A:ZZ").Find(What:="OPC1.", After:=searchRng, LookAt:=xlPart)

            Dim output, parseArr(), parse As String

            parseArr = Nothing

            parse = searchRng.Value
            output = ""

            ' Check if it needs to be split
            If InStr(parse, "|") > 0 Then
                parseArr = parse.Split("|")
            Else
                parseArr.SetValue(parse, 0)
            End If

            Dim checkStr As String

            For Each checkStr In parseArr

                Dim tag As String
                Dim re As New RegExp

                With re
                    .Global = False
                    .Multiline = False
                    .IgnoreCase = False
                    .Pattern = "OPC1\.\w*"
                End With

                Dim result As Object
                result = re.Execute(checkStr)

                tag = result(0)
                tag = Replace(tag, "OPC1.", "")

                If Is_Sim_Data(wrkBook, tag) Then
                    output = checkStr
                    Exit For
                End If
            Next

            searchRng.Value = output


            If firstMatch Is Nothing Then
                firstMatch = searchRng
            ElseIf searchRng.Value = firstMatch.Value Then
                searchRng = Nothing
            End If

        Loop

    End Sub

    Private Function Is_Sim_Data(ByRef wrkBook As Workbook, str As String) As Boolean

        Dim wrkSht As Worksheet = wrkBook.Sheets("SimData")

        Dim searchRng As Excel.Range
        searchRng = wrkSht.Range("A:A").Find(str)

        Is_Sim_Data = Not (searchRng Is Nothing)

    End Function

    Sub ReplaceOPC1(ByRef wrkBook As Workbook, NewSheetName As String)
        '
        ' Replaces OPC1. in a sheet with the CPU_Name
        '
        wrkBook.Sheets(NewSheetName).UsedRange.Select

        wrkBook.Sheets(NewSheetName).Cells.Replace(What:="OPC1.",
                            Replacement:=CPU_Name & ".",
                            LookAt:=xlPart,
                            SearchOrder:=xlByRows,
                            MatchCase:=False,
                            SearchFormat:=False,
                            ReplaceFormat:=False)
        wrkBook.Sheets(NewSheetName).Range("A1").Select

    End Sub

    Sub Color_Wire_Tabs(ByRef wb As Workbook)

        Dim ws As Worksheet

        For Each ws In wb.Worksheets

            If InStr(ws.Name, "Wire") > 0 Then
                wb.Sheets(ws.Name).Tab.ColorIndex = 24
                If InStr(ws.Name, "Template") > 0 Then
                    wb.Sheets(ws.Name).Tab.ColorIndex = 49
                End If
            End If

        Next ws

    End Sub

    Sub Delete_Wire_Sheets(ByRef wb As Workbook, sheetTemplate As String)

        Dim sheetName As String
        Dim ws As Worksheet

        For i = 1 To 100     'Increase this value if there are ever more than 15 Wire Sheets
            sheetName = Replace(sheetTemplate, " Template", "_") & i

            ' Deletes existing sheets if necessary
            For Each ws In wb.Worksheets
                If ws.Name = sheetName Then
                    wb.Application.DisplayAlerts = False
                    ws.Delete()
                    wb.Application.DisplayAlerts = True
                End If
            Next ws
        Next i

    End Sub

    Sub Create_Basic_Wire_Sheets(ByRef wb As Workbook, typeStr As String, countStr As String)

        Dim sheetTemplate As String = "Wire_" & typeStr & " Template"
        Dim sourceSheet As String = "IOTags - " & typeStr
        Dim minMaxSheet As String = "MinMax - " & typeStr
        Dim minMax As Boolean = Worksheet_Exists(wb, minMaxSheet)
        Dim ws As Worksheet = wb.Sheets(sourceSheet).Select
        Dim RowGap As Integer = GetRowGap(wb, sheetTemplate)
        Dim ItemCount As Integer = ws.Cells.CountIf("A:A", countStr)
        Dim maxItemCount As Integer = GetMaxItems(wb, sheetTemplate)
        Dim ReqSheets As Integer = ws.Cells.RoundUp(ItemCount / maxItemCount, 0)
        Dim itemRng As Excel.Range = ws.Range("A1")
        Dim InMinCol, InMaxCol, OutMinCol, OutMaxCol As Integer
        Dim EUMinCol, EUMaxCol, RawMinCol, RawMaxCol, RawFltCol As Integer

        Call Delete_Wire_Sheets(wb, sheetTemplate)

        If minMax Then

            InMinCol = Find_Column(wb, minMaxSheet, "InputMin")
            InMaxCol = Find_Column(wb, minMaxSheet, "InputMax")
            OutMinCol = Find_Column(wb, minMaxSheet, "OutputMin")
            OutMaxCol = Find_Column(wb, minMaxSheet, "OutputMax")

            If typeStr = "AIn" Then

                EUMinCol = Find_Column(wb, sheetTemplate, "iAI_EU_Min") + 1
                EUMaxCol = Find_Column(wb, sheetTemplate, "iAI_EU_Max") + 1
                RawMinCol = Find_Column(wb, sheetTemplate, "iAI_Raw_Min") + 1
                RawMaxCol = Find_Column(wb, sheetTemplate, "iAI_Raw_Max") + 1
                RawFltCol = Find_Column(wb, sheetTemplate, "iAI_Raw_Flt") + 1

            End If

        End If

        For shtIndex = 1 To ReqSheets

            Dim NewSheetName As String = Strings.Replace(sheetTemplate, " Template", "_") & shtIndex

            wb.Sheets(sheetTemplate).Copy(Before:=wb.Sheets("Wire_AIn Template"))
            wb.ActiveSheet.Unprotect
            wb.ActiveSheet.Name = NewSheetName
            Dim wrkSheet As Worksheet = wb.ActiveSheet.Name
            wrkSheet.Range("A1").Cells.Value = "_Wire_" & typeStr & "_" & shtIndex

            For itemIndex = 1 To maxItemCount

                Dim nextRng As Excel.Range
                nextRng = ws.Range("A:A").Find(countStr, itemRng)

                ' If there is another item to add
                If nextRng.Row > itemRng.Row Then

                    Dim itemNum As String
                    itemNum = Right("00" & itemIndex, 2)

                    Dim searchStr As String
                    searchStr = Replace(countStr, "*", "")
                    searchStr = Replace(searchStr, "?", "\w")

                    Dim re As New RegExp

                    With re
                        .Global = True
                        .Multiline = True
                        .IgnoreCase = False
                        .Pattern = searchStr
                    End With

                    Dim TagName = re.Replace(nextRng.Value, "")

                    wrkSheet.Cells.Replace(What:=Template_Name & itemNum, Replacement:=TagName, LookAt:=xlPart,
                                             SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False)

                    If minMax Then
                        If typeStr = "AIn" Then

                            Dim EU_Min, EU_Max, Raw_Min, Raw_Max, Raw_Flt As Double

                            Dim CurrentRow As Integer = RowGap * (itemIndex - 1) + 1
                            Dim ItemRow As Excel.Range = wb.Sheets(minMaxSheet).Range("A:A").Find(nextRng.Value).Row
                            EU_Min = wb.Sheets(minMaxSheet).Cells(ItemRow, InMinCol).Value
                            EU_Max = wb.Sheets(minMaxSheet).Cells(ItemRow, InMaxCol).Value
                            Raw_Min = wb.Sheets(minMaxSheet).Cells(ItemRow, OutMinCol).Value
                            Raw_Max = wb.Sheets(minMaxSheet).Cells(ItemRow, OutMaxCol).Value
                            Raw_Flt = 0.8 * Raw_Min

                            If Raw_Max > 0 Then
                                wrkSheet.Cells(CurrentRow, EUMinCol).Cells.Value = EU_Min
                                wrkSheet.Cells(CurrentRow, EUMaxCol).Cells.Value = EU_Max
                                wrkSheet.Cells(CurrentRow, RawMinCol).Cells.Value = Raw_Min
                                wrkSheet.Cells(CurrentRow, RawMaxCol).Cells.Value = Raw_Max
                                wrkSheet.Cells(CurrentRow, RawFltCol).Cells.Value = Raw_Flt
                            End If

                        End If
                    End If

                    itemRng = nextRng.Offset(1, 0)

                End If

            Next

            wrkSheet.Range("A1").Select()

            Call ValidateOPC1(wb, NewSheetName)
            Call ReplaceOPC1(wb, NewSheetName)
        Next

    End Sub

    Sub Export_Wire_Data(ByRef wb As Workbook, outFolder As String)

        Dim savePath As String

        wb.Application.DisplayAlerts = False

        savePath = outFolder & "\"

        Save_Sheets(wb, "AIn", savePath)
        Save_Sheets(wb, "DIn", savePath)
        Save_Sheets(wb, "Motor", savePath)
        Save_Sheets(wb, "ValveC", savePath)
        Save_Sheets(wb, "ValveMO", savePath)
        Save_Sheets(wb, "ValveSO", savePath)
        Save_Sheets(wb, "VSD", savePath)

        wb.Application.DisplayAlerts = True

    End Sub

    Private Sub Save_Sheets(ByRef wb As Workbook, typeStr As String, savePath As String)

        Dim i As Integer = 1
        Dim wsName As String = "Wire_" & typeStr & "_" & i

        Do While Worksheet_Exists(wb, wsName)

            Dim book As Workbook

            book = CreateObject("Excel.Application")
            Dim Count As Integer = book.Worksheets.Count
            book.Sheets(wsName).Copy(book.Worksheets(1))

            book.Worksheets(1).Range("A1").EntireRow.Insert
            book.Worksheets(1).Range("A1").Value = ";PICS for Windows - Device Wiring Export V1.10"
            book.Worksheets(1).SaveAs(savePath & wsName & ".wir", xlTextWindows)
            book.Close(False)
            book = Nothing

            i = i + 1
            wsName = "Wire_" & typeStr & "_" & i

        Loop

    End Sub

End Module
