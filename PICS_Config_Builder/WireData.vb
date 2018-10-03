Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Module WireData

    Public Template_Name As String

    Sub Generate_Wire_Data()

        Template_Name = "Object"

        ' build PICS simulator wire sheets from templates
        Call Create_Basic_Wire_Sheets("AIn", "*_Inp_?V")
        Call Create_Basic_Wire_Sheets("DIn", "*_Inp_PV")
        Call Create_Basic_Wire_Sheets("Motor", "*_Out_Run")
        Call Create_Basic_Wire_Sheets("ValveC", "*_Out_CV")
        Call Create_Basic_Wire_Sheets("ValveMO", "*_Out_Open")
        Call Create_Basic_Wire_Sheets("ValveSO", "*_Out")
        Call Create_Basic_Wire_Sheets("VSD", "*_Out_SpeedRef")

        Call Color_Wire_Tabs()

    End Sub

    Sub Create_Basic_Wire_Sheets(ByVal typeStr As String, ByVal countStr As String)

        Dim shtTemplate As String = "Wire_" & typeStr & " Template"
        Dim SourceSheet As String = "IOTags - " & typeStr
        Dim minMaxSheet As String = "MinMax - " & typeStr
        Dim minMax As Boolean = WS_Exists(minMaxSheet)
        Dim ws As Worksheet = XLpicsWB.Sheets(SourceSheet)
        Dim RowGap As Integer = GetRowGap(shtTemplate)
        Dim rng As Range = ws.Range("A:A")
        Dim ItemCount As Integer = ws.Application.WorksheetFunction.CountIf(rng, countStr)
        Dim maxItemCount As Integer = GetMaxItems(shtTemplate)
        Dim ReqSheets As Integer = ws.Application.WorksheetFunction.RoundUp(ItemCount / maxItemCount, 0)
        Dim itemRng As Range = ws.Range("A1")
        Dim InMinCol, InMaxCol, OutMinCol, OutMaxCol As Integer
        Dim EUMinCol, EUMaxCol, RawMinCol, RawMaxCol, RawFltCol As Integer

        Call Delete_Wire_Sheets(shtTemplate)

        If minMax Then

            InMinCol = Find_Column(minMaxSheet, "InputMin")
            InMaxCol = Find_Column(minMaxSheet, "InputMax")
            OutMinCol = Find_Column(minMaxSheet, "OutputMin")
            OutMaxCol = Find_Column(minMaxSheet, "OutputMax")

            If typeStr = "AIn" Then

                EUMinCol = Find_Column(shtTemplate, "iAI_EU_Min") + 1
                EUMaxCol = Find_Column(shtTemplate, "iAI_EU_Max") + 1
                RawMinCol = Find_Column(shtTemplate, "iAI_Raw_Min") + 1
                RawMaxCol = Find_Column(shtTemplate, "iAI_Raw_Max") + 1
                RawFltCol = Find_Column(shtTemplate, "iAI_Raw_Flt") + 1

            End If

        End If

        For shtIndex = 1 To ReqSheets

            Dim NewSheetName As String = Strings.Replace(shtTemplate, " Template", "_") & shtIndex

            XLpicsWB.Sheets(shtTemplate).Copy(Before:=XLpicsWB.Sheets("Wire_AIn Template"))
            XLpicsWB.ActiveSheet.Unprotect
            XLpicsWB.ActiveSheet.Name = NewSheetName
            ws = XLpicsWB.ActiveSheet

            ws.Range("A1").Cells.Value = "_Wire_" & typeStr & "_" & shtIndex

            For i = 1 To maxItemCount

                Dim nextRng As Range = ws.Range("A:A").Find(What:=countStr, After:=itemRng)     ' if nexRng is Nothing - skip loop

                If Not (nextRng Is Nothing) Then        ' c2-10-02-2018: condition statement added since range expression netRng.Row is nothing will crash loop
                    ' If there is another item to add
                    If nextRng.Row > itemRng.Row Then

                        Dim itemNum As String
                        itemNum = Right("00" & i, 2)

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

                        ws.Cells.Replace(What:=Template_Name & itemNum, Replacement:=TagName, LookAt:=XlLookAt.xlPart,
                                             SearchOrder:=XlSearchOrder.xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False)

                        If minMax Then
                            If typeStr = "AIn" Then

                                Dim EU_Min, EU_Max, Raw_Min, Raw_Max, Raw_Flt As Double

                                Dim CurrentRow As Integer = RowGap * (i - 1) + 1
                                Dim ItemRow As Excel.Range = XLpicsWB.Sheets(minMaxSheet).Range("A:A").Find(nextRng.Value).Row
                                EU_Min = XLpicsWB.Sheets(minMaxSheet).Cells(ItemRow, InMinCol).Value
                                EU_Max = XLpicsWB.Sheets(minMaxSheet).Cells(ItemRow, InMaxCol).Value
                                Raw_Min = XLpicsWB.Sheets(minMaxSheet).Cells(ItemRow, OutMinCol).Value
                                Raw_Max = XLpicsWB.Sheets(minMaxSheet).Cells(ItemRow, OutMaxCol).Value
                                Raw_Flt = 0.8 * Raw_Min

                                If Raw_Max > 0 Then
                                    ws.Cells(CurrentRow, EUMinCol).Cells.Value = EU_Min
                                    ws.Cells(CurrentRow, EUMaxCol).Cells.Value = EU_Max
                                    ws.Cells(CurrentRow, RawMinCol).Cells.Value = Raw_Min
                                    ws.Cells(CurrentRow, RawMaxCol).Cells.Value = Raw_Max
                                    ws.Cells(CurrentRow, RawFltCol).Cells.Value = Raw_Flt
                                End If

                            End If
                        End If

                        itemRng = nextRng.Offset(1, 0)

                    End If

                End If      'added for nextRng Is Nothing

            Next i

            Call ValidateOPC1(NewSheetName)
            Call ReplaceOPC1(NewSheetName)

        Next shtIndex

    End Sub

    Private Function GetRowGap(ByVal sht As String) As Integer

        Dim Count As Integer = 0
        Dim ws As Worksheet

        ws = XLpicsWB.Sheets(sht)

        Dim itemRng As Range
        itemRng = ws.Range("B1")

        Do While ExtractNumber(itemRng.Value) = 1
            itemRng = itemRng.Offset(1, 0)
            Count = Count + 1
        Loop

        Return Count

    End Function

    Private Function GetMaxItems(ByVal sheet As String) As Integer

        Dim toSub As String
        Dim lastRng As Range

        lastRng = XLpicsWB.Sheets(sheet).Range("B1").End(XlDirection.xlDown)
        toSub = lastRng.Value

        Return ExtractNumber(toSub)

    End Function

    Private Function ExtractNumber(ByVal str As String) As Integer

        Dim length As Integer
        length = 1

        Do While IsNumeric(Right(str, length))
            length = length + 1
        Loop

        Return CInt(Right(str, length - 1))

    End Function

    Private Function Find_Column(ByRef sheet As String, ByVal str As String) As Integer

        Return XLpicsWB.Sheets(sheet).Range("A:ZZ").Find(str, LookAt:=XlLookAt.xlPart,
                                                               SearchOrder:=XlSearchOrder.xlByRows,
                                                               SearchDirection:=XlSearchDirection.xlNext).Column

    End Function

    Sub ValidateOPC1(ByRef sheetStr As String)
        '
        ' Pre-Pass OPC1 tags for existence
        ' Checks against SimData sheet for existence of OPC1 tag that it is looking to use

        Dim output, parseArr(), parse, checkStr, tag As String
        Dim searchRng As Range, firstMatch As Range = Nothing
        Dim ws As Worksheet = XLpicsWB.Sheets(sheetStr)
        Dim oRE As New RegExp
        Dim oMatch As Object

        With oRE     ' routine RegExp is a concise and flexible notation for finding and replacing patterns of text
            .Global = False
            .Multiline = False
            .IgnoreCase = False
            .Pattern = "OPC1\.\w*"
        End With

        searchRng = ws.Range("A1")

        Do While Not searchRng Is Nothing

            searchRng = ws.Range("A:ZZ").Find(What:="OPC1.", After:=searchRng, LookAt:=XlLookAt.xlPart)
            parseArr = Nothing
            parse = searchRng.Value
            output = ""

            ' Check if it needs to be split
            If InStr(parse, "|") > 0 Then
                parseArr = parse.Split("|")
            Else
                parseArr = {parse}
            End If

            For Each checkStr In parseArr
                oMatch = oRE.Execute(checkStr)
                tag = oMatch.ToString
                tag = Replace(tag, "OPC1.", "")

                If Is_Sim_Data(tag) Then
                    output = checkStr
                    Exit For
                End If

            Next checkStr

            searchRng.Value = output

            If firstMatch Is Nothing Then
                firstMatch = searchRng
            ElseIf searchRng.Value = firstMatch.Value Then
                searchRng = Nothing
            End If

        Loop  'searchRng      

    End Sub

    Sub ReplaceOPC1(ByRef NewSheetName As String)

        ' Replaces OPC1. in a sheet with the CPU_Name
        '
        Dim ws As Worksheet

        ws = XLpicsWB.Sheets(NewSheetName)
        ws.UsedRange.Select()

        ws.Cells.Replace(What:="OPC1.", Replacement:=CPU_Name & ".", LookAt:=XlLookAt.xlPart, SearchOrder:=XlSearchOrder.xlByRows,
                            MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False)

    End Sub

    Private Function Is_Sim_Data(ByRef str As String) As Boolean

        Dim ws As Worksheet = XLpicsWB.Sheets("SimData")

        Dim searchRng As Range
        searchRng = ws.Range("A:A").Find(str)

        Return Not (searchRng Is Nothing)

    End Function

    Sub Color_Wire_Tabs()

        ' assign a color to each wire or template worksheet tabs
        Dim ws As Worksheet

        For Each ws In XLpicsWB.Worksheets
            If InStr(ws.Name, "Wire") > 0 Then
                XLpicsWB.Sheets(ws.Name).Tab.ColorIndex = 24
                If InStr(ws.Name, "Template") > 0 Then
                    XLpicsWB.Sheets(ws.Name).Tab.ColorIndex = 49
                End If
            End If

        Next ws

    End Sub

    Sub Delete_Wire_Sheets(ByVal shtTemplate As String)

        Dim shtName As String
        Dim ws As Worksheet

        For i = 1 To 100     'Increase this value if there are ever more than 15 Wire Sheets
            shtName = Replace(shtTemplate, " Template", "_") & i

            ' Delete existing sheets if necessary
            For Each ws In XLpicsWB.Worksheets
                If ws.Name = shtName Then
                    XLpicsWB.Application.DisplayAlerts = False
                    ws.Delete()
                    XLpicsWB.Application.DisplayAlerts = True
                End If

            Next ws

        Next i

    End Sub

    Sub Export_Wire_Data(ByRef wb As Workbook, ByRef outFolder As String)

        Dim savePath As String

        wb.Application.DisplayAlerts = False

        savePath = outFolder & "\"

        Save_Sheets("AIn", savePath)
        Save_Sheets("DIn", savePath)
        Save_Sheets("Motor", savePath)
        Save_Sheets("ValveC", savePath)
        Save_Sheets("ValveMO", savePath)
        Save_Sheets("ValveSO", savePath)
        Save_Sheets("VSD", savePath)

        wb.Application.DisplayAlerts = True

    End Sub

    Private Sub Save_Sheets(ByVal typeStr As String, ByRef savePath As String)

        Dim i As Integer = 1
        Dim wsName As String = "Wire_" & typeStr & "_" & i
        Dim wb, newBook As Workbook
        Dim ws, newSht As Worksheet

        wb = XLpicsWB

        Do While WS_Exists(wsName)

            newBook = XLApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet)  ' create new workbook 
            ws = wb.Sheets(wsName)      ' select worksheet to copy to new workbook
            ws.Copy(Before:=newBook.Sheets(1))
            newSht = newBook.Sheets(1)
            newSht.Delete()
            newSht = newBook.ActiveSheet
            newSht.Name = wsName        ' rename new worksheet in new workbook

            newBook.SaveAs(savePath & wsName & ".wir", XlFileFormat.xlTextWindows)
            newBook.Close(False)

            i = i + 1
            wsName = "Wire_" & typeStr & "_" & i

        Loop

    End Sub

End Module
