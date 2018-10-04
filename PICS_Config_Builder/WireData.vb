Imports Microsoft.Office.Interop.Excel
Imports System.Text.RegularExpressions
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

    Sub Create_Basic_Wire_Sheets(ByVal templateType As String, ByVal countStr As String)

        Dim WireTemplate As String = "Wire_" & templateType & " Template"
        Dim IOTagsSheet As String = "IOTags - " & templateType
        Dim MinMaxSheet As String = "MinMax - " & templateType
        Dim TagItemNum, searchStr, TagName, NewWireSheetName As String
        Dim InMinCol, InMaxCol, OutMinCol, OutMaxCol As Integer
        Dim EUMinCol, EUMaxCol, RawMinCol, RawMaxCol, RawFltCol As Integer
        Dim CurrentRow As Integer, absCount As Integer = Get_Annunciator_Block_Size(WireTemplate)
        Dim MinMaxShtFound As Boolean = WS_Exists(MinMaxSheet)
        Dim maxItemCount As Integer = GetMaxItems(WireTemplate)
        Dim ws As Worksheet = XLpicsWB.Sheets(IOTagsSheet)
        Dim IOTagsRng As Range = ws.Range("A:A")
        Dim IOTagsRngStart As Range = ws.Range("A1"), IOTagsNextRng, ItemRow As Range
        Dim ItemCount As Integer = ws.Application.WorksheetFunction.CountIf(IOTagsRng, countStr)
        Dim ReqSheets As Integer = ws.Application.WorksheetFunction.RoundUp(ItemCount / maxItemCount, 0)
        Dim re As New RegExp

        Call Delete_Wire_Sheets(WireTemplate)

        If MinMaxShtFound Then  ' find columns

            InMinCol = Find_Column(MinMaxSheet, "InputMin")
            InMaxCol = Find_Column(MinMaxSheet, "InputMax")
            OutMinCol = Find_Column(MinMaxSheet, "OutputMin")
            OutMaxCol = Find_Column(MinMaxSheet, "OutputMax")

            If templateType = "AIn" Then     ' find columns

                EUMinCol = Find_Column(WireTemplate, "iAI_EU_Min") + 1
                EUMaxCol = Find_Column(WireTemplate, "iAI_EU_Max") + 1
                RawMinCol = Find_Column(WireTemplate, "iAI_Raw_Min") + 1
                RawMaxCol = Find_Column(WireTemplate, "iAI_Raw_Max") + 1
                RawFltCol = Find_Column(WireTemplate, "iAI_Raw_Flt") + 1

            End If

        End If

        For shtIndex = 1 To ReqSheets     ' add a new wire worksheet(s) to PICS workbook and populate data into worksheet

            NewWireSheetName = Strings.Replace(WireTemplate, " Template", "_") & shtIndex
            XLpicsWB.Sheets(WireTemplate).Copy(Before:=XLpicsWB.Sheets("Wire_AIn Template"))
            ws = XLpicsWB.ActiveSheet.Unprotect
            ws.Name = NewWireSheetName      ' rename copied worksheet
            ws.Range("A1").Cells.Value = "_Wire_" & templateType & "_" & shtIndex

            For i = 1 To maxItemCount   ' loop for number of annunciators on new wire worksheet to populate with tag item data

                IOTagsNextRng = XLpicsWB.Sheets(IOTagsSheet).Range("A:A").Find(What:=countStr, After:=IOTagsRngStart)     ' if nexRng is Nothing - skip loop

                If Not (IOTagsNextRng Is Nothing) Then        ' c2-10-02-2018: condition statement added since if range expression "netRng.Row" is nothing will crash loop
                    ' Is there another tag item to add to annunciator block
                    If IOTagsNextRng.Row > IOTagsRngStart.Row Then

                        TagItemNum = Right("00" & i, 2)        ' assign a IO tag item #

                        searchStr = Replace(countStr, "*", "")
                        searchStr = Replace(searchStr, "?", "\w")

                        With re
                            .Global = True
                            .Multiline = True
                            .IgnoreCase = False
                            .Pattern = searchStr
                        End With

                        TagName = re.Replace(IOTagsNextRng.Value, "")

                        ws.Cells.Replace(What:=Template_Name & TagItemNum, Replacement:=TagName, LookAt:=XlLookAt.xlPart,
                                             SearchOrder:=XlSearchOrder.xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False)

                        If MinMaxShtFound Then
                            If templateType = "AIn" Then

                                Dim EU_Min, EU_Max, Raw_Min, Raw_Max, Raw_Flt As Double

                                CurrentRow = absCount * (i - 1) + 1
                                ItemRow = XLpicsWB.Sheets(MinMaxSheet).Range("A:A").Find(IOTagsNextRng.Value).Row
                                EU_Min = XLpicsWB.Sheets(MinMaxSheet).Cells(ItemRow, InMinCol).Value
                                EU_Max = XLpicsWB.Sheets(MinMaxSheet).Cells(ItemRow, InMaxCol).Value
                                Raw_Min = XLpicsWB.Sheets(MinMaxSheet).Cells(ItemRow, OutMinCol).Value
                                Raw_Max = XLpicsWB.Sheets(MinMaxSheet).Cells(ItemRow, OutMaxCol).Value
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

                        IOTagsRngStart = IOTagsNextRng.Offset(1, 0)

                    End If

                End If      'added for nextRng Is Nothing

            Next i

            Call ValidateOPC1(NewWireSheetName)
            Call ReplaceOPC1(NewWireSheetName)

        Next shtIndex

    End Sub

    Private Function Get_Annunciator_Block_Size(ByVal WireTemplateName As String) As Integer

        ' count the number of items in the first annunciator block for the template 
        ' type (eg "Wire_type Template" such As "Wire_AIn Template")
        Dim ws As Worksheet
        Dim itemRng As Range
        Dim Count As Integer = 0

        ws = XLpicsWB.Sheets(WireTemplateName)
        itemRng = ws.Range("B1")        ' start of first annunciator block

        Do While ExtractNumber(itemRng.Value) = 1       ' find end of first annunciator block
            itemRng = itemRng.Offset(1, 0)
            Count = Count + 1
        Loop

        Return Count        ' return annunciator block size (# of items) for the template type

    End Function

    Private Function GetMaxItems(ByRef sheet As String) As Integer

        '  Count the number of Annunciator item blocks on the wire template worksheet 
        '  to populate on the new wire worksheet that will be exported to PICS simulator as .wir file
        Dim toSub As String
        Dim lastRng As Range

        lastRng = XLpicsWB.Sheets(sheet).Range("B1").End(XlDirection.xlDown)    ' read last cell in column B1 of wire template worksheet (eg "T24")
        toSub = lastRng.Value  ' pass string to extract # from string (eg "T24" -> 24)

        Return ExtractNumber(toSub)

    End Function

    Private Function ExtractNumber(ByVal str As String) As Integer

        Dim length As Integer = 1

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

    Sub ValidateOPC1(ByRef WireSht As String)

        ' Pre-Pass OPC1 tags for existence
        ' Checks against SimData sheet for existence of OPC1 tag that it is looking to use
        Dim output, parseArr(), parse, checkStr, tag As String
        Dim searchRng As Range, firstMatch As Range = Nothing
        Dim ws As Worksheet = XLpicsWB.Sheets(WireSht)
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

    Sub ReplaceOPC1(ByRef sht As String)

        ' Replaces OPC1. in wire worksheet with the CPU_Name label
        '
        Dim ws As Worksheet

        ws = XLpicsWB.Sheets(sht)
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

    Sub Delete_Wire_Sheets(ByRef shtTemplate As String)

        Dim shtName As String
        Dim ws As Worksheet

        For i = 1 To 100     'Increase this value if there are ever more than 15 Wire Sheets
            shtName = Replace(shtTemplate, " Template", "_") & i

            ' Delete existing worksheets if necessary
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
