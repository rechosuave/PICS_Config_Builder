
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

    Sub Create_Basic_Wire_Sheets(ByVal sTemplateType As String, ByVal sCountStr As String)

        Dim sWireTemplateWS As String = "Wire_" & sTemplateType & " Template"    ' build worksheet names
        Dim sIOTagsWS As String = "IOTags - " & sTemplateType
        Dim sMinMaxWS As String = "MinMax - " & sTemplateType
        Dim sTagItemNum, sSearchStr, sTagName, sNewWireWS As String
        Dim iInMinCol, iInMaxCol, iOutMinCol, iOutMaxCol, iRow As Integer
        Dim iEUMinCol, iEUMaxCol, iRawMinCol, iRawMaxCol, iRawFltCol As Integer
        Dim iCurrentRow As Integer, iBlkSize As Integer = Get_Annunciator_Block_Size(sWireTemplateWS)
        Dim bMinMaxShtFound As Boolean = WS_Exists(sMinMaxWS)
        Dim iMaxItemCount As Integer = GetMaxItems(sWireTemplateWS)     'number of annunciator items in wire template worksheet
        Dim ws As Worksheet = XLpicsWB.Sheets(sIOTagsWS)
        Dim rngIOTagsWS As Range = ws.Range("A:A")
        Dim rngIOTagsWS_Start As Range = ws.Range("A1"), rngIOTagsNext As Range
        Dim iItemCount As Integer = ws.Application.WorksheetFunction.CountIf(rngIOTagsWS, sCountStr)
        Dim iReqWireSheets As Integer = ws.Application.WorksheetFunction.RoundUp(iItemCount / iMaxItemCount, 0)
        Dim fEU_Min, fEU_Max, fRaw_Min, fRaw_Max, fRaw_Flt As Double
        Dim oRE As Object

        oRE = CreateObject("vbscript.regexp")    ' create a regular expression

        Call Delete_Wire_Sheets(sWireTemplateWS)

        If bMinMaxShtFound Then  ' find columns

            iInMinCol = Find_Column(sMinMaxWS, "InputMin")
            iInMaxCol = Find_Column(sMinMaxWS, "InputMax")
            iOutMinCol = Find_Column(sMinMaxWS, "OutputMin")
            iOutMaxCol = Find_Column(sMinMaxWS, "OutputMax")

            If sTemplateType = "AIn" Then     ' find columns

                iEUMinCol = Find_Column(sWireTemplateWS, "iAI_EU_Min") + 1
                iEUMaxCol = Find_Column(sWireTemplateWS, "iAI_EU_Max") + 1
                iRawMinCol = Find_Column(sWireTemplateWS, "iAI_Raw_Min") + 1
                iRawMaxCol = Find_Column(sWireTemplateWS, "iAI_Raw_Max") + 1
                iRawFltCol = Find_Column(sWireTemplateWS, "iAI_Raw_Flt") + 1

            End If

        End If

        For shtIndex = 1 To iReqWireSheets     ' add a new wire worksheet(s) to PICS workbook and populate data into worksheet

            sNewWireWS = Strings.Replace(sWireTemplateWS, " Template", "_") & shtIndex
            XLpicsWB.Sheets(sWireTemplateWS).Copy(Before:=XLpicsWB.Sheets("Wire_AIn Template"))
            XLpicsWB.ActiveSheet.Unprotect
            ws = XLpicsWB.ActiveSheet
            ws.Name = sNewWireWS      ' rename copied worksheet
            ws.Range("A1").Cells.Value = "_Wire_" & sTemplateType & "_" & shtIndex

            For i = 1 To iMaxItemCount   ' loop for number of annunciators on new wire worksheet to populate with tag item data

                rngIOTagsNext = XLpicsWB.Sheets(sIOTagsWS).Range("A:A").Find(What:=sCountStr, After:=rngIOTagsWS_Start)     ' if next range is Nothing - skip loop

                If Not (rngIOTagsNext Is Nothing) Then        ' c2-10-02-2018: condition statement added since if range expression "netRng.Row" is nothing will crash loop
                    ' Is there another tag item to add to annunciator block
                    If rngIOTagsNext.Row > rngIOTagsWS_Start.Row Then

                        sTagItemNum = Right("00" & i, 2)        ' assign a IO tag item #

                        sSearchStr = Replace(sCountStr, "*", "")
                        sSearchStr = Replace(sSearchStr, "?", "\w")

                        With oRE
                            .Global = True
                            .Multiline = True
                            .IgnoreCase = False
                            .Pattern = sSearchStr
                        End With

                        sTagName = oRE.Replace(rngIOTagsNext.Value, "")

                        ws.Cells.Replace(What:=Template_Name & sTagItemNum, Replacement:=sTagName, LookAt:=XlLookAt.xlPart,
                                             SearchOrder:=XlSearchOrder.xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False)

                        If bMinMaxShtFound Then
                            If sTemplateType = "AIn" Then

                                iCurrentRow = iBlkSize * (i - 1) + 1
                                iRow = XLpicsWB.Sheets(sMinMaxWS).Range("A:A").Find(rngIOTagsNext.Value).Row
                                fEU_Min = XLpicsWB.Sheets(sMinMaxWS).Cells(iRow, iInMinCol).Value
                                fEU_Max = XLpicsWB.Sheets(sMinMaxWS).Cells(iRow, iInMaxCol).Value
                                fRaw_Min = XLpicsWB.Sheets(sMinMaxWS).Cells(iRow, iOutMinCol).Value
                                fRaw_Max = XLpicsWB.Sheets(sMinMaxWS).Cells(iRow, iOutMaxCol).Value
                                fRaw_Flt = 0.8 * fRaw_Min

                                If fRaw_Max > 0 Then
                                    ws.Cells(iCurrentRow, iEUMinCol).Cells.Value = fEU_Min
                                    ws.Cells(iCurrentRow, iEUMaxCol).Cells.Value = fEU_Max
                                    ws.Cells(iCurrentRow, iRawMinCol).Cells.Value = fRaw_Min
                                    ws.Cells(iCurrentRow, iRawMaxCol).Cells.Value = fRaw_Max
                                    ws.Cells(iCurrentRow, iRawFltCol).Cells.Value = fRaw_Flt
                                End If

                            End If
                        End If

                        rngIOTagsWS_Start = rngIOTagsNext.Offset(1, 0)

                    End If

                End If      'added for nextRng Is Nothing

            Next i

            Call ValidateOPC1(sNewWireWS)
            Call ReplaceOPC1(sNewWireWS)

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

        '  Count the number of Annunciator blocks on the wire template worksheet 
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
        ' Checks SimData worksheet for existence of OPC1 tag that it is looking to use
        Dim output, parseArr(), parse, checkStr, tag As String
        Dim searchRng As Range, firstMatch As Range = Nothing
        Dim ws As Worksheet = XLpicsWB.Sheets(WireSht)
        'Dim oRE As New RegExp
        Dim oRE As Object
        oRE = CreateObject("vbscript.regexp")

        Dim oMatch As Object

        With oRE     ' vbscript object RegExp is a concise and flexible notation for finding and replacing patterns of text
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

            ' Check if it string needs to be split
            If InStr(parse, "|") > 0 Then
                parseArr = parse.Split("|")
            Else
                parseArr = {parse}
            End If

            For Each checkStr In parseArr       ' check if using string 1 or 2 (pipe | delimited)
                oMatch = oRE.Execute(checkStr)
                tag = oMatch(0).value
                tag = Replace(tag, "OPC1.", "")

                If Is_Sim_Data(tag) Then
                    output = checkStr
                    Exit For
                End If

            Next checkStr

            searchRng.Value = output        ' set IO tag name to "" if item block is not used (no annunciator)

            If firstMatch Is Nothing Then       ' end loop is 1st tag found again
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

        For i = 1 To 15     'Increase this value if there are ever more than 15 Wire Sheets
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

        savePath = outFolder & "\"

        Save_Sheets("AIn", savePath)
        Save_Sheets("DIn", savePath)
        Save_Sheets("Motor", savePath)
        Save_Sheets("ValveC", savePath)
        Save_Sheets("ValveMO", savePath)
        Save_Sheets("ValveSO", savePath)
        Save_Sheets("VSD", savePath)

    End Sub

    Private Sub Save_Sheets(ByVal typeStr As String, ByRef savePath As String)

        Dim i As Integer = 1
        Dim wsName As String = "Wire_" & typeStr & "_" & i
        Dim wb, newBook As Workbook
        Dim ws As Worksheet

        wb = XLpicsWB

        Do While WS_Exists(wsName)

            newBook = XLApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet)  ' create a new workbook for each wire file
            newBook.Application.DisplayAlerts = False
            ws = wb.Sheets(wsName)      ' select worksheet to copy to new workbook
            ws.Copy(Before:=newBook.Sheets(1))
            newBook.Worksheets(1).Range("A1").EntireRow.Insert
            newBook.Worksheets(1).Range("A1").Value = ";PICS for Windows - Device Wiring Export V1.10"
            newBook.SaveAs(savePath & wsName & ".wir", XlFileFormat.xlTextWindows)
            newBook.Application.DisplayAlerts = True
            newBook.Close(False)

            i = i + 1
            wsName = "Wire_" & typeStr & "_" & i

        Loop

    End Sub

End Module
