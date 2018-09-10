
Public CPU_Name As String
Public Template_Name As String

Sub Generate_Wire_Data(ByRef x As Integer)

    Call Unhide_All_Sheets()

    CPU_Name = Get_CPU_Name()
    Template_Name = "Object"

    'Call Create_Wire_AIn_Sheets
    Call Create_Basic_Wire_Sheets("AIn", "*_Inp_?V")
    Call Create_Basic_Wire_Sheets("DIn", "*_Inp_PV")
    Call Create_Basic_Wire_Sheets("Motor", "*_Out_Run")
    Call Create_Basic_Wire_Sheets("ValveC", "*_Out_CV")
    Call Create_Basic_Wire_Sheets("ValveMO", "*_Out_Open")
    Call Create_Basic_Wire_Sheets("ValveSO", "*_Out")
    Call Create_Basic_Wire_Sheets("VSD", "*_Out_SpeedRef")

    Call Color_Wire_Tabs()

    Sheets("Instructions").Select

    Call Hide_Sheets()

End Sub

Private Function GetRowGap(sheet As String) As Integer

    Dim Count As Integer
    Count = 0
    
    Dim wrkSht As Worksheet
    Set wrkSht = ThisWorkbook.Sheets(sheet)
    
    Dim itemRng As Range
    Set itemRng = wrkSht.Range("B1")
    
    Do While ExtractNumber(itemRng.Value) = 1
        Set itemRng = itemRng.Offset(1, 0)
        Count = Count + 1
    Loop
    
    GetRowGap = Count

End Function

Private Function GetMaxItems(sheet As String) As Integer

    Dim maxNum As String
    Dim toSub As String
    Dim lastRng As Range
    Set lastRng = ThisWorkbook.Sheets(sheet).Range("B1").End(xlDown)
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

Private Function Find_Column(sheet As String, str As String) As Integer

    Find_Column = ThisWorkbook.Sheets(sheet).Range("A:ZZ").Find(str, LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext).Column

End Function

' Pre-Pass OPC1 tags for existance
' Checks against SimData sheet for existance of OPC1 tag that it is looking to use
Sub ValidateOPC1(sheetStr As String)

    Dim searchRng As Range
    Dim firstMatch As Range
    Dim wrkSht As Worksheet
    Set wrkSht = ThisWorkbook.Sheets(sheetStr)
    
    Set searchRng = wrkSht.Range("A1")
    
    Do While Not searchRng Is Nothing
    
        Set searchRng = wrkSht.Range("A:ZZ").Find(What:="OPC1.", After:=searchRng, LookAt:=xlPart)
        
        Dim output As String
        Dim parseArr As Variant
        Dim parse As String
        
        parse = searchRng.Value
        output = ""
        
        ' Check if it needs to be split
        If InStr(parse, "|") > 0 Then
            parseArr = Split(parse, "|")
        Else
            parseArr = Array(parse)
        End If
        
        Dim checkStr As Variant
        
        For Each checkStr In parseArr
            
            Dim tag As String
            Dim re As New RegExp
            
            With re
                .Global = False
                .MultiLine = False
                .IgnoreCase = False
                .Pattern = "OPC1\.\w*"
            End With
            
            Dim result As Object
            Set result = re.Execute(checkStr)
            
            tag = result(0)
            tag = Replace(tag, "OPC1.", "")
        
            If Is_Sim_Data(tag) Then
                output = checkStr
                Exit For
            End If
        Next
        
        searchRng.Value = output
        
        
        If firstMatch Is Nothing Then
            Set firstMatch = searchRng
        ElseIf searchRng = firstMatch Then
            Set searchRng = Nothing
        End If
    
    Loop

End Sub

Private Function Is_Sim_Data(str As String) As Boolean
    
    Dim wrkSht As Worksheet
    Set wrkSht = ThisWorkbook.Sheets("SimData")
    
    Dim searchRng As Range
    Set searchRng = wrkSht.Range("A:A").Find(str)
    
    Is_Sim_Data = Not (searchRng Is Nothing)
    
    
End Function

Sub ReplaceOPC1(NewSheetName As String)
'
' Replaces OPC1. in a sheet with the CPU_Name
'
    Sheets(NewSheetName).UsedRange.Select
    
    Cells.Replace What:="OPC1.", _
                            Replacement:=CPU_Name & ".", _
                            LookAt:=xlPart, _
                            SearchOrder:=xlByRows, _
                            MatchCase:=False, _
                            SearchFormat:=False, _
                            ReplaceFormat:=False
    Range("A1").Select

End Sub

Sub Color_Wire_Tabs(ByRef x As Integer)

    Dim ws As Worksheet

    For Each ws In ThisWorkbook.Worksheets

        If InStr(ws.Name, "Wire") > 0 Then
            Sheets(ws.Name).Tab.ColorIndex = 24
            If InStr(ws.Name, "Template") > 0 Then
                Sheets(ws.Name).Tab.ColorIndex = 49
            End If
        End If

    Next ws

End Sub

Sub Delete_Wire_Sheets(sheetTemplate As String)

    Dim sheetName As String
    Dim ws As Worksheet
    
    For i = 1 To 100     'Increase this value if there are ever more than 15 Wire Sheets
        sheetName = Replace(sheetTemplate, " Template", "_") & i
    
        ' Deletes existing sheets if necesary
        For Each ws In ThisWorkbook.Worksheets
            If ws.Name = sheetName Then
                Application.DisplayAlerts = False
                ws.Delete
                Application.DisplayAlerts = True
            End If
        Next ws
    Next i
    
End Sub

Sub Create_Basic_Wire_Sheets(typeStr As String, countStr As String)

    Dim sheetTemplate As String
    sheetTemplate = "Wire_" & typeStr & " Template"
    
    Dim sourceSheet As String
    sourceSheet = "IOTags - " & typeStr
    
    Dim minMaxSheet As String
    minMaxSheet = "MinMax - " & typeStr
    
    Dim minMax As Boolean
    minMax = Worksheet_Exists(minMaxSheet)
    
    Dim sSht As Worksheet
    Set sSht = ThisWorkbook.Sheets(sourceSheet)
    
    RowGap = GetRowGap(sheetTemplate)
    
    ItemCount = Application.WorksheetFunction.CountIf(sSht.Range("A:A"), countStr)
    maxItemCount = GetMaxItems(sheetTemplate)
    ReqSheets = Application.WorksheetFunction.RoundUp(ItemCount / maxItemCount, 0)
    
    Dim itemRng As Range
    Set itemRng = sSht.Range("A1")
    
    Call Delete_Wire_Sheets(sheetTemplate)
    
    If minMax Then
    
        InMinCol = Find_Column(minMaxSheet, "InputMin")
        InMaxCol = Find_Column(minMaxSheet, "InputMax")
        OutMinCol = Find_Column(minMaxSheet, "OutputMin")
        OutMaxCol = Find_Column(minMaxSheet, "OutputMax")
        
        If typeStr = "AIn" Then
        
            EUMinCol = Find_Column(sheetTemplate, "iAI_EU_Min") + 1
            EUMaxCol = Find_Column(sheetTemplate, "iAI_EU_Max") + 1
            RawMinCol = Find_Column(sheetTemplate, "iAI_Raw_Min") + 1
            RawMaxCol = Find_Column(sheetTemplate, "iAI_Raw_Max") + 1
            RawFltCol = Find_Column(sheetTemplate, "iAI_Raw_Flt") + 1
        
        End If
    
    End If
    
    For shtIndex = 1 To ReqSheets
    
        Dim NewSheetName As String
        NewSheetName = Replace(sheetTemplate, " Template", "_") & shtIndex
        
        Sheets(sheetTemplate).Copy Before:=Sheets("Wire_AIn Template")
        ActiveSheet.Unprotect
        ActiveSheet.Name = NewSheetName
        
        Range("A1").Cells.Value = "_Wire_" & typeStr & "_" & shtIndex
        For itemIndex = 1 To maxItemCount
            
            Dim nextRng As Range
            Set nextRng = sSht.Range("A:A").Find(countStr, itemRng)
            
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
                    .MultiLine = True
                    .IgnoreCase = False
                    .Pattern = searchStr
                End With
                
                TagName = re.Replace(nextRng.Value, "")
                
                ActiveSheet.Cells.Replace What:=Template_Name & itemNum, Replacement:=TagName, LookAt:=xlPart, _
                    SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
                
                If minMax Then
                    If typeStr = "AIn" Then
                    
                        Dim EU_Min As Double
                        Dim EU_Max As Double
                        Dim Raw_Min As Double
                        Dim Raw_Max As Double
                        Dim Raw_Flt As Double
                        
                        CurrentRow = RowGap * (itemIndex - 1) + 1
                        ItemRow = Worksheets(minMaxSheet).Range("A:A").Find(nextRng.Value).Row
                        EU_Min = Worksheets(minMaxSheet).Cells(ItemRow, InMinCol).Value
                        EU_Max = Worksheets(minMaxSheet).Cells(ItemRow, InMaxCol).Value
                        Raw_Min = Worksheets(minMaxSheet).Cells(ItemRow, OutMinCol).Value
                        Raw_Max = Worksheets(minMaxSheet).Cells(ItemRow, OutMaxCol).Value
                        Raw_Flt = 0.8 * Raw_Min
                        
                        If Raw_Max > 0 Then
                            Cells(CurrentRow, EUMinCol).Cells.Value = EU_Min
                            Cells(CurrentRow, EUMaxCol).Cells.Value = EU_Max
                            Cells(CurrentRow, RawMinCol).Cells.Value = Raw_Min
                            Cells(CurrentRow, RawMaxCol).Cells.Value = Raw_Max
                            Cells(CurrentRow, RawFltCol).Cells.Value = Raw_Flt
                        End If
                        
                    End If
                End If

                Set itemRng = nextRng.Offset(1, 0)
            
            End If

        Next
        
        Range("A1").Select
        
        Call ValidateOPC1(NewSheetName)
        Call ReplaceOPC1(NewSheetName)
    Next

End Sub

Sub Export_Wire_Data(outFolder As String)
    
    Application.DisplayAlerts = False
    
    Dim ws As Worksheet
    Dim book As Workbook
    Dim wsName As String
    Dim savePath As String
    
    savePath = outFolder & "\"
    
    Save_Sheets "AIn", savePath
    Save_Sheets "DIn", savePath
    Save_Sheets "Motor", savePath
    Save_Sheets "ValveC", savePath
    Save_Sheets "ValveMO", savePath
    Save_Sheets "ValveSO", savePath
    Save_Sheets "VSD", savePath

    Application.DisplayAlerts = True

End Sub

Private Function Save_Sheets(typeStr As String, savePath As String)

    Dim i As Integer
    i = 1
    
    Dim wsName As String
    wsName = "Wire_" & typeStr & "_" & i
    
    Do While Worksheet_Exists(wsName)

        Dim book As Workbook
        Set book = Workbooks.Add
        Count = book.Worksheets.Count
        ThisWorkbook.Sheets(wsName).Copy book.Worksheets(1)
        
        book.Worksheets(1).Range("A1").EntireRow.Insert
        book.Worksheets(1).Range("A1").Value = ";PICS for Windows - Device Wiring Export V1.10"
        
        
        book.Worksheets(1).SaveAs savePath & wsName & ".wir", xlTextWindows
        book.Close False
        Set book = Nothing
            
        i = i + 1
        wsName = "Wire_" & typeStr & "_" & i
        
    Loop

End Function
