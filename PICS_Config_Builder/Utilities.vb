
Public Function Clear_Sheet(sheet As String)

    Dim wrkSheet As Worksheet
    Set wrkSheet = ThisWorkbook.Sheets(sheet)
    
    If InStr(sheet, "IOTags") > 0 Then
        wrkSheet.Range("A2:E9999").Clear
    ElseIf InStr(sheet, "IOMem") > 0 Then
        wrkSheet.Range("A2:F9999").Clear
    ElseIf InStr(sheet, "SimData") > 0 Then
        wrkSheet.Range("A2:E9999").Clear
    ElseIf InStr(sheet, "MinMax") > 0 Then
        wrkSheet.Range("A2:E9999").Clear
    ElseIf InStr(sheet, "MemoryData") > 0 Then
        wrkSheet.Range("A2:F9999").Clear
    ElseIf InStr(sheet, "ControlNetData") > 0 Then
        wrkSheet.Range("A2:F9999").Clear
    ElseIf sheet = "IO Sheets" Then
        ' Clear extra headers
        wrkSheet.Range("B1:AG9999").Clear
        ' Clear data
        wrkSheet.Range("A2:AG9999").Clear
    End If

End Function

Public Function Clear_Sheet_Type(typeStr As String)

    Dim sheetCount As Integer
    sheetCount = ThisWorkbook.Sheets.Count
    
    For i = 1 To sheetCount
        Dim wrkSheet As Worksheet
        Set wrkSheet = ThisWorkbook.Sheets(i)
        
        Dim shtName As String
        shtName = wrkSheet.Name
        
        If InStr(shtName, typeStr) > 0 Then
            Clear_Sheet shtName
        End If
        
    Next

End Function

Public Function Clear_All_Sheets(ByRef x As Integer)

    Clear_Sheet_Type("")

End Function


Public Function Reset_Sheet(sheet As String)

    Dim lastSht As Worksheet
    Set lastSht = Application.ActiveSheet
    
    ' Reset selection to A1
    ThisWorkbook.Sheets(sheet).Select
    Range("A1").Select
    
    ' Return to previous sheet
    lastSht.Select

End Function

Public Function Find_Header_Column(sheet As String, header As String) As Integer

    Dim wrkSheet As Worksheet
    Set wrkSheet = ThisWorkbook.Sheets(sheet)
    
    Dim searchRng As Range
    Set searchRng = wrkSheet.Range("A1")
    
    ' Find either the column or nothing
    Do While searchRng <> "" And searchRng.Value <> header
        Set searchRng = searchRng.Offset(0, 1)
    Loop
    
    ' If column found, return it.
    ' Otherwise zero.
    If searchRng.Value = header Then
        Find_Header_Column = searchRng.Column
    Else
        Find_Header_Column = 0
        MsgBox "Column '" & header & "' was not found. Please contact a VBA developer.", vbOKOnly, "Error: Config Header"
    End If

End Function

Public Function Hide_Sheets(ByRef x As Integer)

    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets

        If InStr(ws.Name, "IOTags") > 0 Then
            Sheets(ws.Name).Visible = False
        End If

        If InStr(ws.Name, "IOMem") > 0 Then
            Sheets(ws.Name).Visible = False
        End If

        If InStr(ws.Name, "Wire") > 0 Then
            Sheets(ws.Name).Visible = False
        End If

        If InStr(ws.Name, "MinMax") > 0 Then
            Sheets(ws.Name).Visible = False
        End If

        ' Added to hide everything but instructions, since no one should care about the other sheets
        If Not InStr(ws.Name, "Instructions") > 0 Then
            Sheets(ws.Name).Visible = False
        End If

    Next ws

End Function

Public Function Unhide_All_Sheets(ByRef x As Integer)

    For Each ws In Sheets : ws.Visible = True : Next

End Function

Public Function Worksheet_Exists(sheetName As String) As Boolean

    Worksheet_Exists = False

    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = sheetName Then
            Worksheet_Exists = True
        End If
    Next i

End Function

Public Function Create_Output_Folder() As String

    Dim outFolder As String
    Dim topFolder As String
    Dim subFolder As String
    Dim pathName As String
    Dim CPU_Name As String
    
    pathName = ActiveWorkbook.Path
    CPU_Name = Get_CPU_Name
    
    topFolder = "\PICS_Files"
    subFolder = "\" & CPU_Name & Format(Now(), "_yyyymmdd_HhNnSs")
    
    outFolder = pathName & topFolder
    If Len(Dir(outFolder, vbDirectory)) = 0 Then
       MkDir outFolder
    End If
    
    outFolder = outFolder & subFolder
    If Len(Dir(outFolder, vbDirectory)) = 0 Then
       MkDir outFolder
    End If
    
    Create_Output_Folder = outFolder

End Function

Sub Export_CSV(outFolder As String, sheetStr As String, saveName As String)
    
    savePath = outFolder & "\" & saveName
    
    Dim NewBook As Workbook
    Set NewBook = Workbooks.Add
    
    ThisWorkbook.Sheets(sheetStr).Visible = True
    
    ThisWorkbook.Sheets(sheetStr).Copy Before:=NewBook.Sheets(1)
    
    ThisWorkbook.Sheets(sheetStr).Visible = False
    
    Application.DisplayAlerts = False
    
    ActiveWorkbook.ActiveSheet.SaveAs fileName:=savePath, FileFormat:=xlCSV
    
    ActiveWorkbook.Close (False)
    
    Application.DisplayAlerts = True
    
End Sub

Public Function Get_CPU_Name() As String

    Dim CPU_Name As String
    CPU_Name = Worksheets("Instructions").Range("CPU_PREFIX").Cells.Value

    If CPU_Name = "" Then

        CPU_Name = Application.InputBox("Enter a topic name:", "Topic Name")

        If CPU_Name = "False" Then
            CPU_Name = "OPC1"
        End If

        Worksheets("Instructions").Range("CPU_PREFIX").Cells.Value = CPU_Name

    End If

    Get_CPU_Name = CPU_Name

End Function
