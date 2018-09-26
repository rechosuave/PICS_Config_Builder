
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Module Utilities

    Public Sub Clear_Sheet(ByRef ws As Worksheet)

        If InStr(ws.Name, "IOTags") > 0 Then
            ws.Range("A2:E9999").Clear()
        ElseIf InStr(ws.Name, "IOMem") > 0 Then
            ws.Range("A2:F9999").Clear()
        ElseIf InStr(ws.Name, "SimData") > 0 Then
            ws.Range("A2:E9999").Clear()
        ElseIf InStr(ws.Name, "MinMax") > 0 Then
            ws.Range("A2:E9999").Clear()
        ElseIf InStr(ws.Name, "MemoryData") > 0 Then
            ws.Range("A2:F9999").Clear()
        ElseIf InStr(ws.Name, "ControlNetData") > 0 Then
            ws.Range("A2:F9999").Clear()
        ElseIf ws.Name = "IO Sheets" Then
            ' Clear extra headers
            ws.Range("B1:AG9999").Clear()
            ' Clear data
            ws.Range("A2:AG9999").Clear()
        End If

    End Sub

    Public Sub Clear_Sheet_Type(typeStr As String)

        Dim ws As Worksheet
        Dim sheetCount As Integer = XLpicsWB.Sheets.Count

        For i = 1 To sheetCount
            ws = XLpicsWB.Sheets(i)
            If InStr(ws.Name, typeStr) > 0 Then Call Clear_Sheet(ws)
        Next

    End Sub

    Public Sub Clear_All_Sheets()

        Clear_Sheet_Type("")

    End Sub

    Public Sub Reset_Sheet(ByRef wrkBook As Workbook, sheet As String)

        Dim lastSht As Worksheet
        lastSht = wrkBook.ActiveSheet

        ' Reset selection to A1
        wrkBook.Sheets(sheet).Select
        lastSht.Range("A1").Select()

        ' Return to previous sheet
        lastSht.Select()

    End Sub

    Public Function Find_Header_Column(shtName As String, header As String) As Integer

        Dim ws As Worksheet = XLpicsWB.Sheets(shtName)
        Dim searchRng As Range = ws.Range("A1")

        ' Find either the column or nothing
        Do While searchRng.Value <> "" And searchRng.Value <> header
            searchRng = searchRng.Offset(0, 1)
        Loop

        ' If column found, return it  -- Otherwise zero.
        If searchRng.Value = header Then
            Return searchRng.Column
        Else
            MsgBox("Column '" & header & "' was not found. Please contact a VBA developer.", vbOKOnly, "Error: Config Header")
            Return 0
        End If

    End Function

    Public Sub Hide_Sheets(ByRef wrkBook As Workbook)

        Dim ws As Worksheet
        For Each ws In wrkBook.Worksheets

            If InStr(ws.Name, "IOTags") > 0 Then
                ws.Visible = False
            End If

            If InStr(ws.Name, "IOMem") > 0 Then
                ws.Visible = False
            End If

            If InStr(ws.Name, "Wire") > 0 Then
                ws.Visible = False
            End If

            If InStr(ws.Name, "MinMax") > 0 Then
                ws.Visible = False
            End If

            ' Added to hide everything but instructions, since no one should care about the other sheets
            If Not InStr(ws.Name, "Instructions") > 0 Then
                ws.Visible = False
            End If

        Next ws

    End Sub

    Public Function Create_Output_Folder(ByRef ActiveWorkbook As Workbook) As String

        Dim outFolder As String
        Dim topFolder As String
        Dim subFolder As String
        Dim pathName As String

        pathName = ActiveWorkbook.Path

        topFolder = "\PICS_Files"
        subFolder = "\" & ImportData.CPU_Name & Format(Now(), "_yyyymmdd_HhNnSs")

        outFolder = pathName & topFolder
        If Len(Dir(outFolder, vbDirectory)) = 0 Then
            MkDir(outFolder)
        End If

        outFolder = outFolder & subFolder
        If Len(Dir(outFolder, vbDirectory)) = 0 Then
            MkDir(outFolder)
        End If

        Create_Output_Folder = outFolder

    End Function

    Sub Export_CSV(outFolder As String, sheetStr As String, saveName As String)

        ' Declare variables, create new Excel workbook object
        Dim wb As Workbook = XLApp.Workbooks.Add
        Dim ws As Worksheet = wb.ActiveSheet
        Dim savePath As String = outFolder & "\" & saveName

        ws.Sheets(sheetStr).Visible = True
        ws.Sheets(sheetStr).Copy(Before:=wb.Sheets(1))
        ws.Sheets(sheetStr).Visible = False
        wb.DisplayAlerts = False
        wb.SaveAs(Filename:=savePath, FileFormat:=XlFileFormat.xlCSV)

        wb.Close(False)
        wb.DisplayAlerts = True

    End Sub

    Public Function Get_CPU_Name(ByRef wrkBook As Workbook) As String

        Dim CPU_Name As String = wrkBook.Sheets("Instructions").Range("CPU_PREFIX").Cells.Value

        If CPU_Name = "" Then

            CPU_Name = InputBox("Enter a topic name:", "Topic Name")

            If CPU_Name = "" Then
                CPU_Name = "OPC1"
            End If

            wrkBook.Sheets("Instructions").Range("CPU_PREFIX").Cells.Value = CPU_Name

        End If

        Return CPU_Name

    End Function

End Module