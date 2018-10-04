
Imports System.IO
Imports Microsoft.Office.Interop.Excel
Imports System.Text.RegularExpressions


Module Utilities

    Sub RegExpTest()

        ' Define a regular expression for repeated words
        ' requires Namespace System.Text.RegularExpressions for class definitions
        Dim rx As New Regex("\b(?<word>\w+)\s+(\k<word>)\b")

        ' Define a test string
        Dim text As String = "The the quick brown fox  fox jumps over the lazy dog dog."

        ' Find matches.
        Dim matches As MatchCollection = rx.Matches(text)

        ' Report the number of matches found
        Console.WriteLine("{0} matches found in:" & vbCrLf & "    {1}", matches.Count, text)

        ' Report on each match      
        Dim groups As GroupCollection
        For Each match In matches
            groups = match.Groups
            Console.WriteLine("'{0}' repeated at positions {1} and {2}", groups("word").Value, groups(0).Index, groups(1).Index)
        Next match

        Console.ReadLine()

        '// The example produces the following output to the console:
        '//       3 matches found in
        '//          The the quick brown fox  fox jumps over the lazy dog dog.
        '//       'The' repeated at positions 0 and 4
        '//       'fox' repeated at positions 20 and 25
        '//       'dog' repeated at positions 50 and 54

    End Sub

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

        Dim outFolder, topFolder, subFolder, pathName As String

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

    Sub Export_CSV(ByRef outFolder As String, ByVal sheetStr As String, ByVal saveName As String)

        ' Declare variables, create new Excel workbook object
        Dim wb As Workbook = XLApp.Workbooks.Add(XlWBATemplate.xlWBATWorksheet)  ' create new workbook
        Dim ws As Worksheet = wb.ActiveSheet
        Dim savePath As String = outFolder & "\" & saveName

        ws.Name = sheetStr
        wb.Application.DisplayAlerts = True
        wb.SaveAs(Filename:=savePath, FileFormat:=XlFileFormat.xlCSV)
        wb.Application.DisplayAlerts = True

        wb.Close(False)

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

    Public Function IsFileOpen(ByRef sName As String) As Boolean

        ' check if a file is still open (hanging process)
        Dim blnRetVal As Boolean = False
        Dim fs As FileStream = Nothing

        Try
            fs = File.Open(sName, FileMode.Open, FileAccess.Read, FileShare.None)
        Catch ex As Exception
            blnRetVal = True
        Finally
            'If Not IsNothing(fs) Then : fs.Close() : End If
            If Not IsNothing(fs) Then blnRetVal = False
        End Try

        Return blnRetVal

    End Function

End Module