
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel


Module Templates

    Public XLTemplateWB As Workbook
    Sub Validate_PICS_WB()

        If Not IsValid_WB() Then
            Call Build_IOSheets_WS()
            Call Build_SimData_WS()
            Call Build_MemoryData_WS()
            Call Build_ControlNetData_WS()
            Call Build_IOTags_AIn_WS()
            Call Build_IOTags_DIn_WS()
            Call Build_IOTags_DOut_WS()
            Call Build_IOTags_Motor_WS()
            Call Build_IOTags_ValveC_WS()
            Call Build_IOTags_ValveSO_WS()
            Call Build_IOTags_ValveMO_WS()
            Call Build_IOTags_VSD_WS()

            Call Build_IOMem_AIn_WS()
            Call Build_IOMem_DIn_WS()
            Call Build_IOMem_Motor_WS()
            Call Build_IOMem_ValveC_WS()
            Call Build_IOMem_ValveMO_WS()
            Call Build_IOMem_ValveSO_WS()
            Call Build_IOMem_VSD_WS()

            Call Build_MinMax_AIn_WS()
            Call Build_MinMax_ValveC_WS()
            Call Build_MinMax_VSD_WS()

            Call Build_Wire_WS()
            'Call Build_Wire_AIn_WS()
            'Call Build_Wire_DIn_WS()
            'Call Build_Wire_Motor_WS()
            'Call Build_Wire_ValveC_WS()
            'Call Build_Wire_ValveMO_WS()
            'Call Build_Wire_ValveSO_WS()
            'Call Build_Wire_VSD_WS()
            'Call Record_WS_Tabs_WS()

        End If

    End Sub

    Function IsValid_WB() As Boolean

        'Check if PICS config file is a correctly formatted workbook
        Dim shtName() As String = {"IO Sheets", "SimData", "MemoryData", "ControlNetData",
           "IOTags - AIn", "IOTags - DIn", "IOTags - DOut", "IOTags - Motor", "IOTags - ValveC",
           "IOTags - ValveSO", "IOTags - ValveMO", "IOTags - VSD",
           "IOMem - AIn", "IOMem - DIn", "IOMem - Motor", "IOMem - ValveC", "IOMem - ValveMO",
           "IOMem - ValveSO", "IOMem - VSD",
           "MinMax - AIn", "MinMax - ValveC", "MinMax - VSD",
           "Wire_AIn Template", "Wire_DIn Template", "Wire_Motor Template", "Wire_ValveC Template",
           "Wire_ValveMO Template", "Wire_ValveSO Template", "Wire_VSD Template"}
        Dim EndofArray As Integer = shtName.Count - 1  ' number of elements in array
        Dim shtFound As Boolean = False
        Dim ws As Worksheet

        For i = 0 To EndofArray
            For Each ws In XLpicsWB.Sheets      ' check if workbook has all required worksheets?
                If ws.Name.Equals(shtName(i)) Then
                    shtFound = True
                    Exit For
                End If
            Next
            If Not shtFound Then
                Exit For
            End If
        Next

        Return shtFound

    End Function

    Sub Build_IOSheets_WS()

        'Create a "IO Sheets" worksheet
        Dim shtName As String = "IO Sheets"
        Dim ColName As String = "PLCBaseTag"
        Dim tabColor As Integer = RGB(196, 215, 155)
        Dim shtFound As Boolean = False
        Dim ws As Worksheet
        Dim shtCount As Integer = XLpicsWB.Sheets.Count

        For Each ws In XLpicsWB.Sheets      ' does worksheet exist?
            If ws.Name.Equals(shtName) Then
                shtFound = True
                Exit For
            End If
        Next

        If Not shtFound Then        ' create worksheet
            If shtCount = 1 Then        ' rename sheet 1
                ws = XLpicsWB.ActiveSheet
                ws.Name = shtName
                ws.Range("A2:AA9999").Clear()
            Else
                ws = XLpicsWB.Sheets.Add(After:=XLpicsWB.Sheets(1))
                ws.Name = shtName
            End If
        Else
            ws = XLpicsWB.Sheets(shtName)
        End If

        ws.Range("A1").Value = ColName
        With ws.Tab
            .Color = tabColor
            .TintAndShade = 0
        End With

    End Sub

    Sub Build_SimData_WS()

        'Create a "SimData" worksheet
        Dim shtName As String = "SimData"
        Dim ColName As String = ";PICS Pro for Windows - Variables Export V2.00"
        Dim tabColor As Integer = RGB(255, 153, 0)
        Dim shtFound As Boolean = False
        Dim ws As Worksheet
        Dim shtCount As Integer = XLpicsWB.Sheets.Count

        For Each ws In XLpicsWB.Sheets      ' does worksheet exist?
            If ws.Name.Equals(shtName) Then
                shtFound = True
                Exit For
            End If
        Next

        If Not shtFound Then        ' create worksheet
            ws = XLpicsWB.Sheets.Add(After:=XLpicsWB.Sheets(shtCount))
            ws.Name = shtName
        Else
            Exit Sub  'worksheet already exits - nothing to do
        End If

        ws.Range("A1").Value = ColName
        With ws.Tab
            .Color = tabColor
            .TintAndShade = 0
        End With

    End Sub

    Sub Build_MemoryData_WS()

        'Create a "SimData" worksheet
        Dim shtName As String = "MemoryData"
        Dim ColName As String = "';PICS Pro for Windows - Variables Export V2.00"
        Dim tabColor As Integer = RGB(204, 153, 255)
        Dim shtFound As Boolean = False
        Dim ws As Worksheet
        Dim shtCount As Integer = XLpicsWB.Sheets.Count

        For Each ws In XLpicsWB.Sheets      ' does worksheet exist?
            If ws.Name.Equals(shtName) Then
                shtFound = True
                Exit For
            End If
        Next

        If Not shtFound Then        ' create worksheet
            ws = XLpicsWB.Sheets.Add(After:=XLpicsWB.Sheets(shtCount))
            ws.Name = shtName
        Else
            Exit Sub  'worksheet already exits - nothing to do
        End If

        ws.Range("A1").Value = ColName
        With ws.Tab
            .Color = tabColor
            .TintAndShade = 0
        End With

    End Sub

    Sub Build_ControlNetData_WS()

        'Create a "SimData" worksheet
        Dim shtName As String = "ControlNetData"
        Dim ColName() As String = {"Base Tag", "Name", "Type", "Default Value", "IO Address", "Description"}
        Dim tabColor As Integer = RGB(204, 204, 255)
        Dim shtFound As Boolean = False
        Dim letter As Char
        Dim ws As Worksheet
        Dim shtCount As Integer = XLpicsWB.Sheets.Count

        For Each ws In XLpicsWB.Sheets      ' does worksheet exist?
            If ws.Name.Equals(shtName) Then
                shtFound = True
                Exit For
            End If
        Next

        If Not shtFound Then        ' create worksheet
            ws = XLpicsWB.Sheets.Add(After:=XLpicsWB.Sheets(shtCount))
            ws.Name = shtName
        Else
            Exit Sub  'worksheet already exits - nothing to do
        End If

        For i = 0 To ColName.Count - 1
            letter = Convert.ToChar(65 + i)  ' starting pt is letter "A"
            ws.Range(letter & "1").Value = ColName(i) ' assign value to cells A1, B1,..
        Next
        With ws.Tab
            .Color = tabColor
            .TintAndShade = 0
        End With

    End Sub

    Sub Build_IOTags_AIn_WS()

        Dim shtName As String = "IOTags - AIn"
        Dim ColName() As String = {"Name", "Type", "Default Value", "IO Address", "Description"}
        Dim tabColor As Integer = RGB(255, 153, 0)
        Dim shtFound As Boolean = False
        Dim letter As Char
        Dim ws As Worksheet
        Dim shtCount As Integer = XLpicsWB.Sheets.Count

        For Each ws In XLpicsWB.Sheets      ' does worksheet exist?
            If ws.Name.Equals(shtName) Then
                shtFound = True
                Exit For
            End If
        Next

        If Not shtFound Then        ' create worksheet
            ws = XLpicsWB.Sheets.Add(After:=XLpicsWB.Sheets(shtCount))
            ws.Name = shtName
        Else
            Exit Sub  'worksheet already exits - nothing to do
        End If

        For i = 0 To ColName.Count - 1
            letter = Convert.ToChar(65 + i)  ' starting pt is letter "A"
            ws.Range(letter & "1").Value = ColName(i) ' assign value to cells A1, B1,..
        Next
        With ws.Tab
            .Color = tabColor
            .TintAndShade = 0
        End With

    End Sub

    Sub Build_IOTags_DIn_WS()

        Dim shtName As String = "IOTags - DIn"
        Dim ColName() As String = {"Name", "Type", "Default Value", "IO Address", "Description"}
        Dim tabColor As Integer = RGB(255, 153, 0)
        Dim shtFound As Boolean = False
        Dim letter As Char
        Dim ws As Worksheet
        Dim shtCount As Integer = XLpicsWB.Sheets.Count

        For Each ws In XLpicsWB.Sheets      ' does worksheet exist?
            If ws.Name.Equals(shtName) Then
                shtFound = True
                Exit For
            End If
        Next

        If Not shtFound Then        ' create worksheet
            ws = XLpicsWB.Sheets.Add(After:=XLpicsWB.Sheets(shtCount))
            ws.Name = shtName
        Else
            Exit Sub  'worksheet already exits - nothing to do
        End If

        For i = 0 To ColName.Count - 1
            letter = Convert.ToChar(65 + i)  ' starting pt is letter "A"
            ws.Range(letter & "1").Value = ColName(i) ' assign value to cells A1, B1,..
        Next
        With ws.Tab
            .Color = tabColor
            .TintAndShade = 0
        End With

    End Sub

    Sub Build_IOTags_DOut_WS()

        Dim shtName As String = "IOTags - DOut"
        Dim ColName() As String = {"Name", "Type", "Default Value", "IO Address", "Description"}
        Dim tabColor As Integer = RGB(255, 153, 0)
        Dim shtFound As Boolean = False
        Dim letter As Char
        Dim ws As Worksheet
        Dim shtCount As Integer = XLpicsWB.Sheets.Count

        For Each ws In XLpicsWB.Sheets      ' does worksheet exist?
            If ws.Name.Equals(shtName) Then
                shtFound = True
                Exit For
            End If
        Next

        If Not shtFound Then        ' create worksheet
            ws = XLpicsWB.Sheets.Add(After:=XLpicsWB.Sheets(shtCount))
            ws.Name = shtName
        Else
            Exit Sub  'worksheet already exits - nothing to do
        End If

        For i = 0 To ColName.Count - 1
            letter = Convert.ToChar(65 + i)  ' starting pt is letter "A"
            ws.Range(letter & "1").Value = ColName(i) ' assign value to cells A1, B1,..
        Next
        With ws.Tab
            .Color = tabColor
            .TintAndShade = 0
        End With

    End Sub

    Sub Build_IOTags_Motor_WS()

        Dim shtName As String = "IOTags - Motor"
        Dim ColName() As String = {"Name", "Type", "Default Value", "IO Address", "Description"}
        Dim tabColor As Integer = RGB(255, 153, 0)
        Dim shtFound As Boolean = False
        Dim letter As Char
        Dim ws As Worksheet
        Dim shtCount As Integer = XLpicsWB.Sheets.Count

        For Each ws In XLpicsWB.Sheets      ' does worksheet exist?
            If ws.Name.Equals(shtName) Then
                shtFound = True
                Exit For
            End If
        Next

        If Not shtFound Then        ' create worksheet
            ws = XLpicsWB.Sheets.Add(After:=XLpicsWB.Sheets(shtCount))
            ws.Name = shtName
        Else
            Exit Sub  'worksheet already exits - nothing to do
        End If

        For i = 0 To ColName.Count - 1
            letter = Convert.ToChar(65 + i)  ' starting pt is letter "A"
            ws.Range(letter & "1").Value = ColName(i) ' assign value to cells A1, B1,..
        Next
        With ws.Tab
            .Color = tabColor
            .TintAndShade = 0
        End With

    End Sub

    Sub Build_IOTags_ValveC_WS()

        Dim shtName As String = "IOTags - ValveC"
        Dim ColName() As String = {"Name", "Type", "Default Value", "IO Address", "Description"}
        Dim tabColor As Integer = RGB(255, 153, 0)
        Dim shtFound As Boolean = False
        Dim letter As Char
        Dim ws As Worksheet
        Dim shtCount As Integer = XLpicsWB.Sheets.Count

        For Each ws In XLpicsWB.Sheets      ' does worksheet exist?
            If ws.Name.Equals(shtName) Then
                shtFound = True
                Exit For
            End If
        Next

        If Not shtFound Then        ' create worksheet
            ws = XLpicsWB.Sheets.Add(After:=XLpicsWB.Sheets(shtCount))
            ws.Name = shtName
        Else
            Exit Sub  'worksheet already exits - nothing to do
        End If

        For i = 0 To ColName.Count - 1
            letter = Convert.ToChar(65 + i)  ' starting pt is letter "A"
            ws.Range(letter & "1").Value = ColName(i) ' assign value to cells A1, B1,..
        Next
        With ws.Tab
            .Color = tabColor
            .TintAndShade = 0
        End With

    End Sub

    Sub Build_IOTags_ValveSO_WS()

        Dim shtName As String = "IOTags - ValveSO"
        Dim ColName() As String = {"Name", "Type", "Default Value", "IO Address", "Description"}
        Dim tabColor As Integer = RGB(255, 153, 0)
        Dim shtFound As Boolean = False
        Dim letter As Char
        Dim ws As Worksheet
        Dim shtCount As Integer = XLpicsWB.Sheets.Count

        For Each ws In XLpicsWB.Sheets      ' does worksheet exist?
            If ws.Name.Equals(shtName) Then
                shtFound = True
                Exit For
            End If
        Next

        If Not shtFound Then        ' create worksheet
            ws = XLpicsWB.Sheets.Add(After:=XLpicsWB.Sheets(shtCount))
            ws.Name = shtName
        Else
            Exit Sub  'worksheet already exits - nothing to do
        End If

        For i = 0 To ColName.Count - 1
            letter = Convert.ToChar(65 + i)  ' starting pt is letter "A"
            ws.Range(letter & "1").Value = ColName(i) ' assign value to cells A1, B1,..
        Next
        With ws.Tab
            .Color = tabColor
            .TintAndShade = 0
        End With

    End Sub

    Sub Build_IOTags_ValveMO_WS()

        Dim shtName As String = "IOTags - ValveMO"
        Dim ColName() As String = {"Name", "Type", "Default Value", "IO Address", "Description"}
        Dim tabColor As Integer = RGB(255, 153, 0)
        Dim shtFound As Boolean = False
        Dim letter As Char
        Dim ws As Worksheet
        Dim shtCount As Integer = XLpicsWB.Sheets.Count

        For Each ws In XLpicsWB.Sheets      ' does worksheet exist?
            If ws.Name.Equals(shtName) Then
                shtFound = True
                Exit For
            End If
        Next

        If Not shtFound Then        ' create worksheet
            ws = XLpicsWB.Sheets.Add(After:=XLpicsWB.Sheets(shtCount))
            ws.Name = shtName
        Else
            Exit Sub  'worksheet already exits - nothing to do
        End If

        For i = 0 To ColName.Count - 1
            letter = Convert.ToChar(65 + i)  ' starting pt is letter "A"
            ws.Range(letter & "1").Value = ColName(i) ' assign value to cells A1, B1,..
        Next
        With ws.Tab
            .Color = tabColor
            .TintAndShade = 0
        End With

    End Sub

    Sub Build_IOTags_VSD_WS()

        Dim shtName As String = "IOTags - VSD"
        Dim ColName() As String = {"Name", "Type", "Default Value", "IO Address", "Description"}
        Dim tabColor As Integer = RGB(255, 153, 0)
        Dim shtFound As Boolean = False
        Dim letter As Char
        Dim ws As Worksheet
        Dim shtCount As Integer = XLpicsWB.Sheets.Count

        For Each ws In XLpicsWB.Sheets      ' does worksheet exist?
            If ws.Name.Equals(shtName) Then
                shtFound = True
                Exit For
            End If
        Next

        If Not shtFound Then        ' create worksheet
            ws = XLpicsWB.Sheets.Add(After:=XLpicsWB.Sheets(shtCount))
            ws.Name = shtName
        Else
            Exit Sub  'worksheet already exits - nothing to do
        End If

        For i = 0 To ColName.Count - 1
            letter = Convert.ToChar(65 + i)  ' starting pt is letter "A"
            ws.Range(letter & "1").Value = ColName(i) ' assign value to cells A1, B1,..
        Next
        With ws.Tab
            .Color = tabColor
            .TintAndShade = 0
        End With

    End Sub

    Sub Build_IOMem_AIn_WS()

        Dim shtName As String = "IOMem - AIn"
        Dim ColName() As String = {"Number", "Name", "Type", "Default Value", "IO Address", "Description"}
        Dim tabColor As Integer = RGB(204, 153, 255)
        Dim shtFound As Boolean = False
        Dim letter As Char
        Dim ws As Worksheet
        Dim shtCount As Integer = XLpicsWB.Sheets.Count

        For Each ws In XLpicsWB.Sheets      ' does worksheet exist?
            If ws.Name.Equals(shtName) Then
                shtFound = True
                Exit For
            End If
        Next

        If Not shtFound Then        ' create worksheet
            ws = XLpicsWB.Sheets.Add(After:=XLpicsWB.Sheets(shtCount))
            ws.Name = shtName
        Else
            Exit Sub  'worksheet already exits - nothing to do
        End If

        For i = 0 To ColName.Count - 1
            letter = Convert.ToChar(65 + i)  ' starting pt is letter "A"
            ws.Range(letter & "1").Value = ColName(i) ' assign value to cells A1, B1,..
        Next
        With ws.Tab
            .Color = tabColor
            .TintAndShade = 0
        End With

    End Sub

    Sub Build_IOMem_DIn_WS()

        Dim shtName As String = "IOMem - DIn"
        Dim ColName() As String = {"Number", "Name", "Type", "Default Value", "IO Address", "Description"}
        Dim tabColor As Integer = RGB(204, 153, 255)
        Dim shtFound As Boolean = False
        Dim letter As Char
        Dim ws As Worksheet
        Dim shtCount As Integer = XLpicsWB.Sheets.Count

        For Each ws In XLpicsWB.Sheets      ' does worksheet exist?
            If ws.Name.Equals(shtName) Then
                shtFound = True
                Exit For
            End If
        Next

        If Not shtFound Then        ' create worksheet
            ws = XLpicsWB.Sheets.Add(After:=XLpicsWB.Sheets(shtCount))
            ws.Name = shtName
        Else
            Exit Sub  'worksheet already exits - nothing to do
        End If

        For i = 0 To ColName.Count - 1
            letter = Convert.ToChar(65 + i)  ' starting pt is letter "A"
            ws.Range(letter & "1").Value = ColName(i) ' assign value to cells A1, B1,..
        Next
        With ws.Tab
            .Color = tabColor
            .TintAndShade = 0
        End With

    End Sub

    Sub Build_IOMem_Motor_WS()

        Dim shtName As String = "IOMem - Motor"
        Dim ColName() As String = {"Number", "Name", "Type", "Default Value", "IO Address", "Description"}
        Dim tabColor As Integer = RGB(204, 153, 255)
        Dim shtFound As Boolean = False
        Dim letter As Char
        Dim ws As Worksheet
        Dim shtCount As Integer = XLpicsWB.Sheets.Count

        For Each ws In XLpicsWB.Sheets      ' does worksheet exist?
            If ws.Name.Equals(shtName) Then
                shtFound = True
                Exit For
            End If
        Next

        If Not shtFound Then        ' create worksheet
            ws = XLpicsWB.Sheets.Add(After:=XLpicsWB.Sheets(shtCount))
            ws.Name = shtName
        Else
            Exit Sub  'worksheet already exits - nothing to do
        End If

        For i = 0 To ColName.Count - 1
            letter = Convert.ToChar(65 + i)  ' starting pt is letter "A"
            ws.Range(letter & "1").Value = ColName(i) ' assign value to cells A1, B1,..
        Next
        With ws.Tab
            .Color = tabColor
            .TintAndShade = 0
        End With

    End Sub

    Sub Build_IOMem_ValveC_WS()

        Dim shtName As String = "IOMem - ValveC"
        Dim ColName() As String = {"Number", "Name", "Type", "Default Value", "IO Address", "Description"}
        Dim tabColor As Integer = RGB(204, 153, 255)
        Dim shtFound As Boolean = False
        Dim letter As Char
        Dim ws As Worksheet
        Dim shtCount As Integer = XLpicsWB.Sheets.Count

        For Each ws In XLpicsWB.Sheets      ' does worksheet exist?
            If ws.Name.Equals(shtName) Then
                shtFound = True
                Exit For
            End If
        Next

        If Not shtFound Then        ' create worksheet
            ws = XLpicsWB.Sheets.Add(After:=XLpicsWB.Sheets(shtCount))
            ws.Name = shtName
        Else
            Exit Sub  'worksheet already exits - nothing to do
        End If

        For i = 0 To ColName.Count - 1
            letter = Convert.ToChar(65 + i)  ' starting pt is letter "A"
            ws.Range(letter & "1").Value = ColName(i) ' assign value to cells A1, B1,..
        Next
        With ws.Tab
            .Color = tabColor
            .TintAndShade = 0
        End With

    End Sub

    Sub Build_IOMem_ValveMO_WS()

        Dim shtName As String = "IOMem - ValveMO"
        Dim ColName() As String = {"Number", "Name", "Type", "Default Value", "IO Address", "Description"}
        Dim tabColor As Integer = RGB(204, 153, 255)
        Dim shtFound As Boolean = False
        Dim letter As Char
        Dim ws As Worksheet
        Dim shtCount As Integer = XLpicsWB.Sheets.Count

        For Each ws In XLpicsWB.Sheets      ' does worksheet exist?
            If ws.Name.Equals(shtName) Then
                shtFound = True
                Exit For
            End If
        Next

        If Not shtFound Then        ' create worksheet
            ws = XLpicsWB.Sheets.Add(After:=XLpicsWB.Sheets(shtCount))
            ws.Name = shtName
        Else
            Exit Sub  'worksheet already exits - nothing to do
        End If

        For i = 0 To ColName.Count - 1
            letter = Convert.ToChar(65 + i)  ' starting pt is letter "A"
            ws.Range(letter & "1").Value = ColName(i) ' assign value to cells A1, B1,..
        Next
        With ws.Tab
            .Color = tabColor
            .TintAndShade = 0
        End With

    End Sub

    Sub Build_IOMem_ValveSO_WS()

        Dim shtName As String = "IOMem - ValveSO"
        Dim ColName() As String = {"Number", "Name", "Type", "Default Value", "IO Address", "Description"}
        Dim tabColor As Integer = RGB(204, 153, 255)
        Dim shtFound As Boolean = False
        Dim letter As Char
        Dim ws As Worksheet
        Dim shtCount As Integer = XLpicsWB.Sheets.Count

        For Each ws In XLpicsWB.Sheets      ' does worksheet exist?
            If ws.Name.Equals(shtName) Then
                shtFound = True
                Exit For
            End If
        Next

        If Not shtFound Then        ' create worksheet
            ws = XLpicsWB.Sheets.Add(After:=XLpicsWB.Sheets(shtCount))
            ws.Name = shtName
        Else
            Exit Sub  'worksheet already exits - nothing to do
        End If

        For i = 0 To ColName.Count - 1
            letter = Convert.ToChar(65 + i)  ' starting pt is letter "A"
            ws.Range(letter & "1").Value = ColName(i) ' assign value to cells A1, B1,..
        Next
        With ws.Tab
            .Color = tabColor
            .TintAndShade = 0
        End With

    End Sub

    Sub Build_IOMem_VSD_WS()

        Dim shtName As String = "IOMem - VSD"
        Dim ColName() As String = {"Number", "Name", "Type", "Default Value", "IO Address", "Description"}
        Dim tabColor As Integer = RGB(204, 153, 255)
        Dim shtFound As Boolean = False
        Dim letter As Char
        Dim ws As Worksheet
        Dim shtCount As Integer = XLpicsWB.Sheets.Count

        For Each ws In XLpicsWB.Sheets      ' does worksheet exist?
            If ws.Name.Equals(shtName) Then
                shtFound = True
                Exit For
            End If
        Next

        If Not shtFound Then        ' create worksheet
            ws = XLpicsWB.Sheets.Add(After:=XLpicsWB.Sheets(shtCount))
            ws.Name = shtName
        Else
            Exit Sub  'worksheet already exits - nothing to do
        End If

        For i = 0 To ColName.Count - 1
            letter = Convert.ToChar(65 + i)  ' starting pt is letter "A"
            ws.Range(letter & "1").Value = ColName(i) ' assign value to cells A1, B1,..
        Next
        With ws.Tab
            .Color = tabColor
            .TintAndShade = 0
        End With

    End Sub

    Sub Build_MinMax_AIn_WS()

        Dim shtName As String = "MinMax - AIn"
        Dim ColName() As String = {"Name", "InputMin", "InputMax", "OutputMin", "OutputMax"}
        Dim tabColor As Integer = RGB(0, 0, 0)
        Dim shtFound As Boolean = False
        Dim letter As Char
        Dim ws As Worksheet
        Dim shtCount As Integer = XLpicsWB.Sheets.Count

        For Each ws In XLpicsWB.Sheets      ' does worksheet exist?
            If ws.Name.Equals(shtName) Then
                shtFound = True
                Exit For
            End If
        Next

        If Not shtFound Then        ' create worksheet
            ws = XLpicsWB.Sheets.Add(After:=XLpicsWB.Sheets(shtCount))
            ws.Name = shtName
        Else
            Exit Sub  'worksheet already exits - nothing to do
        End If

        For i = 0 To ColName.Count - 1
            letter = Convert.ToChar(65 + i)  ' starting pt is letter "A"
            ws.Range(letter & "1").Value = ColName(i) ' assign value to cells A1, B1,..
        Next

        'With ws.Tab
        '    .Color = tabColor
        '    .TintAndShade = 0
        'End With

    End Sub

    Sub Build_MinMax_ValveC_WS()

        Dim shtName As String = "MinMax - ValveC"
        Dim ColName() As String = {"Name", "InputMin", "InputMax", "OutputMin", "OutputMax"}
        Dim tabColor As Integer = RGB(0, 0, 0)
        Dim shtFound As Boolean = False
        Dim letter As Char
        Dim ws As Worksheet
        Dim shtCount As Integer = XLpicsWB.Sheets.Count

        For Each ws In XLpicsWB.Sheets      ' does worksheet exist?
            If ws.Name.Equals(shtName) Then
                shtFound = True
                Exit For
            End If
        Next

        If Not shtFound Then        ' create worksheet
            ws = XLpicsWB.Sheets.Add(After:=XLpicsWB.Sheets(shtCount))
            ws.Name = shtName
        Else
            Exit Sub  'worksheet already exits - nothing to do
        End If

        For i = 0 To ColName.Count - 1
            letter = Convert.ToChar(65 + i)  ' starting pt is letter "A"
            ws.Range(letter & "1").Value = ColName(i) ' assign value to cells A1, B1,..
        Next

        'With ws.Tab
        '    .Color = tabColor
        '    .TintAndShade = 0
        'End With

    End Sub

    Sub Build_MinMax_VSD_WS()

        Dim shtName As String = "MinMax - VSD"
        Dim ColName() As String = {"Name", "InputMin", "InputMax", "OutputMin", "OutputMax"}
        Dim tabColor As Integer = RGB(0, 0, 0)
        Dim shtFound As Boolean = False
        Dim letter As Char
        Dim ws As Worksheet
        Dim shtCount As Integer = XLpicsWB.Sheets.Count

        For Each ws In XLpicsWB.Sheets      ' does worksheet exist?
            If ws.Name.Equals(shtName) Then
                shtFound = True
                Exit For
            End If
        Next

        If Not shtFound Then        ' create worksheet
            ws = XLpicsWB.Sheets.Add(After:=XLpicsWB.Sheets(shtCount))
            ws.Name = shtName
        Else
            Exit Sub  'worksheet already exits - nothing to do
        End If

        For i = 0 To ColName.Count - 1
            letter = Convert.ToChar(65 + i)  ' starting pt is letter "A"
            ws.Range(letter & "1").Value = ColName(i) ' assign value to cells A1, B1,..
        Next

        'With ws.Tab
        '    .Color = tabColor
        '    .TintAndShade = 0
        'End With

    End Sub

    Sub Build_Wire_WS()

        Call OpenXLTemplateFN()

        If XLTemplateWB Is Nothing Then ' wire template file not selected
            Exit Sub
        End If

        Dim count As Integer = XLpicsWB.Sheets.Count

        For Each Sheet In XLTemplateWB.Sheets
            Sheet.Copy(After:=XLpicsWB.Sheets(count))

        Next Sheet

        XLTemplateWB.Application.ScreenUpdating = True
        XLTemplateWB.Application.DisplayAlerts = True 'Turn safety alerts back On
        XLTemplateWB.Close(SaveChanges:=False)

    End Sub
    Sub Build_Wire_AIn_WS()


        ' import template from Excel workbook

        'Dim shtName As String = "Wire_AIn Template"
        'Dim ColName() As String = {"Name", "InputMin", "InputMax", "OutputMin", "OutputMax"}
        'Dim tabColor As Integer = RGB(0, 51, 102)
        'Dim shtFound As Boolean = False
        'Dim letter As Char
        'Dim ws As Worksheet
        'Dim shtCount As Integer = XLpicsWB.Sheets.Count

        'For Each ws In XLpicsWB.Sheets      ' does worksheet exist?
        '    If ws.Name.Equals(shtName) Then
        '        shtFound = True
        '        Exit For
        '    End If
        'Next

        'If Not shtFound Then        ' create worksheet
        '    ws = XLpicsWB.Sheets.Add(After:=XLpicsWB.Sheets(shtCount))
        '    ws.Name = shtName
        'Else
        '    Exit Sub  'worksheet already exits - nothing to do
        'End If

        'For i = 0 To ColName.Count - 1
        '    letter = Convert.ToChar(65 + i)  ' starting pt is letter "A"
        '    ws.Range(letter & "1").Value = ColName(i) ' assign value to cells A1, B1,..
        'Next
        'With ws.Tab
        '    .Color = tabColor
        '    .TintAndShade = 0
        'End With

    End Sub

    Sub OpenXLTemplateFN()

        ' Open the Excel Template file (workbook) for Wire file generation

        Dim title = "Open - Select the Excel Template for Wire files"
        Dim filter = "Excel Files (*.xltx),*.xltx"
        Dim fn As String
        Dim response As MsgBoxResult

        MsgBox("Select Excel Wire Template file: " & Chr(34) & " PICS_Wire_Template.xltx" & Chr(34), vbOKOnly)

        Do
            fn = XLApp.GetOpenFilename(FileFilter:=filter, FilterIndex:=2, Title:=title)
            If IsNothing(fn) Then     ' Cancel button pressed
                response = MsgBox("Cancel Operation?", vbYesNo)
                If response = vbYes Then Exit Sub
            End If
        Loop Until Not IsNothing(fn)

        XLTemplateWB = XLApp.Workbooks.Open(fn)

    End Sub

End Module
