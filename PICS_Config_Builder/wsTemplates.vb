
Imports Microsoft.Office.Interop
Imports Microsoft.Office.Interop.Excel

Module wsTemplates

    Sub Build_PICS_WB()

        Call Build_IOTags_AIn()
        'Call Build_IOTags_DIn()
        'Call Build_IOTags_DOut()
        'Call Build_IOTags_Motor()
        'Call Build_ITTags_ValveC()
        'Call Build_IOTags_ValveSO()
        'Call Build_IOTags_ValveMO()
        'Call Build_IOTags_VSD

        'Call Build_IOMem_AIn()
        'Call Build_IOMem_DIn()
        'Call Build_IOMem_Motor()
        'Call Build_IOMem_ValveC()
        'Call Build_IOMem_ValveMO()
        'Call Build_IOMem_ValveSO()
        'Call Build_IOMem_VSD()

        'Call Build_MinMax_AIn()
        'Call Build_MinMax_ValveC()
        'Call Build_MinMax_VSD()

        'Call Build_Wire_AIn()
        'Call Build_Wire_DIn()
        'Call Build_Wire_Motor()
        'Call Build_Wire_ValveC()
        'Call Build_Wire_ValveMO()
        'Call Build_Wire_ValveSO()
        'Call Build_Wire_VSD()

    End Sub

    Sub Build_IOTags_AIn()

        Dim ws As Worksheet
        Dim ColName() As String = {"Name", "Type", "Default Value", "IO Address", "Description", "RGB(255,153,0)"}
        ws = XLpicsWB.Sheets.Add(After:="IOTags - AIn")
        ws.Tab.Color = "RGB(255,153,0)"
        For i = 0 To ColName.Count - 2
            Character = Convert.ToChar(65 + i)
            ws.Range()
        Next
        ws.Range("A1").Value = "Name"
        ws.Range("B1").Value = "Type"
        ws.Range("C1").Value = "Default Value"
        ws.Range("D1").Value = "IO Address"
        ws.Range("E1").Value = "Description"


    End Sub
End Module
