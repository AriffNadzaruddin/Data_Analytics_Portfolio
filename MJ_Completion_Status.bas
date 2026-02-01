Attribute VB_Name = "MJ_Completion_Status"
' Declare workbook and worksheet objects
Public srcWB As Workbook, srcWS As Worksheet
Public masterWB As Workbook, masterWS As Worksheet
Public lastRowSrc As Long, lastTPMaster As Long
Public startCol As Long, endCol As Long
Public i As Long, col As Long

Sub System_MaintenanceJob_Status()
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim srcPath As String
    
    '--- Set the File Path --'
    srcPath = ThisWorkbook.Path & "\source_data\MJ Status.xlsx"
    
    '--- Open Source File ---'
    Set srcWB = Workbooks.Open(srcPath)
    Set srcWS = srcWB.Worksheets("Data Export")
    
    
    '--- Reference Master File ---'
    Set masterWB = ThisWorkbook
    Set masterWS = masterWB.Worksheets("Job Planning")
    
    '--- Build Dictionary from Source File ---'
    lastRowSrc = srcWS.Cells(srcWS.Rows.Count, "C").End(xlUp).Row
    Set MJDict = CreateObject("Scripting.Dictionary")
    
    For i = 2 To lastRowSrc
        MJTag = Trim(srcWS.Cells(i, "C").Value)
        MJStatus = srcWS.Cells(i, "B").Value
        
        If MJTag <> "" Then
            MJDict(MJTag) = MJStatus
        End If
    Next i
    
    '--- Update Master File ---'
    lastRowMaster = masterWS.Cells(masterWS.Rows.Count, "B").End(xlUp).Row
    
    For i = 4 To lastRowMaster
        MJTag = Trim(masterWS.Cells(i, "B").Value)
        
        If MJDict.Exists(MJTag) Then
            masterWS.Cells(i, "C").Value = MJDict(MJTag)
        Else
            masterWS.Cells(i, "C").Value = ""
        End If
    Next i
    
    MsgBox "Master file Maintenance Job Status updated successfully!", vbInformation
    
    '--- Close Source File ---'
    srcWB.Close False

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub


