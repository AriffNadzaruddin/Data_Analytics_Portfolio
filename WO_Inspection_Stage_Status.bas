Attribute VB_Name = "WO_Inspection_Stage_Status"
' Declare workbook and worksheet objects
Public srcWB As Workbook, srcWS As Worksheet
Public masterWB As Workbook, masterWS As Worksheet
Public lastRowSrc As Long, lastTPMaster As Long
Public startCol As Long, endCol As Long
Public WOList As Variant, StageHeaders As Variant
Public comboDict As Object
Public i As Long, col As Long
Public combo As String
    
Sub System_WorkOrder_InspectionStage_Status()

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    Dim srcPath As String
    
    '--- Set the File Path --'
    srcPath = ThisWorkbook.Path & "\source_data\Inspection Stage Status.xlsx"
    
    '--- Open Source File ---
    Set srcWB = Workbooks.Open(srcPath)
    Set srcWS = srcWB.Worksheets("Data Export")
    

    '--- Reference Master File ---
    Set masterWB = ThisWorkbook
    Set masterWS = masterWB.Worksheets("Job Planning")
    
    '--- Build Dictionary from Source File ---
    lastRowSrc = srcWS.Cells(srcWS.Rows.Count, "G").End(xlUp).Row
    Set comboDict = CreateObject("Scripting.Dictionary")
    
    ' Loop through source rows and store Work Order>Stage -> Status in dictionary
    For i = 2 To lastRowSrc
        Dim WOVal As String, StageVal As String, statusVal As String
        WOVal = Trim(srcWS.Cells(i, "G").Value)
        StageVal = Trim(srcWS.Cells(i, "F").Value)
        statusVal = srcWS.Cells(i, "B").Value
        
        If WOVal <> "" And StageVal <> "" Then
            comboDict(WOVal & ">" & StageVal) = statusVal
        End If
    Next i
    
    '--- Get Master File Ranges ---
    lastTPMaster = masterWS.Cells(masterWS.Rows.Count, "G").End(xlUp).Row
    
    ' Define Inspection Stage columns (H to K)
    startCol = masterWS.Range("H3").Column
    endCol = masterWS.Range("K3").Column
    
    ' Load  WO list and Inspection Stage headers into arrays
    WOList = masterWS.Range("G4:G" & lastTPMaster).Value
    StageHeaders = masterWS.Range(masterWS.Cells(3, startCol), masterWS.Cells(3, endCol)).Value
    
    '--- Prepare Result Array ---
    Dim resultArr() As Variant
    ReDim resultArr(1 To UBound(WOList), 1 To UBound(StageHeaders, 2))
    
    '--- Fill Result Array using Dictionary ---
    For i = 1 To UBound(WOList)
        For col = 1 To UBound(StageHeaders, 2)
            If StageHeaders(1, col) <> "" Then
                combo = WOList(i, 1) & ">" & StageHeaders(1, col)
                If comboDict.Exists(combo) Then
                    resultArr(i, col) = comboDict(combo)
                Else
                    resultArr(i, col) = "N/A"
                End If
            End If
        Next col
    Next i
    
    '--- Write Back to Master File ---
    masterWS.Range(masterWS.Cells(4, startCol), masterWS.Cells(lastTPMaster, endCol)).Value = resultArr
            
    MsgBox "Master file Inspection Stage Status updated successfully!", vbInformation
    
    '--- Close Source File ---
    srcWB.Close SaveChanges:=True

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub
