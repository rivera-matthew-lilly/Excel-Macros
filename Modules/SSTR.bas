Attribute VB_Name = "SSTR"
Sub CalculateRatioAndCreateDataBreakdown()

    CalculateRatioOnSplitPlateData
    
    CreateDataBreakdownSheet
End Sub

Sub CreateDataBreakdownSheet()
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim i As Integer, j As Integer
    Dim colOffset As Integer

    ' Check if "Ratio Calculations" sheet exists
    On Error Resume Next
    Set wsSource = Worksheets("Ratio Calculations")
    On Error GoTo 0

    If wsSource Is Nothing Then
        MsgBox "Ratio Calculations sheet does not exist. Please run the ratio calculations first."
        Exit Sub
    End If

    ' Delete "Data Breakdown" sheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    Set wsTarget = Sheets("Data Breakdown")
    If Not wsTarget Is Nothing Then wsTarget.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Create and name new sheet
    Set wsTarget = Worksheets.Add
    wsTarget.Name = "Data Breakdown"

    ' Copy cAMP ranges and add labels
    For j = 1 To 3
        wsTarget.Cells(1, j).Value = "cAMP Rep " & j
        For i = 2 To 21
            wsTarget.Cells(i, j).Value = wsSource.Cells((j - 1) * 2 + 2, i).Value
        Next i
    Next j

    ' Copy SST14 ranges and add labels
    wsTarget.Cells(1, 21).Value = "SST14 Rep 1"
    For i = 2 To 13
        wsTarget.Cells(i, 21).Value = wsSource.Cells(8, i * 2 - 2).Value
    Next i

    wsTarget.Cells(1, 22).Value = "SST14 Rep 2"
    For i = 2 To 13
        wsTarget.Cells(i, 22).Value = wsSource.Cells(8, i * 2 - 1).Value
    Next i
    
    ' Copy Stim and Non-Stim ranges and add labels
    wsTarget.Cells(15, 5).Value = "Stim"
    For i = 16 To 40
        wsTarget.Cells(i, 5).Value = wsSource.Cells(10, i - 14).Value
    Next i

    wsTarget.Cells(15, 7).Value = "Non-Stim"
    For i = 16 To 40
        wsTarget.Cells(i, 7).Value = wsSource.Cells(12, i - 14).Value
    Next i

    ' Copy Peptide ranges and add labels
    For i = 1 To 8
        colOffset = 5 + (i - 1) * 2
        wsTarget.Cells(1, colOffset).Value = "Peptide_" & i & " Rep 1"
        wsTarget.Cells(1, colOffset + 1).Value = "Peptide_" & i & " Rep 2"
        
        For j = 2 To 13
            wsTarget.Cells(j, colOffset).Value = wsSource.Cells(i * 2 + 1, j * 2 - 2).Value
            wsTarget.Cells(j, colOffset + 1).Value = wsSource.Cells(i * 2 + 1, j * 2 - 1).Value
        Next j
    Next i

    ' Notify the user that the macro is completed
    MsgBox "Data Breakdown sheet has been created successfully and all data has been copied."
End Sub


Sub CalculateRatioOnCombinedPlateData()
    Dim ws As Worksheet
    Dim newSheet As Worksheet
    Dim wb As Workbook
    Dim i As Integer, j As Integer
    Dim x As Double, y As Double, ratio As Double

    Set wb = ThisWorkbook
    Set ws = Sheets(1)

    ' Delete "Ratio Calculations" sheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    Set newSheet = Sheets("Ratio Calculations")
    If Not newSheet Is Nothing Then newSheet.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Create and name new sheet
    Set newSheet = Sheets.Add(After:=Sheets(Sheets.Count))
    newSheet.Name = "Ratio Calculations"

    ' Label columns B1:Y1 with numbers 1 to 24
    For i = 1 To 24
        newSheet.Cells(1, i + 1).Value = i
    Next i

    ' Label rows A2:A17 with letters A to P
    For i = 1 To 16
        newSheet.Cells(i + 1, 1).Value = Chr(64 + i)
    Next i

    ' Calculate ratios and fill the table
    For j = 3 To 26
        For i = 34 To 65 Step 2
            x = ws.Cells(i + 1, j).Value
            y = ws.Cells(i, j).Value
            ratio = IIf(y <> 0, (x / y) * 10 ^ 4, 0)
            newSheet.Cells((i - 34) / 2 + 2, j - 1).Value = ratio
        Next i
    Next j
End Sub

Sub CalculateRatioOnSplitPlateData()
    Dim ws As Worksheet
    Dim newSheet As Worksheet
    Dim wb As Workbook
    Dim i As Integer, j As Integer
    Dim x As Double, y As Double, ratio As Double

    Set wb = ThisWorkbook
    Set ws = Sheets(1)

    ' Delete "Ratio Calculations" sheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    Set newSheet = Sheets("Ratio Calculations")
    If Not newSheet Is Nothing Then newSheet.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Create and name new sheet
    Set newSheet = Sheets.Add(After:=Sheets(Sheets.Count))
    newSheet.Name = "Ratio Calculations"

    ' Label columns B1:Y1 with numbers 1 to 24
    For i = 1 To 24
        newSheet.Cells(1, i + 1).Value = i
    Next i

    ' Label rows A2:A17 with letters A to P
    For i = 1 To 16
        newSheet.Cells(i + 1, 1).Value = Chr(64 + i)
    Next i

    ' Calculate ratios and fill the table
    For j = 3 To 26
        For i = 33 To 48 Step 1
            x = ws.Cells(i + 21, j).Value
            y = ws.Cells(i, j).Value
            ratio = IIf(y <> 0, (x / y) * 10 ^ 4, 0)
            newSheet.Cells((i - 33) / 1 + 2, j - 1).Value = ratio
        Next i
    Next j
End Sub
