VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSSTRDataTool 
   Caption         =   "SSTR Data Tool"
   ClientHeight    =   5980
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   11430
   OleObjectBlob   =   "frmSSTRDataTool.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSSTRDataTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Button_Browse_Click()
    Dim fd As FileDialog
    Dim selectedFile As String
    Dim i As Integer
    Dim fileAlreadyExists As Boolean
    
    ' Initialize the FileDialog object as a file picker
    Set fd = Application.FileDialog(msoFileDialogFilePicker)
    
    ' Set the title of the dialog
    fd.Title = "Select a File"
    
    ' Allow the user to select only one file
    fd.AllowMultiSelect = False
    
    ' Show the dialog. If the user selects a file, save the file path
    If fd.Show = -1 Then
        selectedFile = fd.SelectedItems(1)
        
        ' Check if the file is already in the list box
        fileAlreadyExists = False
        For i = 0 To FileListBox.ListCount - 1
            If FileListBox.List(i) = selectedFile Then
                fileAlreadyExists = True
                Exit For
            End If
        Next i
        
        ' If the file is not in the list box, add it; otherwise, show a warning
        If fileAlreadyExists Then
            MsgBox "This file has already been added to the list.", vbExclamation, "Duplicate File"
        Else
            FileListBox.AddItem selectedFile
        End If
    Else
        MsgBox "No file selected"
    End If
    
    ' Cleanup
    Set fd = Nothing
End Sub

Private Sub Button_DeleteSelectedFile_Click()
    Dim selectedIndex As Integer
    
    ' Check if an item is selected
    If FileListBox.ListIndex = -1 Then
        MsgBox "No file selected. Please select a file from the list to remove.", vbExclamation, "No Selection"
        Exit Sub
    End If
    
    ' Get the selected index
    selectedIndex = FileListBox.ListIndex
    
    ' Remove the selected item
    FileListBox.RemoveItem selectedIndex
End Sub

Private Sub Button_Submit_Click()
    ' Hide the main form to avoid conflicts with the progress form
    Me.Hide

    ' Initialize and show the progress form non-modally
    frmProgress.lblStatus.Caption = "Starting processing..."
    frmProgress.Show vbModeless ' Show the form non-modally

    ' Call the method to process the files
    ProcessFilesInListBox
    
    ' Notify the user that the process is completed
    frmProgress.lblStatus.Caption = "Processing complete. Files have been saved with 'Processed' suffix."
    MsgBox "Processing complete. Files have been saved with 'Processed' suffix."
    
    ' Unload the progress form
    Unload frmProgress

    ' Show the main form again
    Me.Show
End Sub

' Update CalculateRatioOnSplitPlateData to accept a workbook as a parameter
Sub CalculateRatioOnSplitPlateData(wb As Workbook)
    Dim ws As Worksheet
    Dim newSheet As Worksheet
    Dim i As Integer, j As Integer
    Dim x As Double, y As Double, ratio As Double

    Set ws = wb.Sheets(1)

    ' Delete "Ratio Calculations" sheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    Set newSheet = wb.Sheets("Ratio Calculations")
    If Not newSheet Is Nothing Then newSheet.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Create and name new sheet
    Set newSheet = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
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

' Update CreateDataBreakdownSheet to accept a workbook as a parameter
Sub CreateDataBreakdownSheet(wb As Workbook)
    Dim wsSource As Worksheet
    Dim wsTarget As Worksheet
    Dim i As Integer, j As Integer
    Dim colOffset As Integer

    ' Check if "Ratio Calculations" sheet exists
    On Error Resume Next
    Set wsSource = wb.Worksheets("Ratio Calculations")
    On Error GoTo 0

    If wsSource Is Nothing Then
        MsgBox "Ratio Calculations sheet does not exist in " & wb.Name & ". Please run the ratio calculations first."
        Exit Sub
    End If

    ' Delete "Data Breakdown" sheet if it exists
    On Error Resume Next
    Application.DisplayAlerts = False
    Set wsTarget = wb.Sheets("Data Breakdown")
    If Not wsTarget Is Nothing Then wsTarget.Delete
    Application.DisplayAlerts = True
    On Error GoTo 0

    ' Create and name new sheet
    Set wsTarget = wb.Worksheets.Add
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
    ' MsgBox "Data Breakdown sheet has been created successfully and all data has been copied for " & wb.Name
End Sub


Private Sub ProcessFilesInListBox()
    Dim i As Integer
    Dim filePath As String
    Dim wb As Workbook
    Dim originalDir As String
    Dim processedDir As String
    Dim savePath As String
    Dim totalFiles As Integer
    Dim userChoice As VbMsgBoxResult

    ' Get the total number of files in the list box
    totalFiles = FileListBox.ListCount

    ' Loop through each file in the list box
    For i = 0 To totalFiles - 1
        filePath = FileListBox.List(i)
        
        ' Update the progress status
        frmProgress.lblStatus.Caption = "Processing file " & (i + 1) & " of " & totalFiles & ": " & filePath
        DoEvents ' This allows the form to update
        
        ' Open the workbook
        Set wb = Workbooks.Open(filePath)
        
        ' Get the directory of the original file
        originalDir = Left(filePath, InStrRev(filePath, "\"))
        
        ' Define the directory to save the processed files
        processedDir = originalDir & "Processed Files\"
        
        ' Create the directory if it does not exist
        If Dir(processedDir, vbDirectory) = "" Then
            MkDir processedDir
        End If
        
        ' Run the macros on the opened workbook
        CalculateRatioOnSplitPlateData wb
        CreateDataBreakdownSheet wb

        ' Define the save path (within the Processed Files directory)
        savePath = processedDir & Mid(filePath, InStrRev(filePath, "\") + 1, InStrRev(filePath, ".") - InStrRev(filePath, "\") - 1) & "Processed" & ".xlsx" ' Change to ".xlsm" if macro-enabled is required

        ' Check if the processed file already exists
        If Dir(savePath) <> "" Then
            ' Prompt the user for action
            userChoice = MsgBox("The file '" & savePath & "' already exists. Do you want to overwrite it?", vbYesNoCancel + vbExclamation, "File Already Exists")

            Select Case userChoice
                Case vbYes
                    ' Overwrite the file
                    wb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook ' Use xlOpenXMLWorkbookMacroEnabled for .xlsm
                Case vbNo
                    ' Skip saving this file
                    MsgBox "Skipping file: " & savePath, vbInformation, "Skipped"
                Case vbCancel
                    ' Cancel the entire process
                    MsgBox "Process canceled by user.", vbCritical, "Canceled"
                    wb.Close SaveChanges:=False
                    Exit Sub
            End Select
        Else
            ' Save the workbook with the new name in the Processed Files directory
            wb.SaveAs Filename:=savePath, FileFormat:=xlOpenXMLWorkbook ' Use xlOpenXMLWorkbookMacroEnabled for .xlsm
        End If
        
        ' Close the workbook
        wb.Close SaveChanges:=False
    Next i
End Sub

