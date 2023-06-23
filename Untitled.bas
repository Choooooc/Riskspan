Attribute VB_Name = "Module1"
Option Explicit

'test here is the update

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''`
'001 - Hide all selected sheets'''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub HideAllSelectedSheets()

'Create variable to hold worksheets
Dim ws As Worksheet

'Ignore error if trying to hide the last worksheet
On Error Resume Next

'Loop through each worksheet in the active workbook
For Each ws In ActiveWindow.SelectedSheets

    'Hide each sheet
    ws.Visible = xlSheetHidden

Next ws

'Allow errors to appear
On Error GoTo 0

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'002 - Unhide all worksheets''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub UnhideAllWorksheets()

'Create variable to hold worksheets
Dim ws As Worksheet

'Loop through each worksheet in the active workbook
For Each ws In ActiveWorkbook.Worksheets

    'Unhide each sheet
    ws.Visible = xlSheetVisible

Next ws

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'003 - Protect all selected worksheets''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub ProtectSelectedWorksheets()

Dim ws As Worksheet
Dim sheetArray As Variant
Dim myPassword As Variant

'Set the password
myPassword = Application.InputBox(prompt:="Enter password", _
    Title:="Password", Type:=2)

'The User clicked Cancel
If myPassword = False Then Exit Sub

'Capture the selected sheets
Set sheetArray = ActiveWindow.SelectedSheets

'Loop through each worksheet in the active workbook
For Each ws In sheetArray

    On Error Resume Next
    
    'Select the worksheet
    ws.Select

    'Protect each worksheet
    ws.Protect Password:=myPassword
    
    On Error GoTo 0
        
Next ws

sheetArray.Select

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'004 - Unprotect all worksheets with a password'''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub UnprotectAllWorksheets()

'Create a variable to hold worksheets
Dim ws As Worksheet

'Create a variable to hold the password
Dim myPassword As Variant

'Set the password
myPassword = Application.InputBox(prompt:="Enter password", _
    Title:="Password", Type:=2)

'The User clicked Cancel
If myPassword = False Then Exit Sub

'Loop through each worksheet in the active workbook
For Each ws In ActiveWindow.SelectedSheets

    'Protect each worksheet
    ws.Unprotect Password:=myPassword
        
Next ws

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'005 - Lock only cells with formulas''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub LockOnlyCellsWithFormulas()

'Create a variable to hold the password
Dim myPassword As Variant

'If more than one worksheet selected exit the macro
If ActiveWindow.SelectedSheets.Count > 1 Then
    
    'Display error message and exit macro
    MsgBox "Select one worksheet and try again"
    Exit Sub

End If

'Set the password
myPassword = Application.InputBox(prompt:="Enter password", _
    Title:="Password", Type:=2)

'The User clicked Cancel
If myPassword = False Then Exit Sub

'All the following to apply to active sheet
With ActiveSheet
   
    'Ignore errors caused by incorrect passwords
    On Error Resume Next
   
    'Unprotect the active sheet
    .Unprotect Password:=myPassword
   
    'If error occured then exit macro
    If Err.Number <> 0 Then
    
        'Display message then exit
        MsgBox "Incorrect password"
        Exit Sub
    
    End If
    
    'Turn error checking back on
    On Error GoTo 0
    
    'Remove lock setting from all cells
    .Cells.Locked = False
   
    'Add lock setting to all cells
    .Cells.SpecialCells(xlCellTypeFormulas).Locked = True
   
    'Protect the active sheet
    .Protect Password:=myPassword
   
End With

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'006 - Hide formulas when protected'''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub HideFormulasWhenProtected()
 
'Create a variable to hold the password
Dim myPassword As String
 
'Set the password
myPassword = "myPassword"
 
'All the following to apply to active sheet
With ActiveSheet
    
   'Unprotect the active sheet
   .Unprotect Password:=myPassword

    'Hide formulas in all cells
   .Cells.FormulaHidden = True

   'Protect the active sheet
   .Protect Password:=myPassword

End With
 
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'007 - Save time stamped backup file''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub SaveTimeStampedBackup()
 
'Create variable to hold the new file path
Dim saveAsName As String
 
'Set the file path
saveAsName = ActiveWorkbook.Path & "\" & _
    Format(Now, "yymmdd-hhmmss") & " " & ActiveWorkbook.Name
 
'Save the workbook
ActiveWorkbook.SaveCopyAs Filename:=saveAsName
 
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'008 - Prepare workbook for saving''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub PrepareWorkbookForSaving()

'Declare the worksheet variable
Dim ws As Worksheet

'Loop through each worksheet in the active workbook
For Each ws In ActiveWorkbook.Worksheets
    
    'Activate each sheet
    ws.Activate
    
    'Close all of groups
    ws.Outline.ShowLevels RowLevels:=1, ColumnLevels:=1

    'Set the view settings to normal
    ActiveWindow.View = xlNormalView
        
    'Remove the gridlines
    ActiveWindow.DisplayGridlines = False

    'Remove the headings on each of the worksheets
    ActiveWindow.DisplayHeadings = False
    
    'Get worksheet to display top left
    ws.Cells(1, 1).Select
    
Next ws

'Find the first visible worksheet and select it
For Each ws In Worksheets

    If ws.Visible = xlSheetVisible Then
        
        'Select the first visible worksheet
        ws.Select
        
        'Once the first visible worksheet is found exit the sub
        Exit For
    
    End If

Next ws

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'009 - Convert merged cells to center across''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub ConvertMergedCellsToCenterAcross()

Dim c As Range
Dim mergedRange As Range

'Loop through all cells in Used range
For Each c In ActiveSheet.UsedRange

    'If merged and single row
    If c.MergeCells = True And c.MergeArea.Rows.Count = 1 Then

        'Set variable for the merged range
        Set mergedRange = c.MergeArea

        'Unmerge the cell and apply Centre Across Selection
        mergedRange.UnMerge
        mergedRange.HorizontalAlignment = xlCenterAcrossSelection

    End If

Next

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'010 - Fit selection to screen''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub FitSelectionToScreen()
 
'To zoom to a specific area, then select the cells
Range("A1:I15").Select
 
'Zoom to selection
ActiveWindow.Zoom = True
 
'Select first cell on worksheet
Range("A1").Select
 
End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'011 - Flip number signage on selected cells''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub FlipNumberSignage()

'Create variable to hold cells in the worksheet
Dim c As Range

'Loop through each cell in selection
For Each c In Selection

    'Test if the cell contents is a number
    If IsNumeric(c) Then

        'Convert signage for each cell
        c.Value = -c.Value
    
    End If

Next c

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'012 - Clear all data cells'''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub ClearAllDataCellsInSelection()

'Clear all hardcoded values in the selected range
Selection.SpecialCells(xlCellTypeConstants).ClearContents

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'013 - Add Prefix to each cell in selected range''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub AddPrefix()

Dim c As Range
Dim prefixValue As Variant

'Display inputbox to collect prefix text
prefixValue = Application.InputBox(prompt:="Enter prefix:", _
    Title:="Prefix", Type:=2)

'The User clicked Cancel
If prefixValue = False Then Exit Sub

'Loop through each cellin selection
For Each c In Selection

    'Add prefix where cell is not a formula or blank
    If Not c.HasFormula And c.Value <> "" Then

        c.Value = prefixValue & c.Value

    End If

Next

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'014 - Add Suffix to each cell in selected range''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub AddSuffix()

Dim c As Range
Dim suffixValue As Variant

'Display inputbox to collect prefix text
suffixValue = Application.InputBox(prompt:="Enter Suffix:", _
    Title:="Suffix", Type:=2)

'The User clicked Cancel
If suffixValue = False Then Exit Sub

'Loop through each cellin selection
For Each c In Selection

    'Add Suffix where cell is not a formula or blank
    If Not c.HasFormula And c.Value <> "" Then

        c.Value = c.Value & suffixValue

    End If

Next

End Sub



''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'015 - Reverse row order''''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub ReverseRows()

'Create variables
Dim rng As Range
Dim rngArray As Variant
Dim tempRng As Variant
Dim i As Long
Dim j As Long
Dim k As Long

'Record the selected range and it's contents
Set rng = Selection
rngArray = rng.Formula

'Loop through all cells and create a temporary array
For j = 1 To UBound(rngArray, 2)
    k = UBound(rngArray, 1)
    For i = 1 To UBound(rngArray, 1) / 2
        tempRng = rngArray(i, j)
        rngArray(i, j) = rngArray(k, j)
        rngArray(k, j) = tempRng
        k = k - 1
    Next
Next

'Apply the array
rng.Formula = rngArray

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'016 - Reverse column order'''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub ReverseColumns()

'Create variables
Dim rng As Range
Dim rngArray As Variant
Dim tempRng As Variant
Dim i As Long
Dim j As Long
Dim k As Long

'Record the selected range and it's contents
Set rng = Selection
rngArray = rng.Formula

'Loop through all cells and create a temporary array
For i = 1 To UBound(rngArray, 1)
    k = UBound(rngArray, 2)
    For j = 1 To UBound(rngArray, 2) / 2
        tempRng = rngArray(i, j)
        rngArray(i, j) = rngArray(i, k)
        rngArray(i, k) = tempRng
        k = k - 1
    Next
Next

'Apply the array
rng.Formula = rngArray

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'017 - Transpose Selection''''''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub TransposeSelection()

'Create variables
Dim rng As Range
Dim rngArray As Variant
Dim i As Long
Dim j As Long
Dim overflowRng As Range
Dim msgAns As Long

'Record the selected range and it's contents
Set rng = Selection
rngArray = rng.Formula

'Test the range and identify if any cells will be overwritten
If rng.Rows.Count > rng.Columns.Count Then

    Set overflowRng = rng.Cells(1, 1). _
        Offset(0, rng.Columns.Count). _
        Resize(rng.Columns.Count, _
        rng.Rows.Count - rng.Columns.Count)

ElseIf rng.Rows.Count < rng.Columns.Count Then

    Set overflowRng = rng.Cells(1, 1).Offset(rng.Rows.Count, 0). _
        Resize(rng.Columns.Count - rng.Rows.Count, rng.Rows.Count)

End If

If rng.Rows.Count <> rng.Columns.Count Then

    If Application.WorksheetFunction.CountA(overflowRng) > 0 Then

    msgAns = MsgBox("Worksheet data in " & overflowRng.Address & _
        " will be overwritten." & vbNewLine & _
        "Do you wish to continue?", vbYesNo)
    
    If msgAns = vbNo Then Exit Sub

    End If

End If

'Clear the rnage
rng.Clear

'Reapply the cells in transposted position
For i = 1 To UBound(rngArray, 1)

    For j = 1 To UBound(rngArray, 2)

        rng.Cells(1, 1).Offset(j - 1, i - 1) = rngArray(i, j)

    Next

Next

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'018 - Create red box around selected areas'''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub AddRedBox()

Dim redBox As Shape
Dim selectedAreas As Range
Dim i As Integer
Dim tempShape As Shape

'Loop through each selected area in active sheet
For Each selectedAreas In Selection.Areas

    'Create a rectangle
    Set redBox = ActiveSheet.Shapes.AddShape(msoShapeRectangle, _
        selectedAreas.Left, selectedAreas.Top, _
        selectedAreas.Width, selectedAreas.Height)

    'Change attributes of shape created
    redBox.Line.ForeColor.RGB = RGB(255, 0, 0)
    redBox.Line.Weight = 2
    redBox.Fill.Visible = msoFalse
    
    'Loop to find a unique shape name
    Do
        i = i + 1
        Set tempShape = Nothing
        
        On Error Resume Next
        Set tempShape = ActiveSheet.Shapes("RedBox_" & i)
        On Error GoTo 0
    
    Loop Until tempShape Is Nothing
    
    'Rename the shape
    redBox.Name = "RedBox_" & i
    
Next

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'019 - Delete all red boxes on active sheet'''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub DeleteRedBox()

Dim shp As Shape

'Loop through each shape on active sheet
For Each shp In ActiveSheet.Shapes

    'Find shapes with a name starting with "RedBox_"
    If Left(shp.Name, 7) = "RedBox_" Then

        'Delete the shape
        shp.Delete

    End If

Next shp

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'020 - Save selected chart as an image''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub ExportSingleChartAsImage()

'Create a variable to hold the path and name of image
Dim imagePath As String
Dim cht As Chart

imagePath = "C:\Users\marks\Documents\myImage.png"
Set cht = ActiveChart

'Export the chart
cht.Export (imagePath)

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'021 - Resize all charts to active chart''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub ResizeAllCharts()

'Create variables to hold chart dimensions
Dim chtHeight As Long
Dim chtWidth As Long

'Create variable to loop through chart objects
Dim chtObj As ChartObject

'Get the size of the first selected chart
chtHeight = ActiveChart.Parent.Height
chtWidth = ActiveChart.Parent.Width

For Each chtObj In ActiveSheet.ChartObjects

    chtObj.Height = chtHeight
    chtObj.Width = chtWidth

Next chtObj

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'022 - Refresh all Pivot Tables in the workbook'''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub RefreshAllPivotTables()

'Refresh all pivot tables
ActiveWorkbook.RefreshAll

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'023 - Pivot Table turn off auto fit columns''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub TurnOffAutofitColumns()

'Create a variable to hold worksheets
Dim ws As Worksheet

'Create a variable to hold pivot tables
Dim pvt As PivotTable

'Loop through each sheet in the activeworkbook
For Each ws In ActiveWorkbook.Worksheets
  
    'Loop through each pivot table in the worksheet
    For Each pvt In ws.PivotTables
    
        'Turn off auto fit columns on PivotTable
        pvt.HasAutoFormat = False
    
    Next pvt
    
Next ws

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'024 - Get color codes for cell fill color''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub GetColorCodeFromCellFill()

'Create variables hold the color data
Dim fillColor As Long
Dim R As Integer
Dim G As Integer
Dim B As Integer
Dim Hex As String

'Get the fill color
fillColor = ActiveCell.Interior.Color

'Convert fill color to RGB
R = (fillColor Mod 256)
G = (fillColor \ 256) Mod 256
B = (fillColor \ 65536) Mod 256

'Convert fill color to Hex
Hex = "#" & Application.WorksheetFunction.Dec2Hex(fillColor)

'Display fill color codes
MsgBox "Color codes for active cell" & vbNewLine & _
    "R:" & R & ", G:" & G & ", B:" & B & vbNewLine & _
    "Hex: " & Hex, Title:="Color Codes"

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'025 - Create a table of contents'''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub CreateTableOfContents()

Dim i As Long
Dim TOCName As String

'Name of the Table of contents
TOCName = "TOC"

'Delete the existing Table of Contents sheet if it exists
On Error Resume Next
Application.DisplayAlerts = False
ActiveWorkbook.Sheets(TOCName).Delete
Application.DisplayAlerts = True
On Error GoTo 0

'Create a new worksheet
ActiveWorkbook.Sheets.Add before:=ActiveWorkbook.Worksheets(1)
ActiveSheet.Name = TOCName

'Loop through the worksheets
For i = 1 To Sheets.Count

    'Create the table of contents
    ActiveSheet.Hyperlinks.Add _
        Anchor:=ActiveSheet.Cells(i, 1), _
        Address:="", _
        SubAddress:="'" & Sheets(i).Name & "'!A1", _
        ScreenTip:=Sheets(i).Name, _
        TextToDisplay:=Sheets(i).Name

Next i


End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'026 - Excel to speak the cell contents'''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub SpeakCellContents()

'Speak the selected cells
Selection.Speak

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'027 - Fix the range of cells which can be scrolled'''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub FixScrollRange()

'This macro will fix the scroll area to the selected cells

If Selection.Cells.Count = 1 Then
    
    'If one cell selected, then reset
    ActiveSheet.ScrollArea = ""

Else

    'Set the scroll area to the selected cells
    ActiveSheet.ScrollArea = Selection.Address

End If

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'028 - Invert sheet selection'''''''''''''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub InvertSheetSelection()

'Create variable to hold list of selected worksheet
Dim selectedList As String

'Create variable to hold worksheets
Dim ws As Worksheet

'Create variable to switch after the first sheet selected
Dim firstSheet As Boolean

'Convert selected sheest to a text string
For Each ws In ActiveWindow.SelectedSheets
    selectedList = selectedList & ws.Name & "[|]"
Next ws

'Set the toggle of first sheet
firstSheet = True

'Loop through each worksheet in the active workbook
For Each ws In ActiveWorkbook.Sheets

    'Check if the worksheet was not previously selected
    If InStr(selectedList, ws.Name & "[|]") = 0 Then
            
        'Check the worksheet is visible
        If ws.Visible = xlSheetVisible Then
            
            'Select the sheet
            ws.Select firstSheet
            
            'First worksheet has been found, toggle to false
            firstSheet = False
        
        End If
    
    End If
    
Next ws

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'029 - Assign a macro to a shortcut key'''''''''''''''''''''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub AssignMacroToShortcut()

'+ = Ctrl
'^ = Shift
'{T} = the shortcut letter

Application.OnKey "+^{T}", "nameOfMacro"

'Reset shortcut to default - repeat without the name of the macro
'Application.OnKey "+%{T}"

End Sub


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'030 - Apply single accounting underline to selected cells''''''''''
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Sub SingleAccountingUnderline()

'Apply single accounting underline to selected cells
Selection.Font.Underline = xlUnderlineStyleSingleAccounting

End Sub

