# ContinuousPrinting

VBA script to print out Excel worksheet continuously

## Install
Follow these steps: 
1. Save 'ContinuousPrinting.bas' to your computer.
1. Import 'ContinuousPrinting.bas' to any Excel workbook.  
Make it sure that the workbook has extension '**.xlsm**', '**.xlam**' (or '.xls').
1. Press Alt + F8 and type '**cp**' (for **C**ontinuous**P**rinting).

You can add this macro to the context menu (right-click menu):
1. Open VBE (press Alt + F11).
1. Open 'VBAProject (WORKBOOK_NAME)' -> 'Misrosoft Excel Object' -> 'ThisWorkbook'.
1. Add the code below.  
If `Private Sub Workbook_Open()` already exists, just add `Call AddToContextMenu_ContinuousPrinting` before `End Sub`.
```VB
Private Sub Workbook_Open()
    Call AddToContextMenu_ContinuousPrinting
End Sub
```
4. Save workbook and reopen.
4. Now, 'ContinuousPrinting' should appear in the menu. This menu remains while the workbook is open.

## How to use
1. Input numbers to use for printing.
1. Click a cell to input the numbers. **Only one cell** can be selected.
1. Click 'Yes' on confirm window.