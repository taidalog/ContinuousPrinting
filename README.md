# ContinuousPrinting

[Japanese README](README.ja.md)

1. [About](#About)
1. [Install](#Install)
1. [How to use](#How-to-use)
1. [Future features](#Future-features)

## About
VBA script to print out Excel worksheet continuously.  
One of the most typical situations for continuous printing may be that you type 1 in A1 cell, VLOOKUP() function works, then you print it out, type 2 in A1 cell, VLOOKUP() function works, ptint it out, 3 in A1, VLOOKUP(), print it out. This script helps repeat this task. You type numbers like '1-3' for 1 to 3, click a cell in which the numbers will be input, then printing will be done for all the numbers.

## Install
Follow these steps: 
1. Save 'ContinuousPrinting.bas' to your computer.
1. Import 'ContinuousPrinting.bas' to any Excel workbook.  
Make it sure that the workbook has extension '**.xlsm**', '**.xlam**' (or '.xls').
1. Press 'Alt + F8' and type '**cp**' (for **C**ontinuous**P**rinting).

You can add this macro to the context menu (right-click menu):
1. Open VBE (press 'Alt + F11').
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

## Future features
- [ ] Support for inputting text, instead of numbers
- [ ] Support for inputting from cell range
