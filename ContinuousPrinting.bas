Attribute VB_Name = "ContinuousPrinting"
Option Explicit
Option Private Module

Public Sub PrintOutContinuously()
    
    Dim promptMessage As String
    promptMessage = _
        "Input numbers to use for printing." & vbCrLf & _
        "    1 to 3 -> '1-3'" & vbCrLf & _
        "    1 and 3 and 5 -> '1,3,5'" & vbCrLf & _
        "    1 and 3, and 5 -> '1-3,5'"
    
    Dim defaultValue As String
    defaultValue = ""
    
    Do
        Do
            ' uses InputBox to ask numbers to use for printing.
            ' repeats asking while input is empty.
            ' exits if 'Cancel' is selected or Esc is typed.
            Dim numberString As String
            numberString = Application.InputBox(promptMessage, Type:=2, Default:=defaultValue)
            
            If numberString = "False" Then Exit Sub
        Loop While numberString = ""
        
        ' converts the number string to number array.
        ' repeats asking while the number string couldn't converted successfully.
        Dim numberArray As Variant
        numberArray = ConvertToNumArray(numberString)
        
        If IsNull(numberArray) Then defaultValue = numberString
        
    Loop While IsNull(numberArray)
    
    ' asks a cell to input the numbers in.
    ' only ONE cell can be selected.
    ' exits if 'Cancel' is selected or Esc is typed.
    Do
        Dim CL As Range
        Set CL = SetRangeWithInputBox("Click a cell to input the numbers." & vbCrLf & "Only ONE cell can be selected.")
        
        If CL Is Nothing Then Exit Sub
        
        If CL.Count > 1 Then
            MsgBox "Click only ONE cell."
        End If
        
    Loop While CL.Count > 1
    
    Dim WS As Worksheet
    Set WS = ActiveSheet
    
    Dim printCount As Long
    printCount = UBound(numberArray) - LBound(numberArray) + 1
    
    Dim confirmationAnswer As Long
    confirmationAnswer = MsgBox("'" & numberString & "'" & vbCrLf & printCount & " sheet(s) will be printed." & vbCrLf & "Start Printing?", vbYesNo)
    If confirmationAnswer = vbNo Then Exit Sub
    
    ' prints out.
    Dim i As Long
    For i = LBound(numberArray) To UBound(numberArray)
        CL.Value = numberArray(i)
        WS.PrintOut
        Debug.Print Now, numberArray(i)
    Next i
    
End Sub

Private Function SetRangeWithInputBox(prompt_message As String) As Range

    On Error GoTo ErrorHandle
    
    Set SetRangeWithInputBox = Application.InputBox(prompt_message, Type:=8)
    Exit Function
    
ErrorHandle:
    Set SetRangeWithInputBox = Nothing
    
End Function

Private Function ConvertToNumArray(num_string As String) As Variant
    
    Dim strArray As Variant
    strArray = Split(num_string, ",")
        
    Dim result() As Variant
    Dim n As Long
    n = 0
    
    Dim i As Long
    For i = LBound(strArray) To UBound(strArray)
        
        ' gets the position of the first hyphen after the first letter.
        Dim firstHyphenPosition As Long
        firstHyphenPosition = InStr(2, strArray(i), "-")
        
        If firstHyphenPosition = 0 Then
            
            If IsNumeric(strArray(i)) Then
                ReDim Preserve result(n)
                result(n) = CLng(strArray(i))
                n = n + 1
            End If
            
            GoTo Continue
        End If
        
        ' splits with the first hyphen after the first letter.
        Dim strFrom As String, strTo As String
        strFrom = Mid(strArray(i), 1, firstHyphenPosition - 1)
        strTo = Mid(strArray(i), firstHyphenPosition + 1)
        
        If Not IsNumeric(strFrom) Or Not IsNumeric(strTo) Then
            Dim errorMessage As String
            errorMessage = strArray(i)
            
            MsgBox "'" & errorMessage & "' can't be used as numbers."
            ConvertToNumArray = Null
            Exit Function
        End If
        
        If Not IsNumeric(strTo) Then GoTo Continue
        
        Dim forFrom As Long, forTo As Long
        forFrom = CLng(strFrom)
        forTo = CLng(strTo)
        
        Dim forStep As Long
        If forFrom <= forTo Then
            forStep = 1
        Else
            forStep = -1
        End If
        
        Dim j As Long
        For j = forFrom To forTo Step forStep
            ReDim Preserve result(n)
            result(n) = j
            n = n + 1
        Next j
        
Continue:
    Next i
    
    ConvertToNumArray = result
    
End Function

Public Sub cp()
    ' alias
    Call PrintOutContinuously
End Sub

Public Sub AddToContextMenu_ContinuousPrinting()
    
    With Application.CommandBars
        Dim i As Long
        For i = 1 To .Count
            
            With .Item(i)
                If .Name = "Cell" Then
                    
                    With .Controls.Add(Type:=msoControlButton, Temporary:=True)
                        .BeginGroup = True
                        .Caption = "&ContinuousPrinting"
                        .OnAction = ThisWorkbook.Name & "!PrintOutContinuously"
                        
                    End With
                    
                End If
            End With
            
        Next i
        
    End With
    
End Sub
