Attribute VB_Name = "PrintOutSerial"
Option Explicit

Sub PrintOutSerial()
    
    '
    Dim msg As String
    msg = _
        "印刷範囲を入力してください。" & vbCrLf & _
        "1 から 3 -> '1-3'" & vbCrLf & _
        "1 と 3 と 5 -> '1,3,5'" & vbCrLf & _
        "1 から 3 と 5 -> '1-3,5'"
    
    '
    Do
        Dim str As String
        str = Application.InputBox(msg, Type:=2) '-18--16,-10--12,-6,-4-2,4-1,6-8
        
        If str = "False" Then Exit Sub
    Loop While str = ""
    
    Dim res As Variant
    res = ConvertToNumArray(str)
    
    Debug.Print ""
    
    '
    Dim CL As Range
    Set CL = SetRangeWithInputBox("印刷範囲を入力するセルをクリックしてください。")
    
    '
    Dim printCount As Long
    printCount = UBound(res) - LBound(res) + 1
    
    '
    Dim confirmationAnswer As Long
    confirmationAnswer = MsgBox(str & vbCrLf & printCount & "枚印刷します。" & vbCrLf & "実行しますか？", vbYesNo)
    If confirmationAnswer = vbNo Then Exit Sub
    
    '
    Dim i As Long
    For i = LBound(res) To UBound(res)
        CL.Value = res(i)
'        ActiveSheet.PrintOut
        Debug.Print res(i)
    Next i
    
End Sub

Function SetRangeWithInputBox(prompt_message As String) As Range

    On Error GoTo ErrorHandle
    
    Set SetRangeWithInputBox = Application.InputBox(prompt_message, Type:=8)
    Exit Function
    
ErrorHandle:
    Set SetRangeWithInputBox = Nothing
    
End Function

Function ConvertToNumArray(num_string As String) As Variant
    
    Dim strArray As Variant
    strArray = Split(num_string, ",")
        
    Dim result() As Variant
    Dim n As Long
    n = 0
    
    Dim i As Long
    For i = LBound(strArray) To UBound(strArray)
        
        ' split with the first hyphen after the first letter
        Dim splitPosition As Long
        splitPosition = InStr(2, strArray(i), "-")
        
        If splitPosition = 0 Then
            ReDim Preserve result(n)
            result(n) = CLng(strArray(i))
            n = n + 1
            GoTo Continue
        End If
        
        Dim strFrom As String, strTo As String
        strFrom = Mid(strArray(i), 1, splitPosition - 1)
        strTo = Mid(strArray(i), splitPosition + 1)
        
        If Not IsNumeric(strFrom) Then GoTo Continue
        If Not IsNumeric(strTo) Then GoTo Continue
        
        Dim forStep As Long
        If CLng(strFrom) <= CLng(strTo) Then
            forStep = 1
        Else
            forStep = -1
        End If
        
        Dim j As Long
        For j = CLng(strFrom) To CLng(strTo) Step forStep
            ReDim Preserve result(n)
            result(n) = j
            n = n + 1
        Next j
        
Continue:
    Next i
    
    ConvertToNumArray = result
    
End Function
