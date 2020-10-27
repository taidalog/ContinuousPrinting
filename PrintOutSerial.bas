Attribute VB_Name = "PrintOutSerial"
Option Explicit

Sub PrintSerial()
    
    Dim msg As String
    msg = _
        "印刷範囲を指定してください。" & vbCrLf & _
        "'1 から 3' -> '1-3'" & vbCrLf & _
        "'1 と 3 と 5' -> '1,3,5'" & vbCrLf & _
        "'1 と 3 から 5' -> '1,3-5'"
    
    Dim str As String
    str = Application.InputBox(msg, Type:=2)
    
    If str = "False" Then Exit Sub
    
    Debug.Print str
    
    Dim tmp As String
    tmp = str
    tmp = Replace(tmp, " ", "")
    tmp = Replace(tmp, ",", "")
    tmp = Replace(tmp, "-", "")
    
    Debug.Print tmp
    
    If Not IsNumeric(tmp) Then Exit Sub
    
    
    Dim v As Variant
    v = Split(str, ",")
    
    Dim printCount As Long
    printCount = 0
    
    Dim i As Long
    For i = LBound(v) To UBound(v)
        
        Dim tmp2 As String
        tmp2 = Replace(v(i), "-", "")
        
        If v(i) = tmp2 Then
            If Not IsNumeric(v(i)) Then GoTo Continue
            
            Debug.Print CLng(v(i))
            printCount = printCount + 1
        Else
            Dim v2 As Variant
            v2 = Split(v(i), "-")
            
            If UBound(v2) > 1 Then GoTo Continue
            If Not IsNumeric(v2(0)) Then GoTo Continue
            If Not IsNumeric(v2(1)) Then GoTo Continue
            
            Debug.Print CLng(v2(0))
            Debug.Print CLng(v2(1))
            
            Dim j As Long
            For j = CLng(v2(0)) To CLng(v2(1))
                printCount = printCount + 1
            Next j
            
        End If
        
Continue:
    Next i
    
    MsgBox str & vbCrLf & printCount & "枚印刷します。" & vbCrLf & "実行しますか？", vbYesNo
    
End Sub
