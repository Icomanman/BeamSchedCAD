

Sub extractData(ByRef data As Range) 'ByRef range As range,
    On Error Resume Next
    Set data = Application.InputBox("Please select the data range.", "Data Range", "B1", Type:=8)
    If data Is Nothing Then
        MsgBox "No input was selected!"
        Exit Sub
    End If
End Sub

Sub run_kr()

    Dim kr() As Double
    Dim Pn() As Double
    Dim Mn() As Double
    Dim i As Integer
    Dim pairCount As Integer
    Dim dataCol As Integer
    Dim dataRange As Range 'Ranges always start at 1
    Dim data() As Variant
    
    Call extractData(dataRange)
    
    pairCount = dataRange.Rows.count
    dataCol = dataRange.Columns.count
    ReDim data(pairCount, dataCol)
    data = dataRange.Value
    
    ReDim Pn(pairCount)
    ReDim Mn(pairCount, 1) '0: Mn2, 1: Mn3
    ReDim kr(pairCount, 2) '0: Kn, 1: Rn2, 2: Rn3
    
    For i = 0 To pairCount - 1
        
        Pn(i) = data(i, 1)
        If dataCol < 4 Then
            Mn(i, 0) = 0
        Else
            Mn(i, 0) = data(i, 4)
        End If
        
        If dataCol < 5 Then
            Mn(i, 1) = 0
        Else
            Mn(i, 1) = data(i, 5)
        End If
        
    Next i
    

End Sub



