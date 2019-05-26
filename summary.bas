
Dim dForcesArr(9, 4) As Double
Dim pVal As Object
Dim rawData() As Variant

Dim WF As Object

Sub Summary()
    Set WF = WorksheetFunction
    Dim dataOrigin, dataRange, outputRange As Range
    Dim pCol, v2Col, v3Col, m2Col, m3Col As Range
    Dim i, j, k, rowCount As Integer
    Dim txt As String
    
    Set dataOrigin = Cells(3, 1)
    
    rowCount = Range(dataOrigin, dataOrigin.End(xlDown)).Count
    Set dataRange = Range(dataOrigin, Cells(3 + rowCount - 1, 10))
    Set outputRange = Range("M6: Q15")
    
    ReDim rawData(rowCount, dataRange.Columns.Count)
    rawData = dataRange.Value
    
    dForcesArr(0, 0) = WF.Min(dataRange.Columns(5))
    dForcesArr(1, 0) = WF.Max(dataRange.Columns(5))
    dForcesArr(2, 1) = WF.Min(dataRange.Columns(6))
    dForcesArr(3, 1) = WF.Max(dataRange.Columns(6))
    dForcesArr(4, 2) = WF.Min(dataRange.Columns(7))
    dForcesArr(5, 2) = WF.Max(dataRange.Columns(7))
    dForcesArr(6, 3) = WF.Min(dataRange.Columns(9))
    dForcesArr(7, 3) = WF.Max(dataRange.Columns(9))
    dForcesArr(8, 4) = WF.Min(dataRange.Columns(10))
    dForcesArr(9, 4) = WF.Max(dataRange.Columns(10))
    
    For i = 0 To rowCount - 1
        For j = 0 To 1
            'pMin and pMax
            If dForcesArr(j, 0) = dataRange(i, 5) Then
                dForcesArr(j, j + 1) = dataRange(i, 6 + j)
                dForcesArr(j, j + 3) = dataRange(i, 9 + j)
            End If
            'v2Min and v2Max
            If dForcesArr(j + 2, 1) = dataRange(i, 6) Then
                dForcesArr(j + 2, 0) = dataRange(i, 5)
                dForcesArr(j + 2, 2) = dataRange(i, 7)
                dForcesArr(j + 2, 3) = dataRange(i, 9)
                dForcesArr(j + 2, 4) = dataRange(i, 10)
            End If
            'v3Min and v3Max
            If dForcesArr(j + 4, 2) = dataRange(i, 7) Then
                dForcesArr(j + 4, 0) = dataRange(i, 5)
                dForcesArr(j + 4, 1) = dataRange(i, 6)
                dForcesArr(j + 4, 3) = dataRange(i, 9)
                dForcesArr(j + 4, 4) = dataRange(i, 10)
            End If
            'm2Min and m2Max
            If dForcesArr(j + 6, 3) = dataRange(i, 9) Then
                dForcesArr(j + 6, 0) = dataRange(i, 5)
                dForcesArr(j + 6, 1) = dataRange(i, 6)
                dForcesArr(j + 6, 2) = dataRange(i, 7)
                dForcesArr(j + 6, 4) = dataRange(i, 10)
            End If
            'm3Min and m3Max
            If dForcesArr(j + 8, 4) = dataRange(i, 10) Then
                dForcesArr(j + 8, 0) = dataRange(i, 5)
                dForcesArr(j + 8, 1) = dataRange(i, 6)
                dForcesArr(j + 8, 2) = dataRange(i, 7)
                dForcesArr(j + 8, 3) = dataRange(i, 9)
            End If
        Next j
    Next i
    
    txt = WF.Max(dataRange.Columns(5))
    
    'MsgBox txt
    outputRange = dForcesArr

    Set WF = Nothing
End Sub




