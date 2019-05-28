Option Explicit

Private src As Object

Sub verify_count(ByVal count As Integer)
    
    Dim msg As String
    
    If count = Range("L1").Value Then
        msg = "Data range is verified. Ok."
    Else
        msg = "Error. Please Check Data Range. IGOT!!!"
    End If
    MsgBox msg
    
End Sub


Sub clear_output()

    Range("N3: N5").Value = ""
    Range("P3: P5").Value = ""

End Sub


Sub summary()

    Set src = Worksheets("Source")
    
    Dim dForcesArr(2, 1) As Variant
    Dim v2Min, m3Min, v2Max, m3Max As Double
    Dim rawData() As Variant

    Dim WF As Object
    Set WF = WorksheetFunction
    
    Dim dataOrigin, dataRange, v2Range, m3Range As Range
    Dim v2Col, m3Col As Range
    Dim i, j, k, rowCount As Integer
    Dim txt As String
    
    Set dataOrigin = src.Range("A2")
    
    rowCount = src.Range(dataOrigin, dataOrigin.End(xlDown)).count
    verify_count (rowCount)
    
    Set dataRange = src.Range(dataOrigin, Cells(2 + rowCount - 1, 10))
    Set v2Range = src.Range("P3: P5")
    Set m3Range = src.Range("N3: N5")
    
    ReDim rawData(rowCount, dataRange.Columns.count)
    rawData = dataRange.Value
    
    v2Min = WF.Min(dataRange.Columns(6))
    m3Min = WF.Min(dataRange.Columns(10))
    v2Max = WF.Min(dataRange.Columns(6))
    m3Max = WF.Max(dataRange.Columns(10))
    
    dForcesArr(0, 0) = m3Min
    If m3Max = 0 Then 'Cantilever
        dForcesArr(1, 0) = m3Max
        dForcesArr(2, 0) = ""
        
        dForcesArr(0, 1) = ""
        dForcesArr(1, 1) = WF.Max(Abs(v2Min), Abs(v2Max))
        dForcesArr(2, 1) = ""
    Else 'Continuous
        dForcesArr(1, 0) = 0
        dForcesArr(2, 0) = m3Max
        
        dForcesArr(0, 1) = ""
        dForcesArr(1, 1) = WF.Max(Abs(v2Min), Abs(v2Max))
        dForcesArr(2, 1) = ""
    End If
    
    'txt = CStr(m3Min)
    'MsgBox txt
    
    v2Range.Value = WF.index(dForcesArr, 0, 2)
    m3Range.Value = WF.index(dForcesArr, 0, 1)

    Set src = Nothing
    Set WF = Nothing
    Set dataOrigin = Nothing
    Set dataRange = Nothing
    Set v2Range = Nothing
    Set m3Range = Nothing
    
End Sub






