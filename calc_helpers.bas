Option Explicit

Public PI As Double

Private db As Worksheet
Private dbRange As Range
Private dbOrigin As Range

Public Type Beam

    dWidth As Double
    dDepth As Double
    dBarDia As Double
    Cc As Double
    dLinks As Integer
    dMinSpace As Double
    dPhi As Double
    
    dFyMain As Double
    dFySec As Double
    dFc As Double
    
    iBarNos As Integer
    a As Double
    dAst As Double
    dMn As Double
    
End Type

Public Sub Mn(ByRef bm As Beam)
    Dim a As Double
        
    a = bm.dAst * bm.dFyMain / (0.85 * bm.dFc * bm.dWidth)
            
    bm.dMn = (bm.dAst * bm.dFyMain * (bm.dDepth - (a / 2))) / 1000000
    
End Sub

Public Function myBeam(ByVal iRow As Integer) As Beam

    Dim prompt As String
     
    myBeam.dWidth = dbOrigin.Offset(iRow)
    myBeam.dDepth = dbOrigin.Offset(iRow, 1)
    myBeam.dBarDia = dbOrigin.Offset(iRow, 2)
    myBeam.Cc = db.Range("F1")
    myBeam.dLinks = dbOrigin.Offset(iRow, 5)
    myBeam.dMinSpace = db.Range("F2")
    myBeam.dPhi = db.Range("F3")
    myBeam.dFyMain = db.Range("B1")
    myBeam.dFySec = db.Range("B2")
    myBeam.dFc = db.Range("B3")
    
    myBeam.iBarNos = db.Cells(6 + iRow, 4)
    myBeam.dAst = myBeam.iBarNos * 0.25 * (myBeam.dBarDia ^ 2) * PI
        
    myBeam.a = myBeam.dAst * myBeam.dFyMain / (0.85 * myBeam.dFc * myBeam.dWidth)
    myBeam.dMn = (myBeam.dAst * myBeam.dFyMain * (myBeam.dDepth - (myBeam.a / 2))) / 1000000
    
    dbOrigin.Offset(iRow, 4).Value = Round(myBeam.dPhi * myBeam.dMn, 2)

End Function

Public Function max_bar(ByRef bm As Beam) As Integer
    'For single layer
    Const iMinPcs = 2
    Dim dBarMaxCount As Double
    
    dBarMaxCount = (bm.dWidth - 2 * (bm.Cc + bm.dLinks) + bm.dMinSpace) / (bm.dBarDia + bm.dMinSpace)
    max_bar = WorksheetFunction.RoundDown(dBarMaxCount, 0)
    
End Function

Public Function max_double_layer(ByVal max_bar As Integer)
    max_double_layer = max_bar * 2
End Function


Function bar_nos(ByVal maxBar As Integer) As Integer()

    Dim iPcs() As Integer
    Dim i As Integer
    Dim txt As String
    
    ReDim iPcs(maxBar)
    
    For i = 0 To maxBar - 2
        iPcs(i) = i + 2
        
        txt = iPcs(i)
        'MsgBox txt
    Next i
    
    bar_nos = iPcs
        
End Function

Sub fill_DB()
    Dim ws As Object
    Dim rowCount As Integer
    Dim colCount As Integer
    Dim txt As String
    Dim dbData() As Variant
    Dim iMaxBar As Integer
    Dim iPcs() As Integer
    Dim pcsRange As Range
    Dim lastRow As Range
    
    Set ws = WorksheetFunction
    Set db = Worksheets("DB")
    Set dbOrigin = db.Range("A6")
    
    rowCount = Range(dbOrigin, dbOrigin.Offset(-1).End(xlDown)).count
    colCount = Range(dbOrigin.Offset(-1), dbOrigin.Offset(-1).End(xlToRight)).count
    Set dbRange = Range(dbOrigin, db.Cells(rowCount, colCount))

    ReDim dbData(rowCount, colCount)
    PI = ws.PI()
    
    'pcsRange and lastRow are dynamic Ranges
    iMaxBar = max_bar(myBeam(0)) ' dynamic; as function of B
    
    Set lastRow = Cells(dbOrigin.Row + iMaxBar - 1, 1)
    Set pcsRange = Range(Cells(dbOrigin.Row, 4), Cells(dbOrigin.Row + iMaxBar - 2, 4))
    
    iPcs = bar_nos(iMaxBar)
    pcsRange.Value = Application.Transpose(iPcs)
    
    txt = Range("A7").Value
    If IsEmpty(lastRow) = False Then
        MsgBox txt
    End If
    
    lastRow.Select
    
    Set pcsRange = Nothing
    Set lastRow = Nothing
    Set ws = Nothing
    Set db = Nothing
    Set dbOrigin = Nothing
    Set dbRange = Nothing
    
End Sub


