Option Explicit

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
    
    dFyMain As Double
    dFySec As Double
    dFc As Double
    
    iBarNos As Integer
    
End Type

Public Function myBeam(ByVal iRow As Integer) As Beam

    Dim prompt As String
     
    myBeam.dWidth = dbOrigin.Offset(iRow)
    myBeam.dDepth = dbOrigin.Offset(iRow, 1)
    myBeam.dBarDia = dbOrigin.Offset(iRow, 2)
    myBeam.Cc = db.Range("F1")
    myBeam.dLinks = dbOrigin.Offset(iRow, 5)
    myBeam.dMinSpace = db.Range("F2")
    myBeam.dFyMain = db.Range("B1")
    myBeam.dFySec = db.Range("B2")
    myBeam.dFc = db.Range("B3")
    
    myBeam.iBarNos = db.Cells(6 + iRow, 4)

End Function

Public Function max_bar(ByRef bm As Beam) As Integer

    Const iMinPcs = 2
    Dim dBarMaxCount As Double
    
    dBarMaxCount = (bm.dWidth - 2 * (bm.Cc + bm.dLinks) + bm.dMinSpace) / (bm.dBarDia + bm.dMinSpace)
    max_bar = WorksheetFunction.RoundDown(dBarMaxCount, 0)
    
End Function

Public Function Mn(ByRef bm As Beam) As Variant
    Dim a As Double
    Dim Ast As Double
    
    Ast = bm.iBarNo * 16
    
    a = Ast * bm.dFyMain / (0.85 * bm.dFc * bm.dWidth)
            If spacing(k + (4 * j)) < dBarSpac Then
                Mn = "-"
            Else
               Mn = 0.9 * (Ast(k + (4 * j)) * fy * (bm.dDepth(i) - (a(i) / 2))) / 1000000
            End If
End Function

Function bar_nos(ByVal maxBar As Integer) As Integer()

    Dim iPcs() As Integer
    Dim i As Integer
    Dim txt As String
    
    ReDim iPcs(maxBar)
    
    For i = 0 To maxBar - 2
        iPcs(i) = i + 2
        
        txt = iPcs(i)
        MsgBox txt
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


