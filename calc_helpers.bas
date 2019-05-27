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
    minSpace As Double
    
    dFyMain As Double
    dFySec As Double
    dFc As Double
    
End Type

Public Function myBeam(ByVal iRow As Integer) As Beam

    Dim prompt As String
     
    myBeam.dWidth = dbOrigin.Offset(iRow)
    myBeam.dDepth = dbOrigin.Offset(iRow, 1)
    myBeam.dBarDia = dbOrigin.Offset(iRow, 2)
    myBeam.Cc = db.Range("F1")
    myBeam.dLinks = dbOrigin.Offset(iRow, 5)
    myBeam.minSpace = db.Range("F2")
    myBeam.dFyMain = db.Range("B1")
    myBeam.dFySec = db.Range("B2")
    myBeam.dFc = db.Range("B3")

End Function

Public Function maxBar(ByRef myBeam As Beam) As Integer
    Const iMinPcs = 2
    Dim dBarSpac
    maxBar = myBeam.dWidth - 2 * (Beam.Cc)
    
End Function

Public Function Mn(ByRef bm As Beam) As Variant
    Dim a As Double
    a = Ast * bm.dFyMain / (0.85 * bm.dFc * bm.dWidth)
            If spacing(k + (4 * j)) < dBarSpac Then
                Mn = "-"
            Else
               Mn = 0.9 * (Ast(k + (4 * j)) * fy * (bm.dDepth(i) - (a(i) / 2))) / 1000000
            End If
End Function

Sub main()
    Dim rowCount As Integer
    Dim colCount As Integer
    Dim txt As String
    Dim dbData() As Variant
    
    Set db = Worksheets("DB")
    Set dbOrigin = db.Range("A6")
    
    rowCount = Range(dbOrigin, dbOrigin.Offset(-1).End(xlDown)).count
    colCount = Range(dbOrigin.Offset(-1), dbOrigin.Offset(-1).End(xlToRight)).count
    Set dbRange = Range(dbOrigin, db.Cells(rowCount, colCount))

    ReDim dbData(rowCount, colCount)
    
    MsgBox txt
    
    Set db = Nothing
    Set dbOrigin = Nothing
    Set dbRange = Nothing
    
End Sub
