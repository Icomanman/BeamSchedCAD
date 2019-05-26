Option Explicit

Private Ast(11) As Double
Private spacing(11) As Double

Private BarDiam(3) As Integer, pcs(3) As Integer
Private BarArea(3) As Double

Private Links As Integer

Private dWidth As Double

Private a(11) As Double
Private MomentCap(4, 11) As Variant '(height x rebar configuration)
Private ShearCap(4, 3) As Variant

Private i As Integer, j As Integer, k As Integer
Private absRow As Integer, iSet As Integer

Private txt As String

Private fc As Double, fy As Double
Private Cc As Double

Const fysec = 275

'v004 Bar spacing of 40mm converted to dBarSpec with assigned value of 30mm
Const dBarSpac = 30

Sub RebarData(ByVal dWidth As Double, ByVal Cc As Double, ByVal Links As Integer)

    BarDiam(0) = 16
    BarDiam(1) = 20
    BarDiam(2) = 25
    BarDiam(3) = 10
    pcs(0) = 2
    pcs(1) = 3
    pcs(2) = 4
    pcs(3) = 5
    BarArea(0) = 201
    BarArea(1) = 314
    BarArea(2) = 490
    BarArea(3) = 78
    
    For i = 0 To 2
        For j = 0 To 3
            Ast(j + (4 * i)) = BarArea(i) * pcs(j)
            
            spacing(j + (4 * i)) = (dWidth - (2 * (Cc + Links) _
            + (BarDiam(i) * pcs(j)))) / (pcs(j) - 1)
        Next j
    Next i
    
End Sub

Sub BeamCalc()

fy = Range("Q2")
fc = Range("Q3")
Links = Range("Q4")
Cc = Range("Q5")

Dim dHeight(4) As Double

For iSet = 0 To 6

    absRow = 15 + (6 * iSet)
    dWidth = Cells(absRow, 1)
    txt = dWidth & " mm-wide beam with heights of:" & vbCrLf
    
    For i = 0 To 4
        dHeight(i) = Cells(absRow + i, 2)
        txt = txt & dHeight(i) & " mm" & vbCrLf
    Next i
    MsgBox txt
    
    RebarData dWidth, Cc, Links

    CalculateMoments dWidth, dHeight, Ast, absRow
    For i = 0 To 11
        If spacing(i) < dBarSpac Then
            Cells(absRow - 1, 3 + i) = "< " & dBarSpac
        Else
            Cells(absRow - 1, 3 + i) = spacing(i)
        End If
    Next i
    'Range(Cells(absRow - 1, 3), Cells(absRow + 4, 14)) = spacing
    Range(Cells(absRow, 3), Cells(absRow + 4, 14)) = MomentCap
    
    CalculateShears dWidth, dHeight, absRow
    Range(Cells(absRow, 15), Cells(absRow + 4, 18)) = ShearCap

Next iSet

End Sub

Sub CalculateMoments(ByVal dWidth As Double, dHeight() As Double, Ast() As Double, absRow As Integer)

Dim dDepth(4) As Double

For i = 0 To 4
    For j = 0 To 2
        For k = 0 To 3
            dDepth(i) = dHeight(i) - (Cc + Links + (BarDiam(j) / 2))
            a(i) = Ast(k + (4 * j)) * fy / (0.85 * fc * dWidth)
            If spacing(k + (4 * j)) < dBarSpac Then
                MomentCap(i, k + (4 * j)) = "-"
            Else
                MomentCap(i, k + (4 * j)) = 0.9 * (Ast(k + (4 * j)) * fy * (dDepth(i) - (a(i) / 2))) / 1000000
            End If
        Next k
    Next j
Next i

End Sub


Sub CalculateShears(ByVal dWidth As Double, dHeight() As Double, absRow As Integer)

    Dim dDepth(4) As Double 'reckoned from smallest possible value of d (from 25mm Dia)
    Dim LinkSpa(3) As Double
    Dim vc(4) As Double
    Dim vs(4, 3) As Double
    
    LinkSpa(0) = 100
    LinkSpa(1) = 150
    LinkSpa(2) = 200
    'LinkSpa(3) = 300, 'max spacing: dDepth / 2

    For i = 0 To 4

        dDepth(i) = dHeight(i) - (Cc + Links + (BarDiam(2) / 2))
        vc(i) = 0.17 * (fc ^ 0.5) * dWidth * dDepth(i)
    
        LinkSpa(3) = dDepth(i) / 2
    
        For j = 0 To 3
            vs(i, j) = 2 * BarArea(3) * fysec * dDepth(i) / LinkSpa(j)
            ShearCap(i, j) = 0.75 * (vs(i, j) + vc(i)) / 1000
        Next j
    
    Next i

End Sub




