Attribute VB_Name = "MDL_CatmullRom"
Option Explicit

'----------------------------------------------------------------------
'       HK (a.k.a. eXeption)
'----------------------------------------------------------------------

Public Type POINTAPI
    x As Long
    y As Long
End Type


Public Sub ForwardDiffCalc(Ap As POINTAPI, Bp As POINTAPI, Cp As POINTAPI, dp As POINTAPI, T#, d#, x#, y#)
    Dim T2#, T3#
    ' Compute the spline by forward differencing (non-Adaptive or fixed)
    T2 = T * T                                              '{ Square of t }
    T3 = T2 * T                                             '{ Cube of t }
    x = ((Ap.x * T3) + (Bp.x * T2) + (Cp.x * T) + dp.x) / d '{ Calc x value }
    y = ((Ap.y * T3) + (Bp.y * T2) + (Cp.y * T) + dp.y) / d '{ Calc y value }
End Sub

Public Sub ComputeCoeffs_CatmullRom(pt() As POINTAPI, N As Long, Ap As POINTAPI, Bp As POINTAPI, Cp As POINTAPI, dp As POINTAPI)
    Ap.x = -pt(N - 1).x + 3 * pt(N).x - 3 * pt(N + 1).x + pt(N + 2).x
    Bp.x = 2 * pt(N - 1).x - 5 * pt(N).x + 4 * pt(N + 1).x - pt(N + 2).x
    Cp.x = -pt(N - 1).x + pt(N + 1).x
    dp.x = 2 * pt(N).x
    Ap.y = -pt(N - 1).y + 3 * pt(N).y - 3 * pt(N + 1).y + pt(N + 2).y
    Bp.y = 2 * pt(N - 1).y - 5 * pt(N).y + 4 * pt(N + 1).y - pt(N + 2).y
    Cp.y = -pt(N - 1).y + pt(N + 1).y
    dp.y = 2 * pt(N).y
End Sub

' This function merely draws a catmull rom spline to the spcified HDC with anti aliasing enabled.
Public Function Draw_CatmullRom(pict As Object, pt() As POINTAPI, N, nSegmnts, mColor As Long, Optional ByVal Closed As Boolean = False) As Long
    Dim i&, j&, x#, y#, Lx#, Ly#, Ap As POINTAPI, Bp As POINTAPI, Cp As POINTAPI, dp As POINTAPI
    If Closed = False Then
        pt(0) = pt(1)
        pt(N + 1) = pt(N)
    Else
        pt(0) = pt(N - 1)
        pt(N + 1) = pt(2)
    End If
    ReDim pt2(N * nSegmnts)
    For i = 1 To N - 1
        ComputeCoeffs_CatmullRom pt, i, Ap, Bp, Cp, dp
        ForwardDiffCalc Ap, Bp, Cp, dp, 0, 2, Lx, Ly
        For j = 1 To nSegmnts
            ForwardDiffCalc Ap, Bp, Cp, dp, j / nSegmnts, 2, x, y
            pict.Line (Round(Lx), Round(Ly))-(Round(x), Round(y)), mColor
            Lx = x
            Ly = y
        Next
    Next
    pict.Refresh
End Function


' This function is used for the moving part.
' It returns a
Public Function GetCatmullRom_Points(pt() As POINTAPI, N, nSegmnts, pt2() As POINTAPI, Optional ByVal Closed As Boolean = False) As Long
    Dim i&, j&, x#, y#, Lx#, Ly#, Ap As POINTAPI, Bp As POINTAPI, Cp As POINTAPI, dp As POINTAPI
    Dim a&
    If Closed = False Then
        pt(0) = pt(1)
        pt(N + 1) = pt(N)
    Else
        pt(0) = pt(N - 1)
        pt(N + 1) = pt(2)
    End If
    ReDim pt2(N * (nSegmnts + 1))
    a = 0
    For i = 1 To N - 1
        ComputeCoeffs_CatmullRom pt, i, Ap, Bp, Cp, dp
        ForwardDiffCalc Ap, Bp, Cp, dp, 0, 2, Lx, Ly
        pt2(a).x = Round(Lx)
        pt2(a).y = Round(Ly)
        a = a + 1
        For j = 1 To nSegmnts
            ForwardDiffCalc Ap, Bp, Cp, dp, j / nSegmnts, 2, x, y
            pt2(a).x = Round(x)
            pt2(a).y = Round(y)
            a = a + 1
            Lx = x
            Ly = y
        Next
    Next
    GetCatmullRom_Points = IIf(Closed = True, a - 1, a)
End Function
