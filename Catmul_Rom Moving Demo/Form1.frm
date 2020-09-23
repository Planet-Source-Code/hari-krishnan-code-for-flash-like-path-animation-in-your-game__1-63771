VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Catmull-Rom Spline Object moving example.  by eXeption!"
   ClientHeight    =   7785
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9780
   LinkTopic       =   "Form1"
   ScaleHeight     =   7785
   ScaleWidth      =   9780
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   4590
      Top             =   1440
   End
   Begin VB.PictureBox pcvs 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   3120
      Left            =   1125
      ScaleHeight     =   153
      ScaleMode       =   2  'Point
      ScaleWidth      =   288
      TabIndex        =   7
      Top             =   2250
      Width           =   5820
      Begin VB.Shape shp 
         BorderColor     =   &H00000000&
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   675
         Shape           =   3  'Circle
         Top             =   450
         Width           =   195
      End
   End
   Begin VB.PictureBox picToolBar 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      ScaleHeight     =   810
      ScaleWidth      =   9720
      TabIndex        =   0
      Top             =   0
      Width           =   9780
      Begin VB.HScrollBar hs2 
         Height          =   285
         Left            =   5310
         Max             =   100
         Min             =   1
         TabIndex        =   6
         Top             =   450
         Value           =   25
         Width           =   3885
      End
      Begin VB.HScrollBar hs1 
         Height          =   285
         Left            =   5310
         Max             =   1
         Min             =   1000
         TabIndex        =   4
         Top             =   90
         Value           =   50
         Width           =   3885
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Generate path and Move"
         Default         =   -1  'True
         Height          =   465
         Left            =   90
         TabIndex        =   2
         Top             =   180
         Width           =   2490
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Close Spline"
         Height          =   240
         Left            =   2700
         TabIndex        =   1
         Top             =   292
         Value           =   1  'Checked
         Width           =   1185
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Smoothness : "
         Height          =   195
         Index           =   1
         Left            =   4275
         TabIndex        =   5
         Top             =   495
         Width           =   1005
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000014&
         Index           =   1
         X1              =   4050
         X2              =   4050
         Y1              =   90
         Y2              =   765
      End
      Begin VB.Line Line1 
         BorderColor     =   &H80000010&
         Index           =   0
         X1              =   4065
         X2              =   4065
         Y1              =   90
         Y2              =   765
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         Caption         =   "Speed : "
         Height          =   195
         Index           =   0
         Left            =   4680
         TabIndex        =   3
         Top             =   135
         Width           =   600
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Hari Krishnan G. (a.k.a. eXeption)

Dim cpt(20) As POINTAPI, NumPts&
Dim pbez() As POINTAPI, pbexCount As Long, curptidx As Long, ptinc As Long
Dim nCurveSegments&

Private Sub Check1_Click()
    If Check1.Value = vbUnchecked Then
        Timer1.Enabled = False
        ptinc = 1
        curptidx = 0
        Timer1.Enabled = True
    End If
End Sub

Private Sub Command1_Click()
    Dim i, j, k
    
    Timer1.Enabled = False
    
    i = pcvs.ScaleWidth
    j = pcvs.ScaleHeight
    NumPts = 5
    
    For k = 1 To NumPts
        cpt(k).x = Rnd() * (i)
        cpt(k).y = Rnd() * (j)
    Next k
    
    If Check1.Value = vbChecked Then
        NumPts = NumPts + 1
        cpt(NumPts).x = cpt(1).x
        cpt(NumPts).y = cpt(1).y
    End If
    
    drawSpline
    
    Timer1.Enabled = True
End Sub

Sub drawSpline()
    Dim i&
    Timer1.Enabled = False
    pcvs.Cls
    pcvs.ForeColor = &H606060
    For i = 2 To NumPts
      pcvs.Line (Round(cpt(i).x), Round(cpt(i).y))-(Round(cpt(i - 1).x), Round(cpt(i - 1).y))
    Next
    
    Draw_CatmullRom pcvs, cpt, NumPts, nCurveSegments, &HFF0000, Check1.Value
    pcvs.Refresh
    
    pbexCount = GetCatmullRom_Points(cpt, NumPts, nCurveSegments, pbez, Check1.Value)
    
    curptidx = 0
    ptinc = 1
    Timer1.Enabled = True
End Sub

Private Sub Form_Load()
    Randomize

    pbexCount = 0
    curptidx = 0
    ptinc = 1
    
    hs1.Value = Timer1.Interval
    nCurveSegments = 25
    hs2.Value = nCurveSegments
    
    Command1_Click
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    pcvs.Move 3 * Screen.TwipsPerPixelY, picToolBar.Height + 5 * Screen.TwipsPerPixelY, Me.ScaleWidth - 7 * Screen.TwipsPerPixelX, Me.ScaleHeight - 10 * Screen.TwipsPerPixelY - picToolBar.Height
End Sub


Private Sub hs1_Change()
    Timer1.Interval = hs1.Value
End Sub

Private Sub hs1_Scroll()
    hs1_Change
End Sub

Private Sub hs2_Change()
    nCurveSegments = hs2.Value
    If NumPts > 0 Then drawSpline
End Sub

Private Sub Timer1_Timer()
    If pbexCount < 1 Then
        Timer1.Enabled = False
        Exit Sub
    End If
    If Check1.Value = vbUnchecked Then
        If curptidx >= pbexCount Then
            curptidx = pbexCount
            ptinc = -1
        ElseIf curptidx < 0 Then
            curptidx = 0
            ptinc = 1
        End If
    Else
        If curptidx >= pbexCount Then curptidx = curptidx - pbexCount
    End If
    
    shp.Move pbez(curptidx).x - shp.Width / 2, pbez(curptidx).y - shp.Width / 2
    shp.Visible = True
    curptidx = curptidx + ptinc
    DoEvents
End Sub
