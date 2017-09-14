VERSION 5.00
Begin VB.Form frmFourier 
   Caption         =   "Fourier Series Simulator"
   ClientHeight    =   7110
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13755
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6502.058
   ScaleMode       =   0  'User
   ScaleWidth      =   12090.81
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cbo 
      Height          =   405
      IntegralHeight  =   0   'False
      ItemData        =   "frmFourier.frx":0000
      Left            =   17280
      List            =   "frmFourier.frx":0019
      TabIndex        =   10
      Text            =   "Custom"
      Top             =   840
      Width           =   2775
   End
   Begin VB.CommandButton cmdPlot 
      Caption         =   "&Plot"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   17880
      TabIndex        =   9
      Top             =   2040
      Width           =   1695
   End
   Begin VB.PictureBox grpFun 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4065
      Left            =   12720
      ScaleHeight     =   2000
      ScaleLeft       =   -3142
      ScaleMode       =   0  'User
      ScaleTop        =   -1000
      ScaleWidth      =   6284
      TabIndex        =   8
      Top             =   4440
      Width           =   6780
   End
   Begin VB.PictureBox grpNetSin 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2624
      Left            =   6600
      ScaleHeight     =   2000
      ScaleLeft       =   -3142
      ScaleMode       =   0  'User
      ScaleTop        =   -1000
      ScaleWidth      =   6284
      TabIndex        =   7
      Top             =   6670
      Width           =   5461
   End
   Begin VB.PictureBox grpSinSeries 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2624
      Left            =   600
      ScaleHeight     =   2000
      ScaleLeft       =   -3142
      ScaleMode       =   0  'User
      ScaleTop        =   -1000
      ScaleWidth      =   6284
      TabIndex        =   6
      Top             =   6670
      Width           =   5461
   End
   Begin VB.PictureBox grpNetCos 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2624
      Left            =   6600
      ScaleHeight     =   2000
      ScaleLeft       =   -3142
      ScaleMode       =   0  'User
      ScaleTop        =   -1000
      ScaleWidth      =   6284
      TabIndex        =   5
      Top             =   3480
      Width           =   5461
   End
   Begin VB.PictureBox grpCosSeries 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2624
      Left            =   600
      ScaleHeight     =   2000
      ScaleLeft       =   -3142
      ScaleMode       =   0  'User
      ScaleTop        =   -1000
      ScaleWidth      =   6284
      TabIndex        =   4
      Top             =   3480
      Width           =   5461
   End
   Begin VB.Label lblWave 
      Alignment       =   2  'Center
      Caption         =   "Resultant Waveform"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   14040
      TabIndex        =   16
      Top             =   8640
      Width           =   4335
   End
   Begin VB.Label lblWave 
      Alignment       =   2  'Center
      Caption         =   "Resultant Sine Series"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   3
      Left            =   7080
      TabIndex        =   15
      Top             =   9360
      Width           =   4335
   End
   Begin VB.Label lblWave 
      Alignment       =   2  'Center
      Caption         =   "Resultant Cosine Series"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   2
      Left            =   7200
      TabIndex        =   14
      Top             =   6120
      Width           =   4335
   End
   Begin VB.Label lblWave 
      Alignment       =   2  'Center
      Caption         =   "Sine Waves"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   1200
      TabIndex        =   13
      Top             =   9360
      Width           =   4335
   End
   Begin VB.Label lblWave 
      Alignment       =   2  'Center
      Caption         =   "Cosine Waves"
      BeginProperty Font 
         Name            =   "Consolas"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   1080
      TabIndex        =   12
      Top             =   6120
      Width           =   4335
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      Caption         =   "Waveform"
      Height          =   255
      Left            =   17400
      TabIndex        =   11
      Top             =   360
      Width           =   2535
   End
   Begin VB.Line lineLimit 
      BorderColor     =   &H80000003&
      Index           =   3
      X1              =   2499.91
      X2              =   14900.13
      Y1              =   2725.195
      Y2              =   2725.195
   End
   Begin VB.Line lineLimit 
      BorderColor     =   &H80000003&
      Index           =   2
      X1              =   2499.91
      X2              =   14900.13
      Y1              =   1679.927
      Y2              =   1679.927
   End
   Begin VB.Line lineLimit 
      BorderColor     =   &H80000003&
      Index           =   1
      X1              =   2499.91
      X2              =   14900.13
      Y1              =   1516.232
      Y2              =   1516.232
   End
   Begin VB.Line lineLimit 
      BorderColor     =   &H80000003&
      Index           =   0
      X1              =   2499.91
      X2              =   14900.13
      Y1              =   485.597
      Y2              =   485.597
   End
   Begin VB.Label lbln 
      Alignment       =   2  'Center
      Caption         =   "1"
      Height          =   255
      Index           =   0
      Left            =   2844
      TabIndex        =   3
      Top             =   120
      Width           =   455
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "n="
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Sine Series"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   2160
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Cosine Series"
      Height          =   375
      Left            =   480
      TabIndex        =   0
      Top             =   960
      Width           =   2055
   End
   Begin VB.Line lineb 
      Index           =   0
      X1              =   2499.91
      X2              =   2899.86
      Y1              =   2200.275
      Y2              =   2200.275
   End
   Begin VB.Line linea 
      Index           =   0
      X1              =   2499.91
      X2              =   2899.86
      Y1              =   1000.457
      Y2              =   1000.457
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Index           =   0
      Begin VB.Menu mnuexit 
         Caption         =   "Exit"
         Index           =   9
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Index           =   10
      Begin VB.Menu mnuUse 
         Caption         =   "How To Use?"
         Index           =   12
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuAbout 
         Caption         =   "About"
         Index           =   16
      End
   End
End
Attribute VB_Name = "frmFourier"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim i As Byte
Dim an(0 To 24), bn(0 To 24) As Single
Dim bx, by As Single

Private Sub cmdPlot_Click()
    Dim asum, bsum, forfun(-3142 To 3142) As Single
    grpCosSeries.Cls
    grpSinSeries.Cls
    grpNetSin.Cls
    grpNetCos.Cls
    grpFun.Cls
    If cbo.Text <> "Custom" Then
        Call SeeValues
    End If
    For i = 0 To 24
        For bx = -3142 To 3142
            by = -an(i) * Cos(bx * (i + 1) / 500)
            grpCosSeries.PSet (bx, by), RGB(-i * 10 + 250, 0, i * 10)
        Next
        For bx = -3142 To 3142
            by = -bn(i) * Sin(bx * (i + 1) / 500)
            grpSinSeries.PSet (bx, by), RGB(-i * 10 + 250, 0, i * 10)
        Next
        asum = asum + Sqr(an(i) * an(i))
        bsum = bsum + Sqr(bn(i) * bn(i))
    Next
    If asum = 0 Then
        asum = 1
    End If
    If bsum = 0 Then
        bsum = 1
    End If
    For bx = -3142 To 3142
        by = 0
        For i = 0 To 24
            by = by - an(i) * Cos(bx * (i + 1) / 500)
        Next
        by = by / asum * 800
        grpNetCos.PSet (bx, by), vbGreen
        forfun(bx) = asum * by
    Next
    For bx = -3142 To 3142
        by = 0
        For i = 0 To 24
            by = by - bn(i) * Sin(bx * (i + 1) / 500)
        Next
        by = by / bsum * 800
        grpNetSin.PSet (bx, by), vbGreen
        forfun(bx) = forfun(bx) + bsum * by
    Next
    For bx = -3142 To 3142
        by = forfun(bx) / (asum + bsum)
        grpFun.PSet (bx, by), RGB(128, 128, 128)
    Next
End Sub
Private Sub SeeValues()
    Dim j As Byte
    Select Case cbo.Text
        Case "Cosine Wave":
            Call Form_MouseUp(vbLeftButton, 0, 2700, 489)
            Call Form_MouseUp(vbLeftButton, 0, 2700, 2200)
            For j = 1 To 24
                Call Form_MouseUp(vbLeftButton, 0, j * 500 + 2700, 1000)
                Call Form_MouseUp(vbLeftButton, 0, j * 500 + 2700, 2200)
            Next
            cbo.Text = "Cosine Wave"
        Case "Sine Wave":
            Call Form_MouseUp(vbLeftButton, 0, 2700, 2200 - 511)
            Call Form_MouseUp(vbLeftButton, 0, 2700, 1000)
            For j = 1 To 24
                Call Form_MouseUp(vbLeftButton, 0, j * 500 + 2700, 1000)
                Call Form_MouseUp(vbLeftButton, 0, j * 500 + 2700, 2200)
            Next
            cbo.Text = "Sine Wave"
        Case "Wavegroup":
            For j = 0 To 24
                Call Form_MouseUp(vbLeftButton, 0, j * 500 + 2700, 489)
                Call Form_MouseUp(vbLeftButton, 0, j * 500 + 2700, 2200)
            Next
            cbo.Text = "Wavegroup"
        Case "Square Wave":
            For j = 0 To 24 Step 2
                Call Form_MouseUp(vbLeftButton, 0, j * 500 + 2700, (2200 - 511 / (j + 1)))
                Call Form_MouseUp(vbLeftButton, 0, j * 500 + 2700, 1000)
            Next
            For j = 1 To 23 Step 2
                Call Form_MouseUp(vbLeftButton, 0, j * 500 + 2700, 1000)
                Call Form_MouseUp(vbLeftButton, 0, j * 500 + 2700, 2200)
            Next
            cbo.Text = "Square Wave"
        Case "Saw-Tooth Wave":
            For j = 0 To 24 Step 2
                Call Form_MouseUp(vbLeftButton, 0, j * 500 + 2700, 1000)
                Call Form_MouseUp(vbLeftButton, 0, j * 500 + 2700, (2200 - 511 / (j + 1)))
            Next
            For j = 1 To 23 Step 2
                Call Form_MouseUp(vbLeftButton, 0, j * 500 + 2700, 1000)
                Call Form_MouseUp(vbLeftButton, 0, j * 500 + 2700, (2200 + 511 / (j + 1)))
            Next
            cbo.Text = "Saw-Tooth Wave"
        Case "Triangular Wave":
            For j = 0 To 24 Step 2
                Call Form_MouseUp(vbLeftButton, 0, j * 500 + 2700, (1000 + 511 / (j + 1) / (j + 1)))
            Next
            For j = 1 To 23 Step 2
                Call Form_MouseUp(vbLeftButton, 0, j * 500 + 2700, 1000)
            Next
            For j = 0 To 24
                Call Form_MouseUp(vbLeftButton, 0, j * 500 + 2700, 2200)
            Next
            cbo.Text = "Triangular Wave"
        Case "No Waveform":
            For j = 0 To 24
                Call Form_MouseUp(vbLeftButton, 0, j * 500 + 2700, 1000)
                Call Form_MouseUp(vbLeftButton, 0, j * 500 + 2700, 2200)
            Next
            cbo.Text = "No Waveform"
    End Select
End Sub

Private Sub Form_Load()
    linea(0).BorderColor = RGB(250, 0, 10)
    lineb(0).BorderColor = RGB(250, 0, 10)
    lbln(0).ForeColor = RGB(250, 0, 10)
    For i = 1 To 24
        Load linea(i)
        linea(i).X1 = i * 500 + 2500
        linea(i).X2 = i * 500 + 2900
        linea(i).BorderColor = RGB(-i * 10 + 250, 0, i * 10)
        linea(i).Visible = True
        
        Load lbln(i)
        lbln(i).Left = i * 500 + 2500
        lbln(i).Caption = i + 1
        lbln(i).ForeColor = RGB(-i * 10 + 250, 0, i * 10)
        lbln(i).Visible = True
    Next i
    For i = 1 To 24
        Load lineb(i)
        lineb(i).X1 = i * 500 + 2500
        lineb(i).X2 = i * 500 + 2900
        lineb(i).BorderColor = RGB(-i * 10 + 250, 0, i * 10)
        lineb(i).Visible = True
    Next i
    'Call cmdPlot_Click
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.ScaleMode = 0
    If Button = vbLeftButton Then
        If Y > 488 And Y < 1512 Then
            If X > 2500 And X < 24 * 500 + 2900 Then
                i = CInt((X - 2700) / 500)
                If (X - i * 500 - 2500) > 400 Then
                Else
                    'linea(i).Tag = Y
                    an(i) = (1000 - Y) * 2000 / 1024
                    For by = 488 To 1512
                        If (Y - by) * (Y - by) < (Y - 1000) * (Y - 1000) And (1000 - by) * (1000 - by) < (Y - 1000) * (Y - 1000) Then
                            Me.Line (i * 500 + 2500, by)-(i * 500 + 2900, by), RGB(-i * 10 + 250, 0, i * 10)
                        Else
                            Me.Line (i * 500 + 2500, by)-(i * 500 + 2900, by), &H8000000F
                        End If
                    Next
                End If
            End If
        End If
        If Y > 1688 And Y < 2712 Then
            If X > 2500 And X < 24 * 500 + 2900 Then
                i = CInt((X - 2700) / 500)
                If (X - i * 500 - 2500) > 400 Then
                Else
                    'lineb(i).Tag = Y
                    bn(i) = (2200 - Y) * 2000 / 1024
                    For by = 1688 To 2712
                        If (Y - by) * (Y - by) < (Y - 2200) * (Y - 2200) And (2200 - by) * (2200 - by) < (Y - 2200) * (Y - 2200) Then
                            Me.Line (i * 500 + 2500, by)-(i * 500 + 2900, by), RGB(-i * 10 + 250, 0, i * 10)
                        Else
                            Me.Line (i * 500 + 2500, by)-(i * 500 + 2900, by), &H8000000F
                        End If
                    Next
                End If
            End If
        End If
    End If
    cbo.Text = "Custom"
End Sub

Private Sub mnuAbout_Click(Index As Integer)
    frmAbout.Show vbModal
End Sub

Private Sub mnuexit_Click(Index As Integer)
    Unload Me
    Set frmFourier = Nothing
End Sub

Private Sub mnuUse_Click(Index As Integer)
    Call MsgBox("       It's very easy!" & vbCrLf & "   You can choose amplitudes for various cosine and sine terrms in the Fourier series. Click on 'Plot' button to display the waveforms." & vbCrLf & " You can also view some built-in waveforms from dropdown menu.", vbOKOnly + vbInformation, "How to use?")
End Sub
