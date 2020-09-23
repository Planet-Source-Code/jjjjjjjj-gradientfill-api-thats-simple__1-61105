VERSION 5.00
Begin VB.Form frmTest 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00A3E8FC&
   Caption         =   "[ Fastest\Simplest\Smoothest ]  Of All Gradients"
   ClientHeight    =   5805
   ClientLeft      =   180
   ClientTop       =   810
   ClientWidth     =   11850
   FillColor       =   &H000000FF&
   ForeColor       =   &H00FF0000&
   Icon            =   "frmTest.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   387
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   790
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picTop 
      Align           =   1  'Align Top
      Height          =   945
      Left            =   0
      ScaleHeight     =   885
      ScaleWidth      =   11790
      TabIndex        =   0
      Top             =   0
      Width           =   11850
      Begin VB.TextBox txtRepeat 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   10200
         TabIndex        =   15
         Text            =   "1"
         Top             =   240
         Width           =   1095
      End
      Begin VB.ListBox lstType 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   540
         ItemData        =   "frmTest.frx":000C
         Left            =   7080
         List            =   "frmTest.frx":0019
         TabIndex        =   14
         Top             =   240
         Width           =   2055
      End
      Begin VB.OptionButton OptEnd 
         Caption         =   "End"
         Height          =   240
         Left            =   3000
         TabIndex        =   10
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton optStart 
         Caption         =   "Start"
         Height          =   255
         Left            =   3000
         TabIndex        =   9
         Top             =   120
         Value           =   -1  'True
         Width           =   975
      End
      Begin VB.PictureBox picEnd 
         BackColor       =   &H00000000&
         Height          =   225
         Left            =   2400
         ScaleHeight     =   165
         ScaleWidth      =   435
         TabIndex        =   4
         Top             =   360
         Width           =   495
      End
      Begin VB.HScrollBar scrlBlue 
         Height          =   225
         LargeChange     =   50
         Left            =   4200
         Max             =   255
         SmallChange     =   10
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   600
         Width           =   1815
      End
      Begin VB.HScrollBar scrlGreen 
         Height          =   225
         LargeChange     =   50
         Left            =   4200
         Max             =   255
         SmallChange     =   10
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   360
         Width           =   1815
      End
      Begin VB.HScrollBar scrlRed 
         Height          =   225
         LargeChange     =   50
         Left            =   4200
         Max             =   255
         SmallChange     =   10
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   120
         Width           =   1815
      End
      Begin VB.PictureBox picStart 
         BackColor       =   &H000000FF&
         Height          =   225
         Left            =   2400
         ScaleHeight     =   165
         ScaleWidth      =   435
         TabIndex        =   3
         Top             =   120
         Width           =   495
      End
      Begin VB.PictureBox picCol 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         AutoSize        =   -1  'True
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         FillColor       =   &H00FDDBAC&
         ForeColor       =   &H80000008&
         Height          =   405
         Left            =   120
         Picture         =   "frmTest.frx":005B
         ScaleHeight     =   405
         ScaleWidth      =   2145
         TabIndex        =   1
         Top             =   120
         Width           =   2145
      End
      Begin VB.CheckBox chkRight 
         Caption         =   "< Right To Left "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9360
         TabIndex        =   17
         Top             =   600
         Value           =   1  'Checked
         Width           =   2415
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Repeat"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   9360
         TabIndex        =   16
         Top             =   240
         Width           =   780
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "< Green"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6120
         TabIndex        =   13
         Top             =   360
         Width           =   705
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "< Blue"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6120
         TabIndex        =   12
         Top             =   600
         Width           =   555
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "< Red"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6120
         TabIndex        =   11
         Top             =   120
         Width           =   525
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Gradient Type"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   7080
         TabIndex        =   5
         Top             =   0
         Width           =   1200
      End
      Begin VB.Label lbTime 
         AutoSize        =   -1  'True
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   120
         TabIndex        =   2
         Top             =   600
         Width           =   435
      End
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Not Tested in XP"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   240
      TabIndex        =   18
      Top             =   1080
      Width           =   1770
   End
   Begin VB.Menu mnu_Test 
      Caption         =   "Test Performance"
      Begin VB.Menu mnu_Sec 
         Caption         =   "Conduct test On 1 Secc"
         Shortcut        =   ^T
      End
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim gPause As Boolean

Private Sub chkRight_Click()
    lstType_Click
End Sub

Private Sub Form_Load()
    lstType.ListIndex = 2
End Sub

Private Sub Form_Resize()
    lstType_Click
End Sub

Private Sub lstType_Click()
Dim C As New cTiming
Me.Cls
    C.Reset
    SmoothGradient Me.hdc, picStart.BackColor, picEnd.BackColor, 0, 0, ScaleWidth, ScaleHeight, lstType.ListIndex - 1, chkRight, Val(txtRepeat)
    lbTime = "[ " & ScaleWidth & " * " & ScaleHeight & " ] Pix, Rendered in " & Format$(C.Elapsed / 1000, "0.0000 s") & vbCrLf & vbCrLf
    Debug.Print lbTime
    
    ' Positioned Gradient
    SmoothGradient Me.hdc, picStart.BackColor, picEnd.BackColor, 10, 100, 200, ScaleHeight - 150, lstType.ListIndex - 1, chkRight - 1, Val(txtRepeat)

End Sub

Private Sub mnu_Sec_Click()
Dim X As Long
Dim vPix As Double
Dim C As New cTiming
Me.Cls
    C.Reset
    While C.Elapsed < 1000
        X = X + 1
        DrawGradients Me, picStart.BackColor, picEnd.BackColor, 0, 0, ScaleWidth, ScaleHeight, lstType.ListIndex
    Wend
    vPix = X * ScaleHeight * ScaleWidth / 1000000
    MsgBox "Rendered " & X & " gradients of size [ " & ScaleWidth & " * " & ScaleHeight & " ] in 1 Second " & vbCrLf & vbCrLf & "Total Pixel Rendered " & Format(vPix, "00.000") & " Million "
    Debug.Print lbTime

End Sub

Private Sub picCol_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim R As Long, G As Long, B As Long
    If optStart Then picStart.BackColor = picCol.Point(X, Y)
    If OptEnd Then picEnd.BackColor = picCol.Point(X, Y)
    gPause = True
    GetRGB picCol.Point(X, Y), R, G, B
    scrlBlue = B: scrlRed = R: scrlGreen = G
    gPause = False
    lstType_Click
End Sub

Private Sub scrlBlue_Change()
    scrlRed_Change
End Sub

Private Sub scrlGreen_Change()
    scrlRed_Change
End Sub

Private Sub scrlRed_Change()
    If gPause Then Exit Sub
    If OptEnd Then
        picEnd.BackColor = RGB(scrlRed, scrlGreen, scrlBlue)
    Else
        picStart.BackColor = RGB(scrlRed, scrlGreen, scrlBlue)
    End If
    lstType_Click
End Sub

Private Sub scrlRed_Scroll()
    scrlRed_Change
End Sub

Private Sub txtRepeat_Change()
    lstType_Click
End Sub

Private Function GetRGB(ByVal LngCol As Long, R As Long, G As Long, B As Long)
    R = LngCol Mod 256
    G = (LngCol And vbGreen) / 256 'Green
    B = (LngCol And vbBlue) / 65536 'Blue
End Function
