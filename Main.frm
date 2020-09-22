VERSION 5.00
Begin VB.Form Main 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   840
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1230
   LinkTopic       =   "Form1"
   ScaleHeight     =   56
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   82
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   120
      Top             =   240
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   2040
      Top             =   1800
   End
   Begin VB.PictureBox P 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   360
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   33
      TabIndex        =   0
      Top             =   0
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Notepad"
      BeginProperty Font 
         Name            =   "Terminal"
         Size            =   6
         Charset         =   255
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Left            =   0
      TabIndex        =   1
      Top             =   600
      Width           =   1215
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Dim FrameIndex As Integer


Private Sub Form_Load()
    Me.Left = Screen.Width - Me.Width - 500
    Me.Top = 500
    FrameIndex = 100
    CaptureDektop Me
End Sub

Private Sub P_DblClick()
    Shell "c:\windows\notepad.exe", vbNormalFocus
End Sub


Private Sub Timer1_Timer()
    FrameIndex = FrameIndex + 1
    
    With P
        
        .Picture = LoadResPicture(FrameIndex, 0)
        
        For X = 0 To .ScaleWidth
            For Y = 0 To .ScaleHeight
        
                If .Point(X, Y) = vbWhite Then
                    P.PSet (X, Y), Me.Point(X + .Left, Y + .Top)
                End If
                
            Next Y
        Next X
        
    End With
    
    If FrameIndex = 127 Then FrameIndex = 100
End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    ReleaseCapture
    CaptureDektop Me
    SendMessage hWnd, WM_NCLBUTTONDOWN, 2, 0&
End Sub

