VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "CaptionbarFX V5"
   ClientHeight    =   4935
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8535
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   4935
   ScaleWidth      =   8535
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture6 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   3
      Left            =   120
      Picture         =   "Form1.frx":0442
      ScaleHeight     =   435
      ScaleWidth      =   9600
      TabIndex        =   61
      Top             =   7920
      Visible         =   0   'False
      Width           =   9660
   End
   Begin VB.PictureBox Picture6 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   2
      Left            =   120
      Picture         =   "Form1.frx":2904
      ScaleHeight     =   435
      ScaleWidth      =   9600
      TabIndex        =   60
      Top             =   7800
      Visible         =   0   'False
      Width           =   9660
   End
   Begin VB.PictureBox Picture6 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   1
      Left            =   120
      Picture         =   "Form1.frx":4DC6
      ScaleHeight     =   435
      ScaleWidth      =   9600
      TabIndex        =   59
      Top             =   7680
      Visible         =   0   'False
      Width           =   9660
   End
   Begin VB.PictureBox Picture6 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Index           =   0
      Left            =   120
      Picture         =   "Form1.frx":7288
      ScaleHeight     =   435
      ScaleWidth      =   9600
      TabIndex        =   58
      Top             =   7560
      Visible         =   0   'False
      Width           =   9660
   End
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   8520
      Top             =   1200
   End
   Begin VB.Timer Timer2 
      Interval        =   5
      Left            =   8520
      Top             =   720
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   8520
      Top             =   240
   End
   Begin VB.Frame Frame13 
      Caption         =   "Animations"
      Height          =   1335
      Left            =   4440
      TabIndex        =   53
      Top             =   1920
      Width           =   1815
      Begin VB.CommandButton Command16 
         Caption         =   "Animation 4"
         Height          =   255
         Left            =   120
         TabIndex        =   57
         Top             =   960
         Width           =   1575
      End
      Begin VB.CommandButton Command15 
         Caption         =   "Animation 3"
         Height          =   255
         Left            =   120
         TabIndex        =   56
         Top             =   720
         Width           =   1575
      End
      Begin VB.CommandButton Command14 
         Caption         =   "Animation 2"
         Height          =   255
         Left            =   120
         TabIndex        =   55
         Top             =   480
         Width           =   1575
      End
      Begin VB.CommandButton Command13 
         Caption         =   "Animation 1"
         Height          =   255
         Left            =   120
         TabIndex        =   54
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.Frame Frame12 
      Caption         =   "Caption:"
      Height          =   1335
      Left            =   6360
      TabIndex        =   50
      Top             =   1920
      Width           =   2055
      Begin VB.CommandButton Command12 
         Caption         =   "Change"
         Height          =   375
         Left            =   120
         TabIndex        =   52
         Top             =   840
         Width           =   1815
      End
      Begin VB.TextBox Text3 
         Height          =   315
         Left            =   120
         TabIndex        =   51
         Text            =   "CaptionbarFX V5"
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton Command11 
      Caption         =   "R"
      Height          =   255
      Left            =   8640
      TabIndex        =   49
      Top             =   5040
      Width           =   255
   End
   Begin VB.Frame Frame11 
      Caption         =   "Flat-Color"
      Height          =   1455
      Left            =   7200
      TabIndex        =   44
      Top             =   3360
      Width           =   1215
      Begin VB.CommandButton Command10 
         Caption         =   "Change"
         Height          =   375
         Left            =   120
         TabIndex        =   48
         Top             =   960
         Width           =   975
      End
      Begin VB.OptionButton Option19 
         Caption         =   "Blue"
         Height          =   255
         Left            =   120
         TabIndex        =   47
         Top             =   720
         Width           =   975
      End
      Begin VB.OptionButton Option18 
         Caption         =   "Green"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   480
         Width           =   975
      End
      Begin VB.OptionButton Option17 
         Caption         =   "Red"
         Height          =   255
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Value           =   -1  'True
         Width           =   975
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "Pixel Generator"
      Height          =   1455
      Left            =   120
      TabIndex        =   27
      Top             =   3360
      Width           =   6975
      Begin VB.Frame Frame10 
         Caption         =   "Text Color"
         Height          =   1095
         Left            =   3960
         TabIndex        =   41
         Top             =   240
         Width           =   1335
         Begin VB.OptionButton Option16 
            Caption         =   "White"
            Height          =   255
            Left            =   120
            TabIndex        =   43
            Top             =   480
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton Option15 
            Caption         =   "Black"
            Height          =   255
            Left            =   120
            TabIndex        =   42
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "FX-Mode"
         Height          =   1095
         Left            =   2520
         TabIndex        =   37
         Top             =   240
         Width           =   1335
         Begin VB.OptionButton Option14 
            Caption         =   "Mode 2"
            Height          =   255
            Left            =   120
            TabIndex        =   40
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton Option13 
            Caption         =   "Mode 1"
            Height          =   255
            Left            =   120
            TabIndex        =   39
            Top             =   480
            Width           =   1095
         End
         Begin VB.OptionButton Option12 
            Caption         =   "Off"
            Height          =   255
            Left            =   120
            TabIndex        =   38
            Top             =   240
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame8 
         Caption         =   "Back Color:"
         Height          =   1095
         Left            =   1320
         TabIndex        =   33
         Top             =   240
         Width           =   1095
         Begin VB.OptionButton Option11 
            Caption         =   "Blue"
            Height          =   255
            Left            =   120
            TabIndex        =   36
            Top             =   720
            Width           =   615
         End
         Begin VB.OptionButton Option10 
            Caption         =   "Green"
            Height          =   255
            Left            =   120
            TabIndex        =   35
            Top             =   480
            Value           =   -1  'True
            Width           =   855
         End
         Begin VB.OptionButton Option9 
            Caption         =   "Red"
            Height          =   255
            Left            =   120
            TabIndex        =   34
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "Pixel Color:"
         Height          =   1095
         Left            =   120
         TabIndex        =   29
         Top             =   240
         Width           =   1095
         Begin VB.OptionButton Option8 
            Caption         =   "Blue"
            Height          =   195
            Left            =   120
            TabIndex        =   32
            Top             =   720
            Width           =   615
         End
         Begin VB.OptionButton Option7 
            Caption         =   "Green"
            Height          =   255
            Left            =   120
            TabIndex        =   31
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton Option6 
            Caption         =   "Red"
            Height          =   255
            Left            =   120
            TabIndex        =   30
            Top             =   240
            Value           =   -1  'True
            Width           =   735
         End
      End
      Begin VB.CommandButton Command9 
         Caption         =   "Generate pixels"
         Height          =   975
         Left            =   5400
         TabIndex        =   28
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture5 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   120
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   85
      TabIndex        =   26
      Top             =   5040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Frame Frame4 
      Caption         =   "Custom Gradient"
      Height          =   1695
      Left            =   3240
      TabIndex        =   14
      Top             =   120
      Width           =   5175
      Begin VB.Frame Frame5 
         Caption         =   "Colors:"
         Height          =   1335
         Left            =   3480
         TabIndex        =   22
         Top             =   240
         Width           =   1575
         Begin VB.OptionButton Option5 
            Caption         =   "Blue"
            Height          =   255
            Left            =   120
            TabIndex        =   25
            Top             =   960
            Width           =   1335
         End
         Begin VB.OptionButton Option4 
            Caption         =   "Green"
            Height          =   255
            Left            =   120
            TabIndex        =   24
            Top             =   600
            Width           =   1215
         End
         Begin VB.OptionButton Option3 
            Caption         =   "Red"
            Height          =   255
            Left            =   120
            TabIndex        =   23
            Top             =   240
            Value           =   -1  'True
            Width           =   1335
         End
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Up-Down"
         Height          =   255
         Left            =   2280
         TabIndex        =   21
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Left-Right"
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   240
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   2280
         TabIndex        =   17
         Text            =   "200"
         Top             =   960
         Width           =   975
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2280
         TabIndex        =   16
         Text            =   "0"
         Top             =   600
         Width           =   975
      End
      Begin VB.CommandButton Command8 
         Caption         =   "Change"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   1320
         Width           =   3135
      End
      Begin VB.Label Label2 
         Caption         =   "Gradient end color:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   1005
         Width           =   2055
      End
      Begin VB.Label Label1 
         Caption         =   "Gradient begin color:"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   645
         Width           =   1935
      End
   End
   Begin VB.PictureBox Picture4 
      AutoRedraw      =   -1  'True
      Height          =   495
      Left            =   1680
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   173
      TabIndex        =   13
      Top             =   5040
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.PictureBox Picture3 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   120
      Picture         =   "Form1.frx":974A
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   12
      Top             =   6600
      Visible         =   0   'False
      Width           =   9660
   End
   Begin VB.PictureBox Picture2 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   120
      Picture         =   "Form1.frx":E40C
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   10
      Top             =   6120
      Visible         =   0   'False
      Width           =   9660
   End
   Begin VB.Frame Frame3 
      Caption         =   "Add Bitmap"
      Height          =   1695
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   3015
      Begin VB.CommandButton Command7 
         Caption         =   "Bitmap 3"
         Height          =   375
         Left            =   120
         TabIndex        =   11
         Top             =   1200
         Width           =   2775
      End
      Begin VB.CommandButton Command6 
         Caption         =   "Bitmap 2"
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   2775
      End
      Begin VB.CommandButton Command5 
         Caption         =   "Bitmap 1"
         Height          =   375
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Gradient: LEFT to RIGHT"
      Height          =   1335
      Left            =   2280
      TabIndex        =   4
      Top             =   1920
      Width           =   2055
      Begin VB.CommandButton Command4 
         Caption         =   "Change Gradient 2"
         Height          =   375
         Left            =   120
         TabIndex        =   6
         Top             =   840
         Width           =   1815
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Change Gradient 1"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Gradient: UP to DOWN"
      Height          =   1335
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   2055
      Begin VB.CommandButton Command1 
         Caption         =   "Change Gradient 1"
         Height          =   375
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1815
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Change Gradient 2"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   1815
      End
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   495
      Left            =   120
      Picture         =   "Form1.frx":130CE
      ScaleHeight     =   29
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   5640
      Visible         =   0   'False
      Width           =   9660
   End
   Begin VB.Label Label3 
      Caption         =   "Animation picture boxes:"
      Height          =   255
      Left            =   120
      TabIndex        =   62
      Top             =   7200
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************************************************
'                                                                                                          *
'CaptionbarFX V5 project by Peter Hebels, Website "http://www.phsoft.nl"                                   *
'The author of this code cannot be held responsible for any damages may caused by this project.            *
'                                                                                                          *
'***********************************************************************************************************

Dim K As Integer
Dim Animation As Integer

Private Sub Command10_Click()
    StopAnimation
    
    Set WithForm = Form1
    Set MenuFrm = Form2
  
    CaptionbarFX.GradForceColors = True
    CaptionbarFX.GradVerticalGradient = True
    CaptionbarFX.AddBitmap = False
    CaptionbarFX.BitmapDC = Picture1.hdc
    CaptionbarFX.BitmapH = Picture1.ScaleHeight
    CaptionbarFX.BitmapW = Picture1.ScaleWidth
    
    CaptionbarFX.GradForcedText = vbWhite
    
    If Option17.Value = True Then
        CaptionbarFX.GradForcedFirst = vbRed
        CaptionbarFX.GradForcedSecond = vbRed
    ElseIf Option18.Value = True Then
        CaptionbarFX.GradForcedFirst = vbGreen
        CaptionbarFX.GradForcedSecond = vbGreen
    ElseIf Option19.Value = True Then
        CaptionbarFX.GradForcedFirst = vbBlue
        CaptionbarFX.GradForcedSecond = vbBlue
    End If
   
    CaptionbarFX.GradForcedTextA = &HC0C0C0
    CaptionbarFX.GradForcedFirstA = vbBlack
    CaptionbarFX.GradForcedSecondA = vbBlue
    CaptionbarFX.GradientGetCapsFont
        
    CaptionbarFX.RedrawBar Me
End Sub

Private Sub Command11_Click()
    Form1.Width = 8655
    Form1.Height = 5340
End Sub

Private Sub Command12_Click()
    StopAnimation
    Me.Caption = Text3.Text
    CaptionbarFX.RedrawBar Me
End Sub

Private Sub Command13_Click()
    StopAnimation
    Animation = 1
    Timer1.Enabled = True
End Sub

Private Sub Command14_Click()
    StopAnimation
    Animation = 2
    Timer1.Enabled = True
End Sub

Private Sub Command15_Click()
    StopAnimation
    Animation = 3
    Timer1.Enabled = True
End Sub

Private Sub Command16_Click()
    StopAnimation
    K = -1
    Timer3.Enabled = True
End Sub

Private Sub Command9_Click()
    Dim I As Long
    Dim K As Long
    Dim TheColor As Long
    Dim TheBackColor As Long
    
    Set WithForm = Form1
    Set MenuFrm = Form2
    
    StopAnimation
    
    If Option6.Value = True Then
        TheColor = vbRed
    ElseIf Option7.Value = True Then
        TheColor = vbGreen
    ElseIf Option8.Value = True Then
        TheColor = vbBlue
    End If
    
    If Option9.Value = True Then
        TheBackColor = vbRed
    ElseIf Option10.Value = True Then
        TheBackColor = vbGreen
    ElseIf Option11.Value = True Then
        TheBackColor = vbBlue
    End If
    
    Picture5.Cls
    
    Picture5.BackColor = TheBackColor
    
    For I = 0 To Picture5.ScaleWidth
        Val1 = Int((K * Rnd) + I)
        For K = 0 To Picture5.ScaleHeight
            Val2 = Int((K * Rnd) + K)
            
            If Option13.Value = True Then
                TheColor = TheColor + 1
            ElseIf Option14.Value = True Then
                TheColor = TheColor * 1
            End If
            
            SetPixel Picture5.hdc, Val1, Val2, TheColor
        Next
        DoEvents
    Next
    
    Picture4.ScaleMode = Form1.ScaleMode
    Picture4.Left = 0
    Picture4.Width = Screen.Width + 120
    Picture4.ScaleMode = 3
    Picture4.Cls
    
    TilePicture Picture5.hdc, Picture4.hdc, Picture5.ScaleWidth, Picture5.ScaleHeight, Picture4.Width, Picture4.Height
    
    CaptionbarFX.GradForceColors = True
    CaptionbarFX.GradVerticalGradient = False
    CaptionbarFX.AddBitmap = True
    CaptionbarFX.BitmapDC = Picture4.hdc
    CaptionbarFX.BitmapH = Picture4.ScaleHeight
    CaptionbarFX.BitmapW = Picture4.ScaleWidth
        
    If Option15.Value = True Then
        CaptionbarFX.GradForcedText = vbBlack
    ElseIf Option16.Value = True Then
        CaptionbarFX.GradForcedText = vbWhite
    End If
    
    CaptionbarFX.GradForcedFirst = vbRed
    CaptionbarFX.GradForcedSecond = vbRed
    
    CaptionbarFX.GradForcedTextA = &HC0C0C0
    CaptionbarFX.GradForcedFirstA = vbBlack
    CaptionbarFX.GradForcedSecondA = vbBlue
    
    CaptionbarFX.GradientGetCapsFont
        
    CaptionbarFX.RedrawBar Me
End Sub

Private Sub Form_Load()
  'Identify the mainform and the menu to be used
  Set WithForm = Form1
  Set MenuFrm = Form2 'Name of the menu has to be 'MnuPop'
  
    'Options for CaptionbarFX
    CaptionbarFX.GradForceColors = True
    CaptionbarFX.GradVerticalGradient = True 'Use horizontal or vertical gradient
    
    CaptionbarFX.AddBitmap = False 'Use a nice bitmap in the captionbar
    CaptionbarFX.BitmapDC = Picture1.hdc 'If the above is true, with picturebox its hdc has to be used?
    'Note that the picturebox you use must have Autosize to true and Autoredraw to true, otherwise strange sideffects will occour.
    CaptionbarFX.BitmapH = Picture1.ScaleHeight 'Height of bitmap
    CaptionbarFX.BitmapW = Picture1.ScaleWidth 'Width of bitmap
    
    'Textcolors
    CaptionbarFX.GradForcedText = vbWhite
    CaptionbarFX.GradForcedFirst = vbRed
    CaptionbarFX.GradForcedSecond = vbBlue
   
    CaptionbarFX.GradForcedTextA = &HC0C0C0
    CaptionbarFX.GradForcedFirstA = vbBlack
    CaptionbarFX.GradForcedSecondA = vbBlue

    'Call this and the settings will be used!!
    CaptionbarFX.GradientGetCapsFont
    CaptionbarFX.GradientCaption Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    StopAnimation
    'Always call this function, otherwise VB will crash!!
    CaptionbarFX.GradientReleaseForm Me
    
    'Now unload the forms, Do not use END
    Unload Form1
    Unload Form2
End Sub

Private Sub Command1_Click()
    StopAnimation
    
    Set WithForm = Form1
    Set MenuFrm = Form2

    CaptionbarFX.GradForceColors = True
    CaptionbarFX.GradVerticalGradient = True
    CaptionbarFX.AddBitmap = False
    CaptionbarFX.BitmapDC = Picture1.hdc
    CaptionbarFX.BitmapH = Picture1.ScaleHeight
    CaptionbarFX.BitmapW = Picture1.ScaleWidth
    
    CaptionbarFX.GradForcedText = vbWhite
    CaptionbarFX.GradForcedFirst = vbGreen
    CaptionbarFX.GradForcedSecond = vbBlue
   
    CaptionbarFX.GradForcedTextA = &HC0C0C0
    CaptionbarFX.GradForcedFirstA = vbBlack
    CaptionbarFX.GradForcedSecondA = vbBlue
    CaptionbarFX.GradientGetCapsFont
        
    'Redraw the CaptionBar.
    CaptionbarFX.RedrawBar Me
End Sub

Private Sub Command2_Click()
    StopAnimation
    
    Set WithForm = Form1
    Set MenuFrm = Form2
  
    CaptionbarFX.GradForceColors = True
    CaptionbarFX.GradVerticalGradient = True
    CaptionbarFX.AddBitmap = False
    CaptionbarFX.BitmapDC = Picture1.hdc
    CaptionbarFX.BitmapH = Picture1.ScaleHeight
    CaptionbarFX.BitmapW = Picture1.ScaleWidth
    
    CaptionbarFX.GradForcedText = vbWhite
    CaptionbarFX.GradForcedFirst = vbMagenta
    CaptionbarFX.GradForcedSecond = vbBlue
   
    CaptionbarFX.GradForcedTextA = &HC0C0C0
    CaptionbarFX.GradForcedFirstA = vbBlack
    CaptionbarFX.GradForcedSecondA = vbBlue
    CaptionbarFX.GradientGetCapsFont
        
    CaptionbarFX.RedrawBar Me
End Sub

Private Sub Command3_Click()
StopAnimation
      Set WithForm = Form1
      Set MenuFrm = Form2
  
    CaptionbarFX.GradForceColors = True
    CaptionbarFX.GradVerticalGradient = False
    CaptionbarFX.AddBitmap = False
    CaptionbarFX.BitmapDC = Picture1.hdc
    CaptionbarFX.BitmapH = Picture1.ScaleHeight
    CaptionbarFX.BitmapW = Picture1.ScaleWidth
    
    CaptionbarFX.GradForcedText = vbWhite
    CaptionbarFX.GradForcedFirst = vbMagenta
    CaptionbarFX.GradForcedSecond = vbBlue
   
    CaptionbarFX.GradForcedTextA = &HC0C0C0
    CaptionbarFX.GradForcedFirstA = vbBlack
    CaptionbarFX.GradForcedSecondA = vbBlue

    CaptionbarFX.GradientGetCapsFont
        
    CaptionbarFX.RedrawBar Me
End Sub

Private Sub Command4_Click()
StopAnimation
    Set WithForm = Form1
    Set MenuFrm = Form2
  
    CaptionbarFX.GradForceColors = True
    CaptionbarFX.GradVerticalGradient = False
    CaptionbarFX.AddBitmap = False
    CaptionbarFX.BitmapDC = Picture1.hdc
    CaptionbarFX.BitmapH = Picture1.ScaleHeight
    CaptionbarFX.BitmapW = Picture1.ScaleWidth
    
    CaptionbarFX.GradForcedText = vbWhite
    CaptionbarFX.GradForcedFirst = vbRed
    CaptionbarFX.GradForcedSecond = vbBlue
   
    CaptionbarFX.GradForcedTextA = &HC0C0C0
    CaptionbarFX.GradForcedFirstA = vbBlack
    CaptionbarFX.GradForcedSecondA = vbBlue

    CaptionbarFX.GradientGetCapsFont
        
    CaptionbarFX.RedrawBar Me
End Sub

Private Sub Command5_Click()
    StopAnimation
    
    Set WithForm = Form1
    Set MenuFrm = Form2
  
    Picture4.ScaleMode = Form1.ScaleMode
    Picture4.Left = 0
    Picture4.Width = Screen.Width + 120
    Picture4.ScaleMode = 3
    Picture4.Cls
    
    TilePicture Picture1.hdc, Picture4.hdc, Picture1.ScaleWidth, Picture1.ScaleHeight, Picture4.Width, Picture4.Height
    
    CaptionbarFX.GradForceColors = True
    CaptionbarFX.GradVerticalGradient = False
    CaptionbarFX.AddBitmap = True
    CaptionbarFX.BitmapDC = Picture4.hdc
    CaptionbarFX.BitmapH = Picture4.ScaleHeight
    CaptionbarFX.BitmapW = Picture4.ScaleWidth
    
    CaptionbarFX.GradForcedText = vbWhite
    CaptionbarFX.GradForcedFirst = vbRed
    CaptionbarFX.GradForcedSecond = vbRed
   
    CaptionbarFX.GradForcedTextA = &HC0C0C0
    CaptionbarFX.GradForcedFirstA = vbBlack
    CaptionbarFX.GradForcedSecondA = vbBlue

    CaptionbarFX.GradientGetCapsFont
        
    CaptionbarFX.RedrawBar Me
End Sub

Private Sub Command6_Click()
    StopAnimation
    
    Set WithForm = Form1
    Set MenuFrm = Form2
  
    Picture4.ScaleMode = Form1.ScaleMode
    Picture4.Left = 0
    Picture4.Width = Screen.Width + 120
    Picture4.ScaleMode = 3
    Picture4.Cls
    
    TilePicture Picture2.hdc, Picture4.hdc, Picture2.ScaleWidth, Picture2.ScaleHeight, Picture4.Width, Picture4.Height
    
    CaptionbarFX.GradForceColors = True
    CaptionbarFX.GradVerticalGradient = False
    CaptionbarFX.AddBitmap = True
    CaptionbarFX.BitmapDC = Picture4.hdc
    CaptionbarFX.BitmapH = Picture4.ScaleHeight
    CaptionbarFX.BitmapW = Picture4.ScaleWidth
    
    CaptionbarFX.GradForcedText = vbBlack
    CaptionbarFX.GradForcedFirst = vbRed
    CaptionbarFX.GradForcedSecond = vbRed
   
    CaptionbarFX.GradForcedTextA = vbBlack
    CaptionbarFX.GradForcedFirstA = vbBlack
    CaptionbarFX.GradForcedSecondA = vbBlack

    CaptionbarFX.GradientGetCapsFont
        
    CaptionbarFX.RedrawBar Me
End Sub

Private Sub Command7_Click()
    StopAnimation
      
    Set WithForm = Form1
    Set MenuFrm = Form2
  
    Picture4.ScaleMode = Form1.ScaleMode
    Picture4.Left = 0
    Picture4.Width = Screen.Width + 120
    Picture4.ScaleMode = 3
    Picture4.Cls
    
    TilePicture Picture3.hdc, Picture4.hdc, Picture3.ScaleWidth, Picture3.ScaleHeight, Picture4.Width, Picture4.Height
      
    CaptionbarFX.GradForceColors = True
    CaptionbarFX.GradVerticalGradient = False
    CaptionbarFX.AddBitmap = True
    CaptionbarFX.BitmapDC = Picture4.hdc
    CaptionbarFX.BitmapH = Picture4.ScaleHeight
    CaptionbarFX.BitmapW = Picture4.ScaleWidth
    
    CaptionbarFX.GradForcedText = vbWhite
    CaptionbarFX.GradForcedFirst = vbGreen
    CaptionbarFX.GradForcedSecond = vbGreen
   
    CaptionbarFX.GradForcedTextA = &HC0C0C0
    CaptionbarFX.GradForcedFirstA = vbBlack
    CaptionbarFX.GradForcedSecondA = vbBlue

    CaptionbarFX.GradientGetCapsFont
        
    CaptionbarFX.RedrawBar Me
End Sub

Private Sub Command8_Click()
    StopAnimation
    
    If Text1.Text > 255 Then
        MsgBox "Please enter a number between 0 and 255.", vbExclamation, "Error"
        Exit Sub
    ElseIf Text2.Text > 255 Then
        MsgBox "Please enter a number between 0 and 255.", vbExclamation, "Error"
        Exit Sub
    ElseIf Text1.Text < 0 Then
        MsgBox "Please enter a number between 0 and 255.", vbExclamation, "Error"
        Exit Sub
    ElseIf Text2.Text < 0 Then
        MsgBox "Please enter a number between 0 and 255.", vbExclamation, "Error"
        Exit Sub
    End If
    
    Set WithForm = Form1
    Set MenuFrm = Form2
      
    CaptionbarFX.GradForceColors = True
    
    If Option2.Value = True Then
        CaptionbarFX.GradVerticalGradient = True
    Else
        CaptionbarFX.GradVerticalGradient = False
    End If
    
    CaptionbarFX.AddBitmap = False
    CaptionbarFX.BitmapDC = Picture1.hdc
    CaptionbarFX.BitmapH = Picture1.ScaleHeight
    CaptionbarFX.BitmapW = Picture1.ScaleWidth
    
    CaptionbarFX.GradForcedText = vbWhite
    
    If Option3.Value = True Then
        CaptionbarFX.GradForcedFirst = Text1.Text
        CaptionbarFX.GradForcedSecond = Text2.Text
    ElseIf Option4.Value = True Then
        CaptionbarFX.GradForcedFirst = vbGreen + Text1.Text
        CaptionbarFX.GradForcedSecond = vbGreen + Text2.Text
    ElseIf Option5.Value = True Then
        CaptionbarFX.GradForcedFirst = vbBlue + Text1.Text
        CaptionbarFX.GradForcedSecond = vbBlue + Text2.Text
    End If
    
    CaptionbarFX.GradForcedTextA = &HC0C0C0
    CaptionbarFX.GradForcedFirstA = vbBlack
    CaptionbarFX.GradForcedSecondA = vbBlue
    CaptionbarFX.GradientGetCapsFont
        
    CaptionbarFX.RedrawBar Me
End Sub

'--------------------------Animation timers---------------------
Private Sub Timer1_Timer()
    K = K + 10
    
    If Animation = 1 Then
        If K = 220 Then
            Timer1.Enabled = False
            Timer2.Enabled = True
        End If
      
        CaptionbarFX.GradForceColors = True
        CaptionbarFX.GradVerticalGradient = False
        CaptionbarFX.AddBitmap = False
            
        CaptionbarFX.GradForcedText = vbWhite
        CaptionbarFX.GradForcedFirst = K
        CaptionbarFX.GradForcedSecond = 220 - K
       
        CaptionbarFX.GradForcedTextA = &HC0C0C0
        CaptionbarFX.GradForcedFirstA = vbBlack
        CaptionbarFX.GradForcedSecondA = vbBlue
    
        CaptionbarFX.GradientGetCapsFont
            
        CaptionbarFX.RedrawBar Form1
    
    ElseIf Animation = 2 Then
        If K = 220 Then
            Timer1.Enabled = False
            Timer2.Enabled = True
        End If
      
        CaptionbarFX.GradForceColors = True
        CaptionbarFX.GradVerticalGradient = True
        CaptionbarFX.AddBitmap = False
            
        CaptionbarFX.GradForcedText = vbWhite
        CaptionbarFX.GradForcedFirst = K
        CaptionbarFX.GradForcedSecond = 220 - K
       
        CaptionbarFX.GradForcedTextA = &HC0C0C0
        CaptionbarFX.GradForcedFirstA = vbBlack
        CaptionbarFX.GradForcedSecondA = vbBlue
    
        CaptionbarFX.GradientGetCapsFont
            
        CaptionbarFX.RedrawBar Form1
    
    ElseIf Animation = 3 Then
        If K = 220 Then
            Timer1.Enabled = False
            Timer2.Enabled = True
        End If
      
        CaptionbarFX.GradForceColors = True
        CaptionbarFX.GradVerticalGradient = False
        CaptionbarFX.AddBitmap = False
            
        CaptionbarFX.GradForcedText = vbWhite
        CaptionbarFX.GradForcedFirst = K - vbGreen
        CaptionbarFX.GradForcedSecond = 220 - K - vbBlue
       
        CaptionbarFX.GradForcedTextA = &HC0C0C0
        CaptionbarFX.GradForcedFirstA = vbBlack
        CaptionbarFX.GradForcedSecondA = vbBlue
    
        CaptionbarFX.GradientGetCapsFont
            
        CaptionbarFX.RedrawBar Form1
    End If
End Sub

Private Sub Timer2_Timer()
    K = K - 10
    
    If Animation = 1 Then
        If K = 0 Then
            Timer2.Enabled = False
            Timer1.Enabled = True
        End If
    
        CaptionbarFX.GradForceColors = True
        CaptionbarFX.GradVerticalGradient = False
        CaptionbarFX.AddBitmap = False
            
        CaptionbarFX.GradForcedText = vbWhite
        CaptionbarFX.GradForcedFirst = K
        CaptionbarFX.GradForcedSecond = 220 - K
       
        CaptionbarFX.GradForcedTextA = &HC0C0C0
        CaptionbarFX.GradForcedFirstA = vbBlack
        CaptionbarFX.GradForcedSecondA = vbBlue
    
        CaptionbarFX.GradientGetCapsFont
            
        CaptionbarFX.RedrawBar Form1
    ElseIf Animation = 2 Then
        If K = 0 Then
            Timer2.Enabled = False
            Timer1.Enabled = True
        End If
    
        CaptionbarFX.GradForceColors = True
        CaptionbarFX.GradVerticalGradient = True
        CaptionbarFX.AddBitmap = False
            
        CaptionbarFX.GradForcedText = vbWhite
        CaptionbarFX.GradForcedFirst = K
        CaptionbarFX.GradForcedSecond = 210 - K
        
        CaptionbarFX.GradForcedTextA = &HC0C0C0
        CaptionbarFX.GradForcedFirstA = vbBlack
        CaptionbarFX.GradForcedSecondA = vbBlue
        
        CaptionbarFX.GradientGetCapsFont
            
        CaptionbarFX.RedrawBar Form1
        
        ElseIf Animation = 3 Then
        
        If K = 0 Then
            Timer2.Enabled = False
            Timer1.Enabled = True
        End If
        
        CaptionbarFX.GradForceColors = True
        CaptionbarFX.GradVerticalGradient = False
        CaptionbarFX.AddBitmap = False
            
        CaptionbarFX.GradForcedText = vbWhite
        CaptionbarFX.GradForcedFirst = K - vbGreen
        CaptionbarFX.GradForcedSecond = 210 - K - vbBlue
        
        CaptionbarFX.GradForcedTextA = &HC0C0C0
        CaptionbarFX.GradForcedFirstA = vbBlack
        CaptionbarFX.GradForcedSecondA = vbBlue
        
        CaptionbarFX.GradientGetCapsFont
            
        CaptionbarFX.RedrawBar Form1
    End If
End Sub
'--------------------------End Animation timers---------------------

Public Sub StopAnimation()
    Timer1.Enabled = False
    Timer2.Enabled = False
    Timer3.Enabled = False
    K = 0
End Sub

Private Sub Timer3_Timer()
    If K = 3 Then K = -1
    K = K + 1

    Set WithForm = Form1
    Set MenuFrm = Form2
  
    Picture6(K).ScaleMode = Form1.ScaleMode
    Picture6(K).Left = 0
    Picture6(K).Width = Screen.Width + 120
    Picture6(K).ScaleMode = 3
    Picture6(K).Cls
    
    CaptionbarFX.GradForceColors = True
    CaptionbarFX.GradVerticalGradient = False
    CaptionbarFX.AddBitmap = True
    CaptionbarFX.BitmapDC = Picture6(K).hdc
    CaptionbarFX.BitmapH = Picture6(K).ScaleHeight
    CaptionbarFX.BitmapW = Picture6(K).ScaleWidth
    
    CaptionbarFX.GradForcedText = vbWhite
    CaptionbarFX.GradForcedFirst = vbRed
    CaptionbarFX.GradForcedSecond = vbRed
   
    CaptionbarFX.GradForcedTextA = &HC0C0C0
    CaptionbarFX.GradForcedFirstA = vbBlack
    CaptionbarFX.GradForcedSecondA = vbBlue

    CaptionbarFX.GradientGetCapsFont
        
    CaptionbarFX.RedrawBar Me
End Sub
