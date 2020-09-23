VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.Menu MnuPop 
      Caption         =   "PopupMen"
      Begin VB.Menu Mnuresize 
         Caption         =   "Restore"
      End
      Begin VB.Menu MnuMaximize 
         Caption         =   "Maximize"
      End
      Begin VB.Menu MnuMinimize 
         Caption         =   "Minimize"
      End
      Begin VB.Menu MnuClose 
         Caption         =   "Close"
      End
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Menu controls

Private Sub MnuClose_Click()
    Unload Form1
End Sub

Private Sub MnuMaximize_Click()
    Form1.WindowState = vbMaximized
End Sub

Private Sub MnuMinimize_Click()
    Form1.WindowState = vbMinimized
End Sub

Private Sub Mnuresize_Click()
    Form1.WindowState = vbNormal
End Sub
