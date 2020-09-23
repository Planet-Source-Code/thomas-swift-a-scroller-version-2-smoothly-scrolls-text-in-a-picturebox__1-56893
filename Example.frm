VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Scroller Example"
   ClientHeight    =   3480
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7095
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3480
   ScaleWidth      =   7095
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check2 
      Caption         =   "Stop At Top Vertical"
      Height          =   195
      Left            =   2670
      TabIndex        =   7
      Top             =   3150
      Width           =   1755
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Scroll Horizontal Continuously"
      Height          =   210
      Left            =   1545
      TabIndex        =   6
      Top             =   2430
      Value           =   1  'Checked
      Width           =   2460
   End
   Begin Project1.Scroller Scroller1 
      Height          =   1800
      Left            =   30
      TabIndex        =   5
      Top             =   480
      Width           =   7020
      _ExtentX        =   12383
      _ExtentY        =   3175
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Stop Centered"
      Height          =   255
      Left            =   4185
      TabIndex        =   4
      Top             =   2408
      Width           =   1365
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Upwards Scroll 2"
      Height          =   255
      Left            =   3600
      TabIndex        =   3
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Upwards Scroll 1"
      Height          =   255
      Left            =   1680
      TabIndex        =   2
      Top             =   2760
      Width           =   1815
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H0000FF00&
      Caption         =   "Scroll 2"
      Height          =   255
      Left            =   3600
      TabIndex        =   1
      Top             =   120
      Width           =   1815
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Scroll 1"
      Height          =   255
      Left            =   1680
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Check1_Click()
If Check1.Value = 1 Then
    Scroller1.ScrollContinuous = True
    Else
    Scroller1.ScrollContinuous = False
    End If
End Sub

Private Sub Check2_Click()
If Check2.Value = 1 Then
    Scroller1.TopStop = True
    Else
    Scroller1.TopStop = False
    End If
End Sub

'Format for using the ocx:
'scroller1.DoScroll <text to scroll>, <speed (pixels/sec)>, <pause (ms)>, <overirde - if text is already scrolling should it stop it?>, <put a 1 here if you want an upwards scroll>

Private Sub Command1_Click()
    
    Scroller1.DoScroll "Scroll Type 1 [override]", 2, 0, True
End Sub

Private Sub Command2_Click()
    Scroller1.DoScroll "Scroll Type 2 [override]", 2, 1000, True
End Sub

Private Sub Command3_Click()
    Scroller1.DoScroll "Upwards Scroll", 2, 0, True, 1
End Sub

Private Sub Command4_Click()
    Scroller1.DoScroll "Upwards Scroll [override]", 2, 1000, True, 1
End Sub

Private Sub Command5_Click()
Scroller1.StopCenter
End Sub

Private Sub Form_Load()
    Scroller1.BackColor = RGB(0, 0, 0)
    Scroller1.TextColor = &HFF00&    'RGB(100, 180, 100)
    
    Scroller1.Font = "Arial"
    Scroller1.FontBold = True
    'Scroller1.FontItalic = True
    Scroller1.FontSize = 9
    'Scroller1.FontStrikethru = True
    'Scroller1.FontUnderline = True
    
    Scroller1.ScrollContinuous = True
    Scroller1.DoScroll "Text Scroller 2.0 updated by Thomas Swift", 1.5, 0, False
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

    End
End Sub
