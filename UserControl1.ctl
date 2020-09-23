VERSION 5.00
Begin VB.UserControl Scroller 
   ClientHeight    =   1080
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3960
   ScaleHeight     =   72
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   264
   ToolboxBitmap   =   "UserControl1.ctx":0000
   Begin VB.Timer STimer 
      Enabled         =   0   'False
      Index           =   0
      Interval        =   1
      Left            =   120
      Top             =   1080
   End
   Begin VB.PictureBox BackB 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   249
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   600
      Width           =   3735
   End
   Begin VB.PictureBox Draw 
      Appearance      =   0  'Flat
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   120
      ScaleHeight     =   25
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   249
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   3735
   End
End
Attribute VB_Name = "Scroller"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

Private tScrollText As String
Private tScrollSpeed As Integer
Private tScrollPause As Long
Private tOverride As Boolean
Private Direction As Integer '0=left, 1=up

Private Scrolling As Boolean
Private CX As Integer
Private CY As Integer
Private METext As String
Private Terminate As Boolean
Private ContinuousScroll As Boolean
Private StopTop As Boolean
Public Property Let Font(Setting As String)
    BackB.Font = Setting
End Property
Public Property Let FontSize(Setting As String)
    BackB.FontSize = Setting
End Property
Public Property Let FontBold(Setting As Boolean)
    BackB.FontBold = Setting
End Property
Public Property Let FontItalic(Setting As Boolean)
    BackB.FontItalic = Setting
End Property
Public Property Let FontStrikethru(Setting As Boolean)
    BackB.FontStrikethru = Setting
End Property
Public Property Let FontUnderline(Setting As Boolean)
    BackB.FontUnderline = Setting
End Property
Public Property Let BackColor(nBackCol As Long)
    BackB.BackColor = nBackCol
End Property
Public Property Let TextColor(nTxtCol As Long)
    BackB.ForeColor = nTxtCol
End Property
Public Property Let ScrollContinuous(KeepScrolling As Boolean)
    ContinuousScroll = KeepScrolling
End Property
Public Property Let TopStop(StopTop1 As Boolean)
StopTop = StopTop1
Draw.Cls
End Property
Public Property Get DText(nTxt As String)
   nTxt = METext
End Property
Sub PS(Interval As Long) 'Pauses for specified interval
    Start = GetTickCount
        Do While Start + Interval > GetTickCount
            DoEvents
        Loop
End Sub
Public Sub StopCenter()
Terminate = True
For i = 0 To 20
    If STimer(i).Enabled = True Then
        STimer(i).Enabled = False
    End If
Next
BackB.Cls
BackB.CurrentX = ScaleWidth / 2 - (Len(tScrollText) * 6) / 2
BackB.Print tScrollText
BitBlt Draw.hDC, 0, 0, BackB.ScaleWidth, BackB.ScaleHeight, BackB.hDC, 0, 0, SRCCOPY
End Sub
Public Sub DoScroll(METext2 As String, ScrollSpeed As Integer, ScrollPause As Long, Override As Boolean, Optional ScrollDir As Integer)
Dim METext As String
If METext2 = "" Then Exit Sub
Terminate = False
tScrollText = ""
CY = 1
CX = 1
METext = METext2


    
    If ScrollDir = 1 Then 'check scroll direction
        Direction = 1
    Else
        Direction = 0
    End If
    If Not Scrolling Then
        tScrollText = METext
        tScrollSpeed = ScrollSpeed
        tScrollPause = ScrollPause
        tOverride = Override
    ElseIf Scrolling And Override Then
        tScrollText = METext
        tScrollSpeed = ScrollSpeed
        tScrollPause = ScrollPause
        tOverride = Override
    Else
        Exit Sub
    End If
    For i = 0 To 20
        If STimer(i).Enabled = False Then
            STimer(i).Enabled = True
            Exit Sub
        End If
    Next

End Sub
Private Sub STimer_Timer(Index As Integer)
    If Terminate = True Then Exit Sub
    If Direction = 1 Then
        Call MainLoopUp
    Else
        Call MainLoop
    End If
    If ContinuousScroll = False Then STimer(Index).Enabled = False
    If StopTop = True Then STimer(Index).Enabled = False
End Sub
Private Sub UserControl_Initialize()
    BackB.FontSize = 8
    BackB.Font = "Fixedsys"
    BackB.CurrentX = 0
    BackB.CurrentY = 0
    BackB.BackColor = RGB(0, 0, 0)
    BackB.ForeColor = RGB(240, 240, 240)
    For i = 1 To 20
        Load STimer(i)
        STimer(i).Enabled = False
        STimer(i).Interval = 1
    Next
End Sub
Private Sub MainLoopUp()
    CX = ScaleWidth / 2 - (Len(tScrollText) * 8) / 2
    CY = Draw.ScaleHeight
    Scrolling = True
    Do While CY > 1
        If Terminate = True Then Exit Sub
        BackB.Cls
        BackB.CurrentX = CX
        BackB.CurrentY = CY
        BackB.Print tScrollText
        DoEvents
        BitBlt Draw.hDC, 0, 0, BackB.ScaleWidth, BackB.ScaleHeight, BackB.hDC, 0, 0, SRCCOPY
        CY = CY - tScrollSpeed
        PS 20
    Loop
    If StopTop = True Then Exit Sub
    PS tScrollPause
    
    Do While CY > -16
        If Terminate = True Then Exit Sub
        BackB.Cls
        BackB.CurrentX = CX
        BackB.CurrentY = CY
        BackB.Print tScrollText
        DoEvents
        BitBlt Draw.hDC, 0, 0, BackB.ScaleWidth, BackB.ScaleHeight, BackB.hDC, 0, 0, SRCCOPY
        CY = CY - tScrollSpeed
        PS 20
    Loop
    Scrolling = False
End Sub

Private Sub MainLoop()
    
    CX = Draw.ScaleWidth
    CY = 0

    Scrolling = True
    Do While CX > 1
        If Terminate = True Then Exit Sub
        If StopTop = True Then
        BackB.Cls
        Exit Sub
        End If
        BackB.Cls
        BackB.CurrentX = CX
        BackB.CurrentY = CY
        BackB.Print tScrollText
        DoEvents
        BitBlt Draw.hDC, 0, 0, BackB.ScaleWidth, BackB.ScaleHeight, BackB.hDC, 0, 0, SRCCOPY
        CX = CX - tScrollSpeed
        PS 20
    Loop
    
    PS tScrollPause
    
    Do While CX > 0 - (Len(tScrollText) * 8)
        If Terminate = True Then Exit Sub
        If StopTop = True Then
        BackB.Cls
        Exit Sub
        End If
        BackB.Cls
        BackB.CurrentX = CX
        BackB.CurrentY = CY
        BackB.Print tScrollText
        DoEvents
        BitBlt Draw.hDC, 0, 0, BackB.ScaleWidth, BackB.ScaleHeight, BackB.hDC, 0, 0, SRCCOPY
        CX = CX - tScrollSpeed
        PS 20
    Loop

    Scrolling = False
End Sub
Private Sub UserControl_Resize()
    Draw.Left = 0
    Draw.Top = 0
    Draw.Width = ScaleWidth
    Draw.Height = ScaleHeight
    BackB.Top = ScaleHeight + 1
    BackB.Width = ScaleWidth
    BackB.Height = ScaleHeight
End Sub

