VERSION 5.00
Begin VB.Form Slider 
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Slider 2000 v2.1"
   ClientHeight    =   6585
   ClientLeft      =   705
   ClientTop       =   420
   ClientWidth     =   7350
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   7350
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   5820
      Top             =   4200
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   2565
      Left            =   5430
      TabIndex        =   23
      Top             =   465
      Width           =   1410
   End
   Begin VB.PictureBox SwapGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Left            =   5070
      ScaleHeight     =   1200
      ScaleWidth      =   1200
      TabIndex        =   22
      Top             =   5085
      Width           =   1200
   End
   Begin VB.CommandButton ShuffleBtn 
      Caption         =   "Shuffle"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   5430
      TabIndex        =   21
      Top             =   3600
      Width           =   1380
   End
   Begin VB.PictureBox PicGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   19
      Left            =   3870
      ScaleHeight     =   1200
      ScaleWidth      =   1200
      TabIndex        =   20
      Top             =   5085
      Width           =   1200
   End
   Begin VB.PictureBox PicGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   18
      Left            =   2670
      ScaleHeight     =   1200
      ScaleWidth      =   1200
      TabIndex        =   19
      Top             =   5085
      Width           =   1200
   End
   Begin VB.PictureBox PicGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   17
      Left            =   1470
      ScaleHeight     =   1200
      ScaleWidth      =   1200
      TabIndex        =   18
      Top             =   5085
      Width           =   1200
   End
   Begin VB.PictureBox PicGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   16
      Left            =   270
      ScaleHeight     =   1200
      ScaleWidth      =   1200
      TabIndex        =   17
      Top             =   5085
      Width           =   1200
   End
   Begin VB.PictureBox PicGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   15
      Left            =   3870
      ScaleHeight     =   1200
      ScaleWidth      =   1200
      TabIndex        =   16
      Top             =   3885
      Width           =   1200
   End
   Begin VB.PictureBox PicGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   14
      Left            =   2670
      ScaleHeight     =   1200
      ScaleWidth      =   1200
      TabIndex        =   15
      Top             =   3885
      Width           =   1200
   End
   Begin VB.PictureBox PicGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   13
      Left            =   1470
      ScaleHeight     =   1200
      ScaleWidth      =   1200
      TabIndex        =   14
      Top             =   3885
      Width           =   1200
   End
   Begin VB.PictureBox PicGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   12
      Left            =   270
      ScaleHeight     =   1200
      ScaleWidth      =   1200
      TabIndex        =   13
      Top             =   3885
      Width           =   1200
   End
   Begin VB.PictureBox PicGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   1
      Left            =   1470
      ScaleHeight     =   1200
      ScaleWidth      =   1200
      TabIndex        =   1
      Top             =   285
      Width           =   1200
   End
   Begin VB.PictureBox PicGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   0
      Left            =   270
      ScaleHeight     =   1200
      ScaleWidth      =   1200
      TabIndex        =   0
      Top             =   285
      Width           =   1200
   End
   Begin VB.PictureBox PicGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   11
      Left            =   3870
      ScaleHeight     =   1200
      ScaleWidth      =   1200
      TabIndex        =   12
      Top             =   2685
      Width           =   1200
   End
   Begin VB.PictureBox PicGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   10
      Left            =   2670
      ScaleHeight     =   1200
      ScaleWidth      =   1200
      TabIndex        =   11
      Top             =   2685
      Width           =   1200
   End
   Begin VB.PictureBox PicGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   9
      Left            =   1470
      ScaleHeight     =   1200
      ScaleWidth      =   1200
      TabIndex        =   10
      Top             =   2685
      Width           =   1200
   End
   Begin VB.PictureBox PicGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   8
      Left            =   270
      ScaleHeight     =   1200
      ScaleWidth      =   1200
      TabIndex        =   9
      Top             =   2685
      Width           =   1200
   End
   Begin VB.PictureBox PicGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   7
      Left            =   3870
      ScaleHeight     =   1200
      ScaleWidth      =   1200
      TabIndex        =   8
      Top             =   1485
      Width           =   1200
   End
   Begin VB.PictureBox PicGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   6
      Left            =   2670
      ScaleHeight     =   1200
      ScaleWidth      =   1200
      TabIndex        =   7
      Top             =   1485
      Width           =   1200
   End
   Begin VB.PictureBox PicGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   5
      Left            =   1470
      ScaleHeight     =   1200
      ScaleWidth      =   1200
      TabIndex        =   6
      Top             =   1485
      Width           =   1200
   End
   Begin VB.PictureBox PicGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   4
      Left            =   270
      ScaleHeight     =   1200
      ScaleWidth      =   1200
      TabIndex        =   5
      Top             =   1485
      Width           =   1200
   End
   Begin VB.PictureBox PicGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   3
      Left            =   3870
      ScaleHeight     =   1200
      ScaleWidth      =   1200
      TabIndex        =   4
      Top             =   285
      Width           =   1200
   End
   Begin VB.PictureBox PicGrid 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1200
      Index           =   2
      Left            =   2670
      ScaleHeight     =   1200
      ScaleWidth      =   1200
      TabIndex        =   3
      Top             =   285
      Width           =   1200
   End
   Begin VB.PictureBox PuzzleImage 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   615
      Left            =   2025
      ScaleHeight     =   555
      ScaleWidth      =   2370
      TabIndex        =   2
      Top             =   105
      Visible         =   0   'False
      Width           =   2430
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Caption         =   "Puzzles"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3090
      Left            =   5265
      TabIndex        =   24
      Top             =   150
      Width           =   1725
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00000000&
      FillStyle       =   0  'Solid
      Height          =   1230
      Left            =   5055
      Top             =   5070
      Width           =   1230
   End
   Begin VB.Shape Shape1 
      FillStyle       =   0  'Solid
      Height          =   6030
      Left            =   255
      Top             =   270
      Width           =   4830
   End
   Begin VB.Label TimeLbl 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   " - - : - - "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   5355
      TabIndex        =   25
      Top             =   4680
      Width           =   675
   End
End
Attribute VB_Name = "Slider"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'# -------------------------------------------- #
'# Slider Game Example                          #
'# -------------------------------------------- #
'# email: natburrow@yahoo.com                   #
'# -------------------------------------------- #

Public MovedPeice As Boolean
Public Shuffling As Boolean
Public Complete As Boolean
Public Speed As Integer
Public Countdown As Integer

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

If ShuffleBtn.FontBold = True Then ShuffleBtn.FontBold = False

End Sub

Private Sub File1_Click()

PuzzleImage.Cls
PuzzleImage = LoadPicture(App.Path + "\" + File1.List(File1.ListIndex))

Call InitGrid

End Sub

Private Sub Form_Load()

Randomize Timer

File1.Path = App.Path
File1.Pattern = "*.jpg"

If File1.ListCount = 0 Then
    MsgBox "No valid images were found!"
End

Else

RNDIMG = Int(Rnd * File1.ListCount)
PuzzleImage = LoadPicture(App.Path + "\" + File1.List(RNDIMG))

File1.ListIndex = RNDIMG

Call InitGrid

End If

Speed = 10
Timer1.Enabled = False

'DetectSnd

End Sub


Public Sub PicGrid_Click(Index As Integer)



MovedPeice = False


Call MoveUp(Index)
If MovedPeice = True Then Exit Sub

Call MoveDown(Index)
If MovedPeice = True Then Exit Sub

Call MoveLeft(Index)
If MovedPeice = True Then Exit Sub

Call MoveRight(Index)
If MovedPeice = True Then Exit Sub

If Shuffling = False Then
    SoundFile$ = App.Path + "\nomove.wav"
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    X% = sndPlaySound(SoundFile$, wFlags%)
End If

End Sub

Public Sub MoveUp(i As Integer)

Dim PicGridX As Long
Dim PicGridY As Long
Dim SwpGridX As Long
Dim SwpGridY As Long


PicGridX = PicGrid(i).Top
PicGridY = PicGrid(i).Left
SwpGridX = SwapGrid.Top + PicGrid(i).Height
SwpGridY = SwapGrid.Left '+ PicGrid(i).Width



If PicGridX = SwpGridX And PicGridY = SwpGridY Then  ' It`s ok to move this peice up
    
swpx = SwapGrid.Top
swpy = SwapGrid.Left
picx = PicGrid(i).Top
picy = PicGrid(i).Left

SwapGrid.Visible = False

SwapGrid.Top = 0
SwapGrid.Left = 0


For d = picx To swpx Step -Speed

    PicGrid(i).Top = d
DoEvents

Next d

If Shuffling = False Then
    SoundFile$ = App.Path + "\move.wav"
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    X% = sndPlaySound(SoundFile$, wFlags%)

End If

SwapGrid.Top = picx
SwapGrid.Left = picy
SwapGrid.Visible = True
PicGrid(i).Top = swpx
PicGrid(i).Left = swpy
    
    
    
        MovedPeice = True
    Else


End If

End Sub

Public Sub MoveDown(i As Integer)

Dim PicGridX As Long
Dim PicGridY As Long
Dim SwpGridX As Long
Dim SwpGridY As Long


PicGridX = PicGrid(i).Top + PicGrid(i).Height
PicGridY = PicGrid(i).Left
SwpGridX = SwapGrid.Top
SwpGridY = SwapGrid.Left

If PicGridX = SwpGridX And PicGridY = SwpGridY Then  ' It`s ok to move this peice up
    
swpx = SwapGrid.Top
swpy = SwapGrid.Left
picx = PicGrid(i).Top
picy = PicGrid(i).Left


SwapGrid.Visible = False

SwapGrid.Top = 0
SwapGrid.Left = 0


For d = picx To swpx Step Speed

    PicGrid(i).Top = d
DoEvents

Next d

If Shuffling = False Then
    SoundFile$ = App.Path + "\move.wav"
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    X% = sndPlaySound(SoundFile$, wFlags%)
End If

SwapGrid.Top = picx
SwapGrid.Left = picy
SwapGrid.Visible = True


SwapGrid.Top = picx
SwapGrid.Left = picy
PicGrid(i).Top = swpx
PicGrid(i).Left = swpy
    
    
    MovedPeice = True
    
    Else


End If


End Sub

Public Sub MoveLeft(i As Integer)

Dim PicGridX As Long
Dim PicGridY As Long
Dim SwpGridX As Long
Dim SwpGridY As Long


PicGridX = PicGrid(i).Top
PicGridY = PicGrid(i).Left
SwpGridX = SwapGrid.Top
SwpGridY = SwapGrid.Left + PicGrid(i).Width



If PicGridY = SwpGridY And PicGridX = SwpGridX Then  ' It`s ok to move this peice Left
    
swpx = SwapGrid.Top
swpy = SwapGrid.Left
picx = PicGrid(i).Top
picy = PicGrid(i).Left


SwapGrid.Visible = False

SwapGrid.Top = 0
SwapGrid.Left = 0


For d = picy To swpy Step -Speed

    PicGrid(i).Left = d
DoEvents

Next d

If Shuffling = False Then
    SoundFile$ = App.Path + "\move.wav"
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    X% = sndPlaySound(SoundFile$, wFlags%)
End If

SwapGrid.Top = picx
SwapGrid.Left = picy
SwapGrid.Visible = True


SwapGrid.Top = picx
SwapGrid.Left = picy
PicGrid(i).Top = swpx
PicGrid(i).Left = swpy
    
    
        MovedPeice = True
    Else


End If




End Sub


Public Sub MoveRight(i As Integer)

Dim PicGridX As Long
Dim PicGridY As Long
Dim SwpGridX As Long
Dim SwpGridY As Long


PicGridX = PicGrid(i).Top
PicGridY = PicGrid(i).Left + PicGrid(i).Width
SwpGridX = SwapGrid.Top '+ PicGrid(i).Height
SwpGridY = SwapGrid.Left

If PicGridY = SwpGridY And PicGridX = SwpGridX Then  ' It`s ok to move this peice up

    
swpx = SwapGrid.Top
swpy = SwapGrid.Left
picx = PicGrid(i).Top
picy = PicGrid(i).Left


SwapGrid.Visible = False

SwapGrid.Top = 0
SwapGrid.Left = 0


For d = picy To swpy Step Speed

    PicGrid(i).Left = d
DoEvents

Next d

If Shuffling = False Then
    SoundFile$ = App.Path + "\move.wav"
    wFlags% = SND_ASYNC Or SND_NODEFAULT
    X% = sndPlaySound(SoundFile$, wFlags%)
End If

SwapGrid.Top = picx
SwapGrid.Left = picy
SwapGrid.Visible = True


SwapGrid.Top = picx
SwapGrid.Left = picy
PicGrid(i).Top = swpx
PicGrid(i).Left = swpy
    
    
        MovedPeice = True
    Else


End If


End Sub

Sub InitGrid()

Call BitBlt(PicGrid(0).hDC, 0, 0, 1200, 1200, PuzzleImage.hDC, 0, 0, SRCCOPY)
    PicGrid(0).Refresh
 Call BitBlt(PicGrid(1).hDC, 0, 0, 1200, 1200, PuzzleImage.hDC, 80, 0, SRCCOPY)
    PicGrid(1).Refresh
Call BitBlt(PicGrid(2).hDC, 0, 0, 1200, 1200, PuzzleImage.hDC, 160, 0, SRCCOPY)
    PicGrid(2).Refresh
Call BitBlt(PicGrid(3).hDC, 0, 0, 1200, 1200, PuzzleImage.hDC, 240, 0, SRCCOPY)
    PicGrid(3).Refresh

Call BitBlt(PicGrid(4).hDC, 0, 0, 1200, 1200, PuzzleImage.hDC, 0, 80, SRCCOPY)
    PicGrid(4).Refresh
Call BitBlt(PicGrid(5).hDC, 0, 0, 1200, 1200, PuzzleImage.hDC, 80, 80, SRCCOPY)
    PicGrid(5).Refresh
Call BitBlt(PicGrid(6).hDC, 0, 0, 1200, 1200, PuzzleImage.hDC, 160, 80, SRCCOPY)
    PicGrid(6).Refresh
Call BitBlt(PicGrid(7).hDC, 0, 0, 1200, 1200, PuzzleImage.hDC, 240, 80, SRCCOPY)
    PicGrid(7).Refresh

Call BitBlt(PicGrid(8).hDC, 0, 0, 1200, 1200, PuzzleImage.hDC, 0, 160, SRCCOPY)
    PicGrid(8).Refresh
Call BitBlt(PicGrid(9).hDC, 0, 0, 1200, 1200, PuzzleImage.hDC, 80, 160, SRCCOPY)
    PicGrid(9).Refresh
Call BitBlt(PicGrid(10).hDC, 0, 0, 1200, 1200, PuzzleImage.hDC, 160, 160, SRCCOPY)
    PicGrid(10).Refresh
Call BitBlt(PicGrid(11).hDC, 0, 0, 1200, 1200, PuzzleImage.hDC, 240, 160, SRCCOPY)
    PicGrid(11).Refresh

Call BitBlt(PicGrid(12).hDC, 0, 0, 1200, 1200, PuzzleImage.hDC, 0, 240, SRCCOPY)
    PicGrid(12).Refresh
Call BitBlt(PicGrid(13).hDC, 0, 0, 1200, 1200, PuzzleImage.hDC, 80, 240, SRCCOPY)
    PicGrid(13).Refresh
Call BitBlt(PicGrid(14).hDC, 0, 0, 1200, 1200, PuzzleImage.hDC, 160, 240, SRCCOPY)
    PicGrid(14).Refresh
Call BitBlt(PicGrid(15).hDC, 0, 0, 1200, 1200, PuzzleImage.hDC, 240, 240, SRCCOPY)
    PicGrid(15).Refresh

Call BitBlt(PicGrid(16).hDC, 0, 0, 1200, 1200, PuzzleImage.hDC, 0, 320, SRCCOPY)
    PicGrid(16).Refresh
Call BitBlt(PicGrid(17).hDC, 0, 0, 1200, 1200, PuzzleImage.hDC, 80, 320, SRCCOPY)
    PicGrid(17).Refresh
Call BitBlt(PicGrid(18).hDC, 0, 0, 1200, 1200, PuzzleImage.hDC, 160, 320, SRCCOPY)
    PicGrid(18).Refresh
Call BitBlt(PicGrid(19).hDC, 0, 0, 1200, 1200, PuzzleImage.hDC, 240, 320, SRCCOPY)
    PicGrid(19).Refresh


End Sub

Private Sub Picture1_Click()

End Sub



Private Sub ShuffleBtn_Click()

Speed = 50

Randomize Timer
X = Int(Rnd * 35) + 35

shuffle = 0
Shuffling = True

Do
shuffle = shuffle + 1
For peice = 0 To 19
    PicGrid_Click (peice)
Next peice
DoEvents
Loop Until shuffle = X
Shuffling = False

Speed = 10
Countdown = 300
TimeLbl = "5:00"
Timer1.Enabled = True

End Sub

Private Sub ShuffleBtn_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

ShuffleBtn.FontBold = True

End Sub

Private Sub Timer1_Timer()


Countdown = Countdown - 1

Dim Minute As Integer
Dim Second As String

Select Case Countdown
    Case Is < 60
        Minute = 0
    Case Is < 120
        Minute = 1
    Case Is < 180
        Minute = 2
    Case Is < 240
        Minute = 3
    Case Is < 300
        Minute = 4
    Case Is < 360
        Minute = 5

End Select
    


Second = ""
Second = Str(Format(Countdown Mod 60))
    
Select Case Val(Second)

Case Is < 10
    Second = "0" & Right(Str(Second), 1)
Case Else
    Second = Right(Str(Second), 2)
End Select

    
    
    TimeLbl.Caption = Minute & ":" + Second

If Countdown = -1 Then
  
  Timer1.Enabled = False
  SoundFile$ = App.Path + "\notime.wav"
  wFlags% = SND_ASYNC Or SND_NODEFAULT
  X% = sndPlaySound(SoundFile$, wFlags%)

TimeLbl.Caption = "- - : - -"
End If



End Sub
