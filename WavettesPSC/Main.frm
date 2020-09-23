VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "avettes ~"
   ClientHeight    =   6600
   ClientLeft      =   150
   ClientTop       =   150
   ClientWidth     =   10680
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   440
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   712
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton optBit 
      Caption         =   "16 bit"
      Height          =   255
      Index           =   1
      Left            =   3135
      TabIndex        =   42
      Top             =   60
      Value           =   -1  'True
      Width           =   825
   End
   Begin VB.OptionButton optBit 
      Caption         =   "8 bit"
      Height          =   255
      Index           =   0
      Left            =   2475
      TabIndex        =   41
      Top             =   60
      Width           =   660
   End
   Begin Project1.Container fraMisc 
      Height          =   585
      Left            =   3090
      TabIndex        =   26
      Top             =   5715
      Width           =   3525
      _ExtentX        =   6218
      _ExtentY        =   1032
      BackColor       =   -2147483633
      BorderColorDark =   4210752
      Caption         =   "Miscellaneous"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Picture         =   "Main.frx":0442
      Begin VB.CheckBox chkReverse 
         Caption         =   "Reverse"
         Height          =   255
         Left            =   2445
         TabIndex        =   29
         Top             =   210
         Width           =   930
      End
      Begin VB.CheckBox chkABS 
         Caption         =   "Absolute"
         Height          =   270
         Left            =   1440
         TabIndex        =   28
         Top             =   210
         Width           =   930
      End
      Begin VB.CheckBox chkShape 
         Caption         =   "Basic shape"
         Height          =   225
         Left            =   150
         TabIndex        =   27
         Top             =   225
         Width           =   1215
      End
   End
   Begin Project1.Container fraRepeat 
      Height          =   1140
      Left            =   4320
      TabIndex        =   22
      Top             =   4500
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   2011
      BackColor       =   -2147483633
      BorderColorDark =   4210752
      Caption         =   "Repeat"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Picture         =   "Main.frx":045E
      Begin VB.CheckBox chkRepeat 
         Caption         =   "Repeat"
         Height          =   210
         Left            =   150
         TabIndex        =   25
         Top             =   825
         Width           =   945
      End
      Begin VB.HScrollBar HSRepeat 
         Height          =   210
         Left            =   105
         Max             =   32
         Min             =   2
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   555
         Value           =   2
         Width           =   1065
      End
      Begin VB.Label LabRepeat 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LabReat"
         Height          =   240
         Left            =   375
         TabIndex        =   24
         Top             =   240
         Width           =   450
      End
   End
   Begin VB.CommandButton cmdAdd2Presets 
      Caption         =   "Add to Presets"
      Height          =   300
      Left            =   9000
      TabIndex        =   20
      Top             =   30
      Width           =   1260
   End
   Begin VB.CommandButton cmdPresets 
      Caption         =   "Presets"
      Height          =   285
      Left            =   7920
      TabIndex        =   19
      Top             =   30
      Width           =   930
   End
   Begin Project1.Container fraPlay 
      Height          =   2640
      Left            =   5715
      TabIndex        =   13
      Top             =   3000
      Width           =   1140
      _ExtentX        =   2011
      _ExtentY        =   4657
      BackColor       =   -2147483633
      BorderColorDark =   4210752
      Caption         =   "Play"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Picture         =   "Main.frx":047A
      Begin VB.CommandButton cmdPlay 
         Caption         =   "Play"
         Height          =   540
         Left            =   210
         TabIndex        =   43
         Top             =   390
         Width           =   705
      End
      Begin VB.CommandButton cmdRepStop 
         Caption         =   "Loop Play"
         Height          =   555
         Index           =   0
         Left            =   210
         TabIndex        =   15
         Top             =   1140
         Width           =   705
      End
      Begin VB.CommandButton cmdRepStop 
         Caption         =   "STOP"
         Height          =   555
         Index           =   1
         Left            =   210
         TabIndex        =   14
         Top             =   1905
         Width           =   705
      End
      Begin VB.Shape ShapeCmdRepStop 
         BorderColor     =   &H00404040&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H8000000F&
         FillStyle       =   0  'Solid
         Height          =   720
         Index           =   1
         Left            =   120
         Top             =   1830
         Width           =   870
      End
      Begin VB.Shape ShapeCmdRepStop 
         BorderColor     =   &H00404040&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H8000000F&
         FillStyle       =   0  'Solid
         Height          =   750
         Index           =   0
         Left            =   135
         Top             =   1035
         Width           =   870
      End
   End
   Begin Project1.Container fraFormulae 
      Height          =   3300
      Left            =   6960
      TabIndex        =   6
      Top             =   3000
      Width           =   3330
      _ExtentX        =   5874
      _ExtentY        =   5821
      BackColor       =   -2147483633
      BorderColorDark =   4210752
      Caption         =   "Formulae"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Picture         =   "Main.frx":0496
      Begin VB.ListBox List1 
         Height          =   2790
         Left            =   150
         TabIndex        =   7
         Top             =   300
         Width           =   2985
      End
   End
   Begin Project1.Container fraRamps 
      Height          =   2640
      Left            =   2835
      TabIndex        =   5
      Top             =   3000
      Width           =   1380
      _ExtentX        =   2434
      _ExtentY        =   4657
      BackColor       =   -2147483633
      BorderColorDark =   4210752
      Caption         =   "Ramps"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Picture         =   "Main.frx":04B2
      Begin VB.CommandButton cmdRamp 
         Caption         =   "None"
         Height          =   330
         Left            =   240
         TabIndex        =   12
         Top             =   2100
         Width           =   870
      End
      Begin VB.PictureBox picRamp 
         AutoRedraw      =   -1  'True
         Height          =   345
         Index           =   1
         Left            =   120
         MousePointer    =   9  'Size W E
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   72
         TabIndex        =   9
         Top             =   1530
         Width           =   1140
      End
      Begin VB.PictureBox picRamp 
         AutoRedraw      =   -1  'True
         Height          =   345
         Index           =   0
         Left            =   120
         MousePointer    =   9  'Size W E
         ScaleHeight     =   19
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   72
         TabIndex        =   8
         Top             =   615
         Width           =   1140
      End
      Begin VB.Label LabFrac 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   1
         Left            =   255
         TabIndex        =   11
         Top             =   1215
         Width           =   780
      End
      Begin VB.Label LabFrac 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   240
         Index           =   0
         Left            =   255
         TabIndex        =   10
         Top             =   300
         Width           =   780
      End
   End
   Begin Project1.Container fraEcho 
      Height          =   1470
      Left            =   4320
      TabIndex        =   4
      Top             =   3000
      Width           =   1290
      _ExtentX        =   2275
      _ExtentY        =   2593
      BackColor       =   -2147483633
      BorderColorDark =   4210752
      Caption         =   "Echoes"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Picture         =   "Main.frx":04CE
      Begin VB.CheckBox chkStagger 
         Caption         =   "+Stagger"
         Height          =   195
         Left            =   165
         TabIndex        =   21
         Top             =   1140
         Width           =   975
      End
      Begin VB.CheckBox chkEcho 
         Caption         =   "Echo"
         Height          =   270
         Left            =   165
         TabIndex        =   18
         Top             =   810
         Width           =   855
      End
      Begin VB.HScrollBar HSEcho 
         Height          =   210
         Left            =   105
         Max             =   32
         Min             =   1
         TabIndex        =   16
         TabStop         =   0   'False
         Top             =   555
         Value           =   1
         Width           =   1065
      End
      Begin VB.Label LabEcho 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LabEcho"
         Height          =   225
         Left            =   390
         TabIndex        =   17
         Top             =   240
         Width           =   450
      End
   End
   Begin VB.PictureBox PIC2 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      Height          =   2385
      Left            =   8190
      ScaleHeight     =   155
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   136
      TabIndex        =   3
      Top             =   510
      Width           =   2100
      Begin VB.Image Image1 
         Height          =   600
         Left            =   135
         Picture         =   "Main.frx":04EA
         Top             =   30
         Width           =   1920
      End
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save WAV (16 bit mono)"
      Height          =   300
      Left            =   375
      TabIndex        =   2
      Top             =   30
      Width           =   2010
   End
   Begin VB.PictureBox PIC 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   2385
      Left            =   360
      ScaleHeight     =   155
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   508
      TabIndex        =   0
      Top             =   510
      Width           =   7680
   End
   Begin Project1.Container fraParams 
      Height          =   3315
      Left            =   360
      TabIndex        =   30
      Top             =   2985
      Width           =   2370
      _ExtentX        =   4180
      _ExtentY        =   5847
      BackColor       =   -2147483633
      BorderColorDark =   4210752
      Caption         =   "Params"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   16711680
      Picture         =   "Main.frx":11F4
      Begin VB.CommandButton cmdMaxAmp 
         Caption         =   "Max"
         Height          =   255
         Left            =   1710
         TabIndex        =   35
         Top             =   330
         Width           =   495
      End
      Begin VB.HScrollBar HS 
         Height          =   210
         Index           =   0
         LargeChange     =   1000
         Left            =   120
         Min             =   1
         SmallChange     =   100
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   690
         Value           =   1
         Width           =   2085
      End
      Begin VB.HScrollBar HS 
         Height          =   210
         Index           =   1
         LargeChange     =   100
         Left            =   120
         Max             =   8000
         Min             =   5
         TabIndex        =   33
         TabStop         =   0   'False
         Top             =   1410
         Value           =   10
         Width           =   2085
      End
      Begin VB.HScrollBar HS 
         Height          =   210
         Index           =   2
         Left            =   135
         Max             =   100
         Min             =   1
         TabIndex        =   32
         TabStop         =   0   'False
         Top             =   2100
         Value           =   10
         Width           =   2085
      End
      Begin VB.HScrollBar HS 
         Height          =   210
         Index           =   3
         Left            =   135
         Max             =   4
         Min             =   1
         TabIndex        =   31
         TabStop         =   0   'False
         Top             =   2790
         Value           =   4
         Width           =   2085
      End
      Begin VB.Label LabTest 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "LabTest"
         ForeColor       =   &H00000000&
         Height          =   270
         Left            =   570
         TabIndex        =   40
         Top             =   2985
         Width           =   1170
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Amplitude"
         Height          =   255
         Index           =   0
         Left            =   165
         TabIndex        =   39
         Top             =   330
         Width           =   1500
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Frequency"
         Height          =   255
         Index           =   1
         Left            =   165
         TabIndex        =   38
         Top             =   1050
         Width           =   1950
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Duration"
         Height          =   255
         Index           =   2
         Left            =   165
         TabIndex        =   37
         Top             =   1755
         Width           =   1950
      End
      Begin VB.Label Lab 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SampleRate"
         Height          =   255
         Index           =   3
         Left            =   195
         TabIndex        =   36
         Top             =   2430
         Width           =   1950
      End
   End
   Begin VB.Label LabF 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BorderStyle     =   1  'Fixed Single
      Caption         =   "LabF"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3975
      TabIndex        =   1
      Top             =   45
      Width           =   3705
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' ~Wavettes~  by  Robert Rayment  May 2006

' Update 31 May VBAccelerator code inserted to prevent
'           'crash on shutdown' on some PCs with XP themes.
' Update  4 May  Compile with no advanced options checked

' Some code adapted from Ulli's prog at PSC CodeId=64845
' Container UC by Eric Madison, PSC CodeId=40130

' Formulae:_
' Strings at  Sub SetFunctions
' Evaluate at Function EvalFunc (Module1.bas)
' If a formula is added or deleted then
' both these routines have to be modified.

' See notes in Pretext.txt for rules on
' editting by hand.

Option Explicit

' For XP manifest
Private Declare Sub InitCommonControls Lib "comctl32" ()
Private Declare Function LoadLibrary Lib "Kernel32" Alias "LoadLibraryA" ( _
    ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "Kernel32" ( _
   ByVal hLibModule As Long) As Long

Private m_hMod As Long

'------- Highlighting
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef _
   lpPoint As POINTAPI) As Long

Private Declare Function WindowFromPoint Lib "user32.dll" ( _
     ByVal xPoint As Long, _
     ByVal yPoint As Long) As Long

Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Pnt As POINTAPI
Private aHiLit As Boolean
'--------

'-------  Shape Controls
Private Declare Function CreateRoundRectRgn Lib "gdi32.dll" _
   (ByVal x1 As Long, ByVal y1 As Long, _
   ByVal x2 As Long, ByVal y2 As Long, _
   ByVal X3 As Long, ByVal Y3 As Long) As Long

' X1,Y1  X2,Y2  Top left & Bottom right coords of rectangle.
' For whole control X1 & Y1 = 0
' X2 & Y2 = Controls width & height
' X3,Y3  width & height of ellipse used to create corners

Private Declare Function SetWindowRgn Lib "user32.dll" _
(ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'-------

Private uAmp As Single, uFrq As Single, uDur As Single
Private aBlock As Boolean
Private aPlay As Boolean
Private aPlayOnce As Boolean
Private aRepeatPlay As Boolean

Private CDL As OSDialog


Private Sub SetFunctions()
'Public Func$()
'Public FuncIndex As Long
Dim k As Long
   ' Could be in an external file
   ReDim Func$(20)
   Func$(1) = "1.  Sin(x)"
   Func$(2) = "2.  Sin(x) + Sin(x * pi / 3)"
   Func$(3) = "3.  (x * pi) * Sin(x)"
   Func$(4) = "4.  (Rnd - Rnd)"
   Func$(5) = "5.  Sin(x) + Sin(x)^2 + Sin(x)^3"
   Func$(6) = "6.  Sin(Int(x / pi)) * pi + pi / 2"
   Func$(7) = "7.  Sin(x)^3"
   Func$(8) = "8.  Sin(x^2)^3"
   Func$(9) = "9.  Sawtooth 1"
   Func$(10) = "10. Squarewave"
   Func$(11) = "11. Sin(3 * x) / Tan(x)"
   Func$(12) = "12. (Sin(3 * x) / Tan(x)) * sin(x / 2)"
   Func$(13) = "13. Exp(Sin(x + pi / 2)) * Sin(x)"
   Func$(14) = "14. Exp(Sin(x^2 + pi / 2)) * Sin(x)"
   Func$(15) = "15. Atn(Cos(x^3) + Sin(x^2))"
   Func$(16) = "16. x^Sin(x)"
   'Func$(17) = "17. -4(Cos(x)+Cos(3*x)/9+Cos(5*x)/25)/pi"
   Func$(17) = "17. Sawtooth 2"
   'Func$(18) = "18. 1/pi#+Sin(x)/2-(Cos(2*x)/1*3 + Cos(4*x)/3*5 + cos(6*x)/5*7 + +)/pi"
   Func$(18) = "18. Bumps"
   Func$(19) = "19. Cosec(x)"
   
   For k = 1 To 19   ' 20
      List1.AddItem Func$(k)
   Next k
   
End Sub


Private Sub Form_Initialize()
   m_hMod = LoadLibrary("shell32.dll")
   InitCommonControls
   SamplesPerSecond = 44100
   Bitnum = 0  ' 8 bit start
   InitHeader
   LoadWav
End Sub


Private Sub Form_Load()
Dim X As Single
   PathSpec$ = App.Path
   If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"
   CurrPath$ = PathSpec$
   ReDim RampFrac(0 To 2)
   
   X = picRamp(0).ScaleWidth / 2
   picRamp(1).Line (0, 0)-(X, picRamp(1).ScaleHeight)
   picRamp(1).Line (X, picRamp(1).ScaleHeight)-(picRamp(1).ScaleWidth, 0)
   RampFrac(1) = 0.5
   LabFrac(1) = RampFrac(1)
   
   SetFunctions
   
   aPlay = False
   aPlayOnce = False
   aBlock = True
   AmpMult = 1
   aRepeat = False
   RepeatMul = 2
   LabRepeat = "2"
   
   aShape = False
   aABS = False
   aReverse = False
   ' Default to 8 bit
   Bitnum = 0
   optBit(Bitnum).Value = True
   optBit_Click 0
   
   LabTest.Visible = False
   
   Show
   
   ' Starter
   picRamp_MouseDown 0, 1, 0, 10, 0
   chkEcho.Value = Checked
   HSEcho.Value = 7
   HS(0).Value = 23000  ' Amplitude 100*23000/32767 = 70%
   HS(1).Value = 300
   HS(2).Value = 20
   HS(3).Value = 4
   List1.ListIndex = 1
   
   ShapeCtrl PIC, 20, 0
   ShapeCtrl PIC2, 20, 0

   aBlock = False
   aPresets = False
   
End Sub

Private Sub ShapeCtrl(p As Control, Rad As Long, SM As Long)
Dim Reg As Long
' Rad, in this case circular corner radius
' SM 0 for Pixels, 1 for Twips
   If SM = 0 Then
      Reg = CreateRoundRectRgn(0, 0, p.Width, p.Height, Rad, Rad)
   Else
      Reg = CreateRoundRectRgn(1, 1, p.Width \ Screen.TwipsPerPixelX, p.Height \ Screen.TwipsPerPixelY, Rad, Rad)
   End If
   SetWindowRgn p.hWnd, Reg, True
   DeleteObject Reg
End Sub

Private Sub chkReverse_Click()
   aReverse = -chkReverse.Value
   If Not aBlock Then EvalPlot PIC, PIC2, uAmp * AmpMult, uFrq, uDur, aPlay
   aPlayOnce = False
End Sub

Private Sub chkABS_Click()
   aABS = -chkABS.Value
   If Not aBlock Then EvalPlot PIC, PIC2, uAmp * AmpMult, uFrq, uDur, aPlay
   aPlayOnce = False
End Sub

Private Sub cmdMaxAmp_Click()
   HS(0).Value = 32767
   aPlayOnce = False
End Sub

Private Sub chkShape_Click()
' With Dummy freq
   aShape = -chkShape.Value
   If Not aBlock Then EvalPlot PIC, PIC2, uAmp * AmpMult, uFrq, uDur, aPlay
End Sub


' Play  ----------------------------------------------------

Private Sub cmdPlay_Click()
   cmdPlay.Enabled = False
   aPlay = True
   aPlayOnce = True
   EvalPlot PIC, PIC2, uAmp * AmpMult, uFrq, uDur, aPlay
   DoEvents
   If Not aRepeatPlay Then
      PlayWAV SoundFile(1), 0, SND_MEMORY Or SND_ASYNC
   Else
      PlayWAV SoundFile(1), 0, SND_MEMORY Or SND_ASYNC Or SND_LOOP
   End If
   cmdPlay.Enabled = True
   aPlay = False
End Sub

Private Sub cmdRepStop_Click(Index As Integer)
' NB Repeat (Rnd - Rnd) chosen to repeat pattern but not random numbers
   If Index = 0 Then
      aRepeatPlay = True
      aHiLit = True
      Call cmdPlay_Click
   Else
      StopPlay
      aRepeatPlay = False
   End If
End Sub

' Formulae  ----------------------------------------------------

Private Sub List1_Click()
' Public FuncIndex
   FuncIndex = List1.ListIndex + 1
   Evaluate
   aPlayOnce = False
End Sub

Private Sub Evaluate()
' Public FuncIndex
' Private uAmp As Single, uFrq As Single, uDur As Single
Dim percent As Integer
   uAmp = HS(0).Value
   uFrq = HS(1).Value
   uDur = CSng(HS(2).Value) / 10
   
   LabF = Func$(FuncIndex)
   percent = uAmp * 100 / 32767
   Lab(0) = "Amplitude =" & Str$(percent) & "%"
   Lab(1) = "Frequency = " & CLng(uFrq) & " Hz"
   Lab(2) = "Duration = " & Str$(uDur) & " s"
   Lab(3) = "SampleRate = " & Str$(SamplesPerSecond) & " /s"
   EvalPlot PIC, PIC2, uAmp * AmpMult, uFrq, uDur, aPlay
End Sub

' Params ----------------------------------------------------

Private Sub HS_Scroll(Index As Integer)
   Call HS_Change(Index)
End Sub

Private Sub HS_Change(Index As Integer)
' Public FuncIndex
' Private uAmp As Single, uFrq As Single, uDur As Single
Dim percent As Integer
   uAmp = HS(0).Value
   uFrq = HS(1).Value
   uDur = CSng(HS(2).Value) / 10
   If Index = 3 Then
      Select Case HS(3).Value
      Case 1: SamplesPerSecond = 5012
      Case 2: SamplesPerSecond = 11025
      Case 3: SamplesPerSecond = 22050
      Case 4: SamplesPerSecond = 44100
      End Select
      Header.SRate = SamplesPerSecond
      If Bitnum = 1 Then ' 16 bit
         Header.Blk = 2
         Header.BRate = SamplesPerSecond * 2
      Else   ' 8 bit
         Header.Blk = 1
         Header.BRate = SamplesPerSecond
      End If
   End If
   
   LabF = Func$(FuncIndex)
   percent = uAmp * 100 / 32767
   Lab(0) = "Amplitude =" & Str$(percent) & "%"
   Lab(1) = "Frequency = " & CLng(uFrq) & " Hz"
   Lab(2) = "Duration = " & Str$(uDur) & " s"
   Lab(3) = "SampleRate = " & Str$(SamplesPerSecond) & " /s"
   
   If Not aBlock Then EvalPlot PIC, PIC2, uAmp * AmpMult, uFrq, uDur, aPlay
   aPlayOnce = False
End Sub

Private Sub optBit_Click(Index As Integer)
   Bitnum = Index
   If Index = 0 Then   ' 8 bit
      cmdSave.Caption = "SaveWAV (8 bit mono)"
      Header.Bits = 8
      Header.Blk = 1
      Header.BRate = SamplesPerSecond
   Else   ' 16 bit
      cmdSave.Caption = "SaveWAV (16 bit mono)"
      Header.Bits = 16
      Header.Blk = 2
      Header.BRate = SamplesPerSecond * 2
   End If
   aPlayOnce = False
End Sub

' Ramps ----------------------------------------------------

Private Sub picRamp_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   Call picRamp_MouseMove(Index, Button, Shift, X, Y)
End Sub

Private Sub picRamp_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim X0 As Single, x1 As Single
   If X < 0 Then X = 0
   If X > picRamp(0).ScaleWidth Then X = picRamp(0).ScaleWidth
   If Button = 1 Then
      If Index = 0 Then ' Up/Down
         Ramp = 0
         picRamp(0).Cls
         picRamp(0).BackColor = 0
         picRamp(0).ForeColor = vbCyan
         picRamp(0).Line (0, picRamp(0).ScaleHeight)-(X, 0)
         picRamp(0).Line (X, 0)-(picRamp(0).ScaleWidth, picRamp(0).ScaleHeight)
         RampFrac(0) = Round(X / picRamp(0).ScaleWidth, 2)
         
         ' X = RampFrac(Ramp) * picRamp(Ramp).ScaleWidth
         
         picRamp(1).BackColor = vbButtonFace
         picRamp(1).ForeColor = 0
         
         x1 = RampFrac(1) * picRamp(1).ScaleWidth
         picRamp(1).Line (0, 0)-(x1, picRamp(1).ScaleHeight)
         picRamp(1).Line (x1, picRamp(1).ScaleHeight)-(picRamp(1).ScaleWidth, 0)
      Else     ' Down/Up
         Ramp = 1
         picRamp(1).Cls
         picRamp(1).BackColor = 0
         picRamp(1).ForeColor = vbCyan
         picRamp(1).Line (0, 0)-(X, picRamp(1).ScaleHeight)
         picRamp(1).Line (X, picRamp(1).ScaleHeight)-(picRamp(1).ScaleWidth, 0)
         RampFrac(1) = Round(X / picRamp(1).ScaleWidth, 2)
         
         picRamp(0).BackColor = vbButtonFace
         picRamp(0).ForeColor = 0
         
         X0 = RampFrac(0) * picRamp(0).ScaleWidth
         picRamp(0).Line (0, picRamp(0).ScaleHeight)-(X0, 0)
         picRamp(0).Line (X0, 0)-(picRamp(0).ScaleWidth, picRamp(0).ScaleHeight)
      End If
   
      LabFrac(0) = RampFrac(0)
      LabFrac(1) = RampFrac(1)
      If Not aBlock Then EvalPlot PIC, PIC2, uAmp * AmpMult, uFrq, uDur, aPlay
      aPlayOnce = False
   End If
End Sub

Private Sub picRamp_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
   EvalPlot PIC, PIC2, uAmp * AmpMult, uFrq, uDur, aPlay
End Sub

Private Sub cmdRamp_Click()
' No Ramp
Dim X0 As Single, x1 As Single
   picRamp(0).BackColor = vbButtonFace
   picRamp(0).ForeColor = 0
   X0 = RampFrac(0) * picRamp(0).ScaleWidth
   picRamp(0).Line (0, picRamp(0).ScaleHeight)-(X0, 0)
   picRamp(0).Line (X0, 0)-(picRamp(0).ScaleWidth, picRamp(0).ScaleHeight)
   
   picRamp(1).BackColor = vbButtonFace
   picRamp(1).ForeColor = 0
   x1 = RampFrac(1) * picRamp(1).ScaleWidth
   picRamp(1).Line (0, 0)-(x1, picRamp(1).ScaleHeight)
   picRamp(1).Line (x1, picRamp(1).ScaleHeight)-(picRamp(1).ScaleWidth, 0)
   
   Ramp = 2
   RampFrac(2) = 0
   If Not aBlock Then EvalPlot PIC, PIC2, uAmp * AmpMult, uFrq, uDur, aPlay
   aPlayOnce = False
End Sub

' Echoes ----------------------------------------------------

Private Sub HSEcho_Scroll()
   Call HSEcho_Change
End Sub

Private Sub HSEcho_Change()
   EchoMul = HSEcho.Value
   LabEcho = EchoMul
   If aEcho Then
      If Not aBlock Then EvalPlot PIC, PIC2, uAmp * AmpMult, uFrq, uDur, aPlay
      aPlayOnce = False
   End If
End Sub

Private Sub chkEcho_Click()
   aEcho = -chkEcho.Value
   If aEcho Then
      EchoMul = HSEcho.Value
      LabEcho = EchoMul
   Else
      chkStagger.Value = chkEcho.Value
      aStagger = aEcho
   End If
   If Not aBlock Then EvalPlot PIC, PIC2, uAmp * AmpMult, uFrq, uDur, aPlay
   aPlayOnce = False
End Sub

Private Sub chkStagger_Click()
   aStagger = -chkStagger.Value
   If aStagger Then
      chkEcho.Value = chkStagger.Value
      aEcho = aStagger
   End If
   If Not aBlock Then EvalPlot PIC, PIC2, uAmp * AmpMult, uFrq, uDur, aPlay
   aPlayOnce = False
End Sub

' Repeat ----------------------------------------------------

Private Sub HSRepeat_Scroll()
   Call HSRepeat_Change
End Sub

Private Sub HSRepeat_Change()
   RepeatMul = HSRepeat.Value
   LabRepeat = RepeatMul
   If aRepeat Then
      If Not aBlock Then EvalPlot PIC, PIC2, uAmp * AmpMult, uFrq, uDur, aPlay
      aPlayOnce = False
   End If
End Sub

Private Sub chkRepeat_Click()
   aRepeat = -chkRepeat.Value
   If aRepeat Then
      RepeatMul = HSRepeat.Value
      LabRepeat = RepeatMul
   End If
   If Not aBlock Then EvalPlot PIC, PIC2, uAmp * AmpMult, uFrq, uDur, aPlay
   aPlayOnce = False
End Sub

' Presets ----------------------------------------------------

Public Sub Presets()
'' Presets all Public
' pFIndex     = 1 to 16
' pEcho       = 0 UnChecked, 1 Checked
' pEchoMul  = 4 to 32
' pStagger    = 0 UnChecked, 1 Checked
' pRamp       = 0 or 1 U/D or D/U
' pRampFrac   = 0.0 to 1.0
' pReverse    = 0 UnChecked, 1 Checked
' pAmp        = 0.0 to 1.0
' pFreq       = 10 to 8000
' pDuration   = 0.1 to 10
' pSampleRate = 1 to 4
' pAbs        = 0 UnChecked, 1 Checked
' pRepeat     = 0 UnChecked, 1 Checked
' pRepeatMul  = 2 to 32
' pBitnum     = 0(8 bit), 1(16 bit)


   aBlock = True
   
   List1.ListIndex = pFIndex - 1
   chkEcho.Value = pEcho
   HSEcho.Value = pEchoMul
   If pEcho = 0 Then pStagger = 0
   chkStagger.Value = pStagger
   If pRamp <> 2 Then
      Call picRamp_MouseMove(pRamp, 1, 0, pRampFrac * picRamp(pRamp).ScaleWidth, 0)
   Else
      cmdRamp_Click
   End If
   HS(0).Value = pAmp * 32767
   HS(1).Value = pFreq
   HS(2).Value = 10 * pDuration
   HS(3).Value = pSampleRate
   chkReverse.Value = pReverse
   aReverse = pReverse
   chkABS.Value = pAbs
   aABS = pAbs
   aRepeat = pRepeat
   chkRepeat.Value = pRepeat
   HSRepeat.Value = pRepeatMul
   Bitnum = pBitNum
   optBit(Bitnum).Value = True
   InitHeader
   aBlock = False
   EvalPlot PIC, PIC2, uAmp * AmpMult, uFrq, uDur, aPlay
   aPlayOnce = False
End Sub

Private Sub cmdPresets_Click()
   Load frmPresets   ' Reads values from Presets.txt
                     ' & lastly calls Form1.Presets
   If ReadError Then Unload frmPresets
End Sub

Private Sub cmdAdd2Presets_Click()
Dim A$
Dim fnum
   Unload frmPresets
   ' Name
   ' get all values
   ' add to Presets.txt
   fnum = 0
   If FileExists(PathSpec$ & "Presets.txt") Then
      A$ = InputBox("ENTER WAVETTE NAME", "Add to Presets", , 95, 95)
      If Trim$(A$) = "" Then
         PlayNoName
         MsgBox "No name entered or Cancelled   ", vbInformation, "Add to Presets"
         Exit Sub
      End If
   Else
      fnum = MsgBox("Make new Presets.txt file", vbQuestion + vbYesNo, "Add to Presets")
      If fnum = vbNo Then Exit Sub
   End If
   
   If fnum = vbYes Then
      A$ = InputBox("ENTER WAVETTE NAME", "Add to Presets", , 95, 95)
      If Trim$(A$) = "" Then
         PlayNoName
         MsgBox "No name entered or Cancelled   ", vbInformation, "Add to Presets"
         Exit Sub
      End If
      fnum = FreeFile
      Open PathSpec$ & "Presets.txt" For Output As #fnum
   Else
      fnum = FreeFile
      Open PathSpec$ & "Presets.txt" For Append As #fnum
   End If
   
   Print #fnum,
   Print #fnum, "Name = " & A$
   
   pFIndex = List1.ListIndex + 1
   Print #fnum, "pFIndex =" & Str$(pFIndex)
   
   pEcho = chkEcho.Value
   Print #fnum, "pEcho =" & Str$(pEcho)
   pEchoMul = HSEcho.Value
   Print #fnum, "pEchoMul =" & Str$(pEchoMul)
   pStagger = chkStagger.Value
   Print #fnum, "pStagger =" & Str$(pStagger)
   Print #fnum, "pRamp =" & Str$(Ramp)
   Print #fnum, "pRampFrac =" & Str$(RampFrac(Ramp))
   pReverse = chkReverse.Value
   Print #fnum, "pReverse =" & Str$(pReverse)
   
   pAmp = HS(0).Value / 32767
   pFreq = HS(1).Value
   pDuration = HS(2).Value / 10
   pSampleRate = HS(3).Value
   
   pAbs = chkABS.Value
   Print #fnum, "pAmp =" & Str$(pAmp)
   Print #fnum, "pFreq =" & Str$(pFreq)
   Print #fnum, "pDuration =" & Str$(pDuration)
   Print #fnum, "pSampleRate =" & Str$(pSampleRate)
   Print #fnum, "pAbs =" & Str$(pAbs)
   
   pRepeat = chkRepeat.Value
   Print #fnum, "pRepeat =" & Str$(pRepeat)
   pRepeatMul = HSRepeat.Value
   Print #fnum, "pRepeatMul =" & Str$(pRepeatMul)
   Print #fnum, "pBitnum =" & Str$(Bitnum)
   
   Close #fnum
End Sub


'Save WAV  ----------------------------------------------------

Private Sub cmdSave_Click()
Dim Title$, Filt$, InDir$
Dim fnum As Long
Dim res As Long
   If Not aPlayOnce Then
      PlayFirst
      MsgBox "Play it first !    ", vbInformation, "Save WAV"
      Exit Sub
   End If

   Dim CDL As OSDialog
   Title$ = "Save As wav file"
   Filt$ = "Save wav|*.wav"
   InDir$ = CurrPath$ 'PathSpec$
   FileSpec$ = ""
   Set CDL = New OSDialog
   CDL.ShowSave FileSpec$, Title$, Filt$, InDir$, "", Me.hWnd
   Set CDL = Nothing

   If Len(FileSpec$) = 0 Then Exit Sub 'for cancel
   If FileExists(FileSpec$) Then
      res = MsgBox("Delete" & vbCrLf & FileSpec$ & "  " & vbCrLf & _
            "binary file first", vbQuestion + vbYesNo, "Saving WAV")
      If res = vbYes Then
         Kill FileSpec$ ' Else get appending with existing wav file
      Else
         MsgBox "then file not saved to avoid appending to existing file !     ", vbInformation, "Saving WAV"
      End If
   End If
   CurrPath$ = FileSpec$
   fnum = FreeFile
   Open FileSpec$ For Binary Access Write As fnum
   Put fnum, , SoundFile
   Close fnum
   aPlayOnce = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
   StopPlay
   DoEvents
   FreeLibrary m_hMod
   If aPresets Then
      Unload frmPresets
      Set frmPresets = Nothing
   End If
   Set Form1 = Nothing
End Sub


