VERSION 5.00
Begin VB.Form frmPresets 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Presets"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   495
   ClientWidth     =   3180
   Icon            =   "frmPresets.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4800
   ScaleWidth      =   3180
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4140
      Left            =   105
      TabIndex        =   0
      Top             =   270
      Width           =   2925
   End
End
Attribute VB_Name = "frmPresets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' frmPresets.frm

Option Explicit

Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, _
ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, _
ByVal wi As Long, ByVal ht As Long, ByVal wFlags As Long) As Long

Private Const hWndInsertAfter = -1
Private Const wFlags = &H40 Or &H20

Private Presets$()
Private NLines As Long

' Presets all Public
   ' pFIndex     = 1 to 16
   ' pEcho       = 0 UnChecked, 1 Checked
   ' pEchoMul    = 1 to 32
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

Private Sub Form_Load()
Dim k As Long
Dim fw As Long, fh As Long
Dim fnum As Long
Dim A$
ReadError = False
aPresets = True

On Error GoTo one
   fw = 3270
   fh = 5090
   k = SetWindowPos(frmPresets.hWnd, hWndInsertAfter, 20, 20, _
      fw / Screen.TwipsPerPixelX, fh / Screen.TwipsPerPixelY, wFlags)
   
   ReDim Presets$(1 To 10)
   
   Show
   If FileExists(PathSpec$ & "Presets.txt") Then
      NLines = 1
      fnum = FreeFile
      Open PathSpec$ & "Presets.txt" For Input As #fnum
      Do
         Line Input #fnum, A$
         If Left$(A$, 1) <> "" Then
         If InStr(1, A$, "'") = 0 Then
            Presets$(NLines) = A$
            NLines = NLines + 1
            If NLines > UBound(Presets$(), 1) Then
               ReDim Preserve Presets$(1 To UBound(Presets$(), 1) + 24)
            End If
         End If
         End If
      Loop Until EOF(fnum)
      Close fnum
   Else
      MsgBox " Presets.txt not there    ", vbInformation, "Loading"
      ReadError = True
      Exit Sub
   End If
   NLines = NLines - 1

   For k = 1 To NLines
      If InStr(1, Presets$(k), "Name =") <> 0 Then
         A$ = Mid$(Presets$(k), InStr(Presets$(k), "=") + 1)
         List1.AddItem Trim$(A$)
      End If
   Next k
   Print List1.ListCount
   ' Get Presets.txt
   'List1.AddItem "Foghorn(Big Ship)"
   'List1.AddItem "Foghorn(Smaller Ship)"
   ' etc
   On Error GoTo 0
   Exit Sub
'===========
one:
   MsgBox "Error reading Presets.txt", vbCritical, "Presets"
   ReadError = True
   aPresets = False
   Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
   aPresets = False
   Unload Me
End Sub

Private Sub List1_Click()
Dim A$, C$
Dim B$(1 To 15)
Dim k As Long, j As Long
Dim p As Long
'On Error GoTo two
   A$ = List1.List(List1.ListIndex)
   ' Search Presets$() for A$
   ' then read following 11 values
   For k = 1 To NLines
      p = InStr(1, Presets$(k), "=") + 1
      C$ = Trim$(Mid$(Presets$(k), p))
      If C$ = A$ Then
         For j = 1 To 15   ' 15 items after Name
            B$(j) = Trim$(Mid$(Presets$(k + j), InStr(1, Presets$(k + j), "=") + 1))
         Next j
         Exit For
      End If
   Next k
   
   If k = NLines + 1 Then GoTo two 'Stop
   
   pFIndex = Val(B$(1))
   pEcho = Val(B$(2))
   pEchoMul = Val(B$(3))
   pStagger = Val(B$(4))
   pRamp = Val(B$(5))
   pRampFrac = Val(B$(6))
   pReverse = Val(B$(7))
   pAmp = Val(B$(8))
   pFreq = Val(B$(9))
   pDuration = Val(B$(10))
   pSampleRate = Val(B$(11))
   pAbs = Val(B$(12))
   pRepeat = Val(B$(13))
   pRepeatMul = Val(B$(14))
   pBitNum = Val(B$(15))
        
   Erase B$()
   Form1.Presets
   On Error GoTo 0
   Exit Sub
'===========
two:
   MsgBox "Error extracting values", vbCritical, "Presets"
   aPresets = False
   Unload Me
End Sub
