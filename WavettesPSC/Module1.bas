Attribute VB_Name = "Module1"
'Module1.bas   ~Wavettes~

Option Explicit

' Main routine at Sub EvalPlot
' Evaluate at Function EvalFunc
' Echo at Sub EchoIt
' Ramp at Sub RampIt
' Repeat at Sub RepeatIt

Private Declare Sub CopyMemory Lib "Kernel32" Alias "RtlMoveMemory" _
   (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

' Canonical WAV header
Private Type tHeader          ' Header length 44 bytes
    RIFF      As Long         ' "RIFF"
    LenR      As Long         '  Length following (FileSize-8)
    WAVE      As Long         ' "WAVE"
    fmt       As Long         ' "fmt "
    ChunkSize As Long         '  Chunksize
    
    PCMFormat As Integer      '  PCM format = 1
    NumChan   As Integer      '  Number of channels
    SRate     As Long         '  Sample rate
    BRate     As Long         '  Byte rate
    Blk       As Integer      '  Block align (bytes per sample)
    Bits      As Integer      '  Bits per sample
    
    data      As Long         ' "data"
    LenData   As Long         '  Length of datastream, bytes
End Type

Private ByteSamples() As Byte
Private Samples() As Integer
Private SamplesPerCycle As Double
Private omega As Double

Private PICW As Long
Private PICH As Long
Private NumSamples As Long

' For Echo
Private ysave() As Single

' For polar
Private xc As Single, yc As Single
Private sang As Single
'----------------------------------------------------

Public Header As tHeader
' Used in Form1
Public SamplesPerSecond As Long
Public SoundFile() As Byte

' Associated with each function is
Public Func$()
Public FuncIndex As Long

' Passed params
Public Ramp As Long   ' 0 up/down, 1 Down/Up, 2 None
Public RampFrac() As Single
Public aShape As Boolean
Public aABS As Boolean
Public aEcho As Boolean
Public EchoMul As Single
Public aStagger As Boolean
Public aReverse As Boolean
Public aRepeat As Boolean
Public RepeatMul As Single
Public Bitnum As Integer   ' 0(8bit), 1(16bit)

Public AmpMult As Long  ' Test

Public PathSpec$, CurrPath$, FileSpec$

' Presets
   ' pFIndex     = 1 to 16
   ' pEcho       = 0 UnChecked, 1 Checked
   ' pEchoMul    = 1 to 32
   ' pRamp       = 0.0 to 1.0
   ' pRampFrac   = 0 or 1 U/D or D/U
   ' pReverse    = 0 UnChecked, 1 Checked
   ' pAmp        = 0.0 to 1.0
   ' pFreq       = 10 to 8000
   ' pDuration   = 0.1 to 10
   ' pSampleRate = 1 to 4
   ' pAbs        = 0 UnChecked, 1 Checked
   ' pRepeat     = 0 UnChecked, 1 Checked
   ' pRepeatMul  = 2 to 32
   ' pBitNum     = 0 or 1  8 or 16 bit

   Public pFIndex As Long
   Public pEcho As Long
   Public pEchoMul As Long
   Public pStagger As Long
   Public pRamp As Integer
   Public pRampFrac As Single
   Public pReverse As Long
   Public pAmp As Single
   Public pFreq As Long
   Public pDuration As Single
   Public pSampleRate As Long
   Public pAbs As Long
   Public pRepeat As Long
   Public pRepeatMul As Long
   Public pBitNum As Integer
   
   Public ReadError As Boolean
   Public aPresets As Boolean
   
Public Const pi# = 3.14159625

Public Sub EvalPlot(PIC As PictureBox, PIC2 As PictureBox, uAmp As Single, uFrq As Single, uDur As Single, aPlay As Boolean)
' Public FuncIndex
' Public omega As Double
' Public SamplesPerCycle As Double
Dim X As Single, Y As Single
Dim xp As Double
Dim xintv As Single
Dim yoff As Single

'Dim NumSamples        As Long
Dim Time As Double
Dim Vol As Double
Dim VolMax As Double
Dim DeltaTime As Double
Dim DeltaVol  As Double
  
Dim Duration As Single
Dim Amp As Double
Dim AmpMax As Double
Dim DeltaAmp As Double
Dim Ptr As Long
 
Dim OT As Double
Dim Sample As Single 'Double

' Polar
Dim ang As Single
Dim dx As Single, dy As Single

'Dim ymax As Single     ' Test

   PICW = PIC.ScaleWidth
   PICH = PIC.ScaleHeight
   
   ReDim ysave(PICW)   ' eg 0-508 (509)
   PIC.Picture = LoadPicture
   
   ' For Polar plot
   PIC2.Picture = LoadPicture
   xc = PIC2.ScaleWidth / 2
   yc = PIC2.ScaleHeight / 2
   sang = 2 * pi# / PIC2.ScaleWidth
   
   omega = 2 * pi# * uFrq  ' Public 2pf
   SamplesPerCycle = SamplesPerSecond / uFrq  ' For Sawtooth 1
   
   If aShape Then
      ' Show false shape
      omega = 2 * pi# * 10    ' Fix plot display false but visible
      If FuncIndex = 9 Then
         omega = 2 * pi# * 300
         SamplesPerCycle = SamplesPerSecond / 300  ' For Sawtooth 1
      End If
   End If
   
   yoff = PICH / 2
   Amp = 0.5 * uAmp * (PICH / 32767) ' eg Amp = uAmp* 0.00458/2 so when uAmp = 32767 Amp = 77.5
   If Amp > yoff Then Amp = yoff
   AmpMax = Amp
   xintv = omega / (PICW) ' 512
'ymax = 0  ' test
   DeltaAmp = Amp / PICH 'W
   
   If Ramp = 0 Then Amp = 0
   
   For xp = 0 To PICW - 1
      X = xp * xintv   ' Max x = 512*xintv = 2*pi#*uFrq
      Y = EvalFunc(FuncIndex, Amp, CDbl(X), CLng(xp))
      
      If aEcho Then EchoIt Y, xp, PICW, 0
      If aRepeat Then RepeatIt Y, xp, PICW, 0
      ysave(xp) = Y  ' For Echo
      If Ramp <> 2 Then RampIt Amp, xp, PICW, AmpMax
      ' Polar plot
      ang = sang * xp
      If aReverse Then ang = -ang
      dx = Y * Cos(ang)
      dy = Y * Sin(ang)
      PIC2.Line (xc, yc)-(xc + dx, yc + dy), vbCyan
      Y = yoff + Y
      
      If Y > PICH Then Y = PICH
      
      If aReverse Then
         PIC.Line (PICW - 1 - xp, Y)-(PICW - 1 - xp, yoff), vbCyan
      Else
         PIC.Line (xp, Y)-(xp, yoff), vbCyan
      End If
' Test
'If Y > ymax Then ymax = Y
'Form1.LabTest = "ymax =" & Str$(CInt(ymax))

   Next xp
   PIC2.FillStyle = 0
   PIC2.Circle (xc, yc), 2
   PIC.Line (0, yoff)-(PICW - 1, yoff), vbGreen
   PIC.Refresh
   
   If aPlay Then
      
      omega = 2 * pi# * uFrq  ' Public 2pf
      Duration = uDur
      NumSamples = SamplesPerSecond * Duration
      If NumSamples = 0 Then NumSamples = 1
      
      If Bitnum = 0 Then ReDim ByteSamples(1 To NumSamples) ' 8 bit
      ReDim Samples(1 To NumSamples)      ' 16 bit
      
      Vol = uAmp
      If Vol > 32767 Then Vol = 32767
      VolMax = Vol
      DeltaVol = Vol / NumSamples
      DeltaTime = Duration / NumSamples
      SamplesPerCycle = SamplesPerSecond / uFrq  ' For sawtooth
      Time = 0
      
      If Ramp = 0 Then Vol = 0
      
      For Ptr = 1 To NumSamples
         OT = omega * Time
         Sample = EvalFunc(FuncIndex, Vol, OT, Ptr, 4)
         If Sample > 32767 Then Sample = 32767
         If Sample < -32768 Then Sample = -32768   ' Unsigned Integer 0-65535
         If aEcho Then EchoIt Sample, Ptr, NumSamples, 1
         If aRepeat Then RepeatIt Sample, Ptr, NumSamples, 1
         If Ramp <> 2 Then RampIt Vol, Ptr - 1, NumSamples, VolMax
         ' Sample(Single) -> Samples()(Int) -> SoundFile()(Bytes)
         
         If Sample > 32767 Then Sample = 32767
         If Sample < -32768 Then Sample = -32768   ' Unsigned Integer 0-65535
         Samples(Ptr) = CInt(Sample)
         
         If Bitnum = 0 Then ' 8 bit Rescale 16 bit -32768 -> 32767 to 0 -> 255
            Sample = Sample \ 256
            Sample = Sample + 127
            If Sample > 255 Then Sample = 255
            If Sample < 0 Then Sample = 0   ' Unsigned Byte 0-255
            ByteSamples(Ptr) = CByte(Sample)
         End If
         
         Time = Time + DeltaTime
      Next Ptr
      
      If aReverse Then
         If Bitnum = 0 Then ' 8 bit
            Reverse8bitSound
         Else  ' 16 bit
            Reverse16bitSound
         End If
      End If
      
      With Header
         .LenData = NumSamples * .NumChan * .Bits / 8 ' Samples()(Int) so NumSamples*1*2 is #bytes
                                                        ' .Bits = 16 bits = 2 bytes. Integer data
         .LenR = Len(Header) + .LenData - 8  ' Length following segment (Filesize-8, unless padders on EOF)
         ReDim SoundFile(1 To Len(Header) + .LenData)
         CopyMemory SoundFile(1), Header.RIFF, Len(Header)
         
         If Bitnum = 0 Then ' 8 bit
            CopyMemory SoundFile(1 + Len(Header)), ByteSamples(1), .LenData
         Else   ' 16 bit
            CopyMemory SoundFile(1 + Len(Header)), Samples(1), .LenData
         End If
         ' So sound bytes start at pos Len(Header)
      End With
   
   End If
   
   Erase Samples, ByteSamples
End Sub

Public Function EvalFunc(Index As Long, AV As Double, XT As Double, k1 As Long, Optional k2 As Long = 1) As Double
' Public omega As Double
' Public SamplesPerCycle As Double
' AV:  Amp or Vol
' XT:    x or OT
' k1 for Sawtooth 1
' k2 modify amp
Dim k As Long, j As Long
Dim Sum As Single
   Select Case Index
   Case 1:  EvalFunc = AV * Sin(XT) * k2
   Case 2:  EvalFunc = AV * (Sin(XT) + Sin(XT * pi# / 3)) * k2
   Case 3:  EvalFunc = (AV * (XT * pi#) * Sin(XT)) / omega
   Case 4:  EvalFunc = AV * (Rnd - Rnd)
   Case 5:  EvalFunc = (AV * (Sin(XT) + Sin((XT) ^ 2 + Sin(XT) ^ 3))) / 3
   Case 6:  EvalFunc = (AV * Sin(Int(XT / pi#)) * pi# + pi# / 2) / (1.5 * pi#)
   Case 7:  EvalFunc = AV * Sin(XT) ^ 3 * k2
   Case 8:  EvalFunc = AV * Sin(XT ^ 2) ^ 3 * k2
   Case 9:  EvalFunc = AV / SamplesPerCycle * 2 * (k1 Mod SamplesPerCycle) - AV  ' Sawtooth 1
   Case 10: EvalFunc = AV * Sgn(Sin(XT))   ' Squarewave
   Case 11: EvalFunc = AV * (Sin(3 * XT) / Tan(XT + 0.1)) / omega
   Case 12: EvalFunc = AV * (Sin(3 * XT) / Tan(XT + 0.01)) * Sin(XT / 2)
   Case 13: EvalFunc = AV * Exp(Sin(XT + pi# / 2)) * Sin(XT)
   Case 14: EvalFunc = AV * Exp(Sin(XT ^ 2 + pi# / 2)) * Sin(XT)
   Case 15: EvalFunc = AV * Atn(Cos(XT ^ 3) + Sin(XT ^ 2))
   Case 16: EvalFunc = AV * (XT ^ Sin(XT)) / omega
   Case 17 'Sawtooth 2
      'Func$(17) = "17. -4(Cos(x)+Cos(3*x)/9+Cos(5*x)/25+ +)/pi"
      Sum = 0
      For k = 1 To 6
         j = 2 * k - 1
         Sum = Sum + Cos(j * XT) / (j * j)
      Next k
      EvalFunc = -AV * 4 * Sum / pi#
   Case 18  ' Bumps
      'Func$(18) = "18. 1/pi#+Sin(x)/2-(Cos(2*x)/1*3 + Cos(4*x)/3*5 + cos(6*x)/5*7 + +)/pi"
      Sum = 0
      For k = 1 To 6
         Sum = Sum + Cos(2 * k * XT) / (4 * k * k - 1)
      Next k
      EvalFunc = -AV * (1 / pi# + Sin(XT) / 2 - 2 * Sum / pi#)
   Case 19 ' Cosec(x)
      Sum = Sin(XT)
      If Sum = 0 Then Sum = 40 ' Avoid /0
      EvalFunc = AV * (1 / Sum)
   
   
   End Select
   
   If aABS Then EvalFunc = -Abs(EvalFunc)
End Function

Public Sub EchoIt(YS As Single, ByVal Stp As Variant, Span As Long, PS As Integer)
' Public ysave() As Single
' Public Samples() As Integer
' Public EchoMul set at Sub HSEcho
'                YS      Stp   Span
'Plot:    EchoIt Y,       xp,  PIC.ScaleWidth,     PS = 0
'Sound:   EchoIt Sample, Ptr,  NumSamples,         PS = 1
Dim T As Single
Dim SEoff As Single

   T = Exp(-3 * Stp / Span) * Cos(EchoMul * pi# * Stp / Span)
   YS = YS * T
   
   If aStagger Then
      
      SEoff = 0.5 * Span / EchoMul
      If Stp > SEoff Then
         If PS = 0 Then ' Plot
            YS = ysave(Stp - 0.33 * SEoff + 1) / 2
            YS = YS + (ysave(Stp - SEoff + 1)) / 2
         Else  ' Sound
            YS = CSng(Samples(Stp - 0.33 * SEoff + 1)) / 2
            YS = YS + (CSng(Samples(Stp - SEoff + 1))) / 2
         End If
      End If
   
   End If
   
   If YS > 32767 Then
      YS = 32767
   End If
   If YS < -32768 Then
      YS = -32768
   End If
End Sub

Public Sub RepeatIt(YS As Single, ByVal Stp As Variant, Span As Long, PS As Integer)
' Public ysave() As Single
' Public Samples() As Integer
' Public RepeatMul set at Sub HSRepeat
'                  YS      Stp   Span
'Plot:    RepeatIt Y,       xp,  PIC.ScaleWidth,     PS = 0
'Sound:   RepeatIt Sample, Ptr,  NumSamples,         PS = 1
Dim SEoff As Single
Dim ss As Long
   SEoff = Span / RepeatMul
   If Stp > SEoff Then
      ss = Stp - SEoff + 1
      If ss < 1 Then ss = 1
      If PS = 0 Then ' Plot
         YS = (ysave(ss))
      Else  ' Sound
         YS = (CSng(Samples(ss)))
      End If
      If YS > 32767 Then
         YS = 32767
      End If
      If YS < -32768 Then
         YS = -32768
      End If
   End If
End Sub

Public Sub RampIt(AV As Double, ByVal Stp As Variant, Span As Long, AVMax As Double)
'Plot:      RampIt Amp, xp, PIC.ScaleWidth, AmpMax
'Sound:     RampIt Vol, Ptr, NumSamples, VolMax
' Public Ramp, RampFrac()
   If Ramp = 0 Then  ' AV starts as 0
      If Stp <= Span * RampFrac(0) Then ' Up/Down
         If RampFrac(0) < 0.001 Then
            AV = AVMax
         Else
            AV = AV + AVMax / (RampFrac(0) * Span)
         End If
         If AV > 32767 Then
            AV = 32767
         End If
      Else
         AV = AV - AVMax / (Span * (1 - RampFrac(0)))
         If AV < 0 Then AV = 0
      End If
   
   ElseIf Ramp = 1 Then ' AV starts as Max
      
      If Stp <= Span * RampFrac(1) Then ' Down/Up
         If RampFrac(1) < 0.001 Then
            AV = 0
         Else
            AV = AV - AVMax / (RampFrac(1) * Span)
         End If
         If AV < -32768 Then
            AV = -32768
         End If
         'If AV < 0 Then
         '   AV = 0
         'End If
      Else
         AV = AV + AVMax / (Span * (1 - RampFrac(1)))
         If AV > 32767 Then AV = 32767
      End If
   Else ' None
   End If
End Sub

Public Sub Reverse16bitSound()
Dim irev() As Integer
Dim NoB As Long, k As Long
   ReDim irev(1 To NumSamples)
   NoB = 2 * NumSamples
   CopyMemory irev(1), Samples(1), NoB
   For k = 1 To NumSamples
      Samples(NumSamples - k + 1) = irev(k)
   Next k
   Erase irev()
End Sub

Public Sub Reverse8bitSound()
Dim irev() As Byte
Dim NoB As Long, k As Long
   ReDim irev(1 To NumSamples)
   NoB = NumSamples
   CopyMemory irev(1), ByteSamples(1), NoB
   For k = 1 To NumSamples
      ByteSamples(NumSamples - k + 1) = irev(k)
   Next k
   Erase irev()
End Sub


Public Sub InitHeader()
Dim v As Long
   With Header
      .RIFF = 1179011410
      .WAVE = 1163280727
      .fmt = 544501094
      .ChunkSize = 16
      .PCMFormat = 1
      .NumChan = 1
      .SRate = SamplesPerSecond   ' 44100 etc
      
      If Bitnum = 0 Then ' 8 bit
         .Bits = 8  '  8 Bit
      Else
         .Bits = 16 ' 16 Bit
      End If
      
      ' The number of bytes for one sample including all channels.
      .Blk = .NumChan * .Bits / 8   ' eg (1 chan, 16 bits = 2) (1 chan, 8 bits = 1)
      
      .BRate = .SRate * .Blk
      '.Bits = 16  'Bits Per Sample
      .data = 1635017060
   End With
End Sub

Public Function FileExists(FSpec$) As Boolean
  On Error Resume Next
  FileExists = FileLen(FSpec$)
End Function
