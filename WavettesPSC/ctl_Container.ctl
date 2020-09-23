VERSION 5.00
Begin VB.UserControl Container 
   AutoRedraw      =   -1  'True
   ClientHeight    =   645
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1395
   ControlContainer=   -1  'True
   ScaleHeight     =   645
   ScaleWidth      =   1395
End
Attribute VB_Name = "Container"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'
'
'
'10/24/2002
'Container written by: Eric Madison, PSC CodeId=40130
'
'
'
Option Explicit

'api for borders
Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
Private Declare Function InflateRect Lib "user32" (lpRect As RECT, ByVal X As Long, ByVal Y As Long) As Long

'api for caption rect and printing caption
Private Declare Function DrawTextEx Lib "user32" Alias "DrawTextExA" (ByVal hdc As Long, ByVal lpsz As String, ByVal n As Long, lpRect As RECT, ByVal un As Long, ByVal lpDrawTextParams As Any) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'declare enumerated properties for end user
'   properties window dropdown selections
Public Enum TheBackStyle
    Independent
    AmbientMode
End Enum

Public Enum TheBorderStyle
    Flat
    Bump
    Etch
    None
End Enum

'declare end user public events
Event Click()
Event DblClick()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

'store our font with stdfont to allow end user to
'   select a font at design time or run time
Private WithEvents TheFont As StdFont
Attribute TheFont.VB_VarHelpID = -1

'declare property variables
Private TheBorderStyleX As TheBorderStyle
Private TheBackStyleX As TheBackStyle
Private TheForeColor As OLE_COLOR, TheBackColor As OLE_COLOR
Private TheBorderColorDark As OLE_COLOR, TheBorderColorLight As OLE_COLOR
Private TheBorderWidth As Integer
Private TheCaption As String
Private TheEnabled As Boolean

'raise end user events as they occur
Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If TheBackStyleX = AmbientMode Then 'if parent forms backcolor changes
        If PropertyName = "BackColor" Then
            DrawContainer
        End If
    End If
End Sub

Private Sub TheFont_FontChanged(ByVal PropertyName As String)
    Set UserControl.Font = TheFont
    UserControl.Refresh
End Sub

Private Sub UserControl_Initialize()
    'create instance of the stdfont object and assign
    '   it to the controls font property
    Set TheFont = New StdFont
    Set UserControl.Font = TheFont
End Sub

Private Sub UserControl_InitProperties()
    BackColor = RGB(192, 192, 192)
    BorderColorDark = RGB(128, 128, 128)
    BorderColorLight = RGB(255, 255, 255)
    BorderStyle = 2
    BorderWidth = 1
    Caption = Extender.Name
    Enabled = True
    TheFont.Name = "Arial"
    ForeColor = RGB(0, 0, 0)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        BackColor = .ReadProperty("BackColor", RGB(192, 192, 192))
        BackStyle = .ReadProperty("BackStyle", 0)
        BorderColorDark = .ReadProperty("BorderColorDark", RGB(128, 128, 128))
        BorderColorLight = .ReadProperty("BorderColorLight", RGB(255, 255, 255))
        BorderStyle = .ReadProperty("BorderStyle", 2)
        BorderWidth = .ReadProperty("BorderWidth", 1)
        Caption = .ReadProperty("Caption", Extender.Name)
        Enabled = .ReadProperty("Enabled", True)
        Set Font = .ReadProperty("Font")
        ForeColor = .ReadProperty("ForeColor", RGB(0, 0, 0))
        UserControl.MousePointer = .ReadProperty("MousePointer", vbDefault)
        Set UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
        Set Picture = .ReadProperty("Picture")
    End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "BackColor", BackColor, RGB(192, 192, 192)
        .WriteProperty "BackStyle", BackStyle, 0
        .WriteProperty "BorderColorDark", BorderColorDark, RGB(128, 128, 128)
        .WriteProperty "BorderColorLight", BorderColorLight, RGB(255, 255, 255)
        .WriteProperty "BorderStyle", BorderStyle, 2
        .WriteProperty "BorderWidth", BorderWidth, 1
        .WriteProperty "Caption", Caption, Extender.Name
        .WriteProperty "Enabled", Enabled, True
        .WriteProperty "Font", Font
        .WriteProperty "ForeColor", ForeColor, RGB(0, 0, 0)
        .WriteProperty "MousePointer", MousePointer, vbDefault
        .WriteProperty "MouseIcon", MouseIcon, Nothing
        .WriteProperty "Picture", Picture
    End With
End Sub

Private Sub UserControl_Resize()
    DrawContainer 'make sure to redraw when resizing
End Sub

'set up end user properties
Public Property Get BackColor() As OLE_COLOR
    BackColor = TheBackColor
End Property

Public Property Let BackColor(ByVal NewColor As OLE_COLOR)
    TheBackColor = NewColor
    DrawContainer
    PropertyChanged "BackColor"
End Property

Public Property Get BackStyle() As TheBackStyle
    BackStyle = TheBackStyleX
End Property

Public Property Let BackStyle(ByVal NewStyle As TheBackStyle)
    TheBackStyleX = NewStyle
    DrawContainer
    PropertyChanged "BackStyle"
End Property

Public Property Get BorderColorDark() As OLE_COLOR
    BorderColorDark = TheBorderColorDark
End Property

Public Property Let BorderColorDark(ByVal NewColor As OLE_COLOR)
    TheBorderColorDark = NewColor
    DrawContainer
    PropertyChanged "BorderColorDark"
End Property

Public Property Get BorderColorLight() As OLE_COLOR
    BorderColorLight = TheBorderColorLight
End Property

Public Property Let BorderColorLight(ByVal NewColor As OLE_COLOR)
    TheBorderColorLight = NewColor
    DrawContainer
    PropertyChanged "BorderColorLight"
End Property

Public Property Get BorderStyle() As TheBorderStyle
    BorderStyle = TheBorderStyleX
End Property

Public Property Let BorderStyle(ByVal NewStyle As TheBorderStyle)
    TheBorderStyleX = NewStyle
    DrawContainer
    PropertyChanged "BorderStyle"
End Property

Public Property Get BorderWidth() As Integer
   BorderWidth = TheBorderWidth
End Property

Public Property Let BorderWidth(ByVal NewWidth As Integer)
   If NewWidth > 1 And TheBorderStyleX <> 0 Then
      TheBorderWidth = 1
   Else
      TheBorderWidth = NewWidth
   End If
   DrawContainer
   PropertyChanged "BorderWidth"
End Property

Public Property Get Caption() As String
    Caption = TheCaption
End Property

Public Property Let Caption(ByVal NewCaption As String)
    TheCaption = NewCaption
    DrawContainer
    PropertyChanged "Caption"
End Property

Public Property Get Enabled() As Boolean
    Enabled = TheEnabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    TheEnabled = NewValue
    UserControl.Enabled = TheEnabled
    DrawContainer
    PropertyChanged "Enabled"
End Property

Public Property Get Font() As StdFont
    Set Font = TheFont
End Property

Public Property Set Font(NewFont As StdFont)
    If NewFont Is Nothing Then Exit Property
    With TheFont
        .Bold = NewFont.Bold
        .Charset = NewFont.Charset
        .Italic = NewFont.Italic
        .Name = NewFont.Name
        .Size = NewFont.Size
        .Strikethrough = NewFont.Strikethrough
        .Underline = NewFont.Underline
        .Weight = NewFont.Weight
    End With
    DrawContainer
    PropertyChanged "Font"
End Property

Public Property Get ForeColor() As OLE_COLOR
    ForeColor = TheForeColor
End Property

Public Property Let ForeColor(ByVal NewColor As OLE_COLOR)
    TheForeColor = NewColor
    DrawContainer
    PropertyChanged "ForeColor"
End Property

Public Property Get hdc() As Long
    hdc = UserControl.hdc 'dont need this yet
End Property

Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd 'dont need this yet
End Property

Public Property Get MouseIcon() As StdPicture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal MouseIcon As StdPicture)
    Set UserControl.MouseIcon = MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal MousePointer As MousePointerConstants)
    UserControl.MousePointer = MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get Picture() As Picture
    Set Picture = UserControl.Picture
End Property

Public Property Set Picture(ByVal NewPicture As Picture)
    Set UserControl.Picture = NewPicture
    DrawContainer
    PropertyChanged "Picture"
End Property

Private Sub DrawContainer()
    Dim X As Integer
    Dim r As RECT
    Dim Corner As Integer
    
    With UserControl
        .Cls 'erase everything before redrawing
        .ScaleMode = 3 'make sure to set to pixels for api
        Corner = 5 'change this number if you want greater or lesser curvature on border corners
                   ' 1 for no rounded corner
        'set backcolor
        If TheBackStyleX = 0 Then 'if container has independent backcolor
            .BackColor = TheBackColor 'will have own backcolor
        Else
            .BackColor = Ambient.BackColor 'will change to match parent forms backcolor
        End If
    
        'get borders and caption area
        If TheCaption <> "" Then 'if a caption exists
            r.Top = .TextHeight(TheCaption) / 2
        Else
            r.Top = .ScaleTop
            Corner = 1  ' RR
        End If
        r.Left = .ScaleLeft
        r.Bottom = .ScaleTop + .ScaleHeight
        r.Right = .ScaleLeft + .ScaleWidth
    
        'draw borders
        Select Case TheBorderStyleX
            Case Flat
                .ForeColor = TheBorderColorDark
                For X = 1 To TheBorderWidth
                    RoundRect .hdc, r.Left, r.Top, r.Right, r.Bottom, Corner, Corner
                    InflateRect r, -1, -1
                Next
            Case Bump
                .ForeColor = TheBorderColorDark
                RoundRect .hdc, 1, r.Top + 1, r.Right, r.Bottom, Corner, Corner
                .ForeColor = TheBorderColorLight
                RoundRect .hdc, 0, r.Top, r.Right - 1, r.Bottom - 1, Corner, Corner
            Case Etch
                .ForeColor = TheBorderColorLight
                RoundRect .hdc, 1, r.Top + 1, r.Right, r.Bottom, Corner, Corner
                .ForeColor = TheBorderColorDark
                RoundRect .hdc, 0, r.Top, r.Right - 1, r.Bottom - 1, Corner, Corner
            Case Else
                'no borders
        End Select
    
        'set up caption area
        If TheBackStyleX = AmbientMode Then
            .ForeColor = Ambient.BackColor
            .FillColor = Ambient.BackColor
        Else
            .ForeColor = TheBackColor
            .FillColor = TheBackColor
        End If
        .FillStyle = 0 'set to solid
        
        'draw caption area
        If TheCaption <> "" Then 'if a caption exists
            RoundRect .hdc, 4, 0, .TextWidth(TheCaption) + 8, .TextHeight(TheCaption), 0, 0
            'change x1 and x2 coordinates to accomodate corners as necessary
        End If
        .FillStyle = 1 'reset to transparent so no problems later
        
        'draw caption
        If TheEnabled = False Then 'if container is disabled draw caption inset for disabled appearance
            .ForeColor = RGB(255, 255, 255)
            r.Left = 8: r.Top = 1 'change r.left to accomadate corners as necessary
            DrawTextEx .hdc, TheCaption, Len(TheCaption), r, 0&, 0&
            .ForeColor = RGB(128, 128, 128)
            r.Left = 7
            r.Top = 0 'change r.left to accomodate corners as necessary
            DrawTextEx .hdc, TheCaption, Len(TheCaption), r, 0&, 0&
        Else 'it's enabled
            .ForeColor = TheForeColor
            r.Left = 7: r.Top = 0 'change r.left to accomadate corners as necessary
            DrawTextEx .hdc, TheCaption, Len(TheCaption), r, 0&, 0&
        End If
    End With
End Sub

