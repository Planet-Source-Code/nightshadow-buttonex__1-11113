VERSION 5.00
Begin VB.UserControl ButtonEx 
   AutoRedraw      =   -1  'True
   ClientHeight    =   2265
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3345
   DefaultCancel   =   -1  'True
   ScaleHeight     =   151
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   223
   ToolboxBitmap   =   "ButtonEx.ctx":0000
   Begin VB.PictureBox pictNewPicture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   3
      Top             =   1560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox pictTempHighlight 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1800
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   2
      Top             =   240
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox pictTempDestination 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   1800
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   1
      Top             =   960
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.PictureBox imgPicture 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   240
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   720
      Top             =   120
   End
End
Attribute VB_Name = "ButtonEx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'**************************************************************
'*  FILE:  ButtonEx.ctl                                       *
'*                                                            *
'*  DESCRIPTION:                                              *
'*      Provides a enhanced CommandButton control, including  *
'*      custom graphics as well MouseOver event, etc.         *
'*                                                            *
'*  CHANGE HISTORY:                                           *
'*      Aug 2000    J. Pearson      Initial code              *
'**************************************************************

'//---------------------------------------------------------------------------------------
'// Windows API constants
'//---------------------------------------------------------------------------------------
Private Const BLACKNESS = &H42              '(DWORD) dest = BLACK
Private Const NOTSRCCOPY = &H330008         '(DWORD) dest = (NOT source)
Private Const NOTSRCERASE = &H1100A6        '(DWORD) dest = (NOT src) AND (NOT dest)
Private Const SRCAND = &H8800C6             '(DWORD) dest = source AND dest
Private Const SRCCOPY = &HCC0020            '(DWORD) dest = source
Private Const SRCERASE = &H440328           '(DWORD) dest = source AND (NOT dest )
Private Const SRCINVERT = &H660046          '(DWORD) dest = source XOR dest
Private Const SRCPAINT = &HEE0086           '(DWORD) dest = source OR dest
Private Const WHITENESS = &HFF0062          '(DWORD) dest = WHITE

Private Const BDR_RAISEDINNER = &H4
Private Const BDR_RAISEDOUTER = &H1
Private Const BDR_SUNKENINNER = &H8
Private Const BDR_SUNKENOUTER = &H2

Private Const BDR_RAISED = &H5
Private Const BDR_OUTER = &H3
Private Const BDR_INNER = &HC

Private Const BF_ADJUST = &H2000        'Calculate the space left over.
Private Const BF_FLAT = &H4000          'For flat rather than 3-D borders.
Private Const BF_MONO = &H8000          'For monochrome borders.
Private Const BF_SOFT = &H1000          'Use for softer buttons.
Private Const BF_BOTTOM = &H8
Private Const BF_LEFT = &H1
Private Const BF_RIGHT = &H4
Private Const BF_TOP = &H2
Private Const BF_RECT = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)

Private Const EDGE_RAISED = (BDR_RAISEDOUTER Or BDR_RAISEDINNER)
Private Const EDGE_SUNKEN = (BDR_SUNKENOUTER Or BDR_SUNKENINNER)
Private Const EDGE_BUMP = (BDR_RAISEDOUTER Or BDR_SUNKENINNER)
Private Const EDGE_ETCHED = (BDR_SUNKENOUTER Or BDR_RAISEDINNER)

Private Const DT_CENTER = &H1
Private Const DT_RTLREADING = &H20000
Private Const DT_SINGLELINE = &H20
Private Const DT_VCENTER = &H4

Private Const DST_COMPLEX = &H0
Private Const DST_TEXT = &H1
Private Const DST_PREFIXTEXT = &H2
Private Const DST_ICON = &H3
Private Const DST_BITMAP = &H4

Private Const DSS_NORMAL = &H0
Private Const DSS_UNION = &H10                   '/* Gray string appearance */
Private Const DSS_DISABLED = &H20
Private Const DSS_RIGHT = &H8000

'//---------------------------------------------------------------------------------------
'// Windows API types
'//---------------------------------------------------------------------------------------
Private Type POINTAPI
    X As Long
    Y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

'//---------------------------------------------------------------------------------------
'// Windows API declarations
'//---------------------------------------------------------------------------------------
Private Declare Function BitBlt Lib "gdi32" (ByVal hdcDest As Long, ByVal xDest As Long, ByVal yDest As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hdcSrc As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateBitmap Lib "gdi32" (ByVal nWidth As Long, ByVal nHeight As Long, ByVal nPlanes As Long, ByVal nBitCount As Long, lpBits As Any) As Long
Private Declare Function CreateCompatibleBitmap Lib "gdi32" (ByVal hDC As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DrawEdge Lib "user32" (ByVal hDC As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hDC As Long, lpRect As RECT) As Long
Private Declare Function DrawState Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As Long, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
Private Declare Function DrawStateText Lib "user32" Alias "DrawStateA" (ByVal hDC As Long, ByVal hBrush As Long, ByVal lpDrawStateProc As Long, ByVal lParam As String, ByVal wParam As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal n3 As Long, ByVal n4 As Long, ByVal un As Long) As Long
Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hDC As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long

'//---------------------------------------------------------------------------------------
'// Private constants
'//---------------------------------------------------------------------------------------
Private Enum BorderTypeEnum
    btDown
    btUp
    btOver
End Enum

Private Enum RasterOperationConstants
    roNotSrcCopy = NOTSRCCOPY
    roNotSrcErase = NOTSRCERASE
    roSrcAnd = SRCAND
    roSrcCopy = SRCCOPY
    roSrcErase = SRCERASE
    roSrcInvert = SRCINVERT
    roSrcPaint = SRCPAINT
End Enum

'//---------------------------------------------------------------------------------------
'// Private constants
'//---------------------------------------------------------------------------------------
Private lState As BorderTypeEnum
Private bLeftFocus As Boolean
Private bHasFocus As Boolean

'//---------------------------------------------------------------------------------------
'// Public constants
'//---------------------------------------------------------------------------------------
Public Enum beAppearance
    Flat = 0
    [3D] = 1
End Enum

'//---------------------------------------------------------------------------------------
'// Control property constants
'//---------------------------------------------------------------------------------------
Private Const m_def_Appearance = [3D]
Private Const m_def_BackColor = vbButtonFace
Private Const m_def_Caption = "ButtonEx1"
Private Const m_def_Enabled = True
Private Const m_def_ForeColor = vbButtonText
Private Const m_def_HighlightColor = vbButtonText
Private Const m_def_HighlightPicture = False
Private Const m_def_MousePointer = vbDefault
Private Const m_def_RightToLeft = False
Private Const m_def_ToolTipText = ""
Private Const m_def_TransparentColor = vbBlue
Private Const m_def_WhatsThisHelpID = 0

'//---------------------------------------------------------------------------------------
'// Control property variables
'//---------------------------------------------------------------------------------------
Private m_Appearance As beAppearance
Private m_BackColor As OLE_COLOR
Private m_Caption As String
Private m_Enabled As Boolean
Private m_ForeColor As OLE_COLOR
Private m_Font As Font
Private m_HighlightColor As OLE_COLOR
Private m_HighlightPicture As Boolean
Private m_MouseIcon As Picture
Private m_MousePointer As MousePointerConstants
Private m_Picture As Picture
Private m_RightToLeft As Boolean
Private m_ToolTipText As String
Private m_TransparentColor As OLE_COLOR
Private m_WhatsThisHelpID As Long

'//---------------------------------------------------------------------------------------
'// Control property events
'//---------------------------------------------------------------------------------------
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Public Event KeyDown(KeyCode As Integer, Shift As Integer)
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Public Event KeyPress(KeyAscii As Integer)
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Public Event KeyUp(KeyCode As Integer, Shift As Integer)
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus."
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse."
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus."
Public Event Resize()
Attribute Resize.VB_Description = "Occurs when a form is first displayed or the size of an object changes."

'//---------------------------------------------------------------------------------------
'// Control properties
'//---------------------------------------------------------------------------------------

Public Property Get Appearance() As beAppearance
Attribute Appearance.VB_Description = "Returns/sets whether or not an object is painted with 3-D effects."
    Appearance = m_Appearance
End Property

Public Property Let Appearance(ByVal NewValue As beAppearance)
    m_Appearance = NewValue
        
    Call DrawButton(lState)
    
    PropertyChanged "Appearance"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = m_BackColor
End Property

Public Property Let BackColor(ByVal NewValue As OLE_COLOR)
    m_BackColor = NewValue
    UserControl.BackColor = NewValue
    imgPicture.BackColor = NewValue
    
    Call DrawButton(lState)
    
    PropertyChanged "BackColor"
End Property

Public Property Get Caption() As String
Attribute Caption.VB_Description = "Returns/sets the text displayed in an object."
    Caption = m_Caption
End Property

Public Property Let Caption(ByVal NewValue As String)
    Dim lPlace As Long
    
    m_Caption = NewValue
    
    'set access key
    lPlace = 0
    lPlace = InStr(lPlace + 1, NewValue, "&", vbTextCompare)
    Do While lPlace <> 0
        If Mid$(NewValue, lPlace + 1, 1) <> "&" Then
            UserControl.AccessKeys = Mid$(NewValue, lPlace + 1, 1)
            Exit Do
        Else
            lPlace = lPlace + 1
        End If
    
        lPlace = InStr(lPlace + 1, NewValue, "&", vbTextCompare)
    Loop
    
    Call DrawButton(lState)
    
    PropertyChanged "Caption"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal NewValue As Boolean)
    m_Enabled = NewValue
    UserControl.Enabled = NewValue
    imgPicture.Enabled = NewValue
    
    Call DrawButton(lState)
    
    PropertyChanged "Enabled"
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_ForeColor
End Property

Public Property Let ForeColor(ByVal NewValue As OLE_COLOR)
    m_ForeColor = NewValue
    UserControl.ForeColor = NewValue
    imgPicture.ForeColor = NewValue
    
    Call DrawButton(lState)
    
    PropertyChanged "ForeColor"
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns/sets a Font object used to display text in the object."
    Set Font = m_Font
End Property

Public Property Set Font(ByVal NewValue As Font)
    Set m_Font = NewValue
    Set UserControl.Font = NewValue
    Set imgPicture.Font = NewValue
    
    Call DrawButton(lState)
    
    PropertyChanged "Font"
End Property

Public Property Get HighlightColor() As OLE_COLOR
Attribute HighlightColor.VB_Description = "Returns/sets the highlight color used to display text and graphics when the  mouse is over the object"
    HighlightColor = m_HighlightColor
End Property

Public Property Let HighlightColor(ByVal NewValue As OLE_COLOR)
    m_HighlightColor = NewValue
    
    Call DrawButton(lState)
    
    PropertyChanged "HighlightColor"
End Property

Public Property Get HighlightPicture() As Boolean
Attribute HighlightPicture.VB_Description = "Returns/sets whether or not to highlight the object's picture with the HighlightColor."
    HighlightPicture = m_HighlightPicture
End Property

Public Property Let HighlightPicture(ByVal NewValue As Boolean)
    m_HighlightPicture = NewValue
    PropertyChanged "HighlightPicture"
End Property

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Returns/sets a custom mouse icon."
    Set MouseIcon = m_MouseIcon
End Property

Public Property Set MouseIcon(ByVal NewValue As Picture)
    Set m_MouseIcon = NewValue
    Set UserControl.MouseIcon = NewValue
    Set imgPicture.MouseIcon = NewValue
    
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = m_MousePointer
End Property

Public Property Let MousePointer(ByVal NewValue As MousePointerConstants)
    m_MousePointer = NewValue
    UserControl.MousePointer = NewValue
    imgPicture.MousePointer = NewValue
    
    PropertyChanged "MousePointer"
End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = m_Picture
End Property

Public Property Set Picture(ByVal NewValue As Picture)
    Set m_Picture = NewValue
    Set imgPicture.Picture = NewValue
    
    Call DrawButton(lState)
    
    PropertyChanged "Picture"
End Property

Public Property Get RightToLeft() As Boolean
Attribute RightToLeft.VB_Description = "Determines text display direction and control visual appearance on a bidirectional system."
    RightToLeft = m_RightToLeft
End Property

Public Property Let RightToLeft(ByVal NewValue As Boolean)
    m_RightToLeft = NewValue
    UserControl.RightToLeft = NewValue
    imgPicture.RightToLeft = NewValue
    
    Call DrawButton(lState)
    
    PropertyChanged "RightToLeft"
End Property

Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse cursor is over the control."
    ToolTipText = m_ToolTipText
End Property

Public Property Let ToolTipText(ByVal NewValue As String)
    m_ToolTipText = NewValue
    imgPicture.ToolTipText = NewValue
    
    PropertyChanged "ToolTipText"
End Property

Public Property Get TransparentColor() As OLE_COLOR
Attribute TransparentColor.VB_Description = "Returns/sets the color of the Picture property to make transparent."
    TransparentColor = m_TransparentColor
End Property

Public Property Let TransparentColor(ByVal NewValue As OLE_COLOR)
    m_TransparentColor = NewValue
    
    Call DrawButton(lState)
    
    PropertyChanged "TransparentColor"
End Property

Public Property Get WhatsThisHelpID() As Long
Attribute WhatsThisHelpID.VB_Description = "Returns/sets an associated help context ID for the control."
    WhatsThisHelpID = m_WhatsThisHelpID
End Property

Public Property Let WhatsThisHelpID(ByVal NewValue As Long)
    m_WhatsThisHelpID = NewValue
    imgPicture.WhatsThisHelpID = NewValue
    
    PropertyChanged "WhatsThisHelpID"
End Property

'//---------------------------------------------------------------------------------------
'// Image functions
'//---------------------------------------------------------------------------------------

Private Sub imgPicture_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseDown(Button, Shift, imgPicture.Left + (X \ Screen.TwipsPerPixelX), imgPicture.Top + (Y \ Screen.TwipsPerPixelY))
End Sub

Private Sub imgPicture_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseMove(Button, Shift, imgPicture.Left + (X \ Screen.TwipsPerPixelX), imgPicture.Top + (Y \ Screen.TwipsPerPixelY))
End Sub

Private Sub imgPicture_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call UserControl_MouseUp(Button, Shift, imgPicture.Left + (X \ Screen.TwipsPerPixelX), imgPicture.Top + (Y \ Screen.TwipsPerPixelY))
End Sub

'//---------------------------------------------------------------------------------------
'// Timer functions
'//---------------------------------------------------------------------------------------

Private Sub Timer1_Timer()
    'check for mouse leaving control
    Dim pnt As POINTAPI
    
    GetCursorPos pnt
    ScreenToClient UserControl.hWnd, pnt
    
    If pnt.X < UserControl.ScaleLeft Or _
            pnt.Y < UserControl.ScaleTop Or _
            pnt.X > (UserControl.ScaleLeft + UserControl.ScaleWidth) Or _
            pnt.Y > (UserControl.ScaleTop + UserControl.ScaleHeight) Then
        Timer1.Enabled = False
    
        'left focus
        If lState <> btUp Then
            Call DrawButton(btUp)
        End If
        bLeftFocus = True
    Else
        'gained focus
        If bLeftFocus Then
            Call DrawButton(btDown)
        End If
    End If
End Sub

'//---------------------------------------------------------------------------------------
'// UserControl functions
'//---------------------------------------------------------------------------------------

Private Sub UserControl_InitProperties()
    'Initialize Properties for User Control
    Appearance = m_def_Appearance
    BackColor = m_def_BackColor
    Caption = m_def_Caption
    Enabled = m_def_Enabled
    ForeColor = m_def_ForeColor
    Set Font = Ambient.Font
    HighlightColor = m_def_HighlightColor
    HighlightPicture = m_def_HighlightPicture
    Set MouseIcon = LoadPicture("")
    MousePointer = m_def_MousePointer
    Set Picture = LoadPicture("")
    RightToLeft = m_def_RightToLeft
    ToolTipText = m_def_ToolTipText
    TransparentColor = m_def_TransparentColor
    WhatsThisHelpID = m_def_WhatsThisHelpID
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    'Load property values from storage
    Appearance = PropBag.ReadProperty("Appearance", m_def_Appearance)
    BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    Caption = PropBag.ReadProperty("Caption", m_def_Caption)
    Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    ForeColor = PropBag.ReadProperty("ForeColor", m_def_ForeColor)
    Set Font = PropBag.ReadProperty("Font", Ambient.Font)
    HighlightColor = PropBag.ReadProperty("HighlightColor", m_def_HighlightColor)
    HighlightPicture = PropBag.ReadProperty("HighlightPicture", m_def_HighlightPicture)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    MousePointer = PropBag.ReadProperty("MousePointer", m_def_MousePointer)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    RightToLeft = PropBag.ReadProperty("RightToLeft", m_def_RightToLeft)
    ToolTipText = PropBag.ReadProperty("ToolTipText", m_def_ToolTipText)
    TransparentColor = PropBag.ReadProperty("TransparentColor", m_def_TransparentColor)
    WhatsThisHelpID = PropBag.ReadProperty("WhatsThisHelpID", m_def_WhatsThisHelpID)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    'Write property values to storage
    Call PropBag.WriteProperty("Appearance", m_Appearance, m_def_Appearance)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
    Call PropBag.WriteProperty("Caption", m_Caption, m_def_Caption)
    Call PropBag.WriteProperty("ForeColor", m_ForeColor, m_def_ForeColor)
    Call PropBag.WriteProperty("Font", m_Font, Ambient.Font)
    Call PropBag.WriteProperty("HighlightColor", m_HighlightColor, m_def_HighlightColor)
    Call PropBag.WriteProperty("HighlightPicture", m_HighlightPicture, m_def_HighlightPicture)
    Call PropBag.WriteProperty("Picture", m_Picture, Nothing)
    Call PropBag.WriteProperty("RightToLeft", m_RightToLeft, m_def_RightToLeft)
    Call PropBag.WriteProperty("TransparentColor", m_TransparentColor, m_def_TransparentColor)
    Call PropBag.WriteProperty("MouseIcon", m_MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", m_MousePointer, m_def_MousePointer)
    Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, m_def_ToolTipText)
    Call PropBag.WriteProperty("WhatsThisHelpID", m_WhatsThisHelpID, m_def_WhatsThisHelpID)
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    RaiseEvent Click
End Sub

Private Sub UserControl_AmbientChanged(PropertyName As String)
    If PropertyName = "DisplayAsDefault" Then
        If UserControl.Ambient.DisplayAsDefault Then
            bHasFocus = True
        Else
            bHasFocus = False
        End If
        Call DrawButton(lState)
    End If
End Sub

Private Sub UserControl_Initialize()
    'note: this really sets to 1215x375
    UserControl.Width = 1200
    UserControl.Height = 360
End Sub

Private Sub UserControl_GotFocus()
    bHasFocus = True
    Call DrawButton(lState)
End Sub

Private Sub UserControl_LostFocus()
    bHasFocus = False
    Call DrawButton(lState)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bLeftFocus = False
    
    If Button = vbLeftButton Then
        Call DrawButton(btDown)
    End If
    
    RaiseEvent MouseDown(Button, Shift, X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bLeftFocus = False
    
    If UserControl.Ambient.UserMode = True And Not Timer1.Enabled Then
        'start tracking
        Timer1.Enabled = True
    
    ElseIf Button = 0 Then
        'mouse over (for flat button)
        If lState <> btOver Then
            Call DrawButton(btOver)
        End If

    ElseIf Button = vbLeftButton Then
        If lState <> btDown Then
            Call DrawButton(btDown)
        End If
    End If

    RaiseEvent MouseMove(Button, Shift, X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    bLeftFocus = False
    
    If Button = vbLeftButton Then
        Call DrawButton(btUp)
    End If

    RaiseEvent MouseUp(Button, Shift, X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY)
End Sub

Private Sub UserControl_Resize()
    Call DrawButton(btUp)
    RaiseEvent Resize
End Sub

'//---------------------------------------------------------------------------------------
'// Private functions
'//---------------------------------------------------------------------------------------

Private Sub TransparentBlt_New2(ByVal hDC As Long, ByVal Source As PictureBox, ByRef DestPoint As POINTAPI, ByRef SrcPoint As POINTAPI, ByVal Width As Long, ByVal Height As Long, Optional ByVal TransparentColor As OLE_COLOR = -1, Optional ByVal Clear As Boolean = False, Optional ByVal Resize As Boolean = False, Optional ByVal Refresh As Boolean = False)
    Dim MonoMaskDC As Long
    Dim hMonoMask As Long
    Dim MonoInvDC As Long
    Dim hMonoInv As Long
    Dim ResultDstDC As Long
    Dim hResultDst As Long
    Dim ResultSrcDC As Long
    Dim hResultSrc As Long
    Dim hPrevMask As Long
    Dim hPrevInv As Long
    Dim hPrevSrc As Long
    Dim hPrevDst As Long
    Dim OldBC As Long
    
    If TransparentColor = -1 Then
        TransparentColor = GetPixel(Source.hDC, 1, 1)
    End If
    
    'create monochrome mask and inverse masks
    MonoMaskDC = CreateCompatibleDC(hDC)
    MonoInvDC = CreateCompatibleDC(hDC)
    hMonoMask = CreateBitmap(Width, Height, 1, 1, ByVal 0&)
    hMonoInv = CreateBitmap(Width, Height, 1, 1, ByVal 0&)
    hPrevMask = SelectObject(MonoMaskDC, hMonoMask)
    hPrevInv = SelectObject(MonoInvDC, hMonoInv)
    
    'create keeper DCs and bitmaps
    ResultDstDC = CreateCompatibleDC(hDC)
    ResultSrcDC = CreateCompatibleDC(hDC)
    hResultDst = CreateCompatibleBitmap(hDC, Width, Height)
    hResultSrc = CreateCompatibleBitmap(hDC, Width, Height)
    hPrevDst = SelectObject(ResultDstDC, hResultDst)
    hPrevSrc = SelectObject(ResultSrcDC, hResultSrc)
    
    'copy src to monochrome mask
    OldBC = SetBkColor(Source.hDC, TransparentColor)
    Call BitBlt(MonoMaskDC, 0, 0, Width, Height, Source.hDC, SrcPoint.X, SrcPoint.Y, SRCCOPY)
    TransparentColor = SetBkColor(Source.hDC, OldBC)
    
    'create inverse of mask
    Call BitBlt(MonoInvDC, 0, 0, Width, Height, MonoMaskDC, 0, 0, NOTSRCCOPY)
    
    'get background
    Call BitBlt(ResultDstDC, 0, 0, Width, Height, hDC, DestPoint.X, DestPoint.Y, SRCCOPY)
    
    'AND with Monochrome mask
    Call BitBlt(ResultDstDC, 0, 0, Width, Height, MonoMaskDC, 0, 0, SRCAND)
    
    'get overlapper
    Call BitBlt(ResultSrcDC, 0, 0, Width, Height, Source.hDC, SrcPoint.X, SrcPoint.Y, SRCCOPY)
    
    'AND with inverse monochrome mask
    Call BitBlt(ResultSrcDC, 0, 0, Width, Height, MonoInvDC, 0, 0, SRCAND)
    
    'XOR these two
    Call BitBlt(ResultDstDC, 0, 0, Width, Height, ResultSrcDC, 0, 0, SRCINVERT)
    
    'output results
    Call BitBlt(hDC, DestPoint.X, DestPoint.Y, Width, Height, ResultDstDC, 0, 0, SRCCOPY)
    
    'clean up
    hMonoMask = SelectObject(MonoMaskDC, hPrevMask)
    DeleteObject hMonoMask
    
    hMonoInv = SelectObject(MonoInvDC, hPrevInv)
    DeleteObject hMonoInv
    
    hResultDst = SelectObject(ResultDstDC, hPrevDst)
    DeleteObject hResultDst
    
    hResultSrc = SelectObject(ResultSrcDC, hPrevSrc)
    DeleteObject hResultSrc
    
    DeleteDC MonoMaskDC
    DeleteDC MonoInvDC
    DeleteDC ResultDstDC
    DeleteDC ResultSrcDC
End Sub

Private Sub DrawButton(ByVal BorderType As BorderTypeEnum)
    'draw button around button
    Const clTop As Long = 6
    Const clLeft As Long = 6
    Const clFocusOffset As Long = 4
    Const clDownOffset As Long = 1
    
    Dim rct As RECT
    Dim bFocus As Boolean
    Dim lFormat As Long
    Dim lLeft As Long
    Dim lTop As Long
    Dim lPrevColor As OLE_COLOR
    Dim bUserMode As Boolean
    
    'clear control
    UserControl.Cls
    
    'initialize variable
    bFocus = bHasFocus
    bUserMode = False
    
    'get user mode
    On Local Error Resume Next
    bUserMode = UserControl.Ambient.UserMode
    On Local Error GoTo 0
    
    'get rect
    With rct
        .Left = 0
        .Top = 0
        .Bottom = UserControl.ScaleHeight
        .Right = UserControl.ScaleWidth
    End With
    
    Select Case BorderType
        Case btUp
            If m_Appearance = [3D] Then
                'draw raised border
                If bFocus Then
                    Call DrawEdge(UserControl.hDC, rct, BDR_OUTER, BF_RECT Or BF_ADJUST Or BF_MONO)
                    Call DrawEdge(UserControl.hDC, rct, EDGE_RAISED, BF_RECT)
                Else
                    Call DrawEdge(UserControl.hDC, rct, EDGE_RAISED, BF_RECT)
                End If
            Else
                bFocus = False
            End If
        
        Case btOver
            'draw raised border
            If bFocus Then
                Call DrawEdge(UserControl.hDC, rct, BDR_OUTER, BF_RECT Or BF_ADJUST Or BF_MONO)
                Call DrawEdge(UserControl.hDC, rct, EDGE_RAISED, BF_RECT)
            Else
                Call DrawEdge(UserControl.hDC, rct, EDGE_RAISED, BF_RECT)
            End If
            UserControl.ForeColor = m_HighlightColor
        
        Case btDown
            'draw sunken border
            If bFocus Then
                Call DrawEdge(UserControl.hDC, rct, BDR_OUTER, BF_RECT Or BF_ADJUST Or BF_MONO)
                Call DrawEdge(UserControl.hDC, rct, BDR_SUNKENOUTER, BF_RECT Or BF_FLAT)
            Else
                Call DrawEdge(UserControl.hDC, rct, EDGE_SUNKEN, BF_RECT)
            End If
            UserControl.ForeColor = m_HighlightColor
    End Select

    'calculate caption position
    If imgPicture.Picture <> 0 Then
        lLeft = imgPicture.Left + imgPicture.Width - clLeft
    End If
    
    lLeft = lLeft \ 2 + ((UserControl.ScaleWidth \ 2) - (UserControl.TextWidth(m_Caption) \ 2))
    lTop = (UserControl.ScaleHeight \ 2) - (UserControl.TextHeight(m_Caption) \ 2)
    
    If BorderType = btDown Then
        lLeft = lLeft + clDownOffset
        lTop = lTop + clDownOffset
    End If
    
    'draw caption in button
    lFormat = DST_PREFIXTEXT Or DSS_NORMAL
    If Not m_Enabled Then
        lFormat = lFormat Or DSS_DISABLED
    End If
    If m_RightToLeft Then
        lFormat = lFormat Or DSS_RIGHT
    End If
    
    Call DrawStateText(UserControl.hDC, 0, 0, m_Caption, Len(m_Caption), lLeft, lTop, 0, 0, lFormat)

    If bUserMode Then
        If bFocus Then
            'draw focus rect
            With rct
                .Left = clFocusOffset
                .Top = clFocusOffset
                .Bottom = UserControl.ScaleHeight - clFocusOffset
                .Right = UserControl.ScaleWidth - clFocusOffset
            End With
            lPrevColor = UserControl.ForeColor
            UserControl.ForeColor = vbBlack
            Call DrawFocusRect(UserControl.hDC, rct)
            UserControl.ForeColor = lPrevColor
        End If
    End If

    'move image
    With imgPicture
        If .Picture <> 0 Then
            lLeft = clLeft
            lTop = (UserControl.ScaleHeight \ 2) - (.Height \ 2)
            If lTop < clTop Then
                lTop = clTop
            End If
            
            If BorderType = btDown Then
                lLeft = lLeft + clDownOffset
                lTop = lTop + clDownOffset
            End If
        
            If .Left <> lLeft Then
                .Left = lLeft
            End If
            If .Top <> lTop Then
                .Top = lTop
            End If
        
            Dim ptDest As POINTAPI
            Dim ptSrc As POINTAPI
            
            ptDest.X = .Left
            ptDest.Y = .Top
            ptSrc.X = 0
            ptSrc.Y = 0
            
            pictNewPicture.Cls
            If (BorderType = btDown Or BorderType = btOver Or (Not m_Enabled And BorderType = btUp)) And m_HighlightPicture = True Then
                If m_Enabled Then
                    Call HighlightBltEx(imgPicture, pictNewPicture, pictTempDestination, pictTempHighlight, m_HighlightColor, 0, 0, 0, 0, .Width, .Height)
                Else
                    Call HighlightBltEx(imgPicture, pictNewPicture, pictTempDestination, pictTempHighlight, vbGrayText, 0, 0, 0, 0, .Width, .Height)
                End If
                Call TransparentBlt_New2(UserControl.hDC, pictNewPicture, ptDest, ptSrc, imgPicture.Width, imgPicture.Height, pictNewPicture.BackColor)
            Else
                Call TransparentBlt_New2(UserControl.hDC, imgPicture, ptDest, ptSrc, imgPicture.Width, imgPicture.Height, m_TransparentColor)
            End If
        End If
    End With
    
    'set state
    lState = BorderType
    
    'reset forecolor
    UserControl.ForeColor = m_ForeColor
End Sub

Private Function BitBltEx(ByVal Source As Object, ByVal Destination As Object, ByVal Operation As RasterOperationConstants, Optional ByVal xDest As Long = 0, Optional ByVal yDest As Long = 0, Optional ByVal XSrc As Long = 0, Optional ByVal YSrc As Long = 0, Optional ByVal Width As Long = -1, Optional ByVal Height As Long = -1, Optional ByVal Refresh As Boolean = False) As Boolean
    Dim lReturn As Long
    
    If Width = -1 Then
        Width = Source.Width \ Screen.TwipsPerPixelX
    End If
    If Height = -1 Then
        Height = Source.Height \ Screen.TwipsPerPixelX
    End If
    
    'BitBlt
    lReturn = BitBlt(Destination.hDC, xDest, yDest, Width, Height, Source.hDC, XSrc, YSrc, Operation)
    
    If Refresh Then
        'refresh destination
        Destination.Refresh
    End If
    
    'return result
    If lReturn = 0 Then
        BitBltEx = False
    Else
        BitBltEx = True
    End If
End Function

Private Function MaskBltEx(ByVal Source As Object, ByVal Destination As Object, Optional ByVal MaskColor As OLE_COLOR = -1, Optional ByVal xDest As Long = 0, Optional ByVal yDest As Long = 0, Optional ByVal XSrc As Long = 0, Optional ByVal YSrc As Long = 0, Optional ByVal Width As Long = -1, Optional ByVal Height As Long = -1, Optional ByVal Refresh As Boolean = False) As Boolean
    Dim MonoMaskDC As Long
    Dim hMonoMask As Long
    Dim MonoInvDC As Long
    Dim hMonoInv As Long
    Dim ResultDstDC As Long
    Dim hResultDst As Long
    Dim ResultSrcDC As Long
    Dim hResultSrc As Long
    Dim hPrevMask As Long
    Dim hPrevInv As Long
    Dim hPrevSrc As Long
    Dim hPrevDst As Long
    Dim OldBC As Long
    Dim lReturn As Long
    
    If Width = -1 Then
        Width = Source.Width \ Screen.TwipsPerPixelX
    End If
    If Height = -1 Then
        Height = Source.Height \ Screen.TwipsPerPixelX
    End If
    
    If MaskColor = -1 Then
        MaskColor = GetPixel(Source.hDC, 0, 0)
    End If
    
    'create monochrome mask and inverse masks
    MonoMaskDC = CreateCompatibleDC(Destination.hDC)
    hMonoMask = CreateBitmap(Width, Height, 1, 1, ByVal 0&)
    hPrevMask = SelectObject(MonoMaskDC, hMonoMask)
    
    'copy src to monochrome mask
    OldBC = SetBkColor(Source.hDC, MaskColor)
    lReturn = BitBlt(MonoMaskDC, 0, 0, Width, Height, Source.hDC, XSrc, YSrc, SRCCOPY)
    If lReturn <> 0 Then
        MaskColor = SetBkColor(Source.hDC, OldBC)
        
        'output results
        lReturn = BitBlt(Destination.hDC, xDest, yDest, Width, Height, MonoMaskDC, 0, 0, SRCCOPY)
    End If
    
    'clean up
    hMonoMask = SelectObject(MonoMaskDC, hPrevMask)
    DeleteObject hMonoMask
    DeleteDC MonoMaskDC

    If Refresh Then
        'refresh destination
        Destination.Refresh
    End If
    
    'return result
    If lReturn = 0 Then
        MaskBltEx = False
    Else
        MaskBltEx = True
    End If
End Function

Private Function TransparentBltEx(ByVal Source As Object, ByVal Destination As Object, Optional ByVal TransparentColor As OLE_COLOR = -1, Optional ByVal xDest As Long = 0, Optional ByVal yDest As Long = 0, Optional ByVal XSrc As Long = 0, Optional ByVal YSrc As Long = 0, Optional ByVal Width As Long = -1, Optional ByVal Height As Long = -1, Optional ByVal Refresh As Boolean = False) As Boolean
    Dim MonoMaskDC As Long
    Dim hMonoMask As Long
    Dim MonoInvDC As Long
    Dim hMonoInv As Long
    Dim ResultDstDC As Long
    Dim hResultDst As Long
    Dim ResultSrcDC As Long
    Dim hResultSrc As Long
    Dim hPrevMask As Long
    Dim hPrevInv As Long
    Dim hPrevSrc As Long
    Dim hPrevDst As Long
    Dim OldBC As Long
    Dim lReturn As Long
    
    If Width = -1 Then
        Width = Source.Width \ Screen.TwipsPerPixelX
    End If
    If Height = -1 Then
        Height = Source.Height \ Screen.TwipsPerPixelX
    End If
    
    If TransparentColor = -1 Then
        TransparentColor = GetPixel(Source.hDC, 0, 0)
    End If
    
    'create monochrome mask and inverse masks
    MonoMaskDC = CreateCompatibleDC(Destination.hDC)
    MonoInvDC = CreateCompatibleDC(Destination.hDC)
    hMonoMask = CreateBitmap(Width, Height, 1, 1, ByVal 0&)
    hMonoInv = CreateBitmap(Width, Height, 1, 1, ByVal 0&)
    hPrevMask = SelectObject(MonoMaskDC, hMonoMask)
    hPrevInv = SelectObject(MonoInvDC, hMonoInv)
    
    'create keeper DCs and bitmaps
    ResultDstDC = CreateCompatibleDC(Destination.hDC)
    ResultSrcDC = CreateCompatibleDC(Destination.hDC)
    hResultDst = CreateCompatibleBitmap(Destination.hDC, Width, Height)
    hResultSrc = CreateCompatibleBitmap(Destination.hDC, Width, Height)
    hPrevDst = SelectObject(ResultDstDC, hResultDst)
    hPrevSrc = SelectObject(ResultSrcDC, hResultSrc)
    
    'copy src to monochrome mask
    OldBC = SetBkColor(Source.hDC, TransparentColor)
    lReturn = BitBlt(MonoMaskDC, 0, 0, Width, Height, Source.hDC, XSrc, YSrc, SRCCOPY)
    If lReturn <> 0 Then
        TransparentColor = SetBkColor(Source.hDC, OldBC)
        
        'create inverse of mask
        lReturn = BitBlt(MonoInvDC, 0, 0, Width, Height, MonoMaskDC, 0, 0, NOTSRCCOPY)
        If lReturn <> 0 Then
            'get background
            lReturn = BitBlt(ResultDstDC, 0, 0, Width, Height, Destination.hDC, xDest, yDest, SRCCOPY)
            If lReturn <> 0 Then
                'AND with Monochrome mask
                lReturn = BitBlt(ResultDstDC, 0, 0, Width, Height, MonoMaskDC, 0, 0, SRCAND)
                If lReturn <> 0 Then
                    'get overlapper
                    lReturn = BitBlt(ResultSrcDC, 0, 0, Width, Height, Source.hDC, XSrc, YSrc, SRCCOPY)
                    If lReturn <> 0 Then
                        'AND with inverse monochrome mask
                        lReturn = BitBlt(ResultSrcDC, 0, 0, Width, Height, MonoInvDC, 0, 0, SRCAND)
                        If lReturn <> 0 Then
                            'XOR these two
                            lReturn = BitBlt(ResultDstDC, 0, 0, Width, Height, ResultSrcDC, 0, 0, SRCINVERT)
                            If lReturn <> 0 Then
                                'output results
                                lReturn = BitBlt(Destination.hDC, xDest, yDest, Width, Height, ResultDstDC, 0, 0, SRCCOPY)
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    'clean up
    hMonoMask = SelectObject(MonoMaskDC, hPrevMask)
    DeleteObject hMonoMask
    
    hMonoInv = SelectObject(MonoInvDC, hPrevInv)
    DeleteObject hMonoInv
    
    hResultDst = SelectObject(ResultDstDC, hPrevDst)
    DeleteObject hResultDst
    
    hResultSrc = SelectObject(ResultSrcDC, hPrevSrc)
    DeleteObject hResultSrc
    
    DeleteDC MonoMaskDC
    DeleteDC MonoInvDC
    DeleteDC ResultDstDC
    DeleteDC ResultSrcDC

    If Refresh Then
        'refresh destination
        Destination.Refresh
    End If
    
    'return result
    If lReturn = 0 Then
        TransparentBltEx = False
    Else
        TransparentBltEx = True
    End If
End Function

Private Function HighlightBltEx(ByVal Source As Object, ByVal Destination As Object, ByVal TempDestination As Object, ByVal Highlight As Object, ByVal HighlightColor As OLE_COLOR, Optional ByVal xDest As Long = 0, Optional ByVal yDest As Long = 0, Optional ByVal XSrc As Long = 0, Optional ByVal YSrc As Long = 0, Optional ByVal Width As Long = -1, Optional ByVal Height As Long = -1, Optional ByVal Refresh As Boolean = False) As Boolean
    Highlight.BackColor = HighlightColor
    
    Call MaskBltEx(Source, TempDestination, -1, 0, 0, XSrc, YSrc, Width, Height)
    Call BitBltEx(TempDestination, Highlight, roSrcInvert, 0, 0, 0, 0, Width, Height)
    Call TransparentBltEx(Highlight, Destination, -1, xDest, yDest, 0, 0, Width, Height, Refresh)
End Function

