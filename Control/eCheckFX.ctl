VERSION 5.00
Begin VB.UserControl eCheckFX 
   AutoRedraw      =   -1  'True
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1215
   PropertyPages   =   "eCheckFX.ctx":0000
   ScaleHeight     =   33
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   81
   ToolboxBitmap   =   "eCheckFX.ctx":0045
   Begin VB.PictureBox PicTmp 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   150
      Left            =   900
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   1
      Top             =   285
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.PictureBox PicCheck 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   195
      Left            =   105
      ScaleHeight     =   13
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   13
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   195
   End
End
Attribute VB_Name = "eCheckFX"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "PropPageWizardRun" ,"Yes"
' DM eCheckFX V1
' This is a small replacement for the normal checkbox control in Visual Basic 5 or 6
' The checkbox ActiveX uses mostley Pure VBcode.
' So why make a new checkbox. well basicly I wanted something with a little more style
' and so the user was not restriced to the default style.
' well the control does still use the old default style but with some extra featues
' Please note I no some people are aware that this project may run a little slow.
' but this project was not ment to show how fas to make something in VB. but to make use of it;s own built in functions.
' without reverting to API code. tho in one or two cases I have had to use API. but maniy for extra features such as the Hottracking and Custom Bitmaps
' the rest is pure VB code.


' some of the new features you can find below:
' ----------------------------------------------------------------------------------------------
' Version 1
' ----------------------------------------------------------------------------------------------
' 3 Checkbox styles Old VB Default Style, Modem 3DStyle and a Flat Style
' Support to change check color
' Turn on or off FocusRect
' Change check value with space bar when WHEN FocusRect is enabled
' Change the background color or the checkbox
' Change the foreground color of the caption
' Show check enabled or disbaled
' Alignment options Left and Right.
' Change the checkbox back color. You know the little box

' ----------------------------------------------------------------------------------------------
' Version 2
' ----------------------------------------------------------------------------------------------
' Bug fixes and new features
' ----------------------------------------------------------------------------------------------
' Fixed bug with Focus rect now showing first time.
' Added some new Check styles
' Added feature to add underscore to caption using & sign
' Added checkdrawwidth property
' Added access Key support. note only takes effect if focusrect is enabled for the check
' Added two more box styles, ButtonRaised and Button Lowerd
' Added property to chnage Check Thickness
' Added Hottracking, Feature to chnage hottracking color
' Added feature to support bitmaps
' updated example project with all the new features.
' anyway hope you like it and learn something from it.


' As always if you like to use this in your own project then your are free to do so.
' if you chould add my name dreamvb or ben I whould be very happey
' Thanks

'Written by Ben Jones
'Questions and answers bug, problums,ides:

' if you do get a chnage come and vist my forum, tho it is a little quit at the moment since it only been up 4 days :)

'www.eraystudios.com/forum

Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long

Private m_Backcolor As OLE_COLOR
Private m_CheckColor As OLE_COLOR
Private m_CheckBoxColor As OLE_COLOR
Private m_TempForeColor As OLE_COLOR
Private m_FlatBorderColor As OLE_COLOR
Private m_HotTrackingColor As OLE_COLOR

Private m_checkDrawWidth As Integer
Private m_BoxStyle As nBoxStyle
Private m_Align As nAlign
Private m_CheckStyle As nCheckStyle

Private m_CheckCaption As String
Private m_Checked As Boolean
Private m_ShowRect As Boolean
Private m_Enabled As Boolean
Private m_hottrack As Boolean

Private m_TmpForeColor As OLE_COLOR
Private m_StyleEx As StyleEx

Enum nBoxStyle
    vbDefault
    New3D = 1
    FlatStyle = 2
    ButtonRaised = 3
    ButtonLowerd = 4
    DoubleBorder = 5
    Custom = 6
End Enum

Enum nAlign
    Left = 0
    Right = 1
End Enum

Enum HoverEvt
    HoverIn = 1
    HoverOut = 2
End Enum

Enum nCheckStyle
    Check = 0
    SqrSolid = 1
    SqrOutLine = 2
    CircleSoild = 3
    CircleOutLine = 4
    Grid = 5
    none = 6
End Enum

Enum StyleEx
    OverDrawn = 0
    CustBitmap = 1
End Enum

Private mDrawFocus As Boolean
Private CustomBorders(7) As Long

Event CheckClick()
'Event Declarations:
Event MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Event MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single, HoverEvent As HoverEvt)
Event MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

Function AddCustomBorders(ParamArray CustColors() As Variant)
    'Sub used for custom colors of the checkbox
    For x = 0 To UBound(CustColors)
        CustomBorders(x) = CustColors(x) 'Copy CustColors contents to CustomBorders
    Next
    x = 0
    
    Call DrawBox(m_BoxStyle) 'Must update the check box
    
End Function

Sub FixCaption(lpCaption As String)
Dim c As String * 1, x As Integer
    'This sub is used to print an underscore at the bottom of the text
 
    If Len(lpCaption) = 0 Then Exit Sub 'exit if no length is found
    
    For x = 1 To Len(lpCaption)
        c = Mid(lpCaption, x, 1) 'Get a char from the string
        
        If Mid(lpCaption, x, 1) = "&" Then
            UserControl.FontUnderline = True 'turn on underline
            UserControl.AccessKeys = Mid(lpCaption, x + 1, 1)
        Else
            UserControl.Print c; 'print the char
            UserControl.FontUnderline = False 'turn off underline
        End If
    Next
    
    x = 0
    c = ""
    
End Sub

Private Sub DrawBox(Optional Style As nBoxStyle = vbDefault)
Dim LineColors(8) As OLE_COLOR, mBoxTmpC As OLE_COLOR
Dim Check_Width As Integer, xTextPos As Integer, mTmpBorderC As OLE_COLOR, nForeCol As OLE_COLOR
Dim x As Integer, y As Integer
    
    Check_Width = 13
    PicCheck.Cls
    UserControl.Cls
    
    If m_Enabled Then
        LineColors(8) = m_CheckColor 'Check color
        PicCheck.BackColor = m_CheckBoxColor 'Checkbox color
        mTmpBorderC = m_FlatBorderColor
        mBoxTmpC = BoxColor
    Else
        LineColors(8) = vb3DShadow
        PicCheck.BackColor = vbButtonFace
        mTmpBorderC = vb3DShadow
        mBoxTmpC = vb3DShadow
    End If
    
    'Check box styles
    Select Case Style
        Case vbDefault 'Old default VB Look
            LineColors(0) = vb3DShadow: LineColors(1) = vb3DDKShadow
            LineColors(2) = vb3DShadow: LineColors(3) = vb3DDKShadow
            LineColors(4) = vbButtonFace: LineColors(5) = BoxColor
            LineColors(6) = vbButtonFace: LineColors(7) = BoxColor
        Case New3D 'Like the VB Default one just a little flater
            LineColors(0) = vb3DShadow: LineColors(1) = mBoxTmpC
            LineColors(2) = vb3DShadow: LineColors(3) = mBoxTmpC
            LineColors(4) = vbButtonFace: LineColors(5) = vbWhite
            LineColors(6) = vbButtonFace: LineColors(7) = vbWhite
        Case FlatStyle 'Flat Style using user definded colors
            'PicCheck.DrawStyle = 2
            LineColors(0) = mTmpBorderC: LineColors(1) = mBoxTmpC
            LineColors(2) = mTmpBorderC: LineColors(3) = mBoxTmpC
            LineColors(4) = mBoxTmpC: LineColors(5) = mTmpBorderC
            LineColors(6) = mBoxTmpC: LineColors(7) = mTmpBorderC
        Case ButtonRaised
            LineColors(0) = vbWhite: LineColors(1) = mBoxTmpC
            LineColors(2) = vbWhite: LineColors(3) = mBoxTmpC
            LineColors(4) = vb3DShadow: LineColors(5) = mBoxTmpC
            LineColors(6) = vb3DShadow: LineColors(7) = mBoxTmpC
        Case ButtonLowerd
            LineColors(0) = vb3DShadow: LineColors(1) = mBoxTmpC
            LineColors(2) = vb3DShadow: LineColors(3) = mBoxTmpC
            LineColors(4) = vbWhite: LineColors(5) = mBoxTmpC
            LineColors(6) = vbWhite: LineColors(7) = mBoxTmpC
        Case DoubleBorder
            LineColors(0) = mTmpBorderC: LineColors(1) = mTmpBorderC
            LineColors(2) = mTmpBorderC: LineColors(3) = mTmpBorderC
            LineColors(4) = mTmpBorderC: LineColors(5) = mTmpBorderC
            LineColors(6) = mTmpBorderC: LineColors(7) = mTmpBorderC
        Case Custom
            'CustomBorders
            LineColors(0) = CustomBorders(0): LineColors(1) = CustomBorders(1)
            LineColors(2) = CustomBorders(2): LineColors(3) = CustomBorders(3)
            LineColors(4) = CustomBorders(4): LineColors(5) = CustomBorders(5)
            LineColors(6) = CustomBorders(6): LineColors(7) = CustomBorders(7)
    End Select
    
    'This block of code deals with the drawing of the Check boxes style
    PicCheck.Line (0, 0)-(Check_Width - 1, 0), LineColors(0) 'Top-line1
    PicCheck.Line (1, 1)-(Check_Width - 2, 1), LineColors(1) 'Top-line2
    
    PicCheck.Line (0, Check_Width - 2)-(0, 0), LineColors(2) 'Left-line1
    PicCheck.Line (1, Check_Width - 2)-(1, 1), LineColors(3) 'Left-line2
    
    PicCheck.Line (Check_Width - 2, 1)-(Check_Width - 2, Check_Width - 1), LineColors(4) 'right-line1
    PicCheck.Line (Check_Width - 1, 0)-(Check_Width - 1, Check_Width), LineColors(5) 'Right-Line 2
    PicCheck.Line (1, Check_Width - 2)-(Check_Width - 1, Check_Width - 2), LineColors(6) 'bottom-line1
    PicCheck.Line (0, Check_Width - 1)-(Check_Width - 1, Check_Width - 1), LineColors(7) 'Bottom-Line2
    
    PicCheck.Top = (UserControl.ScaleHeight - UserControl.TextHeight(m_CheckCaption)) \ 2
    
    'Alignment for the checkbox
    If m_Align = Left Then
        PicCheck.Left = 1
        xTextPos = PicCheck.ScaleWidth + 5
    ElseIf m_Align = Right Then
        PicCheck.Left = (UserControl.ScaleWidth - Check_Width)
        xTextPos = 2
    End If

    w = (UserControl.TextWidth(m_CheckCaption) * 2) \ 2 + xTextPos
    yTextPos = PicCheck.Top
    mTextHeight = UserControl.TextHeight(m_CheckCaption)
    
    'If show rect is enabled and we have focus show the focus rect around the caption
    If (m_ShowRect) And (mDrawFocus = True) Then
        For x = xTextPos + 2 To w Step 2.6
            UserControl.PSet (x, yTextPos), 0
            UserControl.PSet (x, yTextPos + mTextHeight), 0
        Next
 
        For y = 0 To mTextHeight Step 2.6
            UserControl.PSet (xTextPos - 1, y + yTextPos), 0
            UserControl.PSet (w + 2, y + yTextPos), 0
        Next
        
        UserControl.PSet (w, yTextPos + mTextHeight), 0
    End If
    
    'Print caption, includeing Eanbled state
    If m_Enabled Then
        UserControl.CurrentY = yTextPos
        UserControl.CurrentX = xTextPos
        UserControl.ForeColor = m_TempForeColor
        Call FixCaption(m_CheckCaption)
    Else
        UserControl.CurrentY = yTextPos
        UserControl.CurrentX = xTextPos
        UserControl.ForeColor = vbWhite
        Call FixCaption(m_CheckCaption)
        UserControl.CurrentY = yTextPos - 1
        UserControl.CurrentX = xTextPos - 1
        UserControl.ForeColor = vb3DShadow
        Call FixCaption(m_CheckCaption)
    End If
    
    
    'Ok i only added this part in quick becuase I was not going to add this
    ' but someone gave me an idea to add support for bitmaps.
    
    If StyleEx = CustBitmap Then
        'it a Bitmap check
        If m_Checked Then
            BitBlt PicCheck.hDC, 0, 0, 13, 13, PicTmp.hDC, 0, 13 * 4, vbSrcCopy
        Else
            BitBlt PicCheck.hDC, 0, 0, 13, 13, PicTmp.hDC, 0, 0, vbSrcCopy
        End If
        
        If Not m_Enabled Then
            If m_Checked Then
                BitBlt PicCheck.hDC, 0, 0, 13, 13, PicTmp.hDC, 0, 13 * 6, vbSrcCopy
            Else
                BitBlt PicCheck.hDC, 0, 0, 13, 13, PicTmp.hDC, 0, 13 * 2, vbSrcCopy
            End If
            
        End If
        Exit Sub
    End If
    
    'Draw the tick for the checkbox if
    If m_Checked Then
        DrawCheck LineColors(8), m_CheckStyle, m_BoxStyle
    End If
    
    PicCheck.Refresh 'Update the control
    
End Sub

Sub DrawCheck(mColor As OLE_COLOR, CheckEffect As nCheckStyle, BoxEffect As nBoxStyle)
Dim m_Area As Integer, m_Offset As Integer
Dim b As Boolean, mCol As Long, x As Integer, y As Integer
    PicCheck.FillStyle = 1
    PicCheck.DrawWidth = m_checkDrawWidth
    x = PicCheck.ScaleWidth
    
    'Get Align and area used for Square style of checkbox
    If BoxEffect = FlatStyle Then
        m_Area = 8: m_Offset = 2
    ElseIf BoxEffect = New3D Then
        m_Area = 7: m_Offset = 2
    Else
        m_Area = 6: m_Offset = 3
    End If
    
    'Draw the selected check style
    Select Case CheckEffect
    Case Check 'Draw the good old Check
        'This sub is used to draw the check
        PicCheck.PSet (x - 4, 3), mColor: PicCheck.PSet (x - 5, 4), mColor
        PicCheck.PSet (x - 4, 4), mColor: PicCheck.PSet (x - 5, 5), mColor
        PicCheck.PSet (x - 4, 5), mColor: PicCheck.PSet (x - 5, 6), mColor
        PicCheck.PSet (x - 6, 5), mColor: PicCheck.PSet (x - 7, 6), mColor
        PicCheck.PSet (x - 6, 6), mColor: PicCheck.PSet (x - 7, 7), mColor
        PicCheck.PSet (x - 6, 7), mColor: PicCheck.PSet (x - 7, 8), mColor
        PicCheck.PSet (x - 8, 7), mColor: PicCheck.PSet (x - 9, 6), mColor
        PicCheck.PSet (x - 8, 8), mColor: PicCheck.PSet (x - 9, 7), mColor
        PicCheck.PSet (x - 8, 9), mColor: PicCheck.PSet (x - 9, 8), mColor
        PicCheck.PSet (x - 10, 5), mColor: PicCheck.PSet (x - 10, 6), mColor: PicCheck.PSet (x - 10, 7), mColor
    Case SqrSolid 'Draw a soild square
        'PicCheck.FillStyle = 0
        PicCheck.Line (m_Offset, m_Offset)-(m_Area + m_Offset, m_Area + m_Offset), mColor, BF
    Case SqrOutLine 'Draw a unfill square
        'PicCheck.FillStyle = 1
        PicCheck.Line (m_Offset, m_Offset)-(m_Area + m_Offset, m_Area + m_Offset), mColor, B
    Case CircleSoild ' Draw a soild Circle
        PicCheck.FillStyle = 0
        PicCheck.FillColor = mColor
        PicCheck.Circle (6, 6), m_Area / 2, mColor
    Case CircleOutLine 'Draw an Circle
         PicCheck.Circle (6, 6), m_Area / 2, mColor
    Case Grid
        For x = m_Offset To m_Area + m_Offset
            For y = m_Offset To m_Area + m_Offset
                b = Not b
                If b Then mCol = PicCheck.BackColor Else mCol = mColor
                PicCheck.PSet (x, y), mCol
            Next
        Next
        
    End Select
    
    PicCheck.DrawWidth = 1

End Sub

Public Property Get BoxColor() As OLE_COLOR
    BoxColor = m_CheckBoxColor
End Property

Public Property Let BoxColor(ByVal vNewColor As OLE_COLOR)
    m_CheckBoxColor = vNewColor
    PropertyChanged "BoxColor"
    Call DrawBox(m_BoxStyle)
End Property

Public Property Get CheckColor() As OLE_COLOR
    CheckColor = m_CheckColor
End Property

Public Property Let CheckColor(ByVal vNewColor As OLE_COLOR)
    m_CheckColor = vNewColor
    PropertyChanged "CheckColor"
    Call DrawBox(m_BoxStyle)
End Property

Public Property Get BoxStyle() As nBoxStyle
    BoxStyle = m_BoxStyle
End Property

Public Property Let BoxStyle(ByVal vNewStyle As nBoxStyle)
    m_BoxStyle = vNewStyle
    PropertyChanged "BoxStyle"
    Call DrawBox(m_BoxStyle)
End Property


Public Property Get Caption() As String
Attribute Caption.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    Caption = m_CheckCaption
End Property

Public Property Let Caption(ByVal vNewCaption As String)
    m_CheckCaption = vNewCaption
    PropertyChanged "Caption"
    Call DrawBox(m_BoxStyle)
End Property

Private Sub Command1_Click()

End Sub

Private Sub PicCheck_GotFocus()
    If Not m_Enabled Then Exit Sub
    If m_ShowRect <> True Then Exit Sub
    mDrawFocus = True
    Call DrawBox(m_BoxStyle)
End Sub

Private Sub PicCheck_LostFocus()
    If Not m_Enabled Then Exit Sub
    If m_ShowRect <> True Then Exit Sub
    mDrawFocus = False
    Call DrawBox(m_BoxStyle)
End Sub

Private Sub PicCheck_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> vbLeftButton Then Exit Sub
    If Not m_Enabled Then Exit Sub

    m_Checked = Not m_Checked
    Call DrawBox(m_BoxStyle)
    RaiseEvent CheckClick
    RaiseEvent MouseDown(Button, Shift, x, y)
    
End Sub

Private Sub PicCheck_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim mTrackState As HoverEvt

    If (x < 0) Or (x > UserControl.ScaleWidth) Or (y < 0) Or (y > UserControl.ScaleHeight) Then
        mTrackState = HoverOut
        If m_hottrack Then m_TempForeColor = m_TmpForeColor
        Call DrawBox(m_BoxStyle)
        RaiseEvent MouseMove(Button, Shift, x, y, mTrackState)
        ReleaseCapture
    ElseIf GetCapture() <> UserControl.hwnd Then
        mTrackState = HoverIn
        If m_hottrack Then m_TmpForeColor = m_TempForeColor: m_TempForeColor = m_HotTrackingColor
        RaiseEvent MouseMove(Button, Shift, x, y, mTrackState)
        SetCapture UserControl.hwnd
        Call DrawBox(m_BoxStyle)
    End If

End Sub

Private Sub PicCheck_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    RaiseEvent MouseUp(Button, Shift, x, y)
    If m_hottrack Then m_TempForeColor = m_TmpForeColor
End Sub

Private Sub UserControl_AccessKeyPress(KeyAscii As Integer)
    'Note this will only make focus if focus rect is turned on for that check
    PicCheck.SetFocus
End Sub

Private Sub UserControl_Initialize()
    m_checkDrawWidth = 1
    m_FlatBorderColor = &HE5A165
    m_HotTrackingColor = vbBlue
    m_CheckStyle = Check
    m_Enabled = True
    m_TempForeColor = 0
    BoxColor = vbWhite
    m_Checked = False
    m_CheckCaption = "eCheckFX"
    UserControl.KeyPreview = True
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If mDrawFocus And KeyAscii = 32 Then
        Call PicCheck_MouseDown(vbLeftButton, 0, 0, 0)
    End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PicCheck_MouseDown(Button, Shift, x, y)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_BoxStyle = PropBag.ReadProperty("BoxStyle", 0)
    m_Checked = PropBag.ReadProperty("Checked", False)
    m_CheckColor = PropBag.ReadProperty("CheckColor", 0)
    m_CheckCaption = PropBag.ReadProperty("Caption", "eCheckFX")
    m_Align = PropBag.ReadProperty("Align", 0)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_TempForeColor = PropBag.ReadProperty("ForeColor", vbBlack)
    
    Set UserControl.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    m_ShowRect = PropBag.ReadProperty("ShowFocusRect", False)
    m_CheckBoxColor = PropBag.ReadProperty("BoxColor", vbWhite)
    m_Enabled = PropBag.ReadProperty("Enabled", True)
    m_FlatBorderColor = PropBag.ReadProperty("BorderColor", &HE5A165)
    m_checkDrawWidth = PropBag.ReadProperty("CheckDrawWidth", 1)
    m_CheckStyle = PropBag.ReadProperty("CheckStyle", 0)
    m_gray = PropBag.ReadProperty("Grayed", 0)
    m_hottrack = PropBag.ReadProperty("HotTracking", 0)
    m_HotTrackingColor = PropBag.ReadProperty("TrackingColor", vbBlue)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    m_StyleEx = PropBag.ReadProperty("StyleEx", 0)
    
End Sub

Private Sub UserControl_Resize()
    Call DrawBox(m_BoxStyle)
End Sub

Private Sub UserControl_Show()
    Call DrawBox(m_BoxStyle)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BoxStyle", m_BoxStyle, 0)
    Call PropBag.WriteProperty("Checked", m_Checked, False)
    Call PropBag.WriteProperty("CheckColor", m_CheckColor, 0)
    Call PropBag.WriteProperty("Caption", m_CheckCaption, "eCheckFx")
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ForeColor", m_TempForeColor, vbBlack)
    Call PropBag.WriteProperty("Align", m_Align, 0)
    Call PropBag.WriteProperty("Font", UserControl.Font, Ambient.Font)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("ShowFocusRect", m_ShowRect, False)
    Call PropBag.WriteProperty("BoxColor", m_CheckBoxColor, vbWhite)
    Call PropBag.WriteProperty("Enabled", m_Enabled, True)
    Call PropBag.WriteProperty("BorderColor", m_FlatBorderColor, &HE5A165)
    Call PropBag.WriteProperty("CheckDrawWidth", m_checkDrawWidth, 1)
    Call PropBag.WriteProperty("CheckStyle", m_CheckStyle, 0)
    Call PropBag.WriteProperty("Grayed", m_gray, 0)
    Call PropBag.WriteProperty("HotTracking", m_hottrack, 0)
    Call PropBag.WriteProperty("TrackingColor", m_HotTrackingColor, vbBlue)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("StyleEx", m_StyleEx, 0)
    
End Sub

Public Property Get Checked() As Boolean
Attribute Checked.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    Checked = m_Checked
End Property

Public Property Let Checked(ByVal vNewCheck As Boolean)
    m_Checked = vNewCheck
    PropertyChanged "Checked"
    Call DrawBox(m_BoxStyle)
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
    Call DrawBox(m_BoxStyle)
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = m_TempForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    m_TempForeColor = New_ForeColor
    PropertyChanged "ForeColor"
    Call DrawBox(m_BoxStyle)
End Property

Public Property Get Align() As nAlign
    Align = m_Align
End Property

Public Property Let Align(ByVal vNewAlign As nAlign)
    m_Align = vNewAlign
    PropertyChanged "Align"
    Call DrawBox(m_BoxStyle)
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = UserControl.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set UserControl.Font = New_Font
    PropertyChanged "Font"
    Call DrawBox(m_BoxStyle)
End Property

Private Sub UserControl_InitProperties()
    Set UserControl.Font = Ambient.Font
    m_TempForeColor = UserControl.ForeColor
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PicCheck_MouseMove(Button, Shift, x, y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Call PicCheck_MouseUp(Button, Shift, x, y)
End Sub

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As MousePointerConstants
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePointerConstants)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get ShowFocusRect() As Boolean
Attribute ShowFocusRect.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
   ShowFocusRect = m_ShowRect
End Property

Public Property Let ShowFocusRect(ByVal vNewValue As Boolean)
    m_ShowRect = vNewValue
    PropertyChanged "ShowFocusRect"
End Property

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal vNewValue As Boolean)
    m_Enabled = vNewValue
    PropertyChanged "Enabled"
    Call DrawBox(m_BoxStyle)
End Property

Public Property Get BorderColor() As OLE_COLOR
    BorderColor = m_FlatBorderColor
End Property

Public Property Let BorderColor(ByVal vBorderColor As OLE_COLOR)
    m_FlatBorderColor = vBorderColor
    PropertyChanged "Enabled"
    Call DrawBox(m_BoxStyle)
End Property

Public Property Get CheckDrawWidth() As Integer
Attribute CheckDrawWidth.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    CheckDrawWidth = m_checkDrawWidth
End Property

Public Property Let CheckDrawWidth(ByVal vNewWidth As Integer)
    m_checkDrawWidth = vNewWidth
    PropertyChanged "CheckDrawWidth"
    Call DrawBox(m_BoxStyle)
End Property

Public Property Get CheckStyle() As nCheckStyle
    CheckStyle = m_CheckStyle
End Property

Public Property Let CheckStyle(ByVal vNewStyle As nCheckStyle)
    m_CheckStyle = vNewStyle
    PropertyChanged "CheckStyle"
    Call DrawBox(m_BoxStyle)
End Property

Public Property Get HotTracking() As Boolean
Attribute HotTracking.VB_ProcData.VB_Invoke_Property = "PropertyPage1"
    HotTracking = m_hottrack
End Property

Public Property Let HotTracking(ByVal vNewHotTrack As Boolean)
    m_hottrack = vNewHotTrack
End Property

Public Property Get TrackingColor() As OLE_COLOR
    TrackingColor = m_HotTrackingColor
End Property

Public Property Let TrackingColor(ByVal vNewTrackColor As OLE_COLOR)
    m_HotTrackingColor = vNewTrackColor
    PropertyChanged "CheckStyle"
End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = PicTmp.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set PicTmp.Picture = New_Picture
    PropertyChanged "Picture"
    Call DrawBox(m_BoxStyle)
End Property

Public Property Get StyleEx() As StyleEx
    StyleEx = m_StyleEx
End Property

Public Property Let StyleEx(ByVal vNewStyleEx As StyleEx)
    m_StyleEx = vNewStyleEx
    PropertyChanged "StyleEx"
    Call DrawBox(m_BoxStyle)
End Property
