VERSION 5.00
Begin VB.Form frmDemo 
   Caption         =   "Demo DM eCheckBoxFX v2"
   ClientHeight    =   6810
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8550
   LinkTopic       =   "Form1"
   ScaleHeight     =   454
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   570
   StartUpPosition =   2  'CenterScreen
   Begin Project1.eCheckFX eCheckFX26 
      Height          =   240
      Left            =   4380
      TabIndex        =   31
      Top             =   5415
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   423
      Caption         =   "eCheckFX"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "demo.frx":0000
      StyleEx         =   1
   End
   Begin Project1.eCheckFX eCheckFX25 
      Height          =   195
      Left            =   3075
      TabIndex        =   30
      Top             =   3300
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   344
      BoxStyle        =   2
      Checked         =   -1  'True
      CheckColor      =   4210752
      Caption         =   "Grid Style"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckStyle      =   5
   End
   Begin Project1.eCheckFX eCheckFX22 
      Height          =   240
      Left            =   4395
      TabIndex        =   27
      Top             =   3960
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   423
      Caption         =   "Check box with HotTracking"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HotTracking     =   -1  'True
      TrackingColor   =   8388608
   End
   Begin Project1.eCheckFX eCheckFX12 
      Height          =   270
      Left            =   120
      TabIndex        =   17
      Top             =   3255
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   476
      Checked         =   -1  'True
      Caption         =   "Soild Square"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckStyle      =   1
   End
   Begin VB.CommandButton cmdabout 
      Caption         =   "About"
      Height          =   495
      Left            =   510
      TabIndex        =   15
      Top             =   5490
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   495
      Left            =   1965
      TabIndex        =   14
      Top             =   5490
      Width           =   1215
   End
   Begin Project1.eCheckFX eCheckFX1 
      Height          =   300
      Left            =   150
      TabIndex        =   0
      Top             =   225
      Width           =   3240
      _ExtentX        =   5715
      _ExtentY        =   529
      Caption         =   "DM Check box with old VB Default Style"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.eCheckFX eCheckFX2 
      Height          =   300
      Left            =   150
      TabIndex        =   1
      Top             =   630
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   529
      BoxStyle        =   1
      Caption         =   "DM Check box with Moden VB Default Style"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.eCheckFX eCheckFX3 
      Height          =   300
      Left            =   150
      TabIndex        =   2
      Top             =   1020
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   529
      Checked         =   -1  'True
      Caption         =   "DM Checkbox This checkbox is checked"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.eCheckFX eCheckFX4 
      Height          =   300
      Left            =   150
      TabIndex        =   3
      Top             =   1455
      Width           =   3495
      _ExtentX        =   6165
      _ExtentY        =   529
      Checked         =   -1  'True
      Caption         =   "DM Checkbox Disbaled"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Enabled         =   0   'False
   End
   Begin Project1.eCheckFX eCheckFX5 
      Height          =   300
      Left            =   150
      TabIndex        =   4
      Top             =   1875
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   529
      Checked         =   -1  'True
      Caption         =   "DM Checkbox Text Align Left"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.eCheckFX eCheckFX6 
      Height          =   300
      Left            =   150
      TabIndex        =   5
      Top             =   2325
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   529
      BoxStyle        =   1
      Checked         =   -1  'True
      Caption         =   "DM Checkbox Text Align Right"
      Align           =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.eCheckFX eCheckFX7 
      Height          =   300
      Left            =   4395
      TabIndex        =   6
      Top             =   180
      Width           =   2610
      _ExtentX        =   4604
      _ExtentY        =   529
      Checked         =   -1  'True
      Caption         =   "This checkbox has focus"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocusRect   =   -1  'True
   End
   Begin Project1.eCheckFX eCheckFX8 
      Height          =   300
      Left            =   4395
      TabIndex        =   7
      Top             =   675
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   529
      Checked         =   -1  'True
      Caption         =   "DM Checkbox Foreground, Background Color"
      BackColor       =   16777215
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocusRect   =   -1  'True
   End
   Begin Project1.eCheckFX eCheckFX9 
      Height          =   300
      Left            =   4395
      TabIndex        =   8
      Top             =   1065
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   529
      Checked         =   -1  'True
      Caption         =   "DM Checkbox Box Color"
      BackColor       =   -2147483638
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocusRect   =   -1  'True
      BoxColor        =   65535
   End
   Begin Project1.eCheckFX eCheckFX10 
      Height          =   300
      Left            =   4395
      TabIndex        =   9
      Top             =   1440
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   529
      Checked         =   -1  'True
      CheckColor      =   16711680
      Caption         =   "DM Checkbox Check Color"
      BackColor       =   -2147483638
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocusRect   =   -1  'True
   End
   Begin Project1.eCheckFX eCheckFX11 
      Height          =   300
      Left            =   4395
      TabIndex        =   10
      Top             =   2475
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   529
      BoxStyle        =   2
      Checked         =   -1  'True
      CheckColor      =   16711680
      Caption         =   "Test Test Test"
      BackColor       =   -2147483638
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocusRect   =   -1  'True
   End
   Begin Project1.eCheckFX eCheckFX13 
      Height          =   300
      Left            =   4395
      TabIndex        =   12
      Top             =   2865
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   529
      BoxStyle        =   2
      Checked         =   -1  'True
      CheckColor      =   12632319
      Caption         =   "Test Test Test"
      BackColor       =   -2147483638
      ForeColor       =   255
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocusRect   =   -1  'True
      BoxColor        =   8421631
      BorderColor     =   255
   End
   Begin Project1.eCheckFX eCheckFX14 
      Height          =   300
      Left            =   4395
      TabIndex        =   13
      Top             =   3240
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   529
      BoxStyle        =   2
      Checked         =   -1  'True
      CheckColor      =   33023
      Caption         =   "Test Test Test"
      BackColor       =   -2147483638
      ForeColor       =   16744703
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocusRect   =   -1  'True
      BoxColor        =   8438015
      BorderColor     =   12640511
   End
   Begin Project1.eCheckFX eCheckFX15 
      Height          =   270
      Left            =   1620
      TabIndex        =   18
      Top             =   3255
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   476
      Checked         =   -1  'True
      Caption         =   "Square Outline"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckStyle      =   2
   End
   Begin Project1.eCheckFX eCheckFX16 
      Height          =   270
      Left            =   120
      TabIndex        =   19
      Top             =   3570
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   476
      Checked         =   -1  'True
      Caption         =   "Soild Circle"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckStyle      =   3
   End
   Begin Project1.eCheckFX eCheckFX17 
      Height          =   270
      Left            =   1620
      TabIndex        =   20
      Top             =   3570
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   476
      Checked         =   -1  'True
      CheckColor      =   33023
      Caption         =   "Circle Outline"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckStyle      =   4
   End
   Begin Project1.eCheckFX eCheckFX18 
      Height          =   270
      Left            =   165
      TabIndex        =   22
      Top             =   4335
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   476
      BoxStyle        =   3
      Checked         =   -1  'True
      Caption         =   "Button Raised"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BoxColor        =   -2147483648
   End
   Begin Project1.eCheckFX eCheckFX19 
      Height          =   270
      Left            =   1620
      TabIndex        =   23
      Top             =   4320
      Width           =   1320
      _ExtentX        =   2328
      _ExtentY        =   476
      BoxStyle        =   4
      Checked         =   -1  'True
      Caption         =   "Button Lowerd"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BoxColor        =   -2147483648
   End
   Begin Project1.eCheckFX eCheckFX20 
      Height          =   270
      Left            =   165
      TabIndex        =   24
      Top             =   4635
      Width           =   1485
      _ExtentX        =   2619
      _ExtentY        =   476
      BoxStyle        =   5
      Checked         =   -1  'True
      Caption         =   "Doubble Border"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderColor     =   33023
   End
   Begin Project1.eCheckFX eCheckFX21 
      Height          =   270
      Left            =   1635
      TabIndex        =   25
      Top             =   4650
      Width           =   900
      _ExtentX        =   1588
      _ExtentY        =   476
      BoxStyle        =   6
      Checked         =   -1  'True
      Caption         =   "Custom"
      BackColor       =   -2147483648
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin Project1.eCheckFX eCheckFX23 
      Height          =   240
      Left            =   4395
      TabIndex        =   28
      Top             =   4335
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   423
      Caption         =   "Check box supporting access &keys"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocusRect   =   -1  'True
   End
   Begin Project1.eCheckFX eCheckFX24 
      Height          =   240
      Left            =   4380
      TabIndex        =   29
      Top             =   4680
      Width           =   3870
      _ExtentX        =   6826
      _ExtentY        =   423
      Caption         =   "Move Over Me"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ShowFocusRect   =   -1  'True
   End
   Begin Project1.eCheckFX eCheckFX27 
      Height          =   240
      Left            =   4380
      TabIndex        =   33
      Top             =   5730
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   423
      Caption         =   "Windows XP Style 1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "demo.frx":2002
      StyleEx         =   1
   End
   Begin Project1.eCheckFX eCheckFX28 
      Height          =   240
      Left            =   4380
      TabIndex        =   34
      Top             =   6045
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   423
      Caption         =   "Some Other Style"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Picture         =   "demo.frx":2C1A
      StyleEx         =   1
   End
   Begin Project1.eCheckFX eCheckFX29 
      Height          =   240
      Left            =   4380
      TabIndex        =   35
      Top             =   6330
      Width           =   1980
      _ExtentX        =   3493
      _ExtentY        =   423
      Checked         =   -1  'True
      Caption         =   "Some Other Style"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CheckStyle      =   5
      Picture         =   "demo.frx":44CC
      StyleEx         =   1
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "Checkbox with custom bitmaps"
      Height          =   195
      Left            =   4380
      TabIndex        =   32
      Top             =   5100
      Width           =   2190
   End
   Begin VB.Label lblOther 
      AutoSize        =   -1  'True
      Caption         =   "Other new things"
      Height          =   195
      Left            =   4395
      TabIndex        =   26
      Top             =   3645
      Width           =   1200
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Some new box styles"
      Height          =   195
      Left            =   165
      TabIndex        =   21
      Top             =   4050
      Width           =   1485
   End
   Begin VB.Label lblStyles 
      AutoSize        =   -1  'True
      Caption         =   "Some new check styles"
      Height          =   195
      Left            =   150
      TabIndex        =   16
      Top             =   2910
      Width           =   1680
   End
   Begin VB.Line Line1 
      X1              =   11
      X2              =   229
      Y1              =   184
      Y2              =   184
   End
   Begin VB.Label Label1 
      Caption         =   "With Flat Style enabled you can use checks like this"
      Height          =   360
      Left            =   4395
      TabIndex        =   11
      Top             =   1905
      Width           =   3195
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdabout_Click()
    MsgBox frmDemo.Caption & vbCrLf & "Written by Ben Jones" & vbCrLf & "  Please Vote", vbInformation, "About"
End Sub

Private Sub cmdExit_Click()
    End
End Sub

Private Sub eCheckFX24_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single, HoverEvent As HoverEvt)
    If HoverEvent = HoverIn Then
        eCheckFX24.Caption = "Hover Event: HoverIn"
    Else
        eCheckFX24.Caption = "Hover Event: HoverOut"
    End If
    
End Sub

Private Sub Form_Load()
    'This is used to add custom colors to a check box
    eCheckFX21.AddCustomBorders vbRed, vbBlack, vbBlue, vbWhite, vbBlue, vbWhite, vbBlack, vbRed
End Sub

Private Sub Form_Paint()
    eCheckFX7.SetFocus
End Sub
