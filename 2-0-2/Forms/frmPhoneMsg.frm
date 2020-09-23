VERSION 5.00
Begin VB.Form frmPhoneMsg 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   " Credo Contact Manager Login"
   ClientHeight    =   4110
   ClientLeft      =   2790
   ClientTop       =   3150
   ClientWidth     =   4845
   Icon            =   "frmPhoneMsg.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2428.323
   ScaleMode       =   0  'User
   ScaleWidth      =   4549.192
   StartUpPosition =   2  'CenterScreen
   Begin Credo_Contact_Manager.XPButton XPButton_Ok 
      Height          =   375
      Left            =   2160
      TabIndex        =   4
      Top             =   3240
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "OK"
      ForeColor       =   -2147483642
      ForeHover       =   12582912
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Cancel 
      Height          =   375
      Left            =   3240
      TabIndex        =   3
      Top             =   3240
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Cancel"
      ForeColor       =   -2147483642
      ForeHover       =   12582912
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Do not show this message again."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   2640
      Value           =   1  'Checked
      Width           =   3855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Phone Message"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   840
      TabIndex        =   2
      Top             =   240
      Width           =   2775
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   4575
   End
   Begin VB.Image Image4 
      Height          =   600
      Left            =   240
      Picture         =   "frmPhoneMsg.frx":030A
      Stretch         =   -1  'True
      Top             =   165
      Width           =   600
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   3720
      MouseIcon       =   "frmPhoneMsg.frx":11D4
      MousePointer    =   99  'Custom
      Picture         =   "frmPhoneMsg.frx":14DE
      ToolTipText     =   " Move "
      Top             =   240
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   4200
      Picture         =   "frmPhoneMsg.frx":1C8C
      ToolTipText     =   " Close "
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "Note:  Phone dialer will only work if a data-fax modem is installed on your PC and a microphone and speakers."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   1095
      Left            =   1320
      TabIndex        =   1
      Top             =   1440
      Width           =   3015
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   480
      Picture         =   "frmPhoneMsg.frx":243A
      Top             =   1560
      Width           =   480
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0FFFF&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   2775
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   4335
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   3015
      Left            =   120
      Top             =   960
      Width           =   4575
   End
End
Attribute VB_Name = "frmPhoneMsg"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wparam As Integer, ByVal iparam As Long) As Long

Public Sub FormDrag(Theform As Form)
    
    ReleaseCapture
    Call SendMessage(Theform.hwnd, &HA1, 2, 0&)

End Sub


Private Sub Form_Load()

'Make some global variables to hold your settings
Dim GlobalPhoneMsg As String

Dim Ini As clsIniFile
Set Ini = New clsIniFile

With Ini
    .File = App.Path & "\Settings.ini"
    GlobalPhoneMsg = .GetValue("CheckedPhoneMsg", "CheckedYN")
    If GlobalPhoneMsg = 1 Then GoTo 1
    If GlobalPhoneMsg = 2 Then GoTo 2
    
End With

1 Check1.Value = Checked
Unload Me
Exit Sub

2 Check1.Value = Unchecked
Exit Sub

End Sub

Private Sub Image2_Click()

    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Unload Me

End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

FormDrag Me

End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

FormDrag Me

End Sub

Private Sub XPButton_Cancel_Click()

    Unload Me

End Sub

Private Sub XPButton_Ok_Click()

Dim Ini As clsIniFile

If Check1 = Checked Then GoTo 1
If Check1 = Unchecked Then GoTo 2

1
Set Ini = New clsIniFile

With Ini
    .File = App.Path & "\Settings.ini"
    .WriteValue "CheckedPhoneMsg", "CheckedYN", 1
End With
Unload Me
Exit Sub

2
Set Ini = New clsIniFile

With Ini
    .File = App.Path & "\Settings.ini"
    .WriteValue "CheckedPhoneMsg", "CheckedYN", 2
End With
Unload Me

End Sub
