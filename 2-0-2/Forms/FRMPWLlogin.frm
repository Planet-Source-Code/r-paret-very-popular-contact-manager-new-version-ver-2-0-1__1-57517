VERSION 5.00
Begin VB.Form frmPWLogin 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   " Credo Contact Manager Login"
   ClientHeight    =   3375
   ClientLeft      =   2790
   ClientTop       =   3150
   ClientWidth     =   4590
   Icon            =   "FRMPWLlogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1994.061
   ScaleMode       =   0  'User
   ScaleWidth      =   4309.761
   StartUpPosition =   2  'CenterScreen
   Begin Credo_Contact_Manager.XPButton XPButton_Cancel 
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      Top             =   2520
      Width           =   1095
      _ExtentX        =   1931
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
   Begin Credo_Contact_Manager.XPButton XPButton_Ok 
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   2520
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
      Caption         =   "OK"
      ForeColor       =   -2147483642
      ForeHover       =   12582912
   End
   Begin VB.CheckBox Check1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Show login dialog"
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
      Left            =   1440
      TabIndex        =   0
      Top             =   1440
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Show Login at Start"
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
      Left            =   720
      TabIndex        =   4
      Top             =   240
      Width           =   2655
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   4335
   End
   Begin VB.Image Image4 
      Height          =   510
      Left            =   240
      Picture         =   "FRMPWLlogin.frx":030A
      Top             =   169
      Width           =   390
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   3480
      MouseIcon       =   "FRMPWLlogin.frx":0DEC
      MousePointer    =   99  'Custom
      Picture         =   "FRMPWLlogin.frx":10F6
      ToolTipText     =   " Move "
      Top             =   240
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   3960
      Picture         =   "FRMPWLlogin.frx":18A4
      ToolTipText     =   " Close "
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   " on start-up of program"
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
      Left            =   1680
      TabIndex        =   1
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   480
      Picture         =   "FRMPWLlogin.frx":2052
      Top             =   1440
      Width           =   480
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0FFFF&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   2055
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   4095
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   2295
      Left            =   120
      Top             =   960
      Width           =   4335
   End
End
Attribute VB_Name = "frmPWLogin"
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
Dim GlobalSettings As String
Dim GlobalSettingsUN As String
Dim GlobalSettingsPW As String

Dim Ini As clsIniFile
Set Ini = New clsIniFile

With Ini
    .File = App.Path & "\Settings.ini"
    GlobalSettings = .GetValue("DataSource", "logon")
    If GlobalSettings = 1 Then GoTo 1
    If GlobalSettings = 2 Then GoTo 2
    
End With

1 Check1.Value = Checked
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

    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
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
    .WriteValue "DataSource", "logon", 1
End With
Unload Me
Exit Sub

2
Set Ini = New clsIniFile

With Ini
    .File = App.Path & "\Settings.ini"
    .WriteValue "DataSource", "logon", 2
End With
Unload Me

End Sub
