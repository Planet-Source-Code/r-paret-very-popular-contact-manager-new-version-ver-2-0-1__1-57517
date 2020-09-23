VERSION 5.00
Begin VB.Form frmLogin2 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   " Credo Contact Manager Login"
   ClientHeight    =   4230
   ClientLeft      =   2790
   ClientTop       =   3150
   ClientWidth     =   5325
   Icon            =   "frmLogin2.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2499.223
   ScaleMode       =   0  'User
   ScaleWidth      =   4999.886
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Credo_Contact_Manager.XPButton XPButton_Cancel 
      Height          =   375
      Left            =   3480
      TabIndex        =   7
      Top             =   3360
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
      ForeColor       =   -2147483630
      ForeHover       =   12582912
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Ok 
      Height          =   375
      Left            =   2640
      TabIndex        =   6
      Top             =   3360
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
      Caption         =   "Ok"
      ForeColor       =   -2147483630
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
      Left            =   1920
      TabIndex        =   4
      Top             =   1440
      Value           =   1  'Checked
      Width           =   2775
   End
   Begin VB.TextBox txtUserName 
      Height          =   345
      Left            =   1920
      TabIndex        =   1
      Top             =   2280
      Width           =   2685
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H80000004&
      Enabled         =   0   'False
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   2760
      Width           =   2685
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   720
      TabIndex        =   8
      Top             =   240
      Width           =   3375
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   4200
      MouseIcon       =   "frmLogin2.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "frmLogin2.frx":0614
      ToolTipText     =   " Move "
      Top             =   240
      Width           =   375
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   4680
      Picture         =   "frmLogin2.frx":0DC2
      ToolTipText     =   " Close "
      Top             =   240
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   240
      Picture         =   "frmLogin2.frx":1570
      Top             =   169
      Width           =   390
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   3135
      Left            =   120
      Top             =   960
      Width           =   5055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   5055
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
      Left            =   2160
      TabIndex        =   5
      Top             =   1800
      Width           =   2535
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   840
      Picture         =   "frmLogin2.frx":2052
      Top             =   1440
      Width           =   480
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "&User Name:"
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
      Height          =   270
      Index           =   0
      Left            =   465
      TabIndex        =   0
      Top             =   2400
      Width           =   1320
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C00000&
      BackStyle       =   0  'Transparent
      Caption         =   "&Password:"
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
      Height          =   270
      Index           =   1
      Left            =   720
      TabIndex        =   2
      Top             =   2880
      Width           =   1080
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0FFFF&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   2895
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   4815
   End
End
Attribute VB_Name = "frmLogin2"
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

    'check for existing password
    
With Ini
    .File = App.Path & "\Settings.ini"
    GlobalSettingsUN = .GetValue("DataSource", "UNdata")
    GlobalSettingsPW = .GetValue("DataSource", "PWdata")
    If GlobalSettingsUN = "" Then
    If GlobalSettingsPW = "" Then
    
       frmFirstLogin.Show
       Unload Me
       Else
       Exit Sub

2 Check1.Value = Unchecked
Address_Contact_Manager.Show
Unload Me
Exit Sub

End If
End If

End With

End Sub

Private Sub Form_Unload(Cancel As Integer)
   

Dim Ini As clsIniFile

If Check1 = Checked Then GoTo 1
If Check1 = Unchecked Then GoTo 2

1
Set Ini = New clsIniFile

With Ini
    .File = App.Path & "\Settings.ini"
    .WriteValue "DataSource", "logon", 1
End With
Exit Sub

2
Set Ini = New clsIniFile

With Ini
    .File = App.Path & "\Settings.ini"
    .WriteValue "DataSource", "logon", 2
End With

End Sub

Private Sub Image3_Click()

    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Unload Me

End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

FormDrag Me

End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

FormDrag Me

End Sub

Private Sub txtUserName_Change()

txtPassword.BackColor = &HFFFFFF
txtPassword.Enabled = True
txtUserName.Enabled = True

End Sub

Private Sub XPButton_Cancel_Click()

    'set the global var to false
    'to denote a failed login
    LoginSucceeded = False
    Unload Me

End Sub

Private Sub XPButton_Ok_Click()

    If txtPassword.Text = "" Then
    'If txtUserName.Text = "" Then
        MsgBox "Invalid User Name or Password, try again!", , "Credo Contact Manager Login"
        SendKeys "{Home}+{End}"
        Exit Sub
 Else
End If

Dim Ini As clsIniFile

'Make some global variables to hold your settings
Dim GlobalSettingsUN As String
Dim GlobalSettingsPW As String

Set Ini = New clsIniFile

    'check for the password
    
With Ini
    .File = App.Path & "\Settings.ini"
    GlobalSettingsUN = .GetValue("DataSource", "UNdata")
    GlobalSettingsPW = .GetValue("DataSource", "PWdata")

    'check for correct password
    If txtUserName.Text = GlobalSettingsUN Then GoTo 1
    GoTo 2
    
1    If txtPassword.Text = GlobalSettingsPW Then
        LoginSucceeded = True
Address_Contact_Manager.Show
Unload Me
       Else
        MsgBox "Invalid User Name or Password, try again!", , "Credo Contact Manager Login"
        SendKeys "{Home}+{End}"
        Exit Sub
        
2       MsgBox "Invalid User Name or Password, try again!", , "Credo Contact Manager Login"
        SendKeys "{Home}+{End}"

        Exit Sub
    End If
    

End With


End Sub
