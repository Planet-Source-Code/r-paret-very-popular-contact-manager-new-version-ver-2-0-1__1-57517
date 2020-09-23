VERSION 5.00
Begin VB.Form frmFirstLogin 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   " Credo Contact Manager Login"
   ClientHeight    =   4470
   ClientLeft      =   2790
   ClientTop       =   3150
   ClientWidth     =   5310
   Icon            =   "frmFirstLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2641.023
   ScaleMode       =   0  'User
   ScaleWidth      =   4985.802
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Credo_Contact_Manager.XPButton XPButton_Cancel 
      Height          =   375
      Left            =   3480
      TabIndex        =   6
      Top             =   3600
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
      Left            =   2400
      TabIndex        =   5
      Top             =   3600
      Width           =   855
      _ExtentX        =   1508
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
   Begin VB.TextBox txtUserName 
      BackColor       =   &H00FFFFFF&
      Height          =   345
      Left            =   1920
      TabIndex        =   1
      Top             =   2520
      Width           =   2685
   End
   Begin VB.TextBox txtPassword 
      BackColor       =   &H80000013&
      Enabled         =   0   'False
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1920
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3000
      Width           =   2685
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "First Login"
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
      TabIndex        =   7
      Top             =   240
      Width           =   3375
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   4200
      MouseIcon       =   "frmFirstLogin.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "frmFirstLogin.frx":0614
      ToolTipText     =   " Move "
      Top             =   240
      Width           =   375
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   4680
      Picture         =   "frmFirstLogin.frx":0DC2
      ToolTipText     =   " Close "
      Top             =   240
      Width           =   375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   5055
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   240
      Picture         =   "frmFirstLogin.frx":1570
      Top             =   169
      Width           =   390
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   3375
      Left            =   120
      Top             =   960
      Width           =   5055
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   840
      Picture         =   "frmFirstLogin.frx":2052
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "This is your first login in Credo Contact Manager, you have to enter a User Name and Password. "
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
      Left            =   1800
      TabIndex        =   4
      Top             =   1320
      Width           =   2895
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
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
      Top             =   2640
      Width           =   1320
   End
   Begin VB.Label lblLabels 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
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
      Top             =   3120
      Width           =   1080
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0FFFF&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   3135
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   4815
   End
End
Attribute VB_Name = "frmFirstLogin"
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

Private Sub Form_Unload(Cancel As Integer)

Dim Ini As clsIniFile
Set Ini = New clsIniFile

With Ini
    .File = App.Path & "\Settings.ini"
    .WriteValue "DataSource", "UNdata", (txtUserName.Text)
    .WriteValue "DataSource", "PWdata", (txtPassword.Text)
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

    'check if text is filled in
    
    'If txtUserName.Text = "" Then
    If txtPassword.Text = "" Then
        MsgBox "Please enter User Name and Password, try again!", , "Credo Contact Manager Login"
       ' txtUserName.SetFocus
       ' txtPassword.SetFocus
        SendKeys "{Home}+{End}"
               
  Else:
Unload Me
Address_Contact_Manager.Show
    
End If

End Sub
