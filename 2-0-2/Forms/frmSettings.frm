VERSION 5.00
Begin VB.Form frmSettings 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   " Credo Contact Manager Login"
   ClientHeight    =   4215
   ClientLeft      =   2790
   ClientTop       =   3150
   ClientWidth     =   6855
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2490.361
   ScaleMode       =   0  'User
   ScaleWidth      =   6436.473
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Print notes when printing all contacts"
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
      TabIndex        =   5
      Top             =   2640
      Width           =   5055
   End
   Begin VB.CheckBox Check2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Print notes when printing a contact"
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
      TabIndex        =   4
      Top             =   2040
      Width           =   5055
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Cancel 
      Height          =   375
      Left            =   4440
      TabIndex        =   2
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
      ForeColor       =   -2147483642
      ForeHover       =   12582912
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Ok 
      Height          =   375
      Left            =   3000
      TabIndex        =   1
      Top             =   3360
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
      Caption         =   "Show password  login on start-up of program"
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
      Width           =   4935
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Settings"
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
      Left            =   1080
      TabIndex        =   3
      Top             =   240
      Width           =   4575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   6615
   End
   Begin VB.Image Image4 
      Height          =   495
      Left            =   240
      Picture         =   "frmSettings.frx":030A
      Top             =   186
      Width           =   825
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   5760
      MouseIcon       =   "frmSettings.frx":18F4
      MousePointer    =   99  'Custom
      Picture         =   "frmSettings.frx":1BFE
      ToolTipText     =   " Move "
      Top             =   240
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   6240
      Picture         =   "frmSettings.frx":23AC
      ToolTipText     =   " Close "
      Top             =   240
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   480
      Picture         =   "frmSettings.frx":2B5A
      Top             =   1440
      Width           =   480
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0FFFF&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   2895
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   6375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   3135
      Left            =   120
      Top             =   960
      Width           =   6615
   End
End
Attribute VB_Name = "frmSettings"
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
Dim GlobalChecked2 As String
Dim GlobalChecked3 As String

Dim NextChecked2
Dim NextChecked3

Dim Ini As clsIniFile
Set Ini = New clsIniFile
With Ini
    .File = App.Path & "\Settings.ini"
    GlobalSettings = .GetValue("DataSource", "logon")
    If GlobalSettings = 1 Then GoTo 1
    If GlobalSettings = 2 Then GoTo 2
End With
Exit Sub

1 Check1.Value = Checked
GoTo NextChecked2
Exit Sub

2 Check1.Value = Unchecked
GoTo NextChecked2
Exit Sub

NextChecked2:
With Ini
    .File = App.Path & "\Settings.ini"
    GlobalChecked2 = .GetValue("CheckedDataSource", "Checked2")
    If GlobalChecked2 = 1 Then GoTo 3
    If GlobalChecked2 = 2 Then GoTo 4
End With
Exit Sub

3 Check2.Value = Checked
GoTo NextChecked3
Exit Sub

4 Check2.Value = Unchecked
GoTo NextChecked3
Exit Sub

NextChecked3:
With Ini
    .File = App.Path & "\Settings.ini"
    GlobalChecked3 = .GetValue("CheckedDataSource", "Checked3")
    If GlobalChecked3 = 1 Then GoTo 5
    If GlobalChecked3 = 2 Then GoTo 6
End With
Exit Sub

5 Check3.Value = Checked
Exit Sub

6 Check3.Value = Unchecked
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

Dim NextChecked2
Dim NextChecked3

Dim Ini As clsIniFile
Set Ini = New clsIniFile

If Check1 = Checked Then GoTo 1
If Check1 = Unchecked Then GoTo 2
Exit Sub

1
With Ini
    .File = App.Path & "\Settings.ini"
    .WriteValue "DataSource", "logon", 1
End With
GoTo NextChecked2
Exit Sub

2
With Ini
    .File = App.Path & "\Settings.ini"
    .WriteValue "DataSource", "logon", 2
End With
GoTo NextChecked2
Exit Sub

NextChecked2:
If Check2 = Checked Then GoTo 3
If Check2 = Unchecked Then GoTo 4
Exit Sub

3
With Ini
    .File = App.Path & "\Settings.ini"
    .WriteValue "CheckedDataSource", "Checked2", 1
End With
GoTo NextChecked3
Exit Sub

4
With Ini
    .File = App.Path & "\Settings.ini"
    .WriteValue "CheckedDataSource", "Checked2", 2
End With
GoTo NextChecked3
Exit Sub

NextChecked3:
If Check3 = Checked Then GoTo 5
If Check3 = Unchecked Then GoTo 6
Exit Sub

5
With Ini
    .File = App.Path & "\Settings.ini"
    .WriteValue "CheckedDataSource", "Checked3", 1
End With
GoTo 7
Exit Sub

6
With Ini
    .File = App.Path & "\Settings.ini"
    .WriteValue "CheckedDataSource", "Checked3", 2
End With
GoTo 7
Exit Sub

7
Unload Me

End Sub
