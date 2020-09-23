VERSION 5.00
Begin VB.Form Register 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Register Credo Contact Manager"
   ClientHeight    =   7590
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6270
   ControlBox      =   0   'False
   Icon            =   "Register.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7590
   ScaleWidth      =   6270
   StartUpPosition =   2  'CenterScreen
   Begin Credo_Contact_Manager.XPButton XPButton_Pay 
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   6840
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Pay For This Software && Get Activationl Key"
      ForeColor       =   32768
      ForeHover       =   33023
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Activate 
      Height          =   285
      Left            =   4800
      TabIndex        =   11
      Top             =   5880
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Activate"
      ForeColor       =   -2147483642
      ForeHover       =   12582912
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Cancel 
      Height          =   375
      Left            =   5040
      TabIndex        =   14
      Top             =   6360
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
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
   Begin Credo_Contact_Manager.XPButton XPButton_Get_Activation_Key 
      Height          =   375
      Left            =   2880
      TabIndex        =   13
      Top             =   6360
      Width           =   2055
      _ExtentX        =   3625
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Get Activation Key"
      ForeColor       =   -2147483642
      ForeHover       =   12582912
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Continue_Unregistered 
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   6360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Continue Unregistered"
      ForeColor       =   -2147483642
      ForeHover       =   12582912
   End
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   4800
      Top             =   1200
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Height          =   4575
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   5775
      Begin VB.Image Image1 
         Height          =   1575
         Left            =   3000
         Picture         =   "Register.frx":030A
         Top             =   1080
         Width           =   1500
      End
      Begin VB.Image Image2 
         Height          =   870
         Left            =   1320
         Picture         =   "Register.frx":1371
         Top             =   720
         Width           =   1680
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C0FFFF&
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00004040&
         Height          =   2535
         Left            =   720
         Shape           =   4  'Rounded Rectangle
         Top             =   240
         Width           =   3975
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "NOT ACTIVATED"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   480
         TabIndex        =   5
         Top             =   3360
         Width           =   4695
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   "Your 35 uses of Credo Contact Manager has expired, You must activate this application!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   570
         Left            =   480
         TabIndex        =   8
         Top             =   3840
         Visible         =   0   'False
         Width           =   4980
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FF00&
         Height          =   255
         Left            =   840
         TabIndex        =   7
         Top             =   3960
         Width           =   495
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "You have         more uses before you must activate Credo Contact Manager."
         ForeColor       =   &H0000FF00&
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   3960
         Width           =   5535
      End
      Begin VB.Shape Shape5 
         BorderColor     =   &H00004040&
         FillColor       =   &H00008080&
         FillStyle       =   0  'Solid
         Height          =   2535
         Left            =   1080
         Shape           =   4  'Rounded Rectangle
         Top             =   600
         Width           =   3975
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00800000&
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   5640
      Width           =   5775
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   2040
         TabIndex        =   0
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00800000&
         BackStyle       =   0  'Transparent
         Caption         =   "Enter ActivationKey:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   0
         TabIndex        =   9
         Top             =   240
         Width           =   1935
      End
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Activate"
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
      TabIndex        =   10
      Top             =   240
      Width           =   4335
   End
   Begin VB.Image Image5 
      Height          =   375
      Left            =   5160
      MouseIcon       =   "Register.frx":21D0
      MousePointer    =   99  'Custom
      Picture         =   "Register.frx":24DA
      ToolTipText     =   " Move "
      Top             =   240
      Width           =   375
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   5640
      Picture         =   "Register.frx":2C88
      ToolTipText     =   " Close "
      Top             =   240
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   6015
   End
   Begin VB.Image Image3 
      Height          =   510
      Left            =   240
      Picture         =   "Register.frx":3436
      Top             =   180
      Width           =   390
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00FFFFFF&
      Height          =   6495
      Left            =   120
      Top             =   960
      Width           =   6015
   End
   Begin VB.Shape Shape3 
      Height          =   495
      Left            =   2640
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000C&
      Height          =   375
      Left            =   1080
      TabIndex        =   2
      Top             =   7560
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H8000000C&
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   7560
      Visible         =   0   'False
      Width           =   2175
   End
End
Attribute VB_Name = "Register"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'This is one way to make a full application, a trial use app.
'I agree that this may not be the best way to do a trial use
'but the way I use it in  my applications it is very effective
'________________________________________________________________________
'To set how many times you want to allow them to use before they register
'Open the Program.ini  file   (it will look like this)
'------------------------------------------------------------------------
'    | 20                                     |
'    | This application is not registered       |
'    | -------------WARNING--------------       |
'    | Do not change or modify this .ini file "" |
'    | " " " " and second line of warning " " " |
'------------------------------------------------------------------------
'The number 20 is how many time you want to give them
'If you want to allow them 30 times change the 20 to 30
'YOU MUST SAVE THE .ini FILE AFTER YOU MAKE ANY CHANGES TO IT!
' You can open and modify the Program.ini file with Notepad!!!!!!!!

'API FUNCTION DECLARATION TO GO TO A WEB PAGE
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_MAXIMIZE = 3

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wparam As Integer, ByVal iparam As Long) As Long

Public Sub FormDrag(Theform As Form)
    
    ReleaseCapture
    Call SendMessage(Theform.hwnd, &HA1, 2, 0&)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

If Label4.Caption = "This application Is Not activated" Then

Dim frm As Form
    For Each frm In Forms
        If frm.Name <> "frmMain" Then Unload frm
    Next frm
    Unload Register
End If

End Sub

Private Sub XPButton_Cancel_Click()

Unload Me 'Quits program

End Sub
Private Sub XPButton_Continue_Unregistered_Click()

If Text1.Text = "" Then
'-------Below rewrites the Program.ini to make it current.-------
Open App.Path & "\Program.ini" For Output As #1
Print #1, Label3
Print #1, Label4
Print #1, "----------------------------WARNING-----------------------------"
Print #1, "DO NOT CHANGE OR MODIFY THIS .dll FILE IT IS NEEDED"
Print #1, "FOR AN APPLICATION TO WORK CORRECTLY (IF CHANGED APP WONT WORK)"
Close #1
'-------Above rewrites the Program.ini to make it current.-------
On Error Resume Next
Unload Register
frmLogin2.Show

Exit Sub
End If

End Sub
Private Sub XPButton_Activate_Click()

If Text1.Text = "CCM2017731" Then 'This is the register code. You can change CCM2017731 to what you want the code to be.
Label4.Caption = "Activated" 'This sets the form up to write the Program.ini that this is registered
'-------Below rewrites the Program.ini to make it current.-------
Open App.Path & "\Program.ini" For Output As #1
Print #1, Label3
Print #1, Label4
Print #1, "----------------------------WARNING-----------------------------"
Print #1, "DO NOT CHANGE OR MODIFY THIS .dll FILE IT IS NEEDED"
Print #1, "FOR AN APPLICATION TO WORK CORRECTLY (IF CHANGED APP WONT WORK)"
Close #1
'-------Above rewrites the Program.ini to make it current.-------
Else 'The number they entered was incorrect
MsgBox "The activation key is incorrect! Try again. ", vbCritical, " Incorrect activation Key"
Label4.Caption = "This application is not activated"
Text1.Text = ""
Text1.SetFocus
'-------Above rewrites the Program.ini to make it current.-------
Exit Sub 'Because the number was wrong this stops them from continuing
End If
Timer1.Enabled = False
On Error Resume Next
Unload Register
frmLogin2.Show

End Sub

Private Sub XPButton_Get_Activation_Key_Click()

Dim strURL As String
    strURL = "http://www.credo-web.co.uk:90/Credo_C-M/credo_c-m/pages/pages/reg_activate2"

ShellExecute 0, "open", strURL, vbNullString, vbNullString, SW_MAXIMIZE

End Sub

Private Sub Form_Load()

On Error Resume Next

'-------below reads the Program.ini and loads the labels with its info-----
Open App.Path & "\Program.ini" For Input As #1
Line Input #1, A
Line Input #1, B
Label1.Caption = A
Label2.Caption = B
Close #1
If Label1.Caption = 0 Then
Label5.Visible = False
Label3.Visible = False
End If
'-------above reads the Program.ini and loads the labels with its info
Label3.Caption = Label1.Caption - 1
'----------------------------------------------
If Label3.Caption = "-1" Then
Label6.Visible = True       'Makes sure the registry info.txt does not go below 0
Label3.Caption = "-1"
End If
'----------------------------------------------
Label4.Caption = Label2.Caption
'----------------------------------------------
If Label4.Caption = "Activated" Then GoTo 2
If XPButton_Continue_Unregistered Then GoTo 1

    'Checks on load if the Reg.ini (b) = Registerd  if so skip register form
Exit Sub

1
frmLogin2.Show
Exit Sub

2
Unload Register
frmLogin2.Show
Exit Sub

End Sub

Private Sub Image4_Click()

Unload Me 'Quits program

End Sub

Private Sub Image5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

FormDrag Me

End Sub

Private Sub Image5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

FormDrag Me

End Sub

'looks for user input and enables and disabled the appropriate buttons
Private Sub Text1_Change()

If Text1.Text = "" Then
XPButton_Activate.Enabled = False
XPButton_Continue_Unregistered.Enabled = True
Else
XPButton_Activate.Enabled = True
XPButton_Continue_Unregistered.Enabled = False
End If

End Sub
'checks to see if registered, if not and trial use is up then it shows the (trial use is over must register)
Private Sub Timer1_Timer()

If Label6.Visible = True Then
XPButton_Continue_Unregistered.Enabled = False
End If
If Label4.Caption = "Activated" Then
Timer1.Enabled = False
frmSplash.Show
Unload Register
End If

End Sub

Private Sub XPButton_Pay_Click()

Dim strURL As String
    strURL = "http://www.credo-web.co.uk:90/Credo_C-M/credo_c-m/pages/buy"

ShellExecute 0, "open", strURL, vbNullString, vbNullString, SW_MAXIMIZE

End Sub
