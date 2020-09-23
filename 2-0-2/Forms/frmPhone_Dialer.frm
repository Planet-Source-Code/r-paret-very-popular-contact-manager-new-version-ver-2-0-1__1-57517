VERSION 5.00
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Begin VB.Form frmPhone_Dialer 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Credo Phone Dialer"
   ClientHeight    =   5460
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4455
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmPhone_Dialer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5460
   ScaleWidth      =   4455
   StartUpPosition =   2  'CenterScreen
   Begin Credo_Contact_Manager.XPButton XPButton_Exit 
      Height          =   495
      Left            =   3120
      TabIndex        =   20
      Top             =   4200
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Exit"
      ForeColor       =   255
      ForeHover       =   32768
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Clear 
      Height          =   495
      Left            =   3120
      TabIndex        =   19
      Top             =   3480
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Clear"
      ForeColor       =   0
      ForeHover       =   8388608
   End
   Begin Credo_Contact_Manager.XPButton XPButton_End 
      Height          =   495
      Left            =   3120
      TabIndex        =   18
      Top             =   2760
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "End"
      ForeColor       =   0
      ForeHover       =   8388608
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Dial 
      Height          =   495
      Left            =   3120
      TabIndex        =   17
      Top             =   2040
      Width           =   855
      _ExtentX        =   1508
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Dial"
      ForeColor       =   0
      ForeHover       =   8388608
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Square 
      Height          =   495
      Left            =   2160
      TabIndex        =   16
      Top             =   4200
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "#"
      ForeColor       =   0
      ForeHover       =   8388608
   End
   Begin Credo_Contact_Manager.XPButton XPButton_0 
      Height          =   495
      Left            =   1320
      TabIndex        =   15
      Top             =   4200
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "0"
      ForeColor       =   0
      ForeHover       =   8388608
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Star 
      Height          =   495
      Left            =   480
      TabIndex        =   14
      Top             =   4200
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "*"
      ForeColor       =   0
      ForeHover       =   8388608
   End
   Begin Credo_Contact_Manager.XPButton XPButton_9 
      Height          =   495
      Left            =   2160
      TabIndex        =   13
      Top             =   3480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "9"
      ForeColor       =   0
      ForeHover       =   8388608
   End
   Begin Credo_Contact_Manager.XPButton XPButton_8 
      Height          =   495
      Left            =   1320
      TabIndex        =   12
      Top             =   3480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "8"
      ForeColor       =   0
      ForeHover       =   8388608
   End
   Begin Credo_Contact_Manager.XPButton XPButton_7 
      Height          =   495
      Left            =   480
      TabIndex        =   11
      Top             =   3480
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "7"
      ForeColor       =   0
      ForeHover       =   8388608
   End
   Begin Credo_Contact_Manager.XPButton XPButton_6 
      Height          =   495
      Left            =   2160
      TabIndex        =   10
      Top             =   2760
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "6"
      ForeColor       =   0
      ForeHover       =   8388608
   End
   Begin Credo_Contact_Manager.XPButton XPButton_5 
      Height          =   495
      Left            =   1320
      TabIndex        =   9
      Top             =   2760
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "5"
      ForeColor       =   8388608
      ForeHover       =   0
   End
   Begin Credo_Contact_Manager.XPButton XPButton_4 
      Height          =   495
      Left            =   480
      TabIndex        =   8
      Top             =   2760
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "4"
      ForeColor       =   0
      ForeHover       =   8388608
   End
   Begin Credo_Contact_Manager.XPButton XPButton_3 
      Height          =   495
      Left            =   2160
      TabIndex        =   7
      Top             =   2040
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "3"
      ForeColor       =   0
      ForeHover       =   8388608
   End
   Begin Credo_Contact_Manager.XPButton XPButton_2 
      Height          =   495
      Left            =   1320
      TabIndex        =   6
      Top             =   2040
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "2"
      ForeColor       =   0
      ForeHover       =   8388608
   End
   Begin Credo_Contact_Manager.XPButton XPButton_1 
      Height          =   495
      Left            =   480
      TabIndex        =   5
      Top             =   2040
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "1"
      ForeColor       =   0
      ForeHover       =   8388608
   End
   Begin VB.Timer Timer2 
      Interval        =   1
      Left            =   5400
      Top             =   1680
   End
   Begin VB.Timer Timer1 
      Interval        =   5
      Left            =   5400
      Top             =   2160
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   5760
      TabIndex        =   4
      Top             =   1320
      Width           =   975
   End
   Begin MSCommLib.MSComm MSComm1 
      Left            =   5280
      Top             =   2640
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DTREnable       =   -1  'True
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000013&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   -10
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Status:"
      Top             =   5160
      Width           =   5055
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      BackColor       =   &H80000018&
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   480
      MaxLength       =   12
      TabIndex        =   1
      Top             =   1440
      Width           =   3495
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1000
      Left            =   0
      ScaleHeight     =   1005
      ScaleWidth      =   4455
      TabIndex        =   0
      Top             =   0
      Width           =   4455
      Begin VB.Image Image4 
         Height          =   375
         Left            =   2880
         MouseIcon       =   "frmPhone_Dialer.frx":0ECA
         MousePointer    =   99  'Custom
         Picture         =   "frmPhone_Dialer.frx":11D4
         ToolTipText     =   " Move "
         Top             =   290
         Width           =   375
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H00FFFFFF&
         Height          =   735
         Left            =   120
         Top             =   120
         Width           =   4215
      End
      Begin VB.Image Image3 
         Height          =   735
         Left            =   240
         Picture         =   "frmPhone_Dialer.frx":1982
         Stretch         =   -1  'True
         Top             =   150
         Width           =   735
      End
      Begin VB.Image Image2 
         Height          =   375
         Left            =   3360
         Picture         =   "frmPhone_Dialer.frx":284C
         ToolTipText     =   " Minimize "
         Top             =   290
         Width           =   375
      End
      Begin VB.Image Image1 
         Height          =   375
         Left            =   3840
         Picture         =   "frmPhone_Dialer.frx":2FFA
         ToolTipText     =   " Close "
         Top             =   290
         Width           =   375
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Phone Dialer"
         BeginProperty Font 
            Name            =   "Comic Sans MS"
            Size            =   14.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   1080
         TabIndex        =   2
         Top             =   300
         Width           =   1695
      End
   End
   Begin VB.Shape Shape3 
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   3735
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1200
      Width           =   3975
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   3975
      Left            =   120
      Top             =   1080
      Width           =   4215
   End
End
Attribute VB_Name = "frmPhone_Dialer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


DefInt A-Z
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wparam As Integer, ByVal iparam As Long) As Long
Dim CancelFlag

Public Sub FormDrag(Theform As Form)
    
    ReleaseCapture
    Call SendMessage(Theform.hwnd, &HA1, 2, 0&)

End Sub

Private Sub Command9_Click()

Text1.Text = Text1.Text & "9"

End Sub

Private Sub Form_Unload(Cancel As Integer)

If MSComm1.PortOpen = True Then

MSComm1.PortOpen = False

End If

End Sub

Private Sub Image1_Click()

Unload Me

End Sub

Private Sub Image2_Click()

Me.WindowState = 1

End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

FormDrag Me

End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

FormDrag Me

End Sub

 Sub Dial(number$)
 
 On Error GoTo NODI
 
  Dim DialString$, FromModem$, dummy
    DialString$ = "ATDT" + number$ + ";" + vbCr
     MSComm1.CommPort = Text3.Text
    MSComm1.Settings = "9600,N,8,1"
    
    On Error Resume Next
    
    If MSComm1.PortOpen = True Then MSComm1.PortOpen = False
    MSComm1.PortOpen = True
    If Err Then
       MsgBox "COM2: not available. Change the CommPort property to another port."
       Exit Sub
    End If
    MSComm1.InBufferCount = 0
    MSComm1.Output = DialString$
    Exit Sub
    
NODI:
    Text2.Text = " Status: No Dial Tone..."
    
End Sub

Private Sub Timer1_Timer()

If Text1.Text = "" Then
XPButton_Dial.Enabled = False
Else: XPButton_Dial.Enabled = True
End If

End Sub

Private Sub Timer2_Timer()

Dim PortErr

Text2.Text = " Status: Please Wait Loading...."

On Error GoTo PortErr

Dim port, X, instring
port = 1
PortinG:
MSComm1.CommPort = port
MSComm1.PortOpen = True

frmPhone_Dialer.MSComm1.Settings = "9600,N,8,1"
    MSComm1.Output = "AT" + Chr$(13)
    X = 1

    Do: DoEvents
        X = X + 1
        If X = 1000 Then MSComm1.Output = "AT" + Chr$(13)
        If X = 2000 Then MSComm1.Output = "AT" + Chr$(13)
        If X = 3000 Then MSComm1.Output = "AT" + Chr$(13)
        If X = 4000 Then MSComm1.Output = "AT" + Chr$(13)
        If X = 5000 Then MSComm1.Output = "AT" + Chr$(13)
        If X = 6000 Then MSComm1.Output = "AT" + Chr$(13)

        If X = 7000 Then
            MSComm1.PortOpen = False
            port = port + 1
            GoTo PortinG:

            If MSComm1.CommPort >= 5 Then
errr:
                MsgBox "Can't Find Modem!"
                GoTo done:
            End If
        End If
    Loop Until MSComm1.InBufferCount >= 2
    instring = MSComm1.Input
    MSComm1.PortOpen = False

  Text3.Text = port
done:
Timer2.Enabled = False
Text2.Text = " Status:"
Exit Sub

PortErr:
MsgBox "No modem detected!"
Unload Me
Exit Sub

End Sub

Private Sub Form_Load()

MSComm1.InputLen = 0

End Sub

Private Sub XPButton_0_Click()

Text1.Text = Text1.Text & "0"

End Sub

Private Sub XPButton_1_Click()

Text1.Text = Text1.Text & "1"

End Sub

Private Sub XPButton_2_Click()

Text1.Text = Text1.Text & "2"

End Sub

Private Sub XPButton_3_Click()

Text1.Text = Text1.Text & "3"

End Sub

Private Sub XPButton_4_Click()

Text1.Text = Text1.Text & "4"

End Sub

Private Sub XPButton_5_Click()

Text1.Text = Text1.Text & "5"

End Sub

Private Sub XPButton_6_Click()

Text1.Text = Text1.Text & "6"

End Sub

Private Sub XPButton_7_Click()

Text1.Text = Text1.Text & "7"

End Sub

Private Sub XPButton_8_Click()

Text1.Text = Text1.Text & "8"

End Sub

Private Sub XPButton_9_Click()

Text1.Text = Text1.Text & "9"

End Sub

Private Sub XPButton_Dial_Click()

Dial Text1.Text
'XPButton_2.Enabled = True
'XPButton_2.Enabled = True
Text2.Text = " Status: Dialing " & Text1 & "...."

End Sub

Private Sub XPButton_End_Click()

'### On Error Resume Next

'MSComm1.Output = "ATH" + vbCr
'MSComm1.PortOpen = False
Unload Me

End Sub

Private Sub XPButton_square_Click()

Text1.Text = Text1.Text & "#"

End Sub

Private Sub XPButton_star_Click()

Text1.Text = Text1.Text & "*"

End Sub

Private Sub XPButton_Exit_Click()

On Error Resume Next

MSComm1.Output = "ATH" + vbCr
MSComm1.PortOpen = False

Unload Me

End Sub

Private Sub XPButton_Clear_Click()

Text1.Text = ""

End Sub
