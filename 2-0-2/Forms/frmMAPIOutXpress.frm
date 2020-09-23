VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMAPIOutXpress 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Credo Instnd Mailer"
   ClientHeight    =   7725
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9765
   Icon            =   "frmMAPIOutXpress.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   9765
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   360
      Top             =   6360
   End
   Begin VB.TextBox txtAttachment 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   2520
      Visible         =   0   'False
      Width           =   5175
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   360
      Top             =   5640
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtMessage 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4605
      Left            =   1560
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   2760
      Width           =   7815
   End
   Begin VB.TextBox txtSubject 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   1
      Top             =   2160
      Width           =   5655
   End
   Begin MSMAPI.MAPIMessages MAPIMessage1 
      Left            =   360
      Top             =   5160
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin VB.TextBox txtSendTo 
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1560
      TabIndex        =   0
      Top             =   1680
      Width           =   5655
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   360
      Top             =   4680
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   0   'False
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
   Begin VB.Image Image_Exit_1 
      Height          =   345
      Left            =   2760
      Picture         =   "frmMAPIOutXpress.frx":11C2
      Top             =   1080
      Width           =   1035
   End
   Begin VB.Image Image_Exit_2 
      Height          =   345
      Left            =   2760
      Picture         =   "frmMAPIOutXpress.frx":24B4
      ToolTipText     =   " Close - Exit "
      Top             =   1080
      Width           =   1035
   End
   Begin VB.Image Image_Send_1 
      Height          =   345
      Left            =   360
      Picture         =   "frmMAPIOutXpress.frx":37A6
      Top             =   1080
      Width           =   1035
   End
   Begin VB.Image Image_Send_2 
      Height          =   345
      Left            =   360
      Picture         =   "frmMAPIOutXpress.frx":4A98
      ToolTipText     =   " Send Email "
      Top             =   1080
      Width           =   1035
   End
   Begin VB.Image Image_New_1 
      Height          =   345
      Left            =   1560
      Picture         =   "frmMAPIOutXpress.frx":5D8A
      ToolTipText     =   " "
      Top             =   1080
      Width           =   1035
   End
   Begin VB.Image Image_New_2 
      Height          =   345
      Left            =   1560
      Picture         =   "frmMAPIOutXpress.frx":707C
      ToolTipText     =   " New Email "
      Top             =   1080
      Width           =   1035
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Send Email"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3480
      TabIndex        =   9
      Top             =   165
      Width           =   2415
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   8160
      MouseIcon       =   "frmMAPIOutXpress.frx":836E
      MousePointer    =   99  'Custom
      Picture         =   "frmMAPIOutXpress.frx":8678
      ToolTipText     =   " Move "
      Top             =   240
      Width           =   375
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   8640
      Picture         =   "frmMAPIOutXpress.frx":8E26
      ToolTipText     =   " Minimize "
      Top             =   240
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   9120
      Picture         =   "frmMAPIOutXpress.frx":95D4
      ToolTipText     =   " Close - Exit "
      Top             =   240
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   240
      Picture         =   "frmMAPIOutXpress.frx":9D82
      Top             =   195
      Width           =   495
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   6615
      Left            =   120
      Top             =   960
      Width           =   9495
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   9495
   End
   Begin VB.Label lblAuthor 
      AutoSize        =   -1  'True
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   1560
      TabIndex        =   8
      Top             =   5160
      Width           =   60
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Attachment:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   240
      TabIndex        =   7
      Top             =   2520
      Visible         =   0   'False
      Width           =   1185
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Message:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   480
      TabIndex        =   5
      Top             =   2880
      Width           =   930
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Subject:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   600
      TabIndex        =   4
      Top             =   2160
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "To:"
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   210
      Left            =   1080
      TabIndex        =   3
      Top             =   1680
      Width           =   315
   End
   Begin VB.Label Label_Mouse_Move 
      BackStyle       =   0  'Transparent
      Height          =   975
      Left            =   0
      TabIndex        =   10
      Top             =   720
      Width           =   4335
   End
End
Attribute VB_Name = "frmMAPIOutXpress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim iStart As Integer
Dim iEnd As Integer

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wparam As Integer, ByVal iparam As Long) As Long

Public Sub FormDrag(Theform As Form)
    
    ReleaseCapture
    Call SendMessage(Theform.hwnd, &HA1, 2, 0&)

End Sub


Private Sub cmdEnd_Click()

    Unload Me
    
End Sub

Private Sub cmdSend_Click()

    MAPISession1.SignOn
    MAPISession1.DownLoadMail = True
DoEvents
        
        MAPIMessage1.SessionID = MAPISession1.SessionID
        MAPIMessage1.Compose
    
        MAPIMessage1.RecipAddress = txtSendTo
        MAPIMessage1.ResolveName
        MAPIMessage1.MsgSubject = txtSubject
        
        
        
                MAPIMessage1.Send False
                
MAPISession1.SignOff

    MsgBox "The message has been sent...", vbInformation, "Credo Instnd Mailer"
    
End Sub

Private Sub Form_Load()

frmMAPIOutXpress.Height = 7725
frmMAPIOutXpress.Width = 9765

End Sub

Private Sub Image_Exit_1_Click()

Unload Me

End Sub

Private Sub Image_Exit_1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Image_Exit_1.Visible = False
Image_Exit_2.Visible = True

End Sub

Private Sub Image_Exit_2_Click()

Unload Me

End Sub

Private Sub Image_New_1_Click()

Dim ctl As Control

Image_New_1.Visible = False
Image_New_2.Visible = True

  For Each ctl In Me.Controls
            If TypeOf ctl Is TextBox Then
                ctl.Text = ""
            End If
        Next
        txtSendTo.SetFocus
End Sub

Private Sub Image_New_1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Image_New_1.Visible = False
Image_New_2.Visible = True

End Sub

Private Sub Image_New_2_Click()

Dim ctl As Control

Image_New_1.Visible = False
Image_New_2.Visible = True

  For Each ctl In Me.Controls
            If TypeOf ctl Is TextBox Then
                ctl.Text = ""
            End If
        Next
        txtSendTo.SetFocus
End Sub

Private Sub Image_Send_1_Click()

Dim ctl As Control

On Error GoTo 2

     If txtSendTo.Text = "" Then GoTo 1
    MAPISession1.SignOn
    MAPISession1.DownLoadMail = False
     DoEvents
                    
   MAPIMessage1.SessionID = MAPISession1.SessionID
   MAPIMessage1.Compose
             
   MAPIMessage1.RecipAddress = txtSendTo
   MAPIMessage1.ResolveName
                    
   MAPIMessage1.MsgSubject = txtSubject
   
   MAPIMessage1.MsgNoteText = txtMessage
                                        
   MAPIMessage1.Send False
                            
   MAPISession1.SignOff
   
   For Each ctl In Me.Controls
          If TypeOf ctl Is TextBox Then
              ctl.Text = ""
            End If
        Next
        txtSendTo.SetFocus
            
    MsgBox "The message has been sent...", vbInformation, "Credo Instnd Mailer"
          Exit Sub

1   MsgBox "Please enter recipient email address!", , "Credo Instnd Mailer"
          Exit Sub
          
2  msgBoxNotInstalled.Show
          Exit Sub

End Sub

Private Sub Image_Send_1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Image_Send_1.Visible = False
Image_Send_2.Visible = True

End Sub

Private Sub Image_Send_2_Click()

Dim ctl As Control

On Error GoTo 2

     If txtSendTo.Text = "" Then GoTo 1
    MAPISession1.SignOn
    MAPISession1.DownLoadMail = False
     DoEvents
                    
   MAPIMessage1.SessionID = MAPISession1.SessionID
   MAPIMessage1.Compose
             
   MAPIMessage1.RecipAddress = txtSendTo
   MAPIMessage1.ResolveName
                    
   MAPIMessage1.MsgSubject = txtSubject
   
   MAPIMessage1.MsgNoteText = txtMessage
                    
   MAPIMessage1.Send False
                            
   MAPISession1.SignOff
   
   For Each ctl In Me.Controls
            If TypeOf ctl Is TextBox Then
                ctl.Text = ""
            End If
        Next
        txtSendTo.SetFocus
            
    MsgBox "The message has been sent...", vbInformation, "Credo Instnd Mailer"
          Exit Sub

1   MsgBox "Please enter recipient email address!", , "Credo Instnd Mailer"
          Exit Sub

2  msgBoxNotInstalled.Show
          Exit Sub

End Sub

Private Sub Image2_Click()

Unload Me

End Sub

Private Sub Image3_Click()

Me.WindowState = 1

End Sub

Private Sub Image4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

FormDrag Me

End Sub

Private Sub Image4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

FormDrag Me

End Sub

Private Sub Image5_Click()

On Error GoTo 1
Dim ReturnValue
If CheckBoxConstants.vbChecked Then
ReturnValue = Shell("C:\Program Files\Credo\Credo_Equine_Man\Modules\Help\Credo_EM_Help1.exe", 1)

Else
ReturnValue = Shell("C:\Program Files\Credo\Credo_Equine_Man\Modules\Help\Credo_EM_Help1.EXE", 1)
End If
Exit Sub

1 msgBoxNotInstalled.Show
Resume Next

End Sub

Private Sub Label_Mouse_Move_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

Image_Send_1.Visible = True
Image_Send_2.Visible = False
Image_New_1.Visible = True
Image_New_2.Visible = False
Image_Exit_1.Visible = True
Image_Exit_2.Visible = False

End Sub

Private Sub txtSendTo_GotFocus()

    txtSendTo.SelStart = 0
    txtSendTo.SelLength = Len(txtSendTo)
    
End Sub
