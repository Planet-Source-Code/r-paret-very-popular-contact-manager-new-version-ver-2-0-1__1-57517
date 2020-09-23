VERSION 5.00
Begin VB.Form frmNotes 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Credo Contact Manager - Notes"
   ClientHeight    =   6000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8070
   Icon            =   "frmNotes.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6000
   ScaleWidth      =   8070
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text_Notes 
      Height          =   4215
      Left            =   240
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   1080
      Width           =   7575
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Edit 
      Height          =   375
      Left            =   5520
      TabIndex        =   4
      Top             =   5400
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
      Caption         =   "Edit"
      ForeColor       =   -2147483642
      ForeHover       =   12582912
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Save 
      Height          =   375
      Left            =   5520
      TabIndex        =   3
      Top             =   5400
      Visible         =   0   'False
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
      Caption         =   "Save"
      ForeColor       =   -2147483642
      ForeHover       =   12582912
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Close 
      Height          =   375
      Left            =   6720
      TabIndex        =   2
      Top             =   5400
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
      Caption         =   "Close"
      ForeColor       =   32768
      ForeHover       =   255
   End
   Begin VB.Image Image4 
      Height          =   375
      Left            =   6480
      MouseIcon       =   "frmNotes.frx":030A
      MousePointer    =   99  'Custom
      Picture         =   "frmNotes.frx":0614
      ToolTipText     =   " Move "
      Top             =   240
      Width           =   375
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   6960
      Picture         =   "frmNotes.frx":0DC2
      ToolTipText     =   " Minimize "
      Top             =   240
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   7440
      Picture         =   "frmNotes.frx":1570
      ToolTipText     =   " Close "
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label_Notes 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Notes"
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
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
   Begin VB.Image Image1 
      Height          =   510
      Left            =   240
      Picture         =   "frmNotes.frx":1D1E
      Top             =   160
      Width           =   510
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   4935
      Left            =   120
      Top             =   960
      Width           =   7815
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   7815
   End
End
Attribute VB_Name = "frmNotes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wparam As Integer, ByVal iparam As Long) As Long


Public Sub FormDrag(Theform As Form)
    
    ReleaseCapture
    Call SendMessage(Theform.hwnd, &HA1, 2, 0&)

End Sub

Private Sub Form_Initialize()

Text_Notes.Text = MakeEnterAddLines()

 Dim ConvSpace
  If TextChanging Then Exit Sub
  TextChanging = True
  ConvSpace = Text_Notes.SelStart
  Text_Notes.Text = Replace(Text_Notes.Text, "¬", " ")
  If ConvSpace > Len(Text_Notes.Text) Then
    Text_Notes.SelStart = Len(Text_Notes.Text)
  Else
    Text_Notes.SelStart = ConvSpace
  End If

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

Private Sub Text_Notes_Change()


 Dim ConvSpace
  If TextChanging Then Exit Sub
  TextChanging = True
  ConvSpace = Text_Notes.SelStart
  Text_Notes.Text = Replace(Text_Notes.Text, "¬", " ")
  If ConvSpace > Len(Text_Notes.Text) Then
    Text_Notes.SelStart = Len(Text_Notes.Text)
  Else
    Text_Notes.SelStart = ConvSpace
  End If
  
  TextChanging = False
  
  
End Sub

Private Sub Text_Notes_KeyDown(KeyCode As Integer, Shift As Integer)

'MakeEnterAddLines

  Select Case KeyCode
           ' If user presses TAB, ENTER, PAGE UP, PAGE DOWN
           Case 13
               ' Disable the keystroke by setting it to 0
               KeyCode = 0
          frmNotes.Text_Notes.Locked = True
           MsgBox "Use spase bar to go to next line!"
          frmNotes.Text_Notes.Locked = False
           Case Else
               Debug.Print KeyCode, Shift
       End Select

End Sub

Private Sub XPButton_Close_Click()

Unload Me

End Sub

Private Sub XPButton_Edit_Click()


XPButton_Edit.Visible = False
XPButton_Save.Visible = True

frmNotes.Text_Notes.Locked = False
frmNotes.Text_Notes.BackColor = &H80000005

End Sub

Private Sub XPButton_Save_Click()

Dim Ini As clsIniFile
Set Ini = New clsIniFile

 Dim ConvSpace
  If TextChanging Then Exit Sub
  TextChanging = True
  ConvSpace = Text_Notes.SelStart
  Text_Notes.Text = Replace(Text_Notes.Text, " ", "¬")
  If ConvSpace > Len(Text_Notes.Text) Then
    Text_Notes.SelStart = Len(Text_Notes.Text)
  Else
    Text_Notes.SelStart = ConvSpace
  End If
  
  TextChanging = False
  
With Ini
    .File = App.Path & "\Data\db\CredoData\" & Trim(Address_Contact_Manager.List1.Text)
    .WriteValue "Notes", "Notes", (Text_Notes.Text)
End With

XPButton_Edit.Visible = True
XPButton_Save.Visible = False

frmNotes.Text_Notes.Locked = True
frmNotes.Text_Notes.Appearance = 0


End Sub
