VERSION 5.00
Begin VB.Form BackupRestore 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Credo Bacup Restore"
   ClientHeight    =   8760
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7470
   DrawStyle       =   1  'Dash
   Icon            =   "Backup-Restore.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8760
   ScaleWidth      =   7470
   StartUpPosition =   2  'CenterScreen
   Begin VB.DriveListBox Drive2 
      Height          =   315
      Left            =   600
      TabIndex        =   31
      Top             =   1800
      Visible         =   0   'False
      Width           =   2415
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Ok 
      Height          =   300
      Left            =   5640
      TabIndex        =   30
      Top             =   3240
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
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
   Begin VB.DirListBox Dir1 
      Height          =   315
      Left            =   3000
      TabIndex        =   29
      Top             =   1800
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   600
      TabIndex        =   28
      Top             =   3000
      Visible         =   0   'False
      Width           =   6255
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Backup_Browse 
      Height          =   300
      Left            =   600
      TabIndex        =   27
      Top             =   3360
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Browse"
      ForeColor       =   -2147483642
      ForeHover       =   12582912
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Restore_Browse 
      Height          =   300
      Left            =   600
      TabIndex        =   26
      Top             =   2160
      Visible         =   0   'False
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   529
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Browse"
      ForeColor       =   -2147483642
      ForeHover       =   12582912
   End
   Begin VB.TextBox txtPattern 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   6360
      TabIndex        =   23
      Top             =   5400
      Width           =   855
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "d/MM/yyyy"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   5129
         SubFormatType   =   3
      EndProperty
      Height          =   285
      Left            =   4680
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   6360
      Width           =   2655
   End
   Begin VB.CheckBox chkRecourse 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Include Subdirs even if empty /E"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   21
      Top             =   5640
      Value           =   1  'Checked
      Width           =   7215
   End
   Begin VB.CheckBox chkAttributeDont 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Only files with Attribute bit set, don't change /A"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   20
      Top             =   5895
      Value           =   1  'Checked
      Width           =   7215
   End
   Begin VB.CheckBox chkAttributeTurnOff 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Only files with Attribute bit set, turn off attribute bit /M"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   19
      Top             =   6150
      Value           =   1  'Checked
      Width           =   7215
   End
   Begin VB.CheckBox chkDate 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Only files changed on or after date: /D:m-d-y"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   18
      Top             =   6390
      Value           =   1  'Checked
      Width           =   4575
   End
   Begin VB.CheckBox chkHidden 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Copies hidden and system files also.  /H"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   17
      Top             =   6645
      Value           =   1  'Checked
      Width           =   7215
   End
   Begin VB.CheckBox chkReadOnly 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Overwrites read-only files  /R"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   16
      Top             =   6900
      Value           =   1  'Checked
      Width           =   7215
   End
   Begin VB.CheckBox chkCopyifExisting 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Copies only files which already exist in destination /U"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   7155
      Value           =   1  'Checked
      Width           =   7215
   End
   Begin VB.CheckBox chkExclude 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Excludes named files /EXCLUDE:File"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   7410
      Value           =   1  'Checked
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   3600
      TabIndex        =   13
      Text            =   "Text2"
      Top             =   7320
      Width           =   3615
   End
   Begin VB.CheckBox chkOverCopy 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Doesn't prompt before overcopying /Y"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   7650
      Value           =   1  'Checked
      Width           =   7215
   End
   Begin VB.CheckBox chkSlashK 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Copies file attributes (doesn't reset ReadOnly) /K"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   7905
      Value           =   1  'Checked
      Width           =   7215
   End
   Begin VB.CheckBox chkSubDir 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Copies subdirectories) /S"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   8160
      Value           =   1  'Checked
      Width           =   7215
   End
   Begin VB.CheckBox chkSlashT 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      Caption         =   "Creates subdirectories but doesn't copy files /T"
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   8400
      Value           =   1  'Checked
      Width           =   7215
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Backup_Restore 
      Height          =   375
      Left            =   600
      TabIndex        =   7
      Top             =   4320
      Width           =   3135
      _extentx        =   5318
      _extenty        =   661
      font            =   "Backup-Restore.frx":030A
      caption         =   "Backup - Restore"
      forecolor       =   -2147483642
      forehover       =   12582912
   End
   Begin VB.OptionButton Option_Restore 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Restore"
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
      Height          =   255
      Left            =   2520
      TabIndex        =   6
      Top             =   3840
      Width           =   1335
   End
   Begin VB.OptionButton Option_Backup 
      BackColor       =   &H00C0FFFF&
      Caption         =   "Backup"
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
      Height          =   255
      Left            =   720
      TabIndex        =   5
      Top             =   3840
      Width           =   1215
   End
   Begin VB.TextBox Text_Restore_Path 
      Height          =   285
      Left            =   600
      TabIndex        =   2
      Text            =   " Choose From Backup or Restore Option to Enter Path"
      Top             =   3000
      Width           =   6255
   End
   Begin VB.TextBox Text_Backup_Path 
      Height          =   285
      Left            =   600
      TabIndex        =   1
      Text            =   " Choose From Backup or Restore Option to Enter Path"
      Top             =   1800
      Width           =   6255
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Exit 
      Height          =   375
      Left            =   5640
      TabIndex        =   0
      Top             =   4320
      Width           =   1215
      _extentx        =   2566
      _extenty        =   661
      font            =   "Backup-Restore.frx":0336
      caption         =   "Exit"
      forecolor       =   255
      forehover       =   49152
   End
   Begin VB.Label Label_Msg 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   495
      Left            =   1440
      TabIndex        =   32
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Label LabelCFM 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "&Copy files matching:"
      ForeColor       =   &H80000008&
      Height          =   315
      Left            =   4680
      TabIndex        =   25
      Top             =   5400
      Width           =   1725
   End
   Begin VB.Label Label_Backup_Err 
      Alignment       =   2  'Center
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   1200
      TabIndex        =   24
      Top             =   1080
      Width           =   5655
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   7215
   End
   Begin VB.Image Image2 
      Height          =   510
      Left            =   240
      Picture         =   "Backup-Restore.frx":0362
      Top             =   170
      Width           =   510
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   6360
      MouseIcon       =   "Backup-Restore.frx":1174
      MousePointer    =   99  'Custom
      Picture         =   "Backup-Restore.frx":147E
      ToolTipText     =   " Move "
      Top             =   240
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   6840
      Picture         =   "Backup-Restore.frx":1C2C
      ToolTipText     =   " Close "
      Top             =   240
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Backup - Restore"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   1320
      TabIndex        =   8
      Top             =   240
      Width           =   4455
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   600
      TabIndex        =   4
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Path"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   1560
      Width           =   6255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00C0FFFF&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   3855
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   6975
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00FFFFFF&
      Height          =   4095
      Left            =   120
      Top             =   960
      Width           =   7215
   End
End
Attribute VB_Name = "BackupRestore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wparam As Integer, ByVal iparam As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Public Sub FormDrag(Theform As Form)
    
    ReleaseCapture
    Call SendMessage(Theform.hwnd, &HA1, 2, 0&)

End Sub

Private Sub Drive1_Change()

On Error Resume Next

Text_Restore_Path.Text = Drive1.Drive & "\"
'### Drive1.Visible = False

End Sub

Private Sub Drive2_Change()

On Error Resume Next

Dir1.Path = Drive2.Drive

End Sub

Private Sub Form_Load()

BackupRestore.Height = 5175

chkSubDir = 1           '/S Copies subdirectories if not empty
chkRecourse = 0         '/E Copies files and directories, including empty ones
chkAttributeDont = 0    '/A Only files with archive attribute set, doesn't change attribute
chkAttributeTurnOff = 0 '/M Only files with archive attribute set, turns off archive attribute
chkDate = 0             '/D Copies files changed on or after specified date. If no date, only those files newer than destination.
chkHidden = 1           '/H Copies hidden and system files also
chkReadOnly = 1         '/R Overwrites read only files
chkCopyifExisting = 0   '/U Only copies files which already exist
chkOverCopy = 1         '/Y Checks before copying over files
chkSlashK = 1          '/K doesn't reset readonly attributes on destination file
chkExclude = 0          '/EXCLUDE Excludes files whose full name contains strings.
chkSlashT = 0           '/T Creates directory structure but doesn't copy files.

Text1.Text = Date - 1
Text2.Text = " "

'NOTES  XCopy switches not included:
' /I destination doesn't exist and copying more than one file assume destination is directory
' /Q don't display file names while copying
' /C Ignore errors
' /N copy using generated short names
' /O copy file ownership information
' /X copy file audit settings
' /P prompts before creating each file
' /V verifies each new file
' /W Prompts for key press before copying
' /F Displays full source and destination file names while copying
' /L Displays file names which would be copied
' /Z Copies networked files in restartable mode
'Note that /EXCLUDE switch doesn't reference an external list of files but uses the text given on the form.


End Sub

Private Sub Image1_Click()

Unload Me

End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

FormDrag Me

End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

FormDrag Me

End Sub

Private Sub Option_Backup_Click()

XPButton_Backup_Restore.Caption = "Backup"
Dir1.Visible = False
Drive2.Visible = False

Label_Backup_Err.Caption = ""
Label1.Caption = "Source"
Label2.Caption = "Destination"
Label3.Caption = "Backup"
XPButton_Backup_Browse.Visible = True
XPButton_Restore_Browse.Visible = False
XPButton_Ok.Visible = False
Text_Backup_Path = Fpath & "\Data\db"
Text_Backup_Path.Locked = True
Text_Backup_Path.Appearance = 0
Text_Restore_Path.Locked = False
Text_Restore_Path.Appearance = 1
Text_Restore_Path = "A:\"
Label_Msg.Caption = ""

End Sub

Private Sub Option_Restore_Click()

XPButton_Backup_Restore.Caption = "Restore"
Drive1.Visible = False

Label_Backup_Err.Caption = ""
Label1.Caption = "Source - Browse to where you saved the 'CredoData' folder"
Label2.Caption = "Destination"
Label3.Caption = "Restore"
XPButton_Backup_Browse.Visible = False
XPButton_Restore_Browse.Visible = True
Text_Backup_Path.Text = ""
'Text_Backup_Path.Text = "A:\CredoData"
Text_Backup_Path.Locked = False
Text_Backup_Path.Appearance = 1
Text_Restore_Path.Locked = True
Text_Restore_Path.Appearance = 0
Text_Restore_Path = Fpath & "\Data\db\CredoData"
Label_Msg.Caption = ""

End Sub

Private Sub XPButton_Backup_Browse_Click()

Drive1.Visible = True
XPButton_Backup_Browse.Visible = False

On Error GoTo ErrHB
Drive1.Drive = "B:"
Exit Sub

ErrHB:
Call Try_Floppy_Drive_A
Exit Sub

End Sub

Private Sub Try_Floppy_Drive_A()

Drive1.Visible = True

On Error GoTo ErrHA
Drive1.Drive = "A:"
Exit Sub

ErrHA:
MsgBox ("No floppy disk in drive, please insert floppy disk!")
Text_Restore_Path = "C:\"
Drive1.Visible = True
Drive1.Drive = "C:"
Exit Sub

End Sub

Private Sub XPButton_Backup_Restore_Click()

XPButton_Exit.Enabled = False
Label_Backup_Err.Caption = ""

Dim BRMSG
If Option_Backup.Value = False Then GoTo opt

opt:
If Option_Restore.Value = False Then GoTo opt2

opt2:
If Option_Backup.Value = True Then GoTo Backup
If Option_Restore.Value = True Then GoTo Restore

GoTo Mess

Backup:
Label_Msg.Caption = "Backing Up....."
DoEvents
Sleep (10)
Call Backup
Exit Sub

Restore:
Label_Msg.Caption = "Restoring....."
DoEvents
Sleep (10)
Call Restore
Exit Sub

Mess: BRMSG = MsgBox("Please choose an option:  Backup or Restore! ", vbInformation, " Backup - Restore")

End Sub

Private Sub Restore()

'If Path lacks a "\", add one to the end
If Right$(Text_Backup_Path, 1) <> "\" Then Text_Backup_Path = Text_Backup_Path & "\"
Text_Backup_Path = UCase$(Text_Backup_Path)
If Right$(Text_Restore_Path, 1) <> "\" Then Text_Restore_Path = Text_Restore_Path & "\"
Text_Restore_Path = UCase$(Text_Restore_Path)
'dstPathBackup = Text_Restore_Path

 'Transfer information into global variables
SlashA = CBool(chkAttributeDont)
SlashD = CBool(chkDate)
SlashE = CBool(chkRecourse)
SlashEX = CBool(chkExclude)
SlashH = CBool(chkHidden)
SlashK = CBool(chkSlashK)
SlashM = CBool(chkAttributeTurnOff)
SlashR = CBool(chkReadOnly)
SlashS = CBool(chkSubDir)
SlashT = CBool(chkSlashT)
SlashU = CBool(chkCopyifExisting)
SlashY = CBool(chkOverCopy)
'ExcludePattern = RTrim(LTrim(Text2.Text))
'IncludePattern = RTrim(LTrim(txtPattern))

Dim R
R = ProcessDirectory(CStr(Text_Backup_Path), CStr(Text_Restore_Path))
   If R < 0 Then
      Label3.Caption = "Restore Error!"
      Label_Msg.Caption = ""
      XPButton_Exit.Enabled = True
   Else
      Label3.Caption = "Restore Successful"
      Label_Msg.Caption = ""
      XPButton_Exit.Enabled = True
   End If
  
End Sub

Private Sub Backup()

'If Path lacks a "\", add one to the end
If Right$(Text_Backup_Path, 1) <> "\" Then Text_Backup_Path = Text_Backup_Path & "\"
Text_Backup_Path = UCase$(Text_Backup_Path)
If Right$(Text_Restore_Path, 1) <> "\" Then Text_Restore_Path = Text_Restore_Path & "\"
Text_Restore_Path = UCase$(Text_Restore_Path)
'dstPathBackup = Text_Restore_Path

'Transfer information into global variables
SlashA = CBool(chkAttributeDont)
SlashD = CBool(chkDate)
SlashE = CBool(chkRecourse)
SlashEX = CBool(chkExclude)
SlashH = CBool(chkHidden)
SlashK = CBool(chkSlashK)
SlashM = CBool(chkAttributeTurnOff)
SlashR = CBool(chkReadOnly)
SlashS = CBool(chkSubDir)
SlashT = CBool(chkSlashT)
SlashU = CBool(chkCopyifExisting)
SlashY = CBool(chkOverCopy)
'ExcludePattern = RTrim(LTrim(Text2.Text))
'IncludePattern = RTrim(LTrim(txtPattern))

Dim R
R = ProcessDirectory(CStr(Text_Backup_Path), CStr(Text_Restore_Path))
   If R < 0 Then
      Label3.Caption = "Backup Error!"
      Label_Msg.Caption = ""
      Label_Backup_Err.Caption = "Backup Error: Backup already exist, delete old backup first! or wrong drive letter! or no flopy in drive! or floppy full! "
      XPButton_Exit.Enabled = True
   Else
      Label3.Caption = "Backup Successful"
      Label_Msg.Caption = ""
      XPButton_Exit.Enabled = True
   End If

End Sub

Private Sub XPButton_Exit_Click()

Unload Me

End Sub

Private Sub XPButton_Ok_Click()

Text_Backup_Path.Text = Dir1.Path
XPButton_Ok.Visible = False
Dir1.Visible = False
Drive2.Visible = False

End Sub

Private Sub XPButton_Restore_Browse_Click()

Drive2.Visible = True
Dir1.Visible = True
Dir1.Height = 2000
XPButton_Ok.Visible = True

End Sub

