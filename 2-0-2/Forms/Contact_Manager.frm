VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Address_Contact_Manager 
   BackColor       =   &H00800000&
   BorderStyle     =   0  'None
   Caption         =   "Contact Manager"
   ClientHeight    =   7950
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9615
   Icon            =   "Contact_Manager.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7950
   ScaleWidth      =   9615
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3720
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   47
      Top             =   8160
      Visible         =   0   'False
      Width           =   855
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Print_All_Contacts 
      Height          =   375
      Left            =   7200
      TabIndex        =   46
      Top             =   3720
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Print All Contacts"
      ForeColor       =   32768
      ForeHover       =   12582912
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   120
      Top             =   7320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Print 
      Height          =   375
      Left            =   600
      TabIndex        =   45
      Top             =   2880
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Print"
      ForeColor       =   -2147483642
      ForeHover       =   12582912
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Settings 
      Height          =   255
      Left            =   7560
      TabIndex        =   44
      Top             =   7320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Microsoft Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Settings"
      ForeColor       =   -2147483642
      ForeHover       =   12582912
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Back_Rest 
      Height          =   375
      Left            =   7200
      TabIndex        =   43
      Top             =   4200
      Visible         =   0   'False
      Width           =   2055
      _ExtentX        =   3625
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
      Caption         =   "Backup/Restore"
      ForeColor       =   32768
      ForeHover       =   12582912
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Notes 
      Height          =   375
      Left            =   7560
      TabIndex        =   41
      Top             =   3240
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "Notes"
      ForeColor       =   -2147483642
      ForeHover       =   12582912
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Web_Site 
      Height          =   375
      Left            =   7560
      TabIndex        =   40
      Top             =   2760
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "Web Site"
      ForeColor       =   -2147483642
      ForeHover       =   12582912
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Email 
      Height          =   375
      Left            =   7560
      TabIndex        =   39
      Top             =   2280
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "Email"
      ForeColor       =   -2147483642
      ForeHover       =   12582912
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Mobile 
      Height          =   375
      Left            =   7560
      TabIndex        =   38
      Top             =   1800
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "Mobile"
      ForeColor       =   -2147483642
      ForeHover       =   12582912
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Phone 
      Height          =   375
      Left            =   7560
      TabIndex        =   37
      Top             =   1320
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "Phone"
      ForeColor       =   -2147483642
      ForeHover       =   12582912
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Exit 
      Height          =   375
      Left            =   600
      TabIndex        =   36
      Top             =   6960
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "Exit"
      ForeColor       =   255
      ForeHover       =   32768
   End
   Begin VB.ListBox List1 
      Height          =   450
      ItemData        =   "Contact_Manager.frx":030A
      Left            =   8760
      List            =   "Contact_Manager.frx":030C
      TabIndex        =   33
      Top             =   960
      Width           =   735
   End
   Begin Credo_Contact_Manager.XPButton XPButton_View 
      Height          =   375
      Left            =   600
      TabIndex        =   32
      Top             =   2160
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "View"
      ForeColor       =   -2147483642
      ForeHover       =   12582912
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Save 
      Height          =   375
      Left            =   600
      TabIndex        =   31
      Top             =   6240
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
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
   Begin Credo_Contact_Manager.XPButton XPButton_Edit 
      Height          =   375
      Left            =   600
      TabIndex        =   30
      Top             =   3600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Enabled         =   0   'False
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
   Begin Credo_Contact_Manager.XPButton XPButton_Delete 
      Height          =   375
      Left            =   600
      TabIndex        =   29
      Top             =   4320
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Delete"
      ForeColor       =   -2147483642
      ForeHover       =   12582912
   End
   Begin VB.TextBox Text_Web_Site 
      Height          =   375
      Left            =   3360
      MaxLength       =   70
      TabIndex        =   13
      Top             =   6960
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.TextBox Text_Company 
      Height          =   375
      Left            =   2640
      MaxLength       =   70
      TabIndex        =   3
      Top             =   2160
      Visible         =   0   'False
      Width           =   6495
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Add 
      Height          =   375
      Left            =   600
      TabIndex        =   26
      Top             =   5760
      Visible         =   0   'False
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Add"
      ForeColor       =   -2147483642
      ForeHover       =   12582912
   End
   Begin Credo_Contact_Manager.XPButton XPButton_Contacts 
      Height          =   375
      Left            =   600
      TabIndex        =   15
      Top             =   1440
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   661
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Contacts"
      ForeColor       =   -2147483642
      ForeHover       =   12582912
   End
   Begin VB.TextBox Text_Town 
      Height          =   375
      Left            =   2640
      MaxLength       =   40
      TabIndex        =   6
      Top             =   4080
      Visible         =   0   'False
      Width           =   3015
   End
   Begin VB.TextBox Text_Fax 
      Height          =   375
      Left            =   2640
      MaxLength       =   25
      TabIndex        =   10
      Top             =   5520
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox Text_Mobile 
      Height          =   375
      Left            =   5880
      MaxLength       =   25
      TabIndex        =   11
      Top             =   5520
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.TextBox Text_Phone 
      Height          =   375
      Left            =   5760
      MaxLength       =   25
      TabIndex        =   9
      Top             =   4800
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.TextBox Text_Email 
      Height          =   375
      Left            =   2640
      MaxLength       =   70
      TabIndex        =   12
      Top             =   6240
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.TextBox Text_County 
      Height          =   375
      Left            =   6000
      MaxLength       =   40
      TabIndex        =   7
      Top             =   4080
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox Text_Post_Code 
      Height          =   375
      Left            =   2640
      MaxLength       =   25
      TabIndex        =   8
      Top             =   4800
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox Text_Address_2 
      Height          =   375
      Left            =   2640
      MaxLength       =   70
      TabIndex        =   5
      Top             =   3360
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.TextBox Text_Address_1 
      Height          =   375
      Left            =   2640
      MaxLength       =   70
      TabIndex        =   4
      Top             =   2880
      Visible         =   0   'False
      Width           =   6495
   End
   Begin VB.TextBox Text_SurName 
      Height          =   375
      Left            =   6000
      MaxLength       =   25
      TabIndex        =   2
      Top             =   1440
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox Text_Name 
      Height          =   375
      Left            =   2640
      MaxLength       =   20
      TabIndex        =   1
      Top             =   1440
      Visible         =   0   'False
      Width           =   3015
   End
   Begin Credo_Contact_Manager.XPButton XPButton_New 
      Height          =   375
      Left            =   600
      TabIndex        =   14
      Top             =   5040
      Width           =   1335
      _ExtentX        =   2355
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
      Caption         =   "New"
      ForeColor       =   -2147483642
      ForeHover       =   12582912
   End
   Begin VB.Label Label_Notes2 
      Caption         =   "Notes"
      Height          =   375
      Left            =   4920
      TabIndex        =   48
      Top             =   8400
      Width           =   1935
   End
   Begin VB.Image Image8 
      Height          =   555
      Left            =   7200
      MouseIcon       =   "Contact_Manager.frx":030E
      MousePointer    =   99  'Custom
      Picture         =   "Contact_Manager.frx":0618
      ToolTipText     =   " About Credo Contact Manager "
      Top             =   5880
      Width           =   555
   End
   Begin VB.Label Label_http 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "http://"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   42
      Top             =   6960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   480
      Picture         =   "Contact_Manager.frx":168A
      Top             =   260
      Width           =   480
   End
   Begin VB.Image Image7 
      Height          =   480
      Left            =   240
      Picture         =   "Contact_Manager.frx":1994
      Stretch         =   -1  'True
      Top             =   190
      Width           =   480
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   " 2004  R. Paret - Credo Software"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   375
      Left            =   3600
      TabIndex        =   34
      Top             =   7440
      Width           =   3135
   End
   Begin VB.Image Image5 
      Height          =   1575
      Left            =   7800
      Picture         =   "Contact_Manager.frx":1C9E
      Top             =   5880
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Image Image6 
      Height          =   870
      Left            =   7380
      Picture         =   "Contact_Manager.frx":2D05
      Top             =   4800
      Visible         =   0   'False
      Width           =   1680
   End
   Begin VB.Label Label_Mobile 
      BackStyle       =   0  'Transparent
      Caption         =   "Mobile Phone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5880
      TabIndex        =   24
      Top             =   5280
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label Label_County 
      BackStyle       =   0  'Transparent
      Caption         =   "County"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   20
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label_Post_Code 
      BackStyle       =   0  'Transparent
      Caption         =   "Post Code"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   21
      Top             =   4560
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00C0FFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "©"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   35
      Top             =   7380
      Width           =   255
   End
   Begin VB.Label Label_Web_Site 
      BackStyle       =   0  'Transparent
      Caption         =   "Web Site"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   28
      Top             =   6720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Label Label_Company 
      BackStyle       =   0  'Transparent
      Caption         =   "Company"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   27
      Top             =   1920
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Label Label_Fax 
      BackStyle       =   0  'Transparent
      Caption         =   "Fax"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   25
      Top             =   5280
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label_Phone 
      BackStyle       =   0  'Transparent
      Caption         =   "Phone"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5760
      TabIndex        =   23
      Top             =   4560
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label_Email 
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   22
      Top             =   6000
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label_Town 
      BackStyle       =   0  'Transparent
      Caption         =   "Town"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   19
      Top             =   3840
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label_Address 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   18
      Top             =   2640
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.Label Label_Surname 
      BackStyle       =   0  'Transparent
      Caption         =   "Surname"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6000
      TabIndex        =   17
      Top             =   1200
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label_Name 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2640
      TabIndex        =   16
      Top             =   1200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H00C0FFFF&
      FillColor       =   &H00FF8080&
      FillStyle       =   0  'Solid
      Height          =   6615
      Left            =   240
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   2055
   End
   Begin VB.Shape Shape5 
      BorderColor     =   &H00C0FFFF&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   6615
      Left            =   1800
      Top             =   1080
      Width           =   5775
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H00C0FFFF&
      FillColor       =   &H00C0FFFF&
      FillStyle       =   0  'Solid
      Height          =   6615
      Left            =   7080
      Shape           =   4  'Rounded Rectangle
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Credo Contact Manager"
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
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   6975
   End
   Begin VB.Image Image3 
      Height          =   375
      Left            =   8040
      MouseIcon       =   "Contact_Manager.frx":3B64
      MousePointer    =   99  'Custom
      Picture         =   "Contact_Manager.frx":3E6E
      ToolTipText     =   " Move "
      Top             =   240
      Width           =   375
   End
   Begin VB.Image Image2 
      Height          =   375
      Left            =   8520
      Picture         =   "Contact_Manager.frx":461C
      ToolTipText     =   " Minimize "
      Top             =   240
      Width           =   375
   End
   Begin VB.Image Image1 
      Height          =   375
      Left            =   9000
      Picture         =   "Contact_Manager.frx":4DCA
      ToolTipText     =   " Close "
      Top             =   240
      Width           =   375
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H00FFFFFF&
      Height          =   6855
      Left            =   120
      Top             =   960
      Width           =   9375
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      Top             =   120
      Width           =   9375
   End
End
Attribute VB_Name = "Address_Contact_Manager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Option Compare Text

Private Declare Sub ReleaseCapture Lib "user32" ()
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wparam As Integer, ByVal iparam As Long) As Long


Public Sub FormDrag(Theform As Form)
    
    ReleaseCapture
    Call SendMessage(Theform.hwnd, &HA1, 2, 0&)

End Sub

Private Sub Form_Load()

List1.Top = 1300
List1.Left = 2500
List1.Height = 6100
List1.Width = 4550

Call Start_Show

Dim Fs As Scripting.FileSystemObject
Dim Fld As Folder
Dim Fls As Files, Fl As File

Set Fs = CreateObject("Scripting.FileSystemObject")
Set Fld = Fs.GetFolder(Fpath & "\Data\db\CredoData\") ' You can specify any folder name with path here
Set Fls = Fld.Files

List1.Clear
For Each Fl In Fls
   List1.AddItem Fl.Name
Next
List1.Refresh

Dim iCount As Integer
 Dim i As Integer
 Dim j As Integer
 Dim temp As String
 iCount = List1.ListCount
 For j = 0 To iCount - 2
   For i = 0 To iCount - 2
     With List1
        If .List(i) > .List(i + 1) Then
            temp = .List(i + 1)
            .List(i + 1) = .List(i)
            .List(i) = temp
        End If
     End With
    Next i
Next j

End Sub

Private Sub Image1_Click()

Dim frm As Form
    For Each frm In Forms
        If frm.Name <> "frmMain" Then Unload frm
    Next frm
    Unload Address_Contact_Manager

End Sub

Private Sub Image2_Click()

Me.WindowState = 1

End Sub

Private Sub Image3_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

FormDrag Me

End Sub

Private Sub Image3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

FormDrag Me

End Sub

Private Sub Label_Address_1_View_Click()

End Sub

Private Sub Image8_Click()

frmAbout.Show

End Sub

Private Sub List1_Click()

XPButton_Delete.Enabled = True
XPButton_Edit.Enabled = True
XPButton_View.Enabled = True
XPButton_Print.Enabled = True
XPButton_Print_All_Contacts.Enabled = True

End Sub

Private Sub Text_Name_Change()

XPButton_Add.Enabled = True

End Sub

Private Sub XPButton_Back_Rest_Click()

BackupRestore.Show

End Sub

Private Sub XPButton_Contacts_Click()

Call Start_Show

XPButton_Contacts.Enabled = False
XPButton_Save.Visible = False
XPButton_Edit.Enabled = True
XPButton_Delete.Enabled = True
XPButton_New.Enabled = True
XPButton_View.Enabled = True
XPButton_Print.Enabled = True
List1.Visible = True

Text_Name.Visible = False
Text_SurName.Visible = False
Text_Company.Visible = False
Text_Address_1.Visible = False
Text_Address_2.Visible = False
Text_Town.Visible = False
Text_County.Visible = False
Text_Post_Code.Visible = False
Text_Email.Visible = False
Text_Phone.Visible = False
Text_Mobile.Visible = False
Text_Fax.Visible = False
Text_Web_Site.Visible = False
Label_Name.Visible = False
Label_Surname.Visible = False
Label_Company.Visible = False
Label_Address.Visible = False
Label_Town.Visible = False
Label_County.Visible = False
Label_Post_Code.Visible = False
Label_Email.Visible = False
Label_Phone.Visible = False
Label_Mobile.Visible = False
Label_Fax.Visible = False
Label_Web_Site.Visible = False

XPButton_Add.Visible = False

Text_Name.Text = ""
Text_SurName.Text = ""
Text_Company.Text = ""
Text_Address_1.Text = ""
Text_Address_2.Text = ""
Text_Town.Text = ""
Text_County.Text = ""
Text_Post_Code.Text = ""
Text_Email.Text = ""
Text_Phone.Text = ""
Text_Mobile.Text = ""
Text_Fax.Text = ""
Text_Web_Site.Text = ""

Dim iCount As Integer
 Dim i As Integer
 Dim j As Integer
 Dim temp As String
 iCount = List1.ListCount
 For j = 0 To iCount - 2
   For i = 0 To iCount - 2
     With List1
        If .List(i) > .List(i + 1) Then
            temp = .List(i + 1)
            .List(i + 1) = .List(i)
            .List(i) = temp
        End If
     End With
    Next i
Next j

  If List1.ListCount = 0 Then
  XPButton_Delete.Enabled = False
  XPButton_Edit.Enabled = False
  XPButton_View.Enabled = False
  Else
  Exit Sub
  End If

End Sub

Private Sub XPButton_Delete_Click()

On Error Resume Next

Dim HighlMSG
If List1.Text = "" Then
HighlMSG = MsgBox("Please highlight name to delete! ", vbInformation, " Highlight Name")
Exit Sub
End If

Dim DelMessage

DelMessage = MsgBox("Delete " & List1.Text & " from your address  book? ", vbYesNo, " Comfirm Delete: " & List1.Text)

If DelMessage = vbNo Then
Exit Sub
Else

Dim strFileName
Dim ListItem

ListItem = List1.Text
 
strFileName = Fpath + "\Data\db\CredoData\" & Trim(ListItem)
    Kill strFileName
    
Dim out

        out = List1.ListIndex
            List1.RemoveItem (out)
            
  End If
  
  If List1.ListCount = 0 Then
  XPButton_Delete.Enabled = False
  XPButton_Edit.Enabled = False
  XPButton_View.Enabled = False
  Else
  Exit Sub
  End If
  
End Sub

Private Sub XPButton_Edit_Click()

Dim HighlMSG
If List1.Text = "" Then
HighlMSG = MsgBox("Please highlight name to edit! ", vbInformation, " Highlight Name")
Exit Sub
End If

Call Start_Hide

XPButton_Delete.Enabled = False
XPButton_New.Enabled = False
XPButton_Edit.Enabled = False
XPButton_View.Enabled = False
XPButton_Print.Enabled = False
XPButton_Save.Visible = True
XPButton_Save.Top = 5760
XPButton_Save.Left = 600

Text_Name.Visible = True
Text_SurName.Visible = True
Text_Company.Visible = True
Text_Address_1.Visible = True
Text_Address_2.Visible = True
Text_Town.Visible = True
Text_County.Visible = True
Text_Post_Code.Visible = True
Text_Email.Visible = True
Text_Phone.Visible = True
Text_Mobile.Visible = True
Text_Fax.Visible = True
Text_Web_Site.Visible = True
Label_Name.Visible = True
Label_Surname.Visible = True
Label_Company.Visible = True
Label_Address.Visible = True
Label_Town.Visible = True
Label_County.Visible = True
Label_Post_Code.Visible = True
Label_Email.Visible = True
Label_Phone.Visible = True
Label_Mobile.Visible = True
Label_Fax.Visible = True
Label_Web_Site.Visible = True

Text_Name.Locked = False
Text_SurName.Locked = False
Text_Company.Locked = False
Text_Address_1.Locked = False
Text_Address_2.Locked = False
Text_Town.Locked = False
Text_County.Locked = False
Text_Post_Code.Locked = False
Text_Email.Locked = False
Text_Phone.Locked = False
Text_Mobile.Locked = False
Text_Fax.Locked = False
Text_Web_Site.Locked = False
Text_Name.Appearance = 1
Text_SurName.Appearance = 1
Text_Company.Appearance = 1
Text_Address_1.Appearance = 1
Text_Address_2.Appearance = 1
Text_Town.Appearance = 1
Text_County.Appearance = 1
Text_Post_Code.Appearance = 1
Text_Email.Appearance = 1
Text_Phone.Appearance = 1
Text_Mobile.Appearance = 1
Text_Fax.Appearance = 1
Text_Web_Site.Appearance = 1
Text_Web_Site.Width = 5775
Text_Web_Site.Left = 3360
Label_http.Visible = True

List1.Visible = False

XPButton_Contacts.Enabled = True

Call data_Edit_Get_Data

End Sub

Private Sub XPButton_Email_Click()

Dim GlobalEmail As String
Dim HighlMSG
Dim MailMSG

If List1.Text = "" Then
HighlMSG = MsgBox("Please highlight name to email! ", vbInformation, " Highlight Name")
Exit Sub
End If

Dim Ini As clsIniFile
Set Ini = New clsIniFile

With Ini
    .File = App.Path & "\Data\db\CredoData\" & Address_Contact_Manager.List1.Text
    GlobalEmail = .GetValue("Email", "Email")
   
End With

If GlobalEmail = "" Then
MailMSG = MsgBox("No email address for this contact! ", vbInformation, " No Email Addr.")
Exit Sub
End If

frmMAPIOutXpress.Show
frmMAPIOutXpress.txtSendTo.Text = GlobalEmail

End Sub

Private Sub XPButton_Exit_Click()

Dim frm As Form
    For Each frm In Forms
        If frm.Name <> "frmMain" Then Unload frm
    Next frm
    Unload Address_Contact_Manager
    
End Sub

Private Sub XPButton_Mobile_Click()

Dim HighlMSG
If List1.Text = "" Then
HighlMSG = MsgBox("Please highlight name to Phone! ", vbInformation, " Highlight Name")
Exit Sub
End If

Dim NrMSG

Dim GlobalMobile As String

Dim Ini As clsIniFile
Set Ini = New clsIniFile

With Ini
    .File = App.Path & "\Data\db\CredoData\" & Address_Contact_Manager.List1.Text
    GlobalMobile = .GetValue("MobilePhone", "MobilePhone")
   
End With

If GlobalMobile = "" Then
NrMSG = MsgBox("No mobile number for this contact! ", vbInformation, " No Mobile Nr.")
Exit Sub
End If

    frmPhone_Dialer.Show
    frmPhone_Dialer.Text1.Text = GlobalMobile
    
    On Error Resume Next
    frmPhoneMsg.Show

End Sub

Private Sub XPButton_New_Click()

Call Start_Hide

    Text_Name.Locked = False
    Text_Name.BackColor = &H80000005
    Text_SurName.Locked = False
    Text_SurName.BackColor = &H80000005

XPButton_Edit.Enabled = False
XPButton_Delete.Enabled = False
XPButton_New.Enabled = False
XPButton_View.Enabled = False
XPButton_Print.Enabled = False
List1.Visible = False

Text_Name.Visible = True
Text_SurName.Visible = True
Text_Company.Visible = True
Text_Address_1.Visible = True
Text_Address_2.Visible = True
Text_Town.Visible = True
Text_County.Visible = True
Text_Post_Code.Visible = True
Text_Email.Visible = True
Text_Phone.Visible = True
Text_Mobile.Visible = True
Text_Fax.Visible = True
Text_Web_Site.Visible = True
Label_Name.Visible = True
Label_Surname.Visible = True
Label_Company.Visible = True
Label_Address.Visible = True
Label_Town.Visible = True
Label_County.Visible = True
Label_Post_Code.Visible = True
Label_Email.Visible = True
Label_Phone.Visible = True
Label_Mobile.Visible = True
Label_Fax.Visible = True
Label_Web_Site.Visible = True

Text_Name.Locked = False
Text_SurName.Locked = False
Text_Company.Locked = False
Text_Address_1.Locked = False
Text_Address_2.Locked = False
Text_Town.Locked = False
Text_County.Locked = False
Text_Post_Code.Locked = False
Text_Email.Locked = False
Text_Phone.Locked = False
Text_Mobile.Locked = False
Text_Fax.Locked = False
Text_Web_Site.Locked = False
Text_Name.Appearance = 1
Text_SurName.Appearance = 1
Text_Company.Appearance = 1
Text_Address_1.Appearance = 1
Text_Address_2.Appearance = 1
Text_Town.Appearance = 1
Text_County.Appearance = 1
Text_Post_Code.Appearance = 1
Text_Email.Appearance = 1
Text_Phone.Appearance = 1
Text_Mobile.Appearance = 1
Text_Fax.Appearance = 1
Text_Web_Site.Appearance = 1
Text_Web_Site.Width = 5775
Text_Web_Site.Left = 3360
Label_http.Visible = True

XPButton_Contacts.Enabled = True

XPButton_Add.Visible = True

End Sub

Private Sub XPButton_Add_Click()

Dim AddNameMSG

Dim strFileName
Dim strResult
Dim add
    add = Text_Name.Text + " " + Text_SurName.Text
        If Text_Name.Text = "" Then
        add = MsgBox("Please enter name", vbInformation, " Enter Name")
         Exit Sub
          End If
        
        If Text_SurName.Text = "" Then
        add = MsgBox("Please enter surname", vbInformation, " Enter Surname")
         Exit Sub
          End If
        
        List1.AddItem (add)

 If Text_Name.Text <> "" Then strFileName = Fpath + "\Data\db\CredoData\" & Trim(Text_Name.Text) + " " + (Text_SurName.Text)
strResult = Dir(strFileName)
If strResult <> "" Then
   ' Kill strFileName
AddNameMSG = MsgBox(" Name " & (Text_Name.Text & " " & Text_SurName.Text) & " already in the contact manager !", vbInformation, " Add Name: " & Text_Name.Text & " " & Text_SurName.Text)
    Exit Sub
End If

Dim Ini As clsIniFile
Set Ini = New clsIniFile

With Ini
    .File = App.Path & "\Data\db\CredoData\" & Trim(Text_Name.Text) + " " + (Text_SurName.Text)
    .WriteValue "Name", "Name", (Text_Name.Text)
    .WriteValue "Surname", "SurName", (Text_SurName.Text)
    .WriteValue "Company", "Company", (Text_Company.Text)
    .WriteValue "Address1", "Address1", (Text_Address_1.Text)
    .WriteValue "Address2", "Address2", (Text_Address_2.Text)
    .WriteValue "Town", "Town", (Text_Town.Text)
    .WriteValue "County", "County", (Text_County.Text)
    .WriteValue "PostCode", "PostCode", (Text_Post_Code.Text)
    .WriteValue "Email", "Email", (Text_Email.Text)
    .WriteValue "Phone", "Phone", (Text_Phone.Text)
    .WriteValue "MobilePhone", "MobilePhone", (Text_Mobile.Text)
    .WriteValue "Fax", "Fax", (Text_Fax.Text)
    .WriteValue "WebSite", "WebSite", (Text_Web_Site.Text)
   
End With

Text_Name.Text = ""
Text_SurName.Text = ""
Text_Company.Text = ""
Text_Address_1.Text = ""
Text_Address_2.Text = ""
Text_Town.Text = ""
Text_County.Text = ""
Text_Post_Code.Text = ""
Text_Email.Text = ""
Text_Phone.Text = ""
Text_Mobile.Text = ""
Text_Fax.Text = ""
Text_Web_Site.Text = ""

Call XPButton_Contacts_Click

End Sub

Private Sub XPButton_Notes_Click()

Dim GlobalNotes As String
Dim HighlMSG
Dim MailMSG

If List1.Text = "" Then
HighlMSG = MsgBox("Please highlight name to view notes! ", vbInformation, " Highlight Name")
Exit Sub
End If

Dim Ini As clsIniFile
Set Ini = New clsIniFile

With Ini
    .File = App.Path & "\Data\db\CredoData\" & Address_Contact_Manager.List1.Text
    GlobalNotes = .GetValue("Notes", "Notes")
   
End With

frmNotes.Show
frmNotes.Text_Notes.Locked = True
frmNotes.Text_Notes.Appearance = 0

If GlobalNotes = "" Then
MailMSG = MsgBox("No notes for this contact! ", vbInformation, " No Notes")
Exit Sub
End If

frmNotes.Text_Notes.Text = GlobalNotes
frmNotes.Label_Notes.Caption = "Notes: " + List1.Text
frmNotes.Text_Notes.Locked = True
frmNotes.Text_Notes.Appearance = 0

End Sub

Private Sub XPButton_Phone_Click()

Dim HighlMSG
If List1.Text = "" Then
HighlMSG = MsgBox("Please highlight name to Phone! ", vbInformation, " Highlight Name")
Exit Sub
End If

Dim NrMSG

Dim GlobalPhone As String

Dim Ini As clsIniFile
Set Ini = New clsIniFile

With Ini
    .File = App.Path & "\Data\db\CredoData\" & Address_Contact_Manager.List1.Text
    GlobalPhone = .GetValue("Phone", "Phone")
   
End With

If GlobalPhone = "" Then
NrMSG = MsgBox("No phone number for this contact! ", vbInformation, " No Phone Nr.")
Exit Sub
End If

    frmPhone_Dialer.Show
    frmPhone_Dialer.Text1.Text = GlobalPhone
    
    On Error Resume Next
    frmPhoneMsg.Show

End Sub

Private Sub XPButton_Print_All_Contacts_Click()

Dim NotesYes
Dim NotesNo
Dim GlobalChecked3 As String

Dim Ini As clsIniFile
Set Ini = New clsIniFile
With Ini
    .File = App.Path & "\Settings.ini"
    GlobalChecked3 = .GetValue("CheckedDataSource", "Checked3")
    If GlobalChecked3 = 1 Then GoTo NotesYes
    If GlobalChecked3 = 2 Then GoTo NotesNo
End With
Exit Sub

NotesYes:
Call XPButton_Print_all_Contacts_Notes_Yes
Exit Sub

NotesNo:
Call XPButton_Print_all_Contacts_Notes_No
Exit Sub

End Sub

Private Sub XPButton_Print_Click()

Dim HighlMSG
If List1.Text = "" Then
HighlMSG = MsgBox("Please highlight name to print out the details! ", vbInformation, " Print Out Contact Details")
Exit Sub
End If

Dim NotesYes
Dim NotesNo
Dim GlobalChecked2 As String

Dim Ini As clsIniFile
Set Ini = New clsIniFile
With Ini
    .File = App.Path & "\Settings.ini"
    GlobalChecked2 = .GetValue("CheckedDataSource", "Checked2")
    If GlobalChecked2 = 1 Then GoTo NotesYes
    If GlobalChecked2 = 2 Then GoTo NotesNo
End With
Exit Sub

NotesYes:
Call XPButton_Print_Notes_Yes
Exit Sub

NotesNo:
Call XPButton_Print_Notes_No
Exit Sub

End Sub

Private Sub XPButton_Settings_Click()

frmSettings.Show

End Sub

Private Sub XPButton_Save_Click()


Dim Ini As clsIniFile
Set Ini = New clsIniFile

With Ini
    .File = App.Path & "\Data\db\CredoData\" & Trim(Text_Name.Text) + " " + (Text_SurName.Text)
    .WriteValue "Name", "Name", (Text_Name.Text)
    .WriteValue "Surname", "SurName", (Text_SurName.Text)
    .WriteValue "Company", "Company", (Text_Company.Text)
    .WriteValue "Address1", "Address1", (Text_Address_1.Text)
    .WriteValue "Address2", "Address2", (Text_Address_2.Text)
    .WriteValue "Town", "Town", (Text_Town.Text)
    .WriteValue "County", "County", (Text_County.Text)
    .WriteValue "PostCode", "PostCode", (Text_Post_Code.Text)
    .WriteValue "Email", "Email", (Text_Email.Text)
    .WriteValue "Phone", "Phone", (Text_Phone.Text)
    .WriteValue "MobilePhone", "MobilePhone", (Text_Mobile.Text)
    .WriteValue "Fax", "Fax", (Text_Fax.Text)
    .WriteValue "WebSite", "WebSite", (Text_Web_Site.Text)
   
End With

Text_Name.Text = ""
Text_SurName.Text = ""
Text_Company.Text = ""
Text_Address_1.Text = ""
Text_Address_2.Text = ""
Text_Town.Text = ""
Text_County.Text = ""
Text_Post_Code.Text = ""
Text_Email.Text = ""
Text_Phone.Text = ""
Text_Mobile.Text = ""
Text_Fax.Text = ""
Text_Web_Site.Text = ""

    Text_Name.Locked = False
    Text_Name.BackColor = &H80000005
    Text_SurName.Locked = False
    Text_SurName.BackColor = &H80000005

Call XPButton_Contacts_Click

End Sub

Private Sub XPButton_View_Click()

Dim HighlMSG
If List1.Text = "" Then
HighlMSG = MsgBox("Please highlight name to view! ", vbInformation, " Highlight Name")
Exit Sub
End If

Call Start_Hide
Call Data_view_Yes

End Sub

Private Sub Data_view_Yes()

Call Data_View_Yes_Customized
Call Data_View_Yes_GetData

Label_Name.Visible = True
Label_Surname.Visible = True
Label_Company.Visible = True
Label_Address.Visible = True
Label_Town.Visible = True
Label_County.Visible = True
Label_Post_Code.Visible = True
Label_Email.Visible = True
Label_Phone.Visible = True
Label_Mobile.Visible = True
Label_Fax.Visible = True
Label_Web_Site.Visible = True
Text_Name.Visible = True
Text_SurName.Visible = True
Text_Company.Visible = True
Text_Address_1.Visible = True
Text_Address_2.Visible = True
Text_Town.Visible = True
Text_County.Visible = True
Text_Post_Code.Visible = True
Text_Email.Visible = True
Text_Phone.Visible = True
Text_Mobile.Visible = True
Text_Fax.Visible = True
Text_Web_Site.Visible = True
Text_Name.BackColor = &H80000005
Text_SurName.BackColor = &H80000005
Text_Web_Site.Width = 6495
Text_Web_Site.Left = 2640
Label_http.Visible = False

End Sub

Private Sub Data_View_Yes_Customized()

Text_Name.Locked = True
Text_SurName.Locked = True
Text_Company.Locked = True
Text_Address_1.Locked = True
Text_Address_2.Locked = True
Text_Town.Locked = True
Text_County.Locked = True
Text_Post_Code.Locked = True
Text_Email.Locked = True
Text_Phone.Locked = True
Text_Mobile.Locked = True
Text_Fax.Locked = True
Text_Web_Site.Locked = True
Text_Name.Appearance = 0
Text_SurName.Appearance = 0
Text_Company.Appearance = 0
Text_Address_1.Appearance = 0
Text_Address_2.Appearance = 0
Text_Town.Appearance = 0
Text_County.Appearance = 0
Text_Post_Code.Appearance = 0
Text_Email.Appearance = 0
Text_Phone.Appearance = 0
Text_Mobile.Appearance = 0
Text_Fax.Appearance = 0
Text_Web_Site.Appearance = 0

XPButton_Delete.Enabled = False
XPButton_New.Enabled = False
XPButton_Edit.Enabled = False
XPButton_View.Enabled = False
XPButton_Save.Visible = False
XPButton_Print.Enabled = False

List1.Visible = False

XPButton_Contacts.Enabled = True

End Sub

Private Sub Data_View_Yes_GetData()

Dim GlobalName As String
Dim GlobalSurname As String
Dim GlobalCompany As String
Dim GlobalAddress1 As String
Dim GlobalAddress2 As String
Dim GlobalTown As String
Dim GlobalCounty As String
Dim GlobalPostCode As String
Dim GlobalEmail As String
Dim GlobalPhone As String
Dim GlobalMobile As String
Dim GlobalFax As String
Dim GlobalWebSite As String

Dim Ini As clsIniFile
Set Ini = New clsIniFile

With Ini
    .File = App.Path & "\Data\db\CredoData\" & List1.Text
    GlobalName = .GetValue("Name", "Name")
    GlobalSurname = .GetValue("Surname", "Surname")
    GlobalCompany = .GetValue("Company", "Company")
    GlobalAddress1 = .GetValue("Address1", "Address1")
    GlobalAddress2 = .GetValue("Address2", "Address2")
    GlobalTown = .GetValue("Town", "Town")
    GlobalCounty = .GetValue("County", "County")
    GlobalPostCode = .GetValue("PostCode", "PostCode")
    GlobalEmail = .GetValue("Email", "Email")
    GlobalPhone = .GetValue("Phone", "Phone")
    GlobalMobile = .GetValue("MobilePhone", "MobilePhone")
    GlobalFax = .GetValue("Fax", "Fax")
    GlobalWebSite = .GetValue("WebSite", "WebSite")
    
    Text_Name = GlobalName
    Text_SurName = GlobalSurname
    Text_Company = GlobalCompany
    Text_Address_1 = GlobalAddress1
    Text_Address_2 = GlobalAddress2
    Text_Town = GlobalTown
    Text_County = GlobalCounty
    Text_Post_Code = GlobalPostCode
    Text_Email = GlobalEmail
    Text_Phone = GlobalPhone
    Text_Mobile = GlobalMobile
    Text_Fax = GlobalFax
    Text_Web_Site = GlobalWebSite
    
End With

End Sub

Private Sub data_Edit_Get_Data()

Dim GlobalName As String
Dim GlobalSurname As String
Dim GlobalCompany As String
Dim GlobalAddress1 As String
Dim GlobalAddress2 As String
Dim GlobalTown As String
Dim GlobalCounty As String
Dim GlobalPostCode As String
Dim GlobalEmail As String
Dim GlobalPhone As String
Dim GlobalMobile As String
Dim GlobalFax As String
Dim GlobalWebSite As String

Dim Ini As clsIniFile
Set Ini = New clsIniFile

With Ini
    .File = App.Path & "\Data\db\CredoData\" & List1.Text
    GlobalName = .GetValue("Name", "Name")
    GlobalSurname = .GetValue("Surname", "Surname")
    GlobalCompany = .GetValue("Company", "Company")
    GlobalAddress1 = .GetValue("Address1", "Address1")
    GlobalAddress2 = .GetValue("Address2", "Address2")
    GlobalTown = .GetValue("Town", "Town")
    GlobalCounty = .GetValue("County", "County")
    GlobalPostCode = .GetValue("PostCode", "PostCode")
    GlobalEmail = .GetValue("Email", "Email")
    GlobalPhone = .GetValue("Phone", "Phone")
    GlobalMobile = .GetValue("MobilePhone", "MobilePhone")
    GlobalFax = .GetValue("Fax", "Fax")
    GlobalWebSite = .GetValue("WebSite", "WebSite")
    
    Text_Name = GlobalName
    Text_Name.Locked = True
    Text_Name.BackColor = &HC0FFFF
    Text_SurName = GlobalSurname
    Text_SurName.Locked = True
    Text_SurName.BackColor = &HC0FFFF
    Text_Company = GlobalCompany
    Text_Address_1 = GlobalAddress1
    Text_Address_2 = GlobalAddress2
    Text_Town = GlobalTown
    Text_County = GlobalCounty
    Text_Post_Code = GlobalPostCode
    Text_Email = GlobalEmail
    Text_Phone = GlobalPhone
    Text_Mobile = GlobalMobile
    Text_Fax = GlobalFax
    Text_Web_Site = GlobalWebSite
End With

End Sub

Private Sub XPButton_Print_Notes_Yes()

Dim GlobalName As String
Dim GlobalSurname As String
Dim GlobalCompany As String
Dim GlobalAddress1 As String
Dim GlobalAddress2 As String
Dim GlobalTown As String
Dim GlobalCounty As String
Dim GlobalPostCode As String
Dim GlobalEmail As String
Dim GlobalPhone As String
Dim GlobalMobile As String
Dim GlobalFax As String
Dim GlobalWebSite As String
Dim GlobalNotes As String

Dim Ini As clsIniFile
Set Ini = New clsIniFile

With Ini
    .File = App.Path & "\Data\db\CredoData\" & List1.Text
    GlobalName = .GetValue("Name", "Name")
    GlobalSurname = .GetValue("Surname", "Surname")
    GlobalCompany = .GetValue("Company", "Company")
    GlobalAddress1 = .GetValue("Address1", "Address1")
    GlobalAddress2 = .GetValue("Address2", "Address2")
    GlobalTown = .GetValue("Town", "Town")
    GlobalCounty = .GetValue("County", "County")
    GlobalPostCode = .GetValue("PostCode", "PostCode")
    GlobalEmail = .GetValue("Email", "Email")
    GlobalPhone = .GetValue("Phone", "Phone")
    GlobalMobile = .GetValue("MobilePhone", "MobilePhone")
    GlobalFax = .GetValue("Fax", "Fax")
    GlobalWebSite = .GetValue("WebSite", "WebSite")
    GlobalNotes = .GetValue("Notes", "Notes")
    
End With
   
Text1.Text = GlobalNotes
   
CommonDialog1.CancelError = True
On Error GoTo CancelButtonPressed

CommonDialog1.ShowPrinter 'Show the printer common dialog box.

MousePointer = 11   ' make the mouse pointer look busy while printing
Printer.Print      ' initialize printer object at beginning of page

'Setup the font.
Printer.FontName = "Arial"
Printer.FontSize = 10
Printer.FontBold = False
 
Printer.Print "     "
Printer.Print "     "
Printer.Print "     "
Printer.Print "     "
Printer.Print "     " & Label_Name & ":  " & GlobalName
Printer.Print "     "
Printer.Print "     " & Label_Surname & ":  " & GlobalSurname
Printer.Print "     "
Printer.Print "     " & Label_Company & ":  " & GlobalCompany
Printer.Print "     "
Printer.Print "     " & Label_Address & ":  " & GlobalAddress1
Printer.Print "     "
Printer.Print "     " & "           " & "     " & GlobalAddress2
Printer.Print "     "
Printer.Print "     " & Label_Town & ":  " & GlobalTown
Printer.Print "     "
Printer.Print "     " & Label_County & ":  " & GlobalCounty
Printer.Print "     "
Printer.Print "     " & Label_Post_Code & ":  " & GlobalPostCode
Printer.Print "     "
Printer.Print "     " & Label_Email & ":  " & GlobalEmail
Printer.Print "     "
Printer.Print "     " & Label_Phone & ":  " & GlobalPhone
Printer.Print "     "
Printer.Print "     " & Label_Mobile & ":  " & GlobalMobile
Printer.Print "     "
Printer.Print "     " & Label_Fax & ":  " & GlobalFax
Printer.Print "     "
Printer.Print "     " & Label_Web_Site & ":  " & GlobalWebSite
Printer.Print "     "
Printer.Print "     " & Label_Notes2 & ": "
TBPrintWrap Text1.Text, 280, 20
Printer.EndDoc
MousePointer = 0
Exit Sub

CancelButtonPressed:
'Print job cancelled

End Sub

Private Sub XPButton_Print_Notes_No()

Dim GlobalName As String
Dim GlobalSurname As String
Dim GlobalCompany As String
Dim GlobalAddress1 As String
Dim GlobalAddress2 As String
Dim GlobalTown As String
Dim GlobalCounty As String
Dim GlobalPostCode As String
Dim GlobalEmail As String
Dim GlobalPhone As String
Dim GlobalMobile As String
Dim GlobalFax As String
Dim GlobalWebSite As String

Dim Ini As clsIniFile
Set Ini = New clsIniFile

With Ini
    .File = App.Path & "\Data\db\CredoData\" & List1.Text
    GlobalName = .GetValue("Name", "Name")
    GlobalSurname = .GetValue("Surname", "Surname")
    GlobalCompany = .GetValue("Company", "Company")
    GlobalAddress1 = .GetValue("Address1", "Address1")
    GlobalAddress2 = .GetValue("Address2", "Address2")
    GlobalTown = .GetValue("Town", "Town")
    GlobalCounty = .GetValue("County", "County")
    GlobalPostCode = .GetValue("PostCode", "PostCode")
    GlobalEmail = .GetValue("Email", "Email")
    GlobalPhone = .GetValue("Phone", "Phone")
    GlobalMobile = .GetValue("MobilePhone", "MobilePhone")
    GlobalFax = .GetValue("Fax", "Fax")
    GlobalWebSite = .GetValue("WebSite", "WebSite")
    
End With
    
CommonDialog1.CancelError = True
On Error GoTo CancelButtonPressed

CommonDialog1.ShowPrinter 'Show the printer common dialog box.

MousePointer = 11   ' make the mouse pointer look busy while printing
Printer.Print      ' initialize printer object at beginning of page

'Setup the font.
Printer.FontName = "Arial"
Printer.FontSize = 10
Printer.FontBold = False
 
Printer.Print "     "
Printer.Print "     "
Printer.Print "     "
Printer.Print "     "
Printer.Print "     " & Label_Name & ":  " & GlobalName
Printer.Print "     "
Printer.Print "     " & Label_Surname & ":  " & GlobalSurname
Printer.Print "     "
Printer.Print "     " & Label_Company & ":  " & GlobalCompany
Printer.Print "     "
Printer.Print "     " & Label_Address & ":  " & GlobalAddress1
Printer.Print "     "
Printer.Print "     " & "           " & "     " & GlobalAddress2
Printer.Print "     "
Printer.Print "     " & Label_Town & ":  " & GlobalTown
Printer.Print "     "
Printer.Print "     " & Label_County & ":  " & GlobalCounty
Printer.Print "     "
Printer.Print "     " & Label_Post_Code & ":  " & GlobalPostCode
Printer.Print "     "
Printer.Print "     " & Label_Email & ":  " & GlobalEmail
Printer.Print "     "
Printer.Print "     " & Label_Phone & ":  " & GlobalPhone
Printer.Print "     "
Printer.Print "     " & Label_Mobile & ":  " & GlobalMobile
Printer.Print "     "
Printer.Print "     " & Label_Fax & ":  " & GlobalFax
Printer.Print "     "
Printer.Print "     " & Label_Web_Site & ":  " & GlobalWebSite
Printer.Print "     "
Printer.EndDoc
MousePointer = 0
Exit Sub

CancelButtonPressed:
'Print job cancelled

End Sub

Private Sub XPButton_Print_all_Contacts_Notes_Yes()

CommonDialog1.CancelError = True
On Error GoTo CancelButtonPressed

CommonDialog1.ShowPrinter 'Show the printer common dialog box.

Dim GlobalName As String
Dim GlobalSurname As String
Dim GlobalCompany As String
Dim GlobalAddress1 As String
Dim GlobalAddress2 As String
Dim GlobalTown As String
Dim GlobalCounty As String
Dim GlobalPostCode As String
Dim GlobalEmail As String
Dim GlobalPhone As String
Dim GlobalMobile As String
Dim GlobalFax As String
Dim GlobalWebSite As String
Dim GlobalNotes As String

Dim Ini As clsIniFile
Set Ini = New clsIniFile
Dim X

 For X = 0 To List1.ListCount - 1 ' Loop all items
   List1.Selected(X) = True ' Select item(x)
  With Ini
    .File = App.Path & "\Data\db\CredoData\" & List1
    GlobalName = .GetValue("Name", "Name")
    GlobalSurname = .GetValue("Surname", "Surname")
    GlobalCompany = .GetValue("Company", "Company")
    GlobalAddress1 = .GetValue("Address1", "Address1")
    GlobalAddress2 = .GetValue("Address2", "Address2")
    GlobalTown = .GetValue("Town", "Town")
    GlobalCounty = .GetValue("County", "County")
    GlobalPostCode = .GetValue("PostCode", "PostCode")
    GlobalEmail = .GetValue("Email", "Email")
    GlobalPhone = .GetValue("Phone", "Phone")
    GlobalMobile = .GetValue("MobilePhone", "MobilePhone")
    GlobalFax = .GetValue("Fax", "Fax")
    GlobalWebSite = .GetValue("WebSite", "WebSite")
    GlobalNotes = .GetValue("Notes", "Notes")
    
End With

Text1.Text = GlobalNotes

MousePointer = 11   ' make the mouse pointer look busy while printing
Printer.Print      ' initialize printer object at beginning of page

Printer.CurrentX = 1440     ' 1440 twips is 1 inch
Printer.CurrentY = 1440     ' so we have 1 inch top and left margins                       ' Establish  location of 1 inch from top, 1 inch from left
    
'Setup the font.
Printer.FontName = "Arial"
Printer.FontSize = 10
Printer.FontBold = False
 
Printer.Print "     "
Printer.Print "     " & Label_Name & ":  " & GlobalName
Printer.Print "     "
Printer.Print "     " & Label_Surname & ":  " & GlobalSurname
Printer.Print "     "
Printer.Print "     " & Label_Company & ":  " & GlobalCompany
Printer.Print "     "
Printer.Print "     " & Label_Address & ":  " & GlobalAddress1
Printer.Print "     "
Printer.Print "     " & "           " & "     " & GlobalAddress2
Printer.Print "     "
Printer.Print "     " & Label_Town & ":  " & GlobalTown
Printer.Print "     "
Printer.Print "     " & Label_County & ":  " & GlobalCounty
Printer.Print "     "
Printer.Print "     " & Label_Post_Code & ":  " & GlobalPostCode
Printer.Print "     "
Printer.Print "     " & Label_Email & ":  " & GlobalEmail
Printer.Print "     "
Printer.Print "     " & Label_Phone & ":  " & GlobalPhone
Printer.Print "     "
Printer.Print "     " & Label_Mobile & ":  " & GlobalMobile
Printer.Print "     "
Printer.Print "     " & Label_Fax & ":  " & GlobalFax
Printer.Print "     "
Printer.Print "     " & Label_Web_Site & ":  " & GlobalWebSite
Printer.Print "     "
Printer.Print "     " & Label_Notes2 & ": "
 TBPrintWrap Text1.Text, 250, 20
Printer.Print "     "
Printer.Print "     "
Printer.EndDoc
  Next X
MousePointer = 0

Exit Sub

CancelButtonPressed:
'Print job cancelled

End Sub

Private Sub XPButton_Print_all_Contacts_Notes_No()

CommonDialog1.CancelError = True
On Error GoTo CancelButtonPressed

CommonDialog1.ShowPrinter 'Show the printer common dialog box.

Dim GlobalName As String
Dim GlobalSurname As String
Dim GlobalCompany As String
Dim GlobalAddress1 As String
Dim GlobalAddress2 As String
Dim GlobalTown As String
Dim GlobalCounty As String
Dim GlobalPostCode As String
Dim GlobalEmail As String
Dim GlobalPhone As String
Dim GlobalMobile As String
Dim GlobalFax As String
Dim GlobalWebSite As String

Dim Ini As clsIniFile
Set Ini = New clsIniFile
Dim X

 For X = 0 To List1.ListCount - 1 ' Loop all items
   List1.Selected(X) = True ' Select item(x)
  With Ini
    .File = App.Path & "\Data\db\CredoData\" & List1
    GlobalName = .GetValue("Name", "Name")
    GlobalSurname = .GetValue("Surname", "Surname")
    GlobalCompany = .GetValue("Company", "Company")
    GlobalAddress1 = .GetValue("Address1", "Address1")
    GlobalAddress2 = .GetValue("Address2", "Address2")
    GlobalTown = .GetValue("Town", "Town")
    GlobalCounty = .GetValue("County", "County")
    GlobalPostCode = .GetValue("PostCode", "PostCode")
    GlobalEmail = .GetValue("Email", "Email")
    GlobalPhone = .GetValue("Phone", "Phone")
    GlobalMobile = .GetValue("MobilePhone", "MobilePhone")
    GlobalFax = .GetValue("Fax", "Fax")
    GlobalWebSite = .GetValue("WebSite", "WebSite")
    
End With

MousePointer = 11   ' make the mouse pointer look busy while printing
Printer.Print      ' initialize printer object at beginning of page

Printer.CurrentX = 1440     ' 1440 twips is 1 inch
Printer.CurrentY = 1440     ' so we have 1 inch top and left margins                       ' Establish  location of 1 inch from top, 1 inch from left
    
'Setup the font.
Printer.FontName = "Arial"
Printer.FontSize = 10
Printer.FontBold = False
 
Printer.Print "     "
Printer.Print "     " & Label_Name & ":  " & GlobalName
Printer.Print "     "
Printer.Print "     " & Label_Surname & ":  " & GlobalSurname
Printer.Print "     "
Printer.Print "     " & Label_Company & ":  " & GlobalCompany
Printer.Print "     "
Printer.Print "     " & Label_Address & ":  " & GlobalAddress1
Printer.Print "     "
Printer.Print "     " & "           " & "     " & GlobalAddress2
Printer.Print "     "
Printer.Print "     " & Label_Town & ":  " & GlobalTown
Printer.Print "     "
Printer.Print "     " & Label_County & ":  " & GlobalCounty
Printer.Print "     "
Printer.Print "     " & Label_Post_Code & ":  " & GlobalPostCode
Printer.Print "     "
Printer.Print "     " & Label_Email & ":  " & GlobalEmail
Printer.Print "     "
Printer.Print "     " & Label_Phone & ":  " & GlobalPhone
Printer.Print "     "
Printer.Print "     " & Label_Mobile & ":  " & GlobalMobile
Printer.Print "     "
Printer.Print "     " & Label_Fax & ":  " & GlobalFax
Printer.Print "     "
Printer.Print "     " & Label_Web_Site & ":  " & GlobalWebSite
Printer.Print "     "
Printer.EndDoc
  Next X
MousePointer = 0

Exit Sub

CancelButtonPressed:
'Print job cancelled

End Sub

Public Sub Start_Hide()

Image5.Visible = False
Image6.Visible = False
Image8.Visible = False
XPButton_Phone.Visible = False
XPButton_Mobile.Visible = False
XPButton_Email.Visible = False
XPButton_Web_Site.Visible = False
XPButton_Notes.Visible = False
XPButton_Back_Rest.Visible = False
XPButton_Settings.Visible = False
XPButton_Print_All_Contacts.Visible = False

End Sub

Public Sub Start_Show()

Image5.Visible = True
Image6.Visible = True
Image8.Visible = True
XPButton_Phone.Visible = True
XPButton_Mobile.Visible = True
XPButton_Email.Visible = True
XPButton_Web_Site.Visible = True
XPButton_Notes.Visible = True
XPButton_Back_Rest.Visible = True
XPButton_Settings.Visible = True
XPButton_Print_All_Contacts.Visible = True

End Sub

Private Sub XPButton_Web_Site_Click()

Dim GlobalWebSite As String
Dim HighlMSG
Dim wwwMSG

If List1.Text = "" Then
HighlMSG = MsgBox("Please highlight name to visit site! ", vbInformation, " Highlight Name")
Exit Sub
End If

Dim Ini As clsIniFile
Set Ini = New clsIniFile

With Ini
    .File = App.Path & "\Data\db\CredoData\" & Address_Contact_Manager.List1.Text
    GlobalWebSite = .GetValue("WebSite", "WebSite")
   
End With

If GlobalWebSite = "" Then
wwwMSG = MsgBox("No web site address for this contact! ", vbInformation, " No Web Site Addr.")
Exit Sub
End If

Dim www As String
www = "http://" + GlobalWebSite

www_Module.ShellExecute Me.hwnd, vbNullString, Trim(www), vbNullString, "c:\", SW_SHOWNORMAL

End Sub

Private Sub TBPrintWrap(ByVal Text As String, ByVal LtMar As Long, ByVal RtMar As Long)
Dim i As Integer
Dim j As Integer
Dim currWord As String


Printer.CurrentX = LtMar
i = 1

Do Until i > Len(Text)
currWord = ""
Do Until i > Len(Text) Or Mid$(Text, i, 1) <= " "

currWord = currWord & Mid$(Text, i, 1)

i = i + 1
Loop

If (Printer.CurrentX + Printer.TextWidth(currWord)) > (Printer.ScaleWidth - RtMar + Printer.ScaleLeft) Then

Printer.Print
Printer.CurrentX = LtMar

End If

Printer.Print currWord;
Do Until i > Len(Text) Or Mid$(Text, i, 1) > " "

Select Case Mid$(Text, i, 1)

Case " "
Printer.Print " ";

Case Chr$(10) 'LF

Printer.Print

Printer.CurrentX = LtMar

Case Chr$(9) 'Tab

j = (Printer.CurrentX) / Printer.TextWidth("0")

j = j + (10 - (j Mod 10))

Printer.CurrentX = (j * Printer.TextWidth("0"))

Case Else

End Select

i = i + 1

Loop

Loop

End Sub
