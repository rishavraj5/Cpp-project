VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "KP New Astro (v1.00 beta)"
   ClientHeight    =   7200
   ClientLeft      =   465
   ClientTop       =   1110
   ClientWidth     =   10875
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   7200
   ScaleWidth      =   10875
   Begin VB.PictureBox picAbout 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   4920
      ScaleHeight     =   495
      ScaleWidth      =   615
      TabIndex        =   51
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
      Begin VB.Label lbls 
         Caption         =   "Blog :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   10
         Left            =   360
         TabIndex        =   58
         Top             =   5400
         Width           =   1575
      End
      Begin VB.Label lbls 
         Alignment       =   2  'Center
         Caption         =   "Visit KP New Astro Blog for source code of this software."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Index           =   14
         Left            =   848
         TabIndex        =   57
         Top             =   3000
         Width           =   9255
      End
      Begin VB.Label lbls 
         Caption         =   "Email :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Index           =   17
         Left            =   360
         TabIndex        =   56
         Top             =   4680
         Width           =   1095
      End
      Begin VB.Label lbls 
         Alignment       =   2  'Center
         Caption         =   "KP New Astro is a Free software and released under GNU general public license (Version 2 or later)."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Index           =   29
         Left            =   203
         TabIndex        =   55
         Top             =   1920
         Width           =   10455
      End
      Begin VB.Label lbls 
         Alignment       =   2  'Center
         Caption         =   $"frmMain.frx":1CFA
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1095
         Index           =   30
         Left            =   203
         TabIndex        =   54
         Top             =   600
         Width           =   10455
      End
      Begin VB.Label lblEmail 
         Caption         =   "kpnewastro@gmail.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1920
         MouseIcon       =   "frmMain.frx":1D8E
         MousePointer    =   99  'Custom
         TabIndex        =   53
         Top             =   4680
         Width           =   3855
      End
      Begin VB.Label lblWeb 
         Caption         =   "kpnewastro.blogspot.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   375
         Left            =   1920
         MouseIcon       =   "frmMain.frx":2658
         MousePointer    =   99  'Custom
         TabIndex        =   52
         Top             =   5400
         Width           =   3855
      End
   End
   Begin MSComDlg.CommonDialog cDlg1 
      Left            =   7560
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox picNatalHorary 
      BorderStyle     =   0  'None
      Height          =   6015
      Left            =   0
      ScaleHeight     =   6015
      ScaleWidth      =   10695
      TabIndex        =   29
      Top             =   480
      Width           =   10695
      Begin VB.TextBox txtDefPlaceCur 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   6840
         TabIndex        =   81
         Text            =   "Default Place"
         Top             =   2760
         Width           =   3735
      End
      Begin VB.TextBox txtTZPMCur 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8880
         MaxLength       =   1
         TabIndex        =   60
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox txtTZHHCur 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9480
         MaxLength       =   2
         TabIndex        =   61
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox txtTZMMCur 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10080
         MaxLength       =   2
         TabIndex        =   62
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox txtHHCur 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8880
         MaxLength       =   2
         TabIndex        =   72
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txtSSCur 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10080
         MaxLength       =   2
         TabIndex        =   74
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txtMMCur 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9480
         MaxLength       =   2
         TabIndex        =   73
         Top             =   2400
         Width           =   495
      End
      Begin VB.TextBox txtYearCur 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10080
         MaxLength       =   4
         TabIndex        =   71
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox txtMonthCur 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9480
         MaxLength       =   2
         TabIndex        =   70
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox txtDayCur 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8880
         MaxLength       =   2
         TabIndex        =   69
         Top             =   2160
         Width           =   495
      End
      Begin VB.ComboBox cmb249 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":2F22
         Left            =   9720
         List            =   "frmMain.frx":2F24
         TabIndex        =   24
         Text            =   "0"
         Top             =   4260
         Width           =   855
      End
      Begin VB.ComboBox cmbRotHor 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":2F26
         Left            =   9720
         List            =   "frmMain.frx":2F4E
         TabIndex        =   25
         Text            =   "01"
         Top             =   4800
         Width           =   855
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Generate KP Horary Chart"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   8040
         TabIndex        =   26
         Top             =   5520
         Width           =   2535
      End
      Begin VB.ComboBox cmbLatNS 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":2F82
         Left            =   4200
         List            =   "frmMain.frx":2F8C
         TabIndex        =   21
         Text            =   "N"
         Top             =   4320
         Width           =   735
      End
      Begin VB.ComboBox cmbLatMin 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":2F96
         Left            =   3360
         List            =   "frmMain.frx":304E
         TabIndex        =   20
         Text            =   "39"
         Top             =   4320
         Width           =   735
      End
      Begin VB.ComboBox cmbLatDeg 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":3142
         Left            =   2520
         List            =   "frmMain.frx":3254
         TabIndex        =   19
         Text            =   "06"
         Top             =   4320
         Width           =   735
      End
      Begin VB.CommandButton cmdGenerate 
         Caption         =   "Generate KP Natal Chart"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   3840
         TabIndex        =   23
         Top             =   5520
         Width           =   2535
      End
      Begin VB.ComboBox cmbLonEW 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":33C0
         Left            =   4200
         List            =   "frmMain.frx":33CA
         TabIndex        =   18
         Text            =   "E"
         Top             =   3960
         Width           =   735
      End
      Begin VB.ComboBox cmbLonMin 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":33D4
         Left            =   3360
         List            =   "frmMain.frx":348C
         TabIndex        =   17
         Text            =   "54"
         Top             =   3960
         Width           =   735
      End
      Begin VB.ComboBox cmbLonDeg 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":3580
         Left            =   2520
         List            =   "frmMain.frx":37A0
         TabIndex        =   16
         Text            =   "079"
         Top             =   3960
         Width           =   735
      End
      Begin VB.ComboBox cmbTzMM 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":3B28
         Left            =   4200
         List            =   "frmMain.frx":3BE0
         TabIndex        =   15
         Text            =   "30"
         Top             =   3600
         Width           =   735
      End
      Begin VB.ComboBox cmbTzHH 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":3CD4
         Left            =   3360
         List            =   "frmMain.frx":3CFC
         TabIndex        =   14
         Text            =   "05"
         Top             =   3600
         Width           =   735
      End
      Begin VB.ComboBox cmbTzPM 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":3D26
         Left            =   2520
         List            =   "frmMain.frx":3D30
         TabIndex        =   13
         Text            =   "+"
         ToolTipText     =   "Use ""+"" Sign for East of GMT"
         Top             =   3600
         Width           =   735
      End
      Begin VB.ComboBox cmbTcMM 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":3D3A
         Left            =   4200
         List            =   "frmMain.frx":3DF2
         TabIndex        =   10
         Text            =   "00"
         Top             =   2640
         Width           =   735
      End
      Begin VB.ComboBox cmbTcHH 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":3EE6
         Left            =   3360
         List            =   "frmMain.frx":3EFC
         TabIndex        =   9
         Text            =   "0"
         Top             =   2640
         Width           =   735
      End
      Begin VB.ComboBox cmbTcPM 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":3F12
         Left            =   2520
         List            =   "frmMain.frx":3F1C
         TabIndex        =   8
         Text            =   "+"
         Top             =   2640
         Width           =   735
      End
      Begin VB.CommandButton cmdAtlas 
         Caption         =   "Atlas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5520
         TabIndex        =   12
         Top             =   3120
         Width           =   855
      End
      Begin VB.TextBox txtBirthPlace 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   11
         Text            =   "Panadura, Sri Lanka"
         Top             =   3120
         Width           =   2895
      End
      Begin VB.ComboBox cmbSS 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":3F26
         Left            =   4200
         List            =   "frmMain.frx":3FDE
         TabIndex        =   7
         Text            =   "00"
         Top             =   2280
         Width           =   735
      End
      Begin VB.ComboBox cmbMM 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":40D2
         Left            =   3360
         List            =   "frmMain.frx":418A
         TabIndex        =   6
         Text            =   "43"
         Top             =   2280
         Width           =   735
      End
      Begin VB.ComboBox cmbHH 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":427E
         Left            =   2520
         List            =   "frmMain.frx":42CA
         TabIndex        =   5
         Text            =   "05"
         Top             =   2280
         Width           =   735
      End
      Begin VB.ComboBox cmbYear 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   4200
         TabIndex        =   4
         Text            =   "1985"
         Top             =   1920
         Width           =   975
      End
      Begin VB.ComboBox cmbMonth 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":432E
         Left            =   3360
         List            =   "frmMain.frx":4356
         TabIndex        =   3
         Text            =   "07"
         Top             =   1920
         Width           =   735
      End
      Begin VB.ComboBox cmbDay 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":438A
         Left            =   2520
         List            =   "frmMain.frx":43EB
         TabIndex        =   2
         Text            =   "31"
         Top             =   1920
         Width           =   735
      End
      Begin VB.ComboBox cmbMF 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":446B
         Left            =   5400
         List            =   "frmMain.frx":4475
         TabIndex        =   1
         Text            =   "Male"
         Top             =   1440
         Width           =   975
      End
      Begin VB.TextBox txtName 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   2520
         MaxLength       =   50
         TabIndex        =   0
         Text            =   "JSW"
         Top             =   1440
         Width           =   2775
      End
      Begin VB.ComboBox cmbRot 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmMain.frx":4487
         Left            =   3480
         List            =   "frmMain.frx":44AF
         TabIndex        =   22
         Text            =   "01"
         Top             =   4800
         Width           =   855
      End
      Begin VB.TextBox txtLatDDCur 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8880
         MaxLength       =   2
         TabIndex        =   66
         Top             =   3720
         Width           =   495
      End
      Begin VB.TextBox txtLatMMCur 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9480
         MaxLength       =   2
         TabIndex        =   67
         Top             =   3720
         Width           =   495
      End
      Begin VB.TextBox txtLatNSCur 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10080
         MaxLength       =   1
         TabIndex        =   68
         Top             =   3720
         Width           =   495
      End
      Begin VB.TextBox txtLonDDCur 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   8880
         MaxLength       =   3
         TabIndex        =   63
         Top             =   3480
         Width           =   495
      End
      Begin VB.TextBox txtLonMMCur 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   9480
         MaxLength       =   2
         TabIndex        =   64
         Top             =   3480
         Width           =   495
      End
      Begin VB.TextBox txtLonEWCur 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   10080
         MaxLength       =   1
         TabIndex        =   65
         Top             =   3480
         Width           =   495
      End
      Begin VB.Label lblWeb2 
         Caption         =   "kpnewastro.blogspot.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmMain.frx":44E3
         MousePointer    =   99  'Custom
         TabIndex        =   80
         Top             =   5760
         Width           =   2295
      End
      Begin VB.Label lblEmail2 
         Caption         =   "kpnewastro@gmail.com"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   120
         MouseIcon       =   "frmMain.frx":4DAD
         MousePointer    =   99  'Custom
         TabIndex        =   79
         Top             =   5400
         Width           =   2055
      End
      Begin VB.Label lbls 
         AutoSize        =   -1  'True
         Caption         =   "Time Zone (HH:MM) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   21
         Left            =   6840
         TabIndex        =   78
         Top             =   3165
         Width           =   1755
      End
      Begin VB.Label lbls 
         AutoSize        =   -1  'True
         Caption         =   "Longitude (DDD:MM) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   19
         Left            =   6840
         TabIndex        =   77
         Top             =   3525
         Width           =   1830
      End
      Begin VB.Label lbls 
         AutoSize        =   -1  'True
         Caption         =   "Latitude (DD:MM) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   12
         Left            =   6840
         TabIndex        =   76
         Top             =   3765
         Width           =   1575
      End
      Begin VB.Label lbls 
         AutoSize        =   -1  'True
         Caption         =   "Time (HH:MM:SS) 24H :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   18
         Left            =   6840
         TabIndex        =   75
         Top             =   2445
         Width           =   1920
      End
      Begin VB.Label lbls 
         AutoSize        =   -1  'True
         Caption         =   "Date (DD/MM/YYYY) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   16
         Left            =   6840
         TabIndex        =   59
         Top             =   2205
         Width           =   1830
      End
      Begin VB.Label lblW 
         Alignment       =   2  'Center
         Caption         =   "Latitude value of your default place is greater than 66 of degrees."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   435
         Left            =   6720
         TabIndex        =   32
         Top             =   840
         Visible         =   0   'False
         Width           =   4005
      End
      Begin VB.Label lbls 
         AutoSize        =   -1  'True
         Caption         =   "Birth Details :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   360
         TabIndex        =   33
         Top             =   960
         Width           =   1125
      End
      Begin VB.Label lbls 
         Caption         =   "1 To 249 :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   20
         Left            =   6960
         TabIndex        =   48
         Top             =   600
         Width           =   855
      End
      Begin VB.Line Line1 
         Index           =   1
         X1              =   6720
         X2              =   10560
         Y1              =   720
         Y2              =   720
      End
      Begin VB.Label lbls 
         AutoSize        =   -1  'True
         Caption         =   "Select a KP Number ( 1 to 249 ) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   13
         Left            =   6840
         TabIndex        =   47
         Top             =   4320
         Width           =   2700
      End
      Begin VB.Label lbls 
         Caption         =   "Which Bhava Should Be Lagna Bhava ? (Rotate Kundali)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   15
         Left            =   6840
         TabIndex        =   46
         Top             =   4740
         Width           =   2850
      End
      Begin VB.Label lbls 
         Caption         =   "Note : This feature cannot be used beyond the polar circle.  Latitude must be less than 66 of fegrees (North Or South)."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   675
         Index           =   11
         Left            =   6840
         TabIndex        =   45
         Top             =   1320
         Width           =   3795
      End
      Begin VB.Label lbls 
         Caption         =   "Which Bhava Should Be Lagna Bhava ? (Rotate Kundali)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Index           =   9
         Left            =   180
         TabIndex        =   44
         Top             =   4740
         Width           =   3210
      End
      Begin VB.Label lbls 
         AutoSize        =   -1  'True
         Caption         =   "Latitude (DD:MM) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   8
         Left            =   180
         TabIndex        =   43
         Top             =   4380
         Width           =   1575
      End
      Begin VB.Label lbls 
         AutoSize        =   -1  'True
         Caption         =   "Longitude (DDD:MM) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   7
         Left            =   180
         TabIndex        =   42
         Top             =   4020
         Width           =   1830
      End
      Begin VB.Label lbls 
         AutoSize        =   -1  'True
         Caption         =   "Time Zone (HH:MM) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   41
         Top             =   3660
         Width           =   1755
      End
      Begin VB.Label lbls 
         AutoSize        =   -1  'True
         Caption         =   "Place of Birth :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   40
         Top             =   3180
         Width           =   1200
      End
      Begin VB.Label lbls 
         AutoSize        =   -1  'True
         Caption         =   "Time Correction (WT/DST) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   39
         Top             =   2700
         Width           =   2325
      End
      Begin VB.Label lbls 
         AutoSize        =   -1  'True
         Caption         =   "Time (HH:MM:SS) 24H :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   38
         Top             =   2340
         Width           =   1920
      End
      Begin VB.Label lbls 
         AutoSize        =   -1  'True
         Caption         =   "Date (DD/MM/YYYY) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   37
         Top             =   1980
         Width           =   1830
      End
      Begin VB.Label lbls 
         AutoSize        =   -1  'True
         Caption         =   "Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   36
         Top             =   1500
         Width           =   570
      End
      Begin VB.Label lbls 
         AutoSize        =   -1  'True
         Caption         =   "KP HORARY"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   24
         Left            =   6840
         TabIndex        =   35
         Top             =   120
         Width           =   2445
      End
      Begin VB.Line Line1 
         Index           =   0
         X1              =   120
         X2              =   6240
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line1 
         Index           =   2
         X1              =   6600
         X2              =   6600
         Y1              =   6000
         Y2              =   120
      End
      Begin VB.Line Line1 
         Index           =   3
         X1              =   120
         X2              =   6480
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Line Line1 
         Index           =   4
         X1              =   6720
         X2              =   10680
         Y1              =   5280
         Y2              =   5280
      End
      Begin VB.Label lbls 
         AutoSize        =   -1  'True
         Caption         =   "KP NATAL"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Index           =   22
         Left            =   240
         TabIndex        =   34
         Top             =   120
         Width           =   2010
      End
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   375
      Left            =   0
      TabIndex        =   28
      Top             =   120
      Width           =   10935
      _ExtentX        =   19288
      _ExtentY        =   661
      TabWidthStyle   =   2
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Natal/Horary"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "KP Chart"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Ruling Planets"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "KP New Astro"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picRulingPlanets 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   3840
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   31
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
      Begin RichTextLib.RichTextBox RT2 
         Height          =   5895
         Left            =   120
         TabIndex        =   50
         Top             =   120
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   10398
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmMain.frx":5677
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.PictureBox picKPChart 
      BorderStyle     =   0  'None
      Height          =   615
      Left            =   3120
      ScaleHeight     =   615
      ScaleWidth      =   615
      TabIndex        =   30
      Top             =   2400
      Visible         =   0   'False
      Width           =   615
      Begin RichTextLib.RichTextBox RT1 
         Height          =   5895
         Left            =   120
         TabIndex        =   49
         Top             =   120
         Width           =   10575
         _ExtentX        =   18653
         _ExtentY        =   10398
         _Version        =   393217
         Enabled         =   -1  'True
         ReadOnly        =   -1  'True
         ScrollBars      =   2
         AutoVerbMenu    =   -1  'True
         TextRTF         =   $"frmMain.frx":56F7
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9360
      TabIndex        =   27
      Top             =   6720
      Width           =   1335
   End
   Begin VB.Line Line2 
      BorderWidth     =   4
      X1              =   120
      X2              =   10680
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Menu mnuFile 
      Caption         =   "File"
      Begin VB.Menu mnuFileNew 
         Caption         =   "New"
         Shortcut        =   ^N
      End
      Begin VB.Menu Sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileSaveAsPDF 
         Caption         =   "Save As a PDF File"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSaveAsRTF 
         Caption         =   "Save As a RTF File"
         Enabled         =   0   'False
      End
      Begin VB.Menu mnuFileSaveAsTXT 
         Caption         =   "Save As a TXT file"
         Enabled         =   0   'False
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "Exit"
         Shortcut        =   ^Q
      End
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Begin VB.Menu mnuToolsSettings 
         Caption         =   "Settings"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "Help"
      Begin VB.Menu mnuHelpKP 
         Caption         =   "KP New Astro Help"
      End
      Begin VB.Menu mnuHelpKPPDF 
         Caption         =   "PDF Help"
      End
      Begin VB.Menu sep3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "About"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'================================================================
' Description..: Krishnamurti (KP) Astrology Software
' Software.....: KP New Astro
' Date.........: 05/12/2010
' Version......: 1.0.xx beta
' Language.....: Visual Basic 6.0 Ent - SP 6
' Tested.......: Windows XP Professional - SP 3
' Copyright....: (C) 2009-2010 JSW
' E-Mail.......: kpnewastro@gmail.com
' Web..........: http://www.kpnewastro.blogspot.com
'================================================================
' Released under GNU general public license (version 2 or later)
'================================================================

Option Explicit

Private isHorary As Boolean     ' Select the Mode
Private sNo249Hor As Integer    ' Select Sub lord int 1 to 249 (not 0 to 248)

Private rotVal As Integer       'Rotation

Private isGeo As Boolean 'Geocentric
Private isTru As Boolean 'True Position

'For Print

'RTF
Private ppName As String
Private ppBDate As String
Private ppBTime As String
Private ppPlaceName As String
Private ppPlaceDetails As String

'PDF
Private ppNamePdf As String
Private ppPlaceNamePdf As String
Private ppBDatePdf As String
Private ppBTimePdf As String
Private ppPlaceDetailsPdf As String

'CHART
Private bNmeM As String
Private bAscM As String
Private bDateM As String
Private bTimeM As String
Private bCorrM As String
Private bPlaceM As String
Private bTzM As String
Private bAyanM As String

'A00_calcCurrent
Private hCusCur(11) As Double
Private pPosCur(11) As Double
Private isReCur(11) As Double

Private stdWeekDayCur As Integer
Private trdWeekDayCur As Integer
Private rsReTimeCur As Double

'A01_calcKPNatal
Private julDateLocal As Double  'Local jul date
Private hCusp(12) As Double, hCus(11) As Double, hCuspI(12) As Double    'House Cusps
Private pPos(11) As Double  'Planetary positions
Private isRe(11) As Double  'Speed of the planet
Private incHouse(11) As Integer    'Included House (Bhava) of the planets
Private kpAyan As Double    'KP Ayanamsa
Private rsReTime As Double  'Sun rise time
Private stdWeekDay As Integer   'Week Day - Standerd
Private trdWeekDay As Integer   'Week day - traditional

Private plntFlg As Long 'Planet flag - true, apparent etc..

'a02_drawChart
Private charDraw(36) As String
Private pdfPX(11, 7) As String

'Private bAscM As String
'Private bAyanM As String
'A03_calcVAspects
Private aspVedic(6) As String   '
Private vedAspP(6) As String    '
Private smeRas(6) As String    'Planet which are in same rashi

'A04_calcWAspects
Private pToHouse(11, 11) As Single      'Aspect planet to House
Private pToPlanet(11, 11) As Single     'Aspect planet to Planet
'Private pToHouseStr(11, 11) As String   'Aspect planet to House as string
'Private pToPlanetStr(11, 11) As String  'Aspect planet to planet as string

'A05_calcDasa
Private mDasL(8) As Integer
Private bDasL(8, 8) As Integer
Private aDasL(8, 8, 8) As Integer
Private sDasL(8, 8, 8, 8) As Integer


Private mDas(8) As Double
Private bDas(8, 8) As Double
Private aDas(8, 8, 8) As Double
Private sDas(8, 8, 8, 8) As Double

'A06_calcSig
Private conSig0(11) As String
Private conSig1(11) As String
Private conSig2(11) As String
Private conSig3(11) As String
Private plntWSig(11) As String

'A07_calc4StepSig
Private ownBhava(8) As String
Private posBhava(8) As Integer

Private stLords(8) As Integer
Private stLordsOwn(8) As String
Private stLordsPos(8) As Integer

Private sbLords(8) As Integer
Private sbLordsOwn(8) As String
Private sbLordsPos(8) As Integer

Private stSbLords(8) As Integer
Private stSbLordsOwn(8) As String
Private stSbLordsPos(8) As Integer

Private emptyBhava As String
Private noPInStar As String
Private selfCons As String
'Load Pesonal details
Private loadedPD() As String
Private Sub A00_calcCurrent()
'This SUB will Calculate Ruling Planets and the Current Planetary & House Positions
'using system time and date (So the system time & date must be correct)

'Load data from dat file
'Dim depLoaded() As String
'depLoaded() = hp_ReadDatFile(App.Path + "\Sett\depdet.dat", "_")

    Dim birthPlace As String
    Dim tzPM As String, tzHH As Integer, tzMM As Integer
    Dim lonDeg As Integer, lonMin As Integer, lonEW As String
    Dim latDeg As Integer, latMin As Integer, latNS As String

    birthPlace = txtDefPlaceCur.Text

    tzPM = Trim(txtTZPMCur.Text)
    tzHH = CInt(txtTZHHCur.Text)
    tzMM = CInt(txtTZMMCur.Text)

    lonDeg = CInt(txtLonDDCur.Text)
    lonMin = CInt(txtLonMMCur.Text)
    lonEW = Trim(txtLonEWCur.Text)

    latDeg = CInt(txtLatDDCur.Text)
    latMin = CInt(txtLatMMCur.Text)
    latNS = Trim(txtLatNSCur.Text)

    'Dim bHrs As Integer, bMin As Integer, bSec As Integer
    '#EXPORT VAR#
    Dim b1 As Byte, b2 As Byte, b3 As Byte

    'Calculation using SWE
    Dim bDay As Integer, bMonth As Integer, bYear As Integer
    bDay = CInt(txtDayCur.Text)
    bMonth = CInt(txtMonthCur.Text)
    bYear = CInt(txtYearCur.Text)

    Dim bHHc As Integer, bMMc As Integer, bSSc As Integer
    bHHc = CInt(txtHHCur.Text)
    bMMc = CInt(txtMMCur.Text)
    bSSc = CInt(txtSSCur.Text)

    Dim bTime As Double         'Standerd birth time as double value
    bTime = CDbl(bHHc) + (CDbl(bMMc) / 60#) + (CDbl(bSSc) / 3600#)

    '****** NO TIME CORRECTION IS INCLUDED FOR THIS CALCULATION
    '    Dim tCorr As Double         'Time correction as double value
    '    tCorr = CDbl(tCorrHrs) + (CDbl(tCorrMin) / 60#)
    '    If tCorrPM = "-" Then tCorr = tCorr * (-1#)

    Dim gmtDiff As Double       'GMT differance as double value
    gmtDiff = CDbl(tzHH) + (CDbl(tzMM) / 60#)
    If tzPM = "-" Then gmtDiff = gmtDiff * (-1#)

    Dim latVal As Double        'Latitude as double value
    latVal = CDbl(latDeg) + (CDbl(latMin) / 60#)
    If isGeo = True Then latVal = kp_GeoCorr(latVal)    'Geo correction
    If UCase(latNS) = "S" Then latVal = latVal * (-1#)

    Dim lonVal As Double        'Longitude as double value
    lonVal = CDbl(lonDeg) + (CDbl(lonMin) / 60#)
    If UCase(lonEW) = "W" Then lonVal = lonVal * (-1#)

    Call swe_set_ephe_path(App.Path + "\Eph")   'Set Swiss Eph Path
    'Call swe_set_sid_mode(5, 0, 0)              'Set sid mode to KP (KP Ayanamsa is used)

    'Julian date calculation
    Dim julDateLocal_ As Double         'Raw julian day according to the birth data
    julDateLocal_ = swe_julday(bYear, bMonth, bDay, bTime, 1)
    'julDateLocal = julDateLocal_       '#EXPORT#

    Dim julDateUT As Double             'Corrected GMT julian date
    julDateUT = julDateLocal_ - (gmtDiff / 24#)    ' + (tCorr / 24#)

    ' NEW KP AYANAMSA
    Dim newKPAyan As Double
    newKPAyan = kp_NKPA(julDateUT)
    'kpAyan = newKPAyan '#EXPORT#

    'KP Ayanamsa Swiss
    'Dim kpAyan_ As Double
    'kpAyan_ = swe_get_ayanamsa_ut(julDateUT)
    'kpAyan = kpAyan_    '#EXPORT#

    'Ayanamsa Diff
    'Dim ayanDiff As Double
    'ayanDiff = newKPAyan - kpAyan_ '( +ve Value)


    'House cusps calculation
    Dim hCusp_c(12) As Double, hCus_c(11) As Double
    Dim ascmc(9) As Double
    Dim hReVal As Long
    hReVal = swe_houses(julDateUT, latVal, lonVal, Asc("P"), hCusp_c(0), ascmc(0))

    '    For b0 = 0 To 12    '#EXPORT#
    '        hCusp(b0) = hCusp_c(b0)
    '    Next

    Dim jk As Byte
    For jk = 0 To 11
        hCus_c(jk) = hp_Rnd0To360v(hCusp_c(jk + 1) - newKPAyan)    'NKPA ** correction
    Next

    For b1 = 0 To 11    '#EXPORT#
        hCusCur(b1) = hCus_c(b1)
    Next

    'Planetary Position]
    Dim pPos_(11) As Double
    Dim isRe_(11) As Double
    Dim pReVal As Long
    Dim pDetails(5) As Double
    Dim errReturn As String

    'Sun (Ravi) 0
    pReVal = swe_calc_ut(julDateUT, 0, plntFlg, pDetails(0), errReturn)
    pPos_(0) = hp_Rnd0To360v(pDetails(0) - newKPAyan)    'NKPA ** correction
    isRe_(0) = pDetails(3)

    'Moon (Chandra) 1
    pReVal = swe_calc_ut(julDateUT, 1, plntFlg, pDetails(0), errReturn)
    pPos_(1) = hp_Rnd0To360v(pDetails(0) - newKPAyan)    'NKPA ** correction
    isRe_(1) = pDetails(3)

    'Mars (Kuja) 2
    pReVal = swe_calc_ut(julDateUT, 4, plntFlg, pDetails(0), errReturn)
    pPos_(2) = hp_Rnd0To360v(pDetails(0) - newKPAyan)    'NKPA ** correction
    isRe_(2) = pDetails(3)

    'Mercury (Budha) 3
    pReVal = swe_calc_ut(julDateUT, 2, plntFlg, pDetails(0), errReturn)
    pPos_(3) = hp_Rnd0To360v(pDetails(0) - newKPAyan)    'NKPA ** correction
    isRe_(3) = pDetails(3)

    'Jupiter (Guru) 4
    pReVal = swe_calc_ut(julDateUT, 5, plntFlg, pDetails(0), errReturn)
    pPos_(4) = hp_Rnd0To360v(pDetails(0) - newKPAyan)    'NKPA ** correction
    isRe_(4) = pDetails(3)

    'Venus (Sukra) 5
    pReVal = swe_calc_ut(julDateUT, 3, plntFlg, pDetails(0), errReturn)
    pPos_(5) = hp_Rnd0To360v(pDetails(0) - newKPAyan)    'NKPA ** correction
    isRe_(5) = pDetails(3)

    'Saturn (Sani) 6
    pReVal = swe_calc_ut(julDateUT, 6, plntFlg, pDetails(0), errReturn)
    pPos_(6) = hp_Rnd0To360v(pDetails(0) - newKPAyan)    'NKPA ** correction
    isRe_(6) = pDetails(3)

    'Asc.Node (Rahu) 7
    If isTru = False Then 'Mean node
        pReVal = swe_calc_ut(julDateUT, 10, plntFlg, pDetails(0), errReturn)
        pPos_(7) = hp_Rnd0To360v(pDetails(0) - newKPAyan)    'NKPA ** correction
        isRe_(7) = pDetails(3) * (-1)
    Else 'true node
        pReVal = swe_calc_ut(julDateUT, 11, plntFlg, pDetails(0), errReturn)
        pPos_(7) = hp_Rnd0To360v(pDetails(0) - newKPAyan)    'NKPA ** correction
        isRe_(7) = pDetails(3) * (-1)
    End If
    
    'Dec.Node (Kethu) 8
    pPos_(8) = pPos_(7) + 180#  'Not NeededNKPA ** correction
    If pPos_(8) >= 360# Then pPos_(8) = pPos_(8) - 360#
    isRe_(8) = 10#

    'Uranus 9
    pReVal = swe_calc_ut(julDateUT, 7, plntFlg, pDetails(0), errReturn)
    pPos_(9) = hp_Rnd0To360v(pDetails(0) - newKPAyan)    'NKPA ** correction
    isRe_(9) = pDetails(3)

    'Neptune 10
    pReVal = swe_calc_ut(julDateUT, 8, plntFlg, pDetails(0), errReturn)
    pPos_(10) = hp_Rnd0To360v(pDetails(0) - newKPAyan)    'NKPA ** correction
    isRe_(10) = pDetails(3)

    'Fortune 11
    pPos_(11) = hCus_c(0) + pPos_(1) - pPos_(0)  'Not NeededNKPA ** correction
    If pPos_(11) < 0 Then pPos_(11) = pPos_(11) + 360#
    If pPos_(11) >= 360# Then pPos_(11) = pPos_(11) - 360#
    isRe_(11) = 10#

    For b2 = 0 To 11    '#EXPORT#
        pPosCur(b2) = pPos_(b2)
    Next

    For b3 = 0 To 11    '#EXPORT#
        isReCur(b3) = isRe_(b3)
    Next

    'Calculation to find Sun rise time
    Dim rsGeoPos(2) As Double
    rsGeoPos(0) = lonVal
    rsGeoPos(1) = latVal
    rsGeoPos(2) = 0#

    'Dim rsJulDayUT As Double
    'rsJulDayUT = rsJulDayLocal - (tzValue / 24.0#)

    'Dim rsJulDayUTMod As Double
    'rsJulDayUTMod = rsJulDayUT - 1.0#

    Dim rsTime(2) As Double
    Dim rsRetVal As Long
    Dim rsErr As String
    rsRetVal = swe_rise_trans(julDateUT, 0, vbNullString, 0, 1, rsGeoPos(0), 0, 0, rsTime(0), rsErr)

    Dim rsTimeLocal As Double   'Local sun rise time
    rsTimeLocal = rsTime(0) + (gmtDiff / 24#)    ' - (tCorr / 24#)

    Dim rsReYear As Long
    Dim rsReMonth As Long
    Dim rsReDay As Long
    Dim rsReTime_ As Double

    Call swe_revjul(rsTimeLocal, 1, rsReYear, rsReMonth, rsReDay, rsReTime_)
    rsReTimeCur = rsReTime_    '#EXPORT#

    'Standard Weekday
    Dim stdWeekDay_ As Integer
    stdWeekDay_ = swe_day_of_week(julDateLocal_)
    stdWeekDayCur = stdWeekDay_    '#EXPORT#

    'Traditional Weekday
    Dim trdWeekDay_ As Integer
    If rsReTime_ <= bTime Then
        trdWeekDay_ = swe_day_of_week(julDateLocal_)
    Else
        trdWeekDay_ = swe_day_of_week(julDateLocal_ - 1#)
    End If
    trdWeekDayCur = trdWeekDay_    '#EXPORT#

End Sub

Private Sub A01_calcKPNatal()
'Compute KP Natal data

'Birth date
    Dim bDay As Integer
    Dim bMonth As Integer
    Dim bYear As Integer
    'Birth time
    Dim bHrs As Integer
    Dim bMin As Integer
    Dim bSec As Integer
    'Time correction
    Dim tCorrHrs As Integer
    Dim tCorrMin As Integer
    Dim tCorrPM As String
    'GMT difference
    Dim gmtDiffHrs As Integer
    Dim gmtDiffMin As Integer
    Dim gmtDiffPM As String

    'Longitude
    Dim lonDeg As Integer
    Dim lonMin As Integer
    Dim lonEW As String
    'Latitude
    Dim latDeg As Integer
    Dim latMin As Integer
    Dim latNS As String

    'KP Horary number
    '    Dim kpHorNum As Integer
    '    Dim kpHorAsc As Double

    'Lord Current Settings For Horary
    'Load data from dat file
    If isHorary = True Then
        'Dim depLoaded_() As String
        'depLoaded_() = hp_ReadDatFile(App.Path + "\Sett\depdet.dat", "_")

        Dim birthPlaceCur As String
        Dim tzPMCur As String, tzHHCur As Integer, tzMMCur As Integer
        Dim lonDegCur As Integer, lonMinCur As Integer, lonEWCur As String
        Dim latDegCur As Integer, latMinCur As Integer, latNSCur As String

        birthPlaceCur = txtDefPlaceCur.Text

        tzPMCur = Trim(txtTZPMCur.Text)
        tzHHCur = CInt(txtTZHHCur.Text)
        tzMMCur = CInt(txtTZMMCur.Text)

        lonDegCur = CInt(txtLonDDCur.Text)
        lonMinCur = CInt(txtLonMMCur.Text)
        lonEWCur = Trim(txtLonEWCur.Text)

        latDegCur = CInt(txtLatDDCur.Text)
        latMinCur = CInt(txtLatMMCur.Text)
        latNSCur = Trim(txtLatNSCur.Text)
    End If

    'Assigning Variables
    If isHorary = False Then
        'Natal
        bDay = CInt(cmbDay.Text)
        bMonth = CInt(cmbMonth.Text)
        bYear = CInt(cmbYear.Text)

        bHrs = CInt(cmbHH.Text)
        bMin = CInt(cmbMM.Text)
        bSec = CInt(cmbSS.Text)

        tCorrHrs = CInt(cmbTcHH.Text)
        tCorrMin = CInt(cmbTcMM.Text)
        tCorrPM = cmbTcPM.Text

        gmtDiffHrs = CInt(cmbTzHH.Text)
        gmtDiffMin = CInt(cmbTzMM.Text)
        gmtDiffPM = cmbTzPM.Text

        lonDeg = CInt(cmbLonDeg.Text)
        lonMin = CInt(cmbLonMin.Text)
        lonEW = cmbLonEW.Text

        latDeg = CInt(cmbLatDeg.Text)
        latMin = CInt(cmbLatMin.Text)
        latNS = cmbLatNS.Text
    Else
        'Horary
        bDay = CInt(txtDayCur.Text)
        bMonth = CInt(txtMonthCur.Text)
        bYear = CInt(txtYearCur.Text)

        bHrs = CInt(txtHHCur.Text)
        bMin = CInt(txtMMCur.Text)
        bSec = CInt(txtSSCur.Text)

        tCorrHrs = 0            'CInt(cmbTcHH.Text)
        tCorrMin = 0            'CInt(cmbTcMM.Text)
        tCorrPM = "+"           ' cmbTcPM.Text

        gmtDiffHrs = tzHHCur    'CInt(cmbTzHH.Text)
        gmtDiffMin = tzMMCur    'CInt(cmbTzMM.Text)
        gmtDiffPM = tzPMCur     'cmbTzPM.Text

        lonDeg = lonDegCur      'CInt(cmbLonDeg.Text)
        lonMin = lonMinCur      'CInt(cmbLonMin.Text)
        lonEW = lonEWCur        'cmbLonEW.Text

        latDeg = latDegCur      'CInt(cmbLatDeg.Text)
        latMin = latMinCur      'CInt(cmbLatMin.Text)
        latNS = latNSCur        'cmbLatNS.Text

    End If

    'Rotate Kundali
    rotVal = CInt(frmMain.cmbRot.Text) - 1

    'For printing
    'RTF
    If isHorary = False Then
        ppName = "NAME" + vbTab + vbTab + ": " + txtName.Text + " (" + cmbMF.Text + ")"
        ppNamePdf = txtName.Text + " (" + cmbMF.Text + ")"
        bNmeM = "NATAL CHART"
        ppPlaceName = "PLACE" + vbTab + vbTab + ": " + txtBirthPlace.Text
        ppPlaceNamePdf = txtBirthPlace.Text
    Else
        ppName = "HORARY NUMBER" + vbTab + ": " + CStr(sNo249Hor) + " / 249"
        ppNamePdf = CStr(sNo249Hor) + " / 249"
        bNmeM = Format(sNo249Hor, "000") + " / 249  "
        ppPlaceName = "PLACE" + vbTab + vbTab + ": " + birthPlaceCur
        ppPlaceNamePdf = birthPlaceCur
    End If
    
    
    Dim geoPrint As String
    If isGeo = True Then
        geoPrint = " (Geocentric)"
    Else
        geoPrint = " (Geographic)"
    End If
    
    ppBDate = "DATE" + vbTab + vbTab + ": " + Format(bDay, "00") + "/" + Format(bMonth, "00") + "/" + Format(bYear, "0000")
    ppBDatePdf = Format(bDay, "00") + "/" + Format(bMonth, "00") + "/" + Format(bYear, "0000")
    ppBTime = "TIME" + vbTab + vbTab + ": " + Format(bHrs, "00") + ":" + Format(bMin, "00") + ":" + Format(bSec, "00") + " H  ( Correction : " + tCorrPM + Format(tCorrHrs, "00") + ":" + Format(tCorrMin, "00") + " )"
    ppBTimePdf = Format(bHrs, "00") + ":" + Format(bMin, "00") + ":" + Format(bSec, "00") + " H  ( Correction : " + tCorrPM + Format(tCorrHrs, "00") + ":" + Format(tCorrMin, "00") + " )"
    ppPlaceDetails = "PLACE DETAILS" + vbTab + ": " + "Lon : " + Format(lonDeg, "000") + lonEW + Format(lonMin, "00") + "  Lat : " + Format(latDeg, "00") + latNS + Format(latMin, "00") + geoPrint + "  T.Zone : " + gmtDiffPM + Format(gmtDiffHrs, "00") + ":" + Format(gmtDiffMin, "00")
    ppPlaceDetailsPdf = Format(lonDeg, "000") + lonEW + Format(lonMin, "00") + "  Lat : " + Format(latDeg, "00") + latNS + Format(latMin, "00") + geoPrint + "  T.Zone : " + gmtDiffPM + Format(gmtDiffHrs, "00") + ":" + Format(gmtDiffMin, "00")

    bDateM = "DATE  : " + Format(bDay, "00") + "/" + Format(bMonth, "00") + "/" + Format(bYear, "0000")
    bTimeM = "TIME  : " + Format(bHrs, "00") + ":" + Format(bMin, "00") + ":" + Format(bSec, "00") + " H"
    bCorrM = "CORR  : " + tCorrPM + Format(tCorrHrs, "00") + ":" + Format(tCorrMin, "00")
    bPlaceM = "PLACE : " + "Lon:" + Format(lonDeg, "000") + lonEW + Format(lonMin, "00") + "  Lat:" + Format(latDeg, "00") + latNS + Format(latMin, "00")
    bTzM = "T.ZONE: " + gmtDiffPM + Format(gmtDiffHrs, "00") + ":" + Format(gmtDiffMin, "00") + " GMT"

    'Calculation using SWE
    Dim bTime As Double         'Standers birth time as double value
    bTime = CDbl(bHrs) + (CDbl(bMin) / 60#) + (CDbl(bSec) / 3600#)

    Dim tCorr As Double         'Time correction as double value
    tCorr = CDbl(tCorrHrs) + (CDbl(tCorrMin) / 60#)
    If tCorrPM = "-" Then tCorr = tCorr * (-1#)

    Dim gmtDiff As Double       'GMT differance as double value
    gmtDiff = CDbl(gmtDiffHrs) + (CDbl(gmtDiffMin) / 60#)
    If gmtDiffPM = "-" Then gmtDiff = gmtDiff * (-1#)

    Dim latVal As Double        'Latitude as double value
    latVal = CDbl(latDeg) + (CDbl(latMin) / 60#)
    If isGeo = True Then latVal = kp_GeoCorr(latVal)    'Geo correction
    If UCase(latNS) = "S" Then latVal = latVal * (-1#)

    Dim lonVal As Double        'Longitude as double value
    lonVal = CDbl(lonDeg) + (CDbl(lonMin) / 60#)
    If UCase(lonEW) = "W" Then lonVal = lonVal * (-1#)

    'Call swe_set_ephe_path(App.Path + "\Eph")       'Set Swiss Eph Path
    'Call swe_set_sid_mode(5, 0, 0)    '2415020.5, 22.3708333333333                    'Set sid mode to KP (KP Ayanamsa is used)

    'Julian date calculation
    Dim julDateLocal_ As Double       'Raw julian day according to the birth data
    julDateLocal_ = swe_julday(bYear, bMonth, bDay, bTime, 1)
    julDateLocal = julDateLocal_    '#EXPORT#

    Dim julDateUT As Double          'Corrected julian date
    julDateUT = julDateLocal_ - (gmtDiff / 24#) + (tCorr / 24#)

    'NUT
    '    Dim nutPReVal As Long
    '    Dim nutPDetails(5) As Double
    '    Dim nutErrReturn As String
    '    nutPReVal = swe_calc_ut(julDateUT, -1, 0, nutPDetails(0), nutErrReturn)

    ' NEW KP AYANAMSA
    Dim newKPAyan As Double
    newKPAyan = kp_NKPA(julDateUT)
    kpAyan = newKPAyan    '#EXPORT#


    'KP Ayanamsa Swiss
    'Dim kpAyan_ As Double
    'kpAyan_ = swe_get_ayanamsa_ut(julDateUT)
    'kpAyan = kpAyan_    '#EXPORT#

    'Ayanamsa Diff
    'Dim ayanDiff As Double
    'ayanDiff = newKPAyan - kpAyan_ '( +ve Value)


    'House cusps calculation
    Dim hCusp_(12) As Double
    Dim hCus_(11) As Double
    Dim ascmc(9) As Double
    Dim hReVal As Long

    Dim fCusNir As Double   'Hor
    Dim horCus() As Double  'Hor
    Dim jz As Byte          'Hor

    If isHorary = False Then    'GLOBAL#####
        hReVal = swe_houses(julDateUT, latVal, lonVal, Asc("P"), hCusp_(0), ascmc(0))
    Else
        fCusNir = kp_Sub249Hor(sNo249Hor)
        horCus = kp_Horar(julDateUT, fCusNir, latVal, newKPAyan)

        For jz = 0 To 12
            hCusp_(jz) = horCus(jz)
        Next
    End If

    'Raw Assig
    Dim jk As Byte
    For jk = 0 To 11
        hCus_(jk) = hp_Rnd0To360v(hCusp_(jk + 1) - newKPAyan)    'NKPA** Correction
    Next

    'Rotate kp_Rot
    Dim hCusR(11) As Double
    Dim h1 As Byte
    For h1 = 0 To 11
        hCusR(h1) = hCus_(hp_Rnd0To11v(h1 + rotVal))
    Next

    Dim b0 As Byte    '#EXPORT VAR#
    For b0 = 0 To 11    '#EXPORT#
        hCuspI(b0 + 1) = hCusR(b0)
        hCusp(b0 + 1) = hCusR(b0)
    Next


    Dim b1 As Byte    '#EXPORT VAR#
    For b1 = 0 To 11    '#EXPORT#  hCuspI
        hCus(b1) = hCusR(b1)
    Next

    'Planetary Position]
    Dim pPos_(11) As Double
    Dim isRe_(11) As Double
    Dim pReVal As Long
    Dim pDetails(5) As Double
    Dim errReturn As String

    'Sun (Ravi) 0
    pReVal = swe_calc_ut(julDateUT, 0, plntFlg, pDetails(0), errReturn)
    pPos_(0) = hp_Rnd0To360v(pDetails(0) - newKPAyan)    'NKPA** Correction
    isRe_(0) = pDetails(3)

    'Moon (Chandra) 1
    pReVal = swe_calc_ut(julDateUT, 1, plntFlg, pDetails(0), errReturn)
    pPos_(1) = hp_Rnd0To360v(pDetails(0) - newKPAyan)    'NKPA** Correction
    isRe_(1) = pDetails(3)

    'Mars (Kuja) 2
    pReVal = swe_calc_ut(julDateUT, 4, plntFlg, pDetails(0), errReturn)
    pPos_(2) = hp_Rnd0To360v(pDetails(0) - newKPAyan)    'NKPA** Correction
    isRe_(2) = pDetails(3)

    'Mercury (Budha) 3
    pReVal = swe_calc_ut(julDateUT, 2, plntFlg, pDetails(0), errReturn)
    pPos_(3) = hp_Rnd0To360v(pDetails(0) - newKPAyan)    'NKPA** Correction
    isRe_(3) = pDetails(3)

    'Jupiter (Guru) 4
    pReVal = swe_calc_ut(julDateUT, 5, plntFlg, pDetails(0), errReturn)
    pPos_(4) = hp_Rnd0To360v(pDetails(0) - newKPAyan)    'NKPA** Correction
    isRe_(4) = pDetails(3)

    'Venus (Sukra) 5
    pReVal = swe_calc_ut(julDateUT, 3, plntFlg, pDetails(0), errReturn)
    pPos_(5) = hp_Rnd0To360v(pDetails(0) - newKPAyan)    'NKPA** Correction
    isRe_(5) = pDetails(3)

    'Saturn (Sani) 6
    pReVal = swe_calc_ut(julDateUT, 6, plntFlg, pDetails(0), errReturn)
    pPos_(6) = hp_Rnd0To360v(pDetails(0) - newKPAyan)    'NKPA** Correction
    isRe_(6) = pDetails(3)

    'Asc.Node (Rahu) 7
    If isTru = False Then 'Mean Node
        pReVal = swe_calc_ut(julDateUT, 10, plntFlg, pDetails(0), errReturn)
        pPos_(7) = hp_Rnd0To360v(pDetails(0) - newKPAyan)    'NKPA** Correction
        isRe_(7) = pDetails(3) * (-1)
    Else 'True Node
        pReVal = swe_calc_ut(julDateUT, 11, plntFlg, pDetails(0), errReturn)
        pPos_(7) = hp_Rnd0To360v(pDetails(0) - newKPAyan)    'NKPA** Correction
        isRe_(7) = pDetails(3) * (-1)
    End If
    'Dec.Node (Kethu) 8
    pPos_(8) = pPos_(7) + 180#  'Not Needed NKPA** Correction
    If pPos_(8) >= 360# Then pPos_(8) = pPos_(8) - 360#
    isRe_(8) = 10#

    'Uranus 9
    pReVal = swe_calc_ut(julDateUT, 7, plntFlg, pDetails(0), errReturn)
    pPos_(9) = hp_Rnd0To360v(pDetails(0) - newKPAyan)    'NKPA** Correction
    isRe_(9) = pDetails(3)

    'Neptune 10
    pReVal = swe_calc_ut(julDateUT, 8, plntFlg, pDetails(0), errReturn)
    pPos_(10) = hp_Rnd0To360v(pDetails(0) - newKPAyan)    'NKPA** Correction
    isRe_(10) = pDetails(3)

    'Fortune 11  hCusCur(0) must be used for Horary
    If isHorary = False Then
        pPos_(11) = hCus_(0) + pPos_(1) - pPos_(0)  'Not Needed NKPA** Correction
        If pPos_(11) < 0# Then pPos_(11) = pPos_(11) + 360#
        If pPos_(11) >= 360# Then pPos_(11) = pPos_(11) - 360#
        isRe_(11) = 10#
    Else
        pPos_(11) = hCusCur(0) + pPos_(1) - pPos_(0)  'Not Needed NKPA** Correction
        If pPos_(11) < 0# Then pPos_(11) = pPos_(11) + 360#
        If pPos_(11) >= 360# Then pPos_(11) = pPos_(11) - 360#
        isRe_(11) = 10#
    End If


    Dim b2 As Byte    '#EXPORT VAR#
    For b2 = 0 To 11    '#EXPORT#
        pPos(b2) = pPos_(b2)
    Next

    Dim b3 As Byte    '#EXPORT VAR#
    For b3 = 0 To 11    '#EXPORT#
        isRe(b3) = isRe_(b3)
    Next

    'kp_IncludedHouse   Ravi=0, Chandra=1,......, Fortune=11
    Dim incHouse_(11) As Integer
    Dim i1 As Byte
    For i1 = 0 To 11
        incHouse_(i1) = kp_IncludedHouse(pPos_(i1))
    Next

    Dim b4 As Byte    '#EXPORT VAR#
    For b4 = 0 To 11    '#EXPORT#
        incHouse(b4) = incHouse_(b4)
    Next

    'Sun rise time
    Dim rsGeoPos(2) As Double
    rsGeoPos(0) = lonVal
    rsGeoPos(1) = latVal
    rsGeoPos(2) = 0#

    'Dim rsJulDayUT As Double
    'rsJulDayUT = rsJulDayLocal - (tzValue / 24.0#)

    'Dim rsJulDayUTMod As Double
    'rsJulDayUTMod = rsJulDayUT - 1.0#

    Dim rsTime(2) As Double
    Dim rsRetVal As Long
    Dim rsErr As String
    rsRetVal = swe_rise_trans(julDateUT, 0, vbNullString, 0, 1, rsGeoPos(0), 0, 0, rsTime(0), rsErr)

    Dim rsTimeLocal As Double   'Local sun rise time
    rsTimeLocal = rsTime(0) + (gmtDiff / 24#) - (tCorr / 24#)

    Dim rsReYear As Long
    Dim rsReMonth As Long
    Dim rsReDay As Long
    Dim rsReTime_ As Double

    Call swe_revjul(rsTimeLocal, 1, rsReYear, rsReMonth, rsReDay, rsReTime_)
    rsReTime = rsReTime_    '#EXPORT#

    'Standard Weekday
    Dim stdWeekDay_ As Integer
    stdWeekDay_ = swe_day_of_week(julDateLocal_)
    stdWeekDay = stdWeekDay_    '#EXPORT#

    'Traditional Weekday
    Dim trdWeekDay_ As Integer
    If rsReTime_ <= bTime Then
        trdWeekDay_ = swe_day_of_week(julDateLocal_)
    Else
        trdWeekDay_ = swe_day_of_week(julDateLocal_ - 1#)
    End If
    trdWeekDay = trdWeekDay_    '#EXPORT#

End Sub


Private Sub A06_calcSig()
'Find the significators

'(01)Bhava Lord
    Dim bhavaLordStr(11) As String
    Dim bhavaLord(11) As Integer
    Dim a1 As Byte
    For a1 = 0 To 11
        bhavaLord(a1) = kp_RashiLordInt(hCus(a1), False)
        bhavaLordStr(a1) = planetName1Tr(kp_RashiLordInt(hCus(a1), False))
    Next

    Dim n1 As Byte    '#EXPOTR VAR#
    For n1 = 0 To 11    '#EXPOTR
        conSig0(n1) = bhavaLordStr(n1)
    Next

    '(02)Planets in Bhava Lord
    Dim planetsInBhavaLordStr(11) As String
    Dim starLordOfPlanet(11) As Integer
    Dim a2 As Byte, a3 As Byte, a4 As Byte
    For a4 = 0 To 11
        starLordOfPlanet(a4) = kp_StarLordInt(pPos(a4), False)
    Next
    For a2 = 0 To 11
        For a3 = 0 To 11
            If bhavaLord(a2) = starLordOfPlanet(a3) Then
                planetsInBhavaLordStr(a2) = planetsInBhavaLordStr(a2) + planetName1Tr(a3) + " "
            End If
        Next
    Next

    Dim n2 As Byte    '#EXPOTR VAR#
    For n2 = 0 To 11    '#EXPOTR
        conSig1(n2) = planetsInBhavaLordStr(n2)
    Next

    '(03)Tenanted Planets In each bhava
    Dim positedPlanets(11) As String
    Dim incBhav(11) As Integer
    Dim a5 As Byte, a6 As Byte, a7 As Byte

    For a5 = 0 To 11
        incBhav(a5) = kp_IncludedHouse(pPos(a5)) - 1
    Next

    For a6 = 0 To 11
        For a7 = 0 To 11
            If incBhav(a7) = a6 Then
                positedPlanets(a6) = positedPlanets(a6) + planetName1Tr(a7) + " "
            End If
        Next
    Next

    Dim n3 As Byte    '#EXPOTR VAR#
    For n3 = 0 To 11    '#EXPOTR
        conSig2(n3) = positedPlanets(n3)
    Next

    '(04)Tenanted Planets In Planets which are posited in each bhava
    Dim a8 As Byte, a9 As Byte, a0 As Byte

    Dim ppPlanets(11) As String
    Dim tmpA() As String

    For a8 = 0 To 11
        tmpA = Split(positedPlanets(a8), " ")
        For a9 = 0 To 11
            If UBound(tmpA) > 0 Then
                For a0 = LBound(tmpA) To UBound(tmpA)
                    If planetName1Tr(kp_StarLordInt(pPos(a9), False)) = tmpA(a0) Then
                        ppPlanets(a8) = ppPlanets(a8) + planetName1Tr(a9) + " "
                    End If
                Next
            End If
        Next
    Next

    Dim n4 As Byte    '#EXPOTR VAR#
    For n4 = 0 To 11    '#EXPOTR
        conSig3(n4) = ppPlanets(n4)
    Next

    'Find planetwise Sig
    Dim pwSig(11) As String
    Dim allSig(11) As String
    Dim tmp() As String
    Dim d2 As Byte, d3 As Byte, d4 As Byte, d5 As Byte


    For d2 = 0 To 11
        allSig(d2) = ppPlanets(d2) + positedPlanets(d2) + planetsInBhavaLordStr(d2) + bhavaLordStr(d2)
    Next

    For d3 = 0 To 11
        tmp = Split(allSig(d3), " ")
        For d4 = 0 To 11
            For d5 = LBound(tmp) To UBound(tmp)
                If planetName1Tr(d4) = tmp(d5) Then
                    pwSig(d4) = pwSig(d4) + CStr(d3 + 1) + " "
                End If
            Next
        Next
    Next

    Dim n5 As Byte    '#EXPOTR VAR#
    For n5 = 0 To 11    '#EXPOTR
        plntWSig(n5) = pwSig(n5)
    Next

    '    Dim z1 As Byte, z2 As Byte
    '
    '    For z1 = 0 To 11
    '        frmAbout.RT1.SelText = Format(z1 + 1, "00") + " " + allSig(z1) + vbNewLine
    '    Next
    '
    '    frmAbout.RT1.SelText = vbNewLine
    '
    '    For z2 = 0 To 11
    '        frmAbout.RT1.SelText = planetName1Tr(z2) + " " + pwSig(z2) + vbNewLine
    '    Next
    '
    '
    '    Load frmAbout
    '    frmAbout.Show
    '
    '
    '    'frmAbout.RT1.SelText = CStr(UBound(plntSp_01)) + vbNewLine
    '    frmAbout.RT1.SelText = vbNewLine
    '    frmAbout.RT1.SelText = vbNewLine

End Sub

Private Sub A03_calcVAspects()
'To calculate Vedic aspects

'Included rashi int of planets   0, 1, 2,....11
    Dim incRashiOfPlanets(11) As Byte
    Dim i1 As Byte
    For i1 = 0 To 11
        incRashiOfPlanets(i1) = kp_RashiInt(pPos(i1))
    Next

    'Included rashi int of cusps   0, 1, 2,....11
    Dim incRashiOfCusps(11) As Byte
    Dim i2 As Byte
    For i2 = 0 To 11
        incRashiOfCusps(i2) = kp_RashiInt(hCus(i2))
    Next

    '    Dim lc As Byte, relRashiOfPlnt(11) As Byte
    '    For lc = 0 To 11
    '        relRashiOfPlnt(lc) = kp_AspRealInt(incRashiOfPlanets(lc), kp_RashiInt(hCus(0)))
    '    Next

    Dim aspectedRashi(11, 3) As Byte
    Dim aspVedic_(6) As String

    'To find aspected Rashi for each house and tenanted cusps
    'No aspect > 12
    Dim i3 As Byte, i4 As Byte
    For i3 = 0 To 8
        For i4 = 0 To 3
            aspectedRashi(i3, i4) = 20
        Next
    Next

    'Dim k0 As Byte, k1 As Byte, k2 As Byte, k3 As Byte, k4 As Byte, k5 As Byte, k6 As Byte
    'Dim pAspV(6) As String

    ' 0. Ravi   -   7
    aspectedRashi(0, 0) = incRashiOfPlanets(0)
    aspectedRashi(0, 1) = hp_Rnd0To11v(incRashiOfPlanets(0) + 6)

    '1. Chandra -   7
    aspectedRashi(1, 0) = incRashiOfPlanets(1)
    aspectedRashi(1, 1) = hp_Rnd0To11v(incRashiOfPlanets(1) + 6)

    '2. Kuja    -   4, 7, 8
    aspectedRashi(2, 0) = incRashiOfPlanets(2)
    aspectedRashi(2, 1) = hp_Rnd0To11v(incRashiOfPlanets(2) + 3)
    aspectedRashi(2, 2) = hp_Rnd0To11v(incRashiOfPlanets(2) + 6)
    aspectedRashi(2, 3) = hp_Rnd0To11v(incRashiOfPlanets(2) + 7)

    '3. Budha   -   7
    aspectedRashi(3, 0) = incRashiOfPlanets(3)
    aspectedRashi(3, 1) = hp_Rnd0To11v(incRashiOfPlanets(3) + 6)

    '4. Guru    -   5, 7, 9
    aspectedRashi(4, 0) = incRashiOfPlanets(4)
    aspectedRashi(4, 1) = hp_Rnd0To11v(incRashiOfPlanets(4) + 4)
    aspectedRashi(4, 2) = hp_Rnd0To11v(incRashiOfPlanets(4) + 6)
    aspectedRashi(4, 3) = hp_Rnd0To11v(incRashiOfPlanets(4) + 8)

    '5. Shukra  -   7
    aspectedRashi(5, 0) = incRashiOfPlanets(5)
    aspectedRashi(5, 1) = hp_Rnd0To11v(incRashiOfPlanets(5) + 6)

    '6. Shani   -   3, 7, 10
    aspectedRashi(6, 0) = incRashiOfPlanets(6)
    aspectedRashi(6, 1) = hp_Rnd0To11v(incRashiOfPlanets(6) + 2)
    aspectedRashi(6, 2) = hp_Rnd0To11v(incRashiOfPlanets(6) + 6)
    aspectedRashi(6, 3) = hp_Rnd0To11v(incRashiOfPlanets(6) + 9)

    '7. Rahu
    aspectedRashi(7, 0) = incRashiOfPlanets(7)

    '8. Kethu
    aspectedRashi(8, 0) = incRashiOfPlanets(8)

    ''''''''''''''''''''''''''''''

    '0. Ravi - 7th
    Dim k0 As Byte
    For k0 = 0 To 11
        If aspectedRashi(0, 1) = incRashiOfCusps(k0) Then aspVedic_(0) = aspVedic_(0) + CStr((k0 + 1)) + "(7th)" + ", "
    Next

    '1. Chandra - 7th
    Dim k1 As Byte
    For k1 = 0 To 11
        If aspectedRashi(1, 1) = incRashiOfCusps(k1) Then aspVedic_(1) = aspVedic_(1) + CStr((k1 + 1)) + "(7th)" + ", "
    Next

    '2. Kuja - 4th,7th,8th
    Dim k2 As Byte
    For k2 = 0 To 11
        If aspectedRashi(2, 1) = incRashiOfCusps(k2) Then aspVedic_(2) = aspVedic_(2) + CStr((k2 + 1)) + "(4th)" + ", "
        If aspectedRashi(2, 2) = incRashiOfCusps(k2) Then aspVedic_(2) = aspVedic_(2) + CStr((k2 + 1)) + "(7th)" + ", "
        If aspectedRashi(2, 3) = incRashiOfCusps(k2) Then aspVedic_(2) = aspVedic_(2) + CStr((k2 + 1)) + "(4th)" + ", "
    Next

    '3. Budha - 7th
    Dim k3 As Byte
    For k3 = 0 To 11
        If aspectedRashi(3, 1) = incRashiOfCusps(k3) Then aspVedic_(3) = aspVedic_(3) + CStr((k3 + 1)) + "(7th)" + ", "
    Next

    '4. Guru - 5th,7th,9th
    Dim k4 As Byte
    For k4 = 0 To 11
        If aspectedRashi(4, 1) = incRashiOfCusps(k4) Then aspVedic_(4) = aspVedic_(4) + CStr((k4 + 1)) + "(5th)" + ", "
        If aspectedRashi(4, 2) = incRashiOfCusps(k4) Then aspVedic_(4) = aspVedic_(4) + CStr((k4 + 1)) + "(7th)" + ", "
        If aspectedRashi(4, 3) = incRashiOfCusps(k4) Then aspVedic_(4) = aspVedic_(4) + CStr((k4 + 1)) + "(9th)" + ", "
    Next

    '5. Sukra - 7th
    Dim k5 As Byte
    For k5 = 0 To 11
        If aspectedRashi(5, 1) = incRashiOfCusps(k5) Then aspVedic_(5) = aspVedic_(5) + CStr((k5 + 1)) + "(7th)" + ", "
    Next

    '6. Sani - 3rd,7th,10th
    Dim k6 As Byte
    For k6 = 0 To 11
        If aspectedRashi(6, 1) = incRashiOfCusps(k6) Then aspVedic_(6) = aspVedic_(6) + CStr((k6 + 1)) + "(3rd)" + ", "
        If aspectedRashi(6, 2) = incRashiOfCusps(k6) Then aspVedic_(6) = aspVedic_(6) + CStr((k6 + 1)) + "(7th)" + ", "
        If aspectedRashi(6, 3) = incRashiOfCusps(k6) Then aspVedic_(6) = aspVedic_(6) + CStr((k6 + 1)) + "(10th)" + ", "
    Next

    '#EXPORT VAR#
    Dim n1 As Byte

    For n1 = 0 To 6    '#EXPORT#
        aspVedic(n1) = aspVedic_(n1)
    Next

    'Vedic aspects planets to planets
    Dim vAspP(6) As String

    '0. Ravi - 7th
    Dim i5 As Byte
    For i5 = 0 To 11
        If aspectedRashi(0, 1) = incRashiOfPlanets(i5) Then vAspP(0) = vAspP(0) + planetName1Tr(i5) + "(7th)" + ", "
    Next

    '1. Chandra - 7th
    Dim i6 As Byte
    For i6 = 0 To 11
        If aspectedRashi(1, 1) = incRashiOfPlanets(i6) Then vAspP(1) = vAspP(1) + planetName1Tr(i6) + "(7th)" + ", "
    Next

    '2. Kuja - 4th,7th,8th
    Dim i7 As Byte
    For i7 = 0 To 11
        If aspectedRashi(2, 1) = incRashiOfPlanets(i7) Then vAspP(2) = vAspP(2) + planetName1Tr(i7) + "(4th)" + ", "
        If aspectedRashi(2, 2) = incRashiOfPlanets(i7) Then vAspP(2) = vAspP(2) + planetName1Tr(i7) + "(7th)" + ", "
        If aspectedRashi(2, 3) = incRashiOfPlanets(i7) Then vAspP(2) = vAspP(2) + planetName1Tr(i7) + "(4th)" + ", "
    Next

    '3. Budha - 7th
    Dim i8 As Byte
    For i8 = 0 To 11
        If aspectedRashi(3, 1) = incRashiOfPlanets(i8) Then vAspP(3) = vAspP(3) + planetName1Tr(i8) + "(7th)" + ", "
    Next

    '4. Guru - 5th,7th,9th
    Dim i9 As Byte
    For i9 = 0 To 11
        If aspectedRashi(4, 1) = incRashiOfPlanets(i9) Then vAspP(4) = vAspP(4) + planetName1Tr(i9) + "(5th)" + ", "
        If aspectedRashi(4, 2) = incRashiOfPlanets(i9) Then vAspP(4) = vAspP(4) + planetName1Tr(i9) + "(7th)" + ", "
        If aspectedRashi(4, 3) = incRashiOfPlanets(i9) Then vAspP(4) = vAspP(4) + planetName1Tr(i9) + "(9th)" + ", "
    Next

    '5. Sukra - 7th
    Dim i10 As Byte
    For i10 = 0 To 11
        If aspectedRashi(5, 1) = incRashiOfPlanets(i10) Then vAspP(5) = vAspP(5) + planetName1Tr(i10) + "(7th)" + ", "
    Next

    '6. Sani - 3rd,7th,10th
    Dim i11 As Byte
    For i11 = 0 To 11
        If aspectedRashi(6, 1) = incRashiOfPlanets(i11) Then vAspP(6) = vAspP(6) + planetName1Tr(i11) + "(3rd)" + ", "
        If aspectedRashi(6, 2) = incRashiOfPlanets(i11) Then vAspP(6) = vAspP(6) + planetName1Tr(i11) + "(7th)" + ", "
        If aspectedRashi(6, 3) = incRashiOfPlanets(i11) Then vAspP(6) = vAspP(6) + planetName1Tr(i11) + "(10th)" + ", "
    Next

    Dim k As Byte   '#EXPORT VAR#
    For k = 0 To 6    '#EXPORT#
        vedAspP(k) = vAspP(k)
    Next

    'Find planets which are in same rashi
    Dim smRas(6) As String
    Dim j1 As Byte, j2 As Byte

    For j1 = 0 To 6
        For j2 = 0 To 11
            If (incRashiOfPlanets(j1) = incRashiOfPlanets(j2)) And (j1 <> j2) Then
                smRas(j1) = smRas(j1) + planetName1Tr(j2) + ", "
            End If
        Next
    Next

    Dim k7 As Byte  '#EXPORT VAR#
    For k7 = 0 To 6    '#EXPORT#
        smeRas(k7) = smRas(k7)
    Next

End Sub

Private Sub A07_calc4StepSig()
'To find the 4 Step Significators
'Kethu = 0, sukra = 1, ....

'STEP 1
'To find own bhava of each planet
'Dim ownBhava(8) As String
    Dim a1 As Byte, a2 As Byte

    For a1 = 0 To 8
        For a2 = 0 To 11
            If a1 = kp_RashiLordInt(hCus(a2), True) Then
                ownBhava(a1) = ownBhava(a1) + CStr(a2 + 1) + " "
            End If
        Next
    Next

    'To find posited bhava of each planet
    'Dim posBhava(8) As Integer
    Dim a3 As Byte

    For a3 = 0 To 8
        posBhava(a3) = kp_IncludedHouse(pPos(planet2Int(a3)))
    Next

    'STEP 2
    '    Dim stLords(8) As Integer
    '    Dim stLordsOwn(8) As String
    '    Dim stLordsPos(8) As Integer
    Dim a4 As Byte, a5 As Byte, a6 As Byte

    For a4 = 0 To 8
        stLords(a4) = kp_StarLordInt(pPos(planet2Int(a4)), True)
    Next

    For a5 = 0 To 8
        stLordsOwn(a5) = ownBhava(stLords(a5))
    Next

    For a6 = 0 To 8
        stLordsPos(a6) = posBhava(stLords(a6))
    Next

    'STEP 3
    '    Dim sbLords(8) As Integer
    '    Dim sbLordsOwn(8) As String
    '    Dim sbLordsPos(8) As Integer
    Dim a7 As Byte, a8 As Byte, a9 As Byte

    For a7 = 0 To 8
        sbLords(a7) = kp_SubLord(pPos(planet2Int(a7)), False)
    Next

    For a8 = 0 To 8
        sbLordsOwn(a8) = ownBhava(sbLords(a8))
    Next

    For a9 = 0 To 8
        sbLordsPos(a9) = posBhava(sbLords(a9))
    Next

    'STEP 4
    '    Dim stSbLords(8) As Integer
    '    Dim stSbLordsOwn(8) As String
    '    Dim stSbLordsPos(8) As Integer
    Dim b1 As Byte, b2 As Byte, b3 As Byte

    For b1 = 0 To 8
        stSbLords(b1) = stLords(sbLords(b1))
    Next

    For b2 = 0 To 8
        stSbLordsOwn(b2) = ownBhava(stSbLords(b2))
    Next

    For b3 = 0 To 8
        stSbLordsPos(b3) = posBhava(stSbLords(b3))
    Next

    'Empty Bhavas
    'Dim emptyBhava As String
    Dim tmp As Integer
    Dim b4 As Byte, b5 As Byte

    For b4 = 1 To 12
        tmp = 0
        For b5 = 0 To 8
            If b4 = posBhava(b5) Then
                tmp = tmp + 1
            End If
        Next
        If tmp < 1 Then
            emptyBhava = emptyBhava + CStr(b4) + " "
        End If
    Next

    'Planets which, has no other planets in their stars
    'Dim noPInStar As String
    Dim tmp1 As Integer
    Dim b6 As Byte, b7 As Byte

    'planetName2Tr
    For b6 = 0 To 8
        tmp1 = 0
        For b7 = 0 To 8
            If b6 = stLords(b7) Then
                tmp1 = tmp1 + 1
            End If
        Next
        If tmp1 < 1 Then
            noPInStar = noPInStar + planetName2Tr(b6) + " "
        End If
    Next

    'Planets in self constallation
    Dim b8 As Byte
    '    Dim selfCons As String

    For b8 = 0 To 8
        If b8 = stLords(b8) Then
            selfCons = selfCons + planetName2Tr(b8) + " "
        End If
    Next

End Sub

Private Sub C02_RulPlanets()
'Find Ruling Planets

    Dim hCusNR(11) As Double 'pPos(11) As Double,
    Dim k1 As Integer

    For k1 = 0 To 11
        Dim k2 As Integer
        k2 = k1 - rotVal
        If k2 < 0 Then k2 = k2 + 11
        hCusNR(k1) = hCus(hp_Rnd0To11v(k2))
    Next

    Dim rulPlanets(29) As String

    rulPlanets(12) = "RULING PLANETS" + vbNewLine

    rulPlanets(13) = vbNewLine
    rulPlanets(14) = txtDayCur.Text + "/" + txtMonthCur.Text + "/" + txtYearCur.Text + " - " + txtHHCur.Text + ":" + txtMMCur.Text + ":" + txtSSCur.Text 'Format(Now, "DD/MM/YYYY - HH:MM:SS") + " H" + vbNewLine
    rulPlanets(15) = vbNewLine

    rulPlanets(16) = "RULING PLANET        " + "NATAL" + Space(8) + "CURRENT" + vbNewLine
    rulPlanets(17) = String(41, "-") + vbNewLine

    rulPlanets(18) = "Day Lord             " + weekDayPlanetNameTr(trdWeekDay) + Space(10) + weekDayPlanetNameTr(trdWeekDayCur) + vbNewLine

    rulPlanets(19) = "Moon Sign Lord       " + planetName1Tr(kp_RashiLordInt(pPos(1), False)) + Space(10) + planetName1Tr(kp_RashiLordInt(pPosCur(1), False)) + vbNewLine
    rulPlanets(20) = "Moon Star Lord       " + planetName1Tr(kp_StarLordInt(pPos(1), False)) + Space(10) + planetName1Tr(kp_StarLordInt(pPosCur(1), False)) + vbNewLine
    rulPlanets(21) = "Moon Sub Lord        " + planetName2Tr(kp_SubLord(pPos(1), False)) + Space(10) + planetName2Tr(kp_SubLord(pPosCur(1), False)) + vbNewLine

    rulPlanets(22) = "Asc. Sign Lord       " + planetName1Tr(kp_RashiLordInt(hCusNR(0), False)) + Space(10) + planetName1Tr(kp_RashiLordInt(hCusCur(0), False)) + vbNewLine
    rulPlanets(23) = "Asc. Star Lord       " + planetName1Tr(kp_StarLordInt(hCusNR(0), False)) + Space(10) + planetName1Tr(kp_StarLordInt(hCusCur(0), False)) + vbNewLine
    rulPlanets(24) = "Asc. Sub Lord        " + planetName2Tr(kp_SubLord(hCusNR(0), False)) + Space(10) + planetName2Tr(kp_SubLord(hCusCur(0), False)) + vbNewLine

    rulPlanets(25) = vbNewLine

    rulPlanets(26) = "DETAILS OF CURRENT POSISIONS OF CUSPS AND PLANETS" + vbNewLine
    rulPlanets(27) = vbNewLine

    rulPlanets(28) = "CS " + "POSITION  " + "SGN     " + "SGL " + "STL " + "SBL " + "SSL" + " | " _
                     + "PLN " + "POSITION  " + "SGN     " + "SGL " + "STL " + "SBL " + "SSL" + vbNewLine
    rulPlanets(29) = String(76, "-") + vbNewLine

    Dim i As Byte
    For i = 0 To 11
        rulPlanets(i) = Format(i + 1, "00") + " " + hp_FormalDegSh(hCusCur(i)) + " " + rashiNameTrSh(kp_RashiInt(hCusCur(i))) + " " + planetName1Tr(kp_RashiLordInt(hCusCur(i), False)) + " " + planetName1Tr(kp_StarLordInt(hCusCur(i), False)) + " " + planetName2Tr(kp_SubLord(hCusCur(i), False)) + " " + planetName2Tr(kp_SubLord(hCusCur(i), True)) + " | " _
                        + UCase(kp_PlanetName(i, isReCur(i))) + " " + hp_FormalDegSh(pPosCur(i)) + " " + rashiNameTrSh(kp_RashiInt(pPosCur(i))) + " " + planetName1Tr(kp_RashiLordInt(pPosCur(i), False)) + " " + planetName1Tr(kp_StarLordInt(pPosCur(i), False)) + " " + planetName2Tr(kp_SubLord(pPosCur(i), False)) + " " + planetName2Tr(kp_SubLord(pPosCur(i), True)) + vbNewLine
    Next

    frmMain.RT2.SelIndent = 300
    frmMain.RT2.Text = vbNullString

    frmMain.RT2.SelText = vbNewLine + vbNewLine

    Dim j As Byte, k As Byte
    For j = 12 To 29
        frmMain.RT2.SelText = rulPlanets(j)
    Next
    For k = 0 To 11
        frmMain.RT2.SelText = rulPlanets(k)
    Next

End Sub


Private Function kp_Horar(jDay As Double, firstCusNir As Double, curLat As Double, kpAyanHor As Double) As Double()
    
    'NUT
    Dim nutPReVal As Long
    Dim nutPDetails(5) As Double
    Dim nutErrReturn As String
    nutPReVal = swe_calc_ut(jDay, -1, 0, nutPDetails(0), nutErrReturn)
    
    Dim sayanaCusp As Double
    sayanaCusp = hp_Rnd0To360v(firstCusNir + kpAyanHor)
    sayanaCusp = sayanaCusp + 0.000000001   'To remove errors
    If sayanaCusp < 0.000000001 Then sayanaCusp = sayanaCusp + 0.000000001 'To remove errors
    
    '1st iteration
    Dim retVal As Long
    Dim ascmc(10) As Double
    Dim cusVal(12) As Double

    Dim allfCus(360) As Double
    'Mean nutPDetails(0) = 23.4525
    Dim a1 As Integer
    For a1 = 0 To 359
        retVal = swe_houses_armc(CDbl(a1), curLat, nutPDetails(0), Asc("P"), cusVal(0), ascmc(0))
        allfCus(a1) = cusVal(1)
    Next
    allfCus(360) = allfCus(0)
        
    Dim lBnd As Integer
    lBnd = 269  'Exception
    Dim a2 As Integer
    For a2 = 0 To 269
        If ((allfCus(a2) <= sayanaCusp) And (allfCus(a2 + 1) > sayanaCusp)) Then
            lBnd = a2
        End If
    Next
    Dim a3 As Integer
    For a3 = 270 To 359
        If ((allfCus(a3) <= sayanaCusp) And (allfCus(a3 + 1) > sayanaCusp)) Then
            lBnd = a3
        End If
    Next

    Dim lBnd2 As Double
    Dim finalARMC As Double
    Const RES = 1000    'Resolution 1
    Const RES2 = 1000   'Resolution 2

    Dim allsCus(RES) As Double
    Dim alltCus(RES2) As Double

    If lBnd <> 269 Then

        Dim retVal_ As Long
        Dim ascMC_(10) As Double
        Dim cusVal_(12) As Double
        Dim a4 As Integer
        For a4 = 0 To RES
            retVal_ = swe_houses_armc(CDbl(lBnd) + 1# * (CDbl(a4 / RES)), curLat, nutPDetails(0), Asc("P"), cusVal_(0), ascMC_(0))
            allsCus(a4) = cusVal_(1)
        Next

        Dim a5 As Integer
        For a5 = 0 To RES - 1
            If ((allsCus(a5) <= sayanaCusp) And (allsCus(a5 + 1) > sayanaCusp)) Then
                lBnd2 = CDbl(lBnd) + 1# * (CDbl(a5 / RES))
            End If
        Next

        Dim retVal__ As Long
        Dim ascMC__(10) As Double
        Dim cusVal__(12) As Double
        Dim a6 As Integer
        For a6 = 0 To RES2
            retVal__ = swe_houses_armc(lBnd2 + (1# / RES) * (CDbl(a6 / RES2)), curLat, nutPDetails(0), Asc("P"), cusVal__(0), ascMC__(0))
            alltCus(a6) = cusVal__(1)
        Next

        Dim a7 As Integer
        For a7 = 0 To RES2 - 1
            If ((alltCus(a7) <= sayanaCusp) And (alltCus(a7 + 1) > sayanaCusp)) Then
                finalARMC = lBnd2 + (1# / RES) * CDbl((a7 + 1) / RES2)
            End If
        Next

    End If

    If lBnd = 269 Then

        Dim retVal2_ As Long
        Dim ascMC2_(10) As Double
        Dim cusVal2_(12) As Double
        Dim a8 As Integer
        For a8 = 0 To RES
            retVal2_ = swe_houses_armc(CDbl(lBnd) + 1# * (CDbl(a8 / RES)), curLat, nutPDetails(0), Asc("P"), cusVal2_(0), ascMC2_(0))
            allsCus(a8) = cusVal2_(1)
        Next

        Dim a9 As Integer
        For a9 = 0 To RES - 1
            If ((allsCus(a9) <= sayanaCusp) And (allsCus(a9 + 1) > sayanaCusp)) Then
                lBnd2 = CDbl(lBnd) + 1# * (CDbl(a9 / RES))
            End If
        Next

        Dim retVal2__ As Long
        Dim ascMC2__(10) As Double
        Dim cusVal2__(12) As Double
        Dim b1 As Integer
        For b1 = 0 To RES2
            retVal2__ = swe_houses_armc(lBnd2 + (1# / RES) * (CDbl(b1 / RES2)), curLat, nutPDetails(0), Asc("P"), cusVal2__(0), ascMC2__(0))
            alltCus(b1) = cusVal2__(1)
        Next

        Dim b2 As Integer
        For b2 = 0 To RES2 - 1
            If ((alltCus(b2) <= sayanaCusp) And (alltCus(b2 + 1) > sayanaCusp)) Then
                finalARMC = lBnd2 + (1# / RES) * CDbl((b2 + 1) / RES2)
            End If
        Next

    End If

    Dim retValf As Long
    Dim ascMCf(10) As Double
    Dim cusValf(12) As Double

    retValf = swe_houses_armc(finalARMC, curLat, nutPDetails(0), Asc("P"), cusValf(0), ascMCf(0))
    
    Dim saCusps(12) As Double
    saCusps(0) = 0#
    Dim b3 As Integer
    For b3 = 1 To 12
        'saCusps(b3) = hp_Rnd0To360v(cusValf(b3) - nutPDetails(2))
        saCusps(b3) = cusValf(b3)
    Next
    
    kp_Horar = saCusps
    
End Function

Private Function kp_IncludedHouse(ByVal planetPos As Double) As Integer
'This function returns the house where the planet is posited

    Dim i As Byte
    Dim j As Byte

    For i = 1 To 11
        If (hCuspI(i) < hCuspI(i + 1)) And (hCuspI(i) <= planetPos And planetPos < hCuspI(i + 1)) Then
            kp_IncludedHouse = CInt(i)
        End If
    Next

    If (hCuspI(12) < hCuspI(1)) And (hCuspI(12) <= planetPos And planetPos < hCuspI(1)) Then
        kp_IncludedHouse = 12
    End If

    For j = 1 To 11
        If (hCuspI(j) > hCuspI(j + 1)) And ((hCuspI(j) <= planetPos And planetPos < 360#) Or (0# <= planetPos And planetPos < hCuspI(j + 1))) Then
            kp_IncludedHouse = j
        End If
    Next

    If (hCuspI(12) > hCuspI(1)) And ((hCuspI(12) <= planetPos And planetPos < 360#) Or (0# <= planetPos And planetPos < hCuspI(1))) Then
        kp_IncludedHouse = 12
    End If

End Function

Private Sub A04_calcWAspects()
'Find Western Aspects

'Private hCus(11) As Double
'Private pPos(11) As Double

'Load aspect settings from dat file
    Dim aspLoaded() As String
    aspLoaded() = hp_ReadDatFile(App.Path + "\Sett\aspdet.dat", "_")

    'Separate the aspect data from loaded string stream
    Dim dataToAdd0() As String, dataToAdd6() As String, dataToAdd12() As String
    Dim dataToAdd1() As String, dataToAdd7() As String, dataToAdd13() As String
    Dim dataToAdd2() As String, dataToAdd8() As String, dataToAdd14() As String

    Dim dataToAdd3() As String, dataToAdd9() As String, dataToAdd15() As String
    Dim dataToAdd4() As String, dataToAdd10() As String, dataToAdd16() As String
    Dim dataToAdd5() As String, dataToAdd11() As String, dataToAdd17() As String

    dataToAdd0() = Split(aspLoaded(0), " ")
    dataToAdd1() = Split(aspLoaded(1), " ")
    dataToAdd2() = Split(aspLoaded(2), " ")
    dataToAdd3() = Split(aspLoaded(3), " ")
    dataToAdd4() = Split(aspLoaded(4), " ")
    dataToAdd5() = Split(aspLoaded(5), " ")
    dataToAdd6() = Split(aspLoaded(6), " ")
    dataToAdd7() = Split(aspLoaded(7), " ")
    dataToAdd8() = Split(aspLoaded(8), " ")

    dataToAdd9() = Split(aspLoaded(9), " ")
    dataToAdd10() = Split(aspLoaded(10), " ")
    dataToAdd11() = Split(aspLoaded(11), " ")
    dataToAdd12() = Split(aspLoaded(12), " ")
    dataToAdd13() = Split(aspLoaded(13), " ")
    dataToAdd14() = Split(aspLoaded(14), " ")
    dataToAdd15() = Split(aspLoaded(15), " ")
    dataToAdd16() = Split(aspLoaded(16), " ")
    dataToAdd17() = Split(aspLoaded(17), " ")

    'Separate Selecterd string inte individual date
    Dim isSel(17) As Boolean
    Dim appVal(17) As Single
    Dim sepVal(17) As Single

    '0
    isSel(0) = CBool(dataToAdd0(0))
    appVal(0) = CSng(dataToAdd0(1))
    sepVal(0) = CSng(dataToAdd0(2))

    '1
    isSel(1) = CBool(dataToAdd1(0))
    appVal(1) = CSng(dataToAdd1(1))
    sepVal(1) = CSng(dataToAdd1(2))

    '2
    isSel(2) = CBool(dataToAdd2(0))
    appVal(2) = CSng(dataToAdd2(1))
    sepVal(2) = CSng(dataToAdd2(2))

    '3
    isSel(3) = CBool(dataToAdd3(0))
    appVal(3) = CSng(dataToAdd3(1))
    sepVal(3) = CSng(dataToAdd3(2))

    '4
    isSel(4) = CBool(dataToAdd4(0))
    appVal(4) = CSng(dataToAdd4(1))
    sepVal(4) = CSng(dataToAdd4(2))

    '5
    isSel(5) = CBool(dataToAdd5(0))
    appVal(5) = CSng(dataToAdd5(1))
    sepVal(5) = CSng(dataToAdd5(2))

    '6
    isSel(6) = CBool(dataToAdd6(0))
    appVal(6) = CSng(dataToAdd6(1))
    sepVal(6) = CSng(dataToAdd6(2))

    '7
    isSel(7) = CBool(dataToAdd7(0))
    appVal(7) = CSng(dataToAdd7(1))
    sepVal(7) = CSng(dataToAdd7(2))

    '8
    isSel(8) = CBool(dataToAdd8(0))
    appVal(8) = CSng(dataToAdd8(1))
    sepVal(8) = CSng(dataToAdd8(2))

    '9
    isSel(9) = CBool(dataToAdd9(0))
    appVal(9) = CSng(dataToAdd9(1))
    sepVal(9) = CSng(dataToAdd9(2))

    '10
    isSel(10) = CBool(dataToAdd10(0))
    appVal(10) = CSng(dataToAdd10(1))
    sepVal(10) = CSng(dataToAdd10(2))

    '11
    isSel(11) = CBool(dataToAdd11(0))
    appVal(11) = CSng(dataToAdd11(1))
    sepVal(11) = CSng(dataToAdd11(2))

    '12
    isSel(12) = CBool(dataToAdd12(0))
    appVal(12) = CSng(dataToAdd12(1))
    sepVal(12) = CSng(dataToAdd12(2))

    '13
    isSel(13) = CBool(dataToAdd13(0))
    appVal(13) = CSng(dataToAdd13(1))
    sepVal(13) = CSng(dataToAdd13(2))

    '14
    isSel(14) = CBool(dataToAdd14(0))
    appVal(14) = CSng(dataToAdd14(1))
    sepVal(14) = CSng(dataToAdd14(2))

    '15
    isSel(15) = CBool(dataToAdd15(0))
    appVal(15) = CSng(dataToAdd15(1))
    sepVal(15) = CSng(dataToAdd15(2))

    '16
    isSel(16) = CBool(dataToAdd16(0))
    appVal(16) = CSng(dataToAdd16(1))
    sepVal(16) = CSng(dataToAdd16(2))

    '17
    isSel(17) = CBool(dataToAdd17(0))
    appVal(17) = CSng(dataToAdd17(1))
    sepVal(17) = CSng(dataToAdd17(2))

    Dim pToHouse_(11, 11) As Single      'Aspect planet to House
    Dim pToPlanet_(11, 11) As Single     'Aspect planet to Planet

    'Find Aspects
    Dim i As Byte, j As Byte, k As Byte, m As Byte
    Dim n As Byte, P As Byte, q As Byte, R As Byte

    'Palnet To House
    For i = 0 To 11         'planet
        For j = 0 To 11     'House
            For k = 0 To 17    'Aspect
                pToHouse_(i, j) = kp_WAspect(pPos(i), hCus(j), aspectPont(k), appVal(k), sepVal(k), isSel(k))
                If pToHouse_(i, j) < 360# Then Exit For
            Next
        Next
    Next

    'Planet To Planet
    For R = 0 To 11
        For q = 0 To 11
            pToPlanet_(R, q) = 400#
        Next
    Next

    For m = 0 To 11         'Planet 1
        For n = 0 To 11     'Planet 2
            For P = 0 To 17    'Aspect
                If m > n Then
                    pToPlanet_(m, n) = kp_WAspect(pPos(m), pPos(n), aspectPont(P), appVal(P), sepVal(P), isSel(P))
                    If pToPlanet_(m, n) < 360# Then Exit For
                End If
            Next
        Next
    Next

    '#EXPORT VAR#
    Dim t1 As Byte, t2 As Byte, t3 As Byte, t4 As Byte

    For t1 = 0 To 11    '#EXPORT#
        For t2 = 0 To 11
            pToHouse(t1, t2) = pToHouse_(t1, t2)
        Next
    Next

    For t3 = 0 To 11    '#EXPORT#
        For t4 = 0 To 11
            pToPlanet(t3, t4) = pToPlanet_(t3, t4)
        Next
    Next

End Sub


Private Sub A05_calcDasa()
'Find Dasa Upto 4 levels

    Dim mPos As Double
    Dim bJulDate As Double

    mPos = pPos(1)
    bJulDate = julDateLocal

    Dim daysPerYear As Double
    daysPerYear = 365.25 '365.2425

    'Find dasa duration for each dasa upto 4 level
    Dim bDasaYears(8, 8) As Double
    Dim k1 As Byte, k2 As Byte
    For k1 = 0 To 8
        For k2 = 0 To 8
            bDasaYears(k1, hp_Rnd0To8v(k1 + k2)) = dasaYears(k1) * (dasaYears(hp_Rnd0To8v(k1 + k2)) / 120#)
        Next
    Next

    Dim aDasaYears(8, 8, 8) As Double
    Dim k3 As Byte, k4 As Byte, k5 As Byte
    For k3 = 0 To 8
        For k4 = 0 To 8
            For k5 = 0 To 8
                aDasaYears(k3, k4, hp_Rnd0To8v(k4 + k5)) = bDasaYears(k3, k4) * (dasaYears(hp_Rnd0To8v(k4 + k5)) / 120#)
            Next
        Next
    Next

    Dim sDasaYears(8, 8, 8, 8) As Double
    Dim k6 As Byte, k7 As Byte, k8 As Byte, k9 As Byte
    For k6 = 0 To 8
        For k7 = 0 To 8
            For k8 = 0 To 8
                For k9 = 0 To 8
                    sDasaYears(k6, k7, k8, hp_Rnd0To8v(k8 + k9)) = aDasaYears(k6, k7, k8) * (dasaYears(hp_Rnd0To8v(k8 + k9)) / 120#)
                Next
            Next
        Next
    Next

    'Fing initial position
    Dim midVal(27) As Double
    Dim i1 As Byte
    For i1 = 0 To 27
        midVal(i1) = 360# * (CDbl(i1) / 27#)
    Next

    Dim uBor As Double
    Dim lBor As Double
    Dim i2 As Byte

    For i2 = 0 To 26
        If (midVal(i2) < mPos And mPos < midVal(i2 + 1)) Then
            lBor = midVal(i2)
            uBor = midVal(i2 + 1)
            Exit For
        End If
    Next

    'Dasa Lords Int s.
    Dim nakLordInt As Integer
    nakLordInt = kp_StarLordInt(mPos, True)

    'MDasa
    Dim i3 As Byte
    Dim mDasaLordInt(8) As Integer

    For i3 = 0 To 8
        mDasaLordInt(i3) = hp_Rnd0To8v(nakLordInt + i3)
    Next

    'BDasa
    Dim i4 As Byte, i5 As Byte
    Dim bDasaLordInt(8, 8) As Integer

    For i4 = 0 To 8
        For i5 = 0 To 8
            bDasaLordInt(i4, i5) = hp_Rnd0To8v(mDasaLordInt(i4) + i5)
        Next
    Next

    'ADasa
    Dim i6 As Byte, i7 As Byte, i8 As Byte
    Dim aDasaLordInt(8, 8, 8) As Integer

    For i6 = 0 To 8
        For i7 = 0 To 8
            For i8 = 0 To 8
                aDasaLordInt(i6, i7, i8) = hp_Rnd0To8v(bDasaLordInt(i6, i7) + i8)
            Next
        Next
    Next

    'SDasa
    Dim i9 As Byte, i10 As Byte, i11 As Byte, i12 As Byte
    Dim sDasaLordInt(8, 8, 8, 8) As Integer

    For i9 = 0 To 8
        For i10 = 0 To 8
            For i11 = 0 To 8
                For i12 = 0 To 8
                    sDasaLordInt(i9, i10, i11, i12) = hp_Rnd0To8v(aDasaLordInt(i9, i10, i11) + i12)
                Next
            Next
        Next
    Next

    '#EXPORT VAR#
    Dim p0 As Byte, p1 As Byte, p2 As Byte, p3 As Byte, p4 As Byte, p5 As Byte, p6 As Byte, p7 As Byte, p8 As Byte, p9 As Byte

    For p0 = 0 To 8    '#EXPORT#
        mDasL(p0) = mDasaLordInt(p0)
    Next

    For p1 = 0 To 8    '#EXPORT#
        For p2 = 0 To 8
            bDasL(p1, p2) = bDasaLordInt(p1, p2)
        Next
    Next

    For p3 = 0 To 8    '#EXPORT#
        For p4 = 0 To 8
            For p5 = 0 To 8
                aDasL(p3, p4, p5) = aDasaLordInt(p3, p4, p5)
            Next
        Next
    Next

    For p6 = 0 To 8    '#EXPORT#
        For p7 = 0 To 8
            For p8 = 0 To 8
                For p9 = 0 To 8
                    sDasL(p6, p7, p8, p9) = sDasaLordInt(p6, p7, p8, p9)
                Next
            Next
        Next
    Next

    'Find Dasa borders starting points
    Dim spndDasa As Double
    spndDasa = ((mPos - lBor) / 13.3333333333333) * dasaYears(nakLordInt) * daysPerYear

    Dim mDasBor As Double
    mDasBor = bJulDate - spndDasa

    'MDasa
    Dim j1 As Byte
    Dim mDasaBor(8) As Double
    mDasaBor(0) = mDasBor
    For j1 = 1 To 8
        mDasaBor(j1) = mDasaBor(j1 - 1) + (dasaYears(mDasaLordInt(j1 - 1)) * daysPerYear)
    Next

    'BDasa
    Dim j2 As Byte, j3 As Byte
    Dim bDasaBor(8, 8) As Double
    For j2 = 0 To 8
        bDasaBor(j2, 0) = mDasaBor(j2)
        For j3 = 1 To 8
            bDasaBor(j2, j3) = bDasaBor(j2, j3 - 1) + bDasaYears(mDasaLordInt(j2), bDasaLordInt(j2, j3 - 1)) * daysPerYear
        Next
    Next

    'ADasa
    Dim j4 As Byte, j5 As Byte, j6 As Byte
    Dim aDasaBor(8, 8, 8) As Double
    For j4 = 0 To 8
        For j5 = 0 To 8
            aDasaBor(j4, j5, 0) = bDasaBor(j4, j5)
            For j6 = 1 To 8
                aDasaBor(j4, j5, j6) = aDasaBor(j4, j5, j6 - 1) + (aDasaYears(mDasaLordInt(j4), bDasaLordInt(j4, j5), aDasaLordInt(j4, j5, j6 - 1)) * daysPerYear)
            Next
        Next
    Next

    'SDasa
    Dim j7 As Byte, j8 As Byte, j9 As Byte, j0 As Byte
    Dim sDasaBor(8, 8, 8, 8) As Double
    For j7 = 0 To 8
        For j8 = 0 To 8
            For j9 = 0 To 8
                sDasaBor(j7, j8, j9, 0) = aDasaBor(j7, j8, j9)
                For j0 = 1 To 8
                    sDasaBor(j7, j8, j9, j0) = sDasaBor(j7, j8, j9, j0 - 1) + (sDasaYears(mDasaLordInt(j7), bDasaLordInt(j7, j8), aDasaLordInt(j7, j8, j9), sDasaLordInt(j7, j8, j9, j0 - 1)) * daysPerYear)
                Next
            Next
        Next
    Next

    '#EXPORT VAR#
    Dim m0 As Byte, m1 As Byte, m2 As Byte, m3 As Byte, m4 As Byte, m5 As Byte, m6 As Byte, m7 As Byte, m8 As Byte, m9 As Byte

    For m0 = 0 To 8    '#EXPORT#
        mDas(m0) = mDasaBor(m0)
    Next

    For m1 = 0 To 8    '#EXPORT#
        For m2 = 0 To 8
            bDas(m1, m2) = bDasaBor(m1, m2)
        Next
    Next

    For m3 = 0 To 8    '#EXPORT#
        For m4 = 0 To 8
            For m5 = 0 To 8
                aDas(m3, m4, m5) = aDasaBor(m3, m4, m5)
            Next
        Next
    Next

    For m6 = 0 To 8    '#EXPORT#
        For m7 = 0 To 8
            For m8 = 0 To 8
                For m9 = 0 To 8
                    sDas(m6, m7, m8, m9) = sDasaBor(m6, m7, m8, m9)
                Next
            Next
        Next
    Next

End Sub


Private Sub A99_doAll()
'Print all things in main RT box

    Call assignGlobal

    Call A00_calcCurrent
    Call A01_calcKPNatal
    Call A02_drawChart
    Call A04_calcWAspects
    Call A03_calcVAspects
    Call A05_calcDasa
    Call A06_calcSig
    Call A07_calc4StepSig

    Call C02_RulPlanets
    'Call C01_PrintDasaEx

    'Print RTF
    Call A97_printRTF

    TabStrip1.Tabs(2).Selected = True

    '    'Clear the dasa extended
    '    RT3.Text = vbNullString
    '    frmMain.TV.Nodes.Clear

    mnuFileSaveAsPDF.Enabled = True
    mnuFileSaveAsTXT.Enabled = True
    mnuFileSaveAsRTF.Enabled = True
    
    isHorary = False   'Reset to Natal mode
    
    'Call swe_close  'Close SWE
    
End Sub

'Private Sub B02_ownBhava()
'
'End Sub
Private Sub A02_drawChart()
'Draw the KP Chart

'Sort all cusps
    Dim allPos(23) As Double, allPos2(23) As Double    'planet and house cusp all together
    Dim i As Byte, j As Byte

    '0 to 11 Cusps
    '12 to 23   Planets

    For i = 0 To 11
        allPos(i) = hCus(i)
        allPos2(i) = hCus(i)
    Next i

    For j = 0 To 11
        allPos(j + 12) = pPos(j)
        allPos2(j + 12) = pPos(j)
    Next j

    Call B01_posSort(allPos, 0, 23)    ' Sorting All Values

    Dim idAr(23) As Integer    'Corresponding ID as integer for sorted doubles
    Dim j1 As Byte, j2 As Byte

    'Determine the cusp and house/planet ID
    For j1 = 0 To 23
        For j2 = 0 To 23
            If Abs(allPos(j1) - allPos2(j2)) < 1E-30 Then
                idAr(j1) = j2
            End If
        Next
    Next

    Dim prPos(23) As String    ' Details print inside the chart boxex
    Dim j3 As Byte
    For j3 = 0 To 23
        prPos(j3) = kp_ChartPrint(allPos(j3), idAr(j3))
    Next
    '    For j33 = 12 To 23
    '        prPos(j33) = kp_ChartPrint(allPos(j33), idAr(j33))
    '    Next


    'Initialize the Chart Space
    Dim PX(11, 7) As String    '(12 box, 8 raws per box)    #### Just change number of raws
    Dim lc1 As Byte
    Dim lc2 As Byte
    For lc1 = 0 To 11
        For lc2 = 0 To 7
            PX(lc1, lc2) = Space(15)
        Next lc2
    Next lc1

    'Find the box of cusp/planet
    '0-30>>box1, 30-60>>box2...etc

    Dim j4 As Byte
    Dim mVal(12) As Double
    For j4 = 0 To 12
        mVal(j4) = 360# * (CDbl(j4) / 12#)
    Next

    Dim selArr(11) As String    'Select the strings to be printed as an array (String separated by ",")
    Dim n1 As Byte, n2 As Byte
    For n1 = 0 To 11
        For n2 = 0 To 23
            If (mVal(n1) <= allPos(n2) And allPos(n2) < mVal(n1 + 1)) Then
                selArr(n1) = selArr(n1) + prPos(n2) + ","
            End If
        Next
    Next

    Dim arrS0() As String, arrS1() As String, arrS2() As String, arrS3() As String
    Dim arrS4() As String, arrS5() As String, arrS6() As String, arrS7() As String
    Dim arrS8() As String, arrS9() As String, arrS10() As String, arrS11() As String

    'Separate string for each box
    arrS0 = Split(selArr(0), ",")
    arrS1 = Split(selArr(1), ",")
    arrS2 = Split(selArr(2), ",")
    arrS3 = Split(selArr(3), ",")

    arrS4 = Split(selArr(4), ",")
    arrS5 = Split(selArr(5), ",")
    arrS6 = Split(selArr(6), ",")
    arrS7 = Split(selArr(7), ",")

    arrS8 = Split(selArr(8), ",")
    arrS9 = Split(selArr(9), ",")
    arrS10 = Split(selArr(10), ",")
    arrS11 = Split(selArr(11), ",")

    'Put strings in correct place

    'Box 0
    Dim n3 As Byte
    If UBound(arrS0) > 0 Then
        For n3 = LBound(arrS0) To UBound(arrS0)
            PX(0, n3) = arrS0(n3)
        Next
    End If

    'Box 1
    Dim n4 As Byte
    If UBound(arrS1) > 0 Then
        For n4 = LBound(arrS1) To UBound(arrS1)
            PX(1, n4) = arrS1(n4)
        Next
    End If

    'Box 2
    Dim n5 As Byte
    If UBound(arrS2) > 0 Then
        For n5 = LBound(arrS2) To UBound(arrS2)
            PX(2, n5) = arrS2(n5)
        Next
    End If

    'Box 3
    Dim n6 As Byte
    If UBound(arrS3) > 0 Then
        For n6 = LBound(arrS3) To UBound(arrS3)
            PX(3, n6) = arrS3(n6)
        Next
    End If

    'Box 4
    Dim n7 As Byte
    If UBound(arrS4) > 0 Then
        For n7 = LBound(arrS4) To UBound(arrS4)
            PX(4, n7) = arrS4(n7)
        Next
    End If

    'Box 5
    Dim n8 As Byte
    If UBound(arrS5) > 0 Then
        For n8 = LBound(arrS5) To UBound(arrS5)
            PX(5, n8) = arrS5(n8)
        Next
    End If

    'Box 6
    Dim n9 As Byte
    If UBound(arrS6) > 0 Then
        For n9 = LBound(arrS6) To UBound(arrS6)
            PX(6, n9) = arrS6(n9)
        Next
    End If

    'Box 7
    Dim k1 As Byte
    If UBound(arrS7) > 0 Then
        For k1 = LBound(arrS7) To UBound(arrS7)
            PX(7, k1) = arrS7(k1)
        Next
    End If

    'Box 8
    Dim k2 As Byte
    If UBound(arrS8) > 0 Then
        For k2 = LBound(arrS8) To UBound(arrS8)
            PX(8, k2) = arrS8(k2)
        Next
    End If

    'Box 9
    Dim k3 As Byte
    If UBound(arrS9) > 0 Then
        For k3 = LBound(arrS9) To UBound(arrS9)
            PX(9, k3) = arrS9(k3)
        Next
    End If

    'Box 10
    Dim k4 As Byte
    If UBound(arrS10) > 0 Then
        For k4 = LBound(arrS10) To UBound(arrS10)
            PX(10, k4) = arrS10(k4)
        Next
    End If

    'Box 11
    Dim k5 As Byte
    If UBound(arrS11) > 0 Then
        For k5 = LBound(arrS11) To UBound(arrS11)
            PX(11, k5) = arrS11(k5)
        Next
    End If

    Dim k6 As Byte, k7 As Byte
    For k6 = 0 To 11
        For k7 = 0 To 7
            If Len(PX(k6, k7)) < 14 Then PX(k6, k7) = Space(15)
        Next
    Next

    'Print middle of the chart
    'Dim bAscM As String
    'Dim bAyanM As String

    bAscM = "ASC   : " + rashiNameTr(kp_RashiInt(hCus(0)))
    '     bDateM = "DATE  : " + Format(cmbDay.Text, "00") + "/" + Format(cmbMonth.Text, "00") + "/" + Format(cmbYear.Text, "0000")
    '     bTimeM = "TIME  : " + Format(cmbHH.Text, "00") + ":" + Format(cmbMM.Text, "00") + ":" + Format(cmbSS.Text, "00") + " H"
    '     bCorrM = "CORR  : " + cmbTcPM.Text + Format(cmbTcHH.Text, "00") + ":" + Format(cmbTcMM.Text, "00")
    '    bPlaceM = "PLACE : " + "Lon:" + Format(cmbLonDeg.Text, "000") + cmbLonEW.Text + Format(cmbLonMin.Text, "00") + ", Lat:" + Format(cmbLatDeg.Text, "00") + cmbLatNS.Text + Format(cmbLatMin.Text, "00")
    '       bTzM = "T.ZONE: " + cmbTzPM.Text + Format(cmbTzHH.Text, "00") + ":" + Format(cmbTzMM.Text, "00") + " GMT"
    bAyanM = "AYAN  : " + hp_FormalDeg(kpAyan)

    'To Draw the chart
    Dim charDraw_(36) As String

    'RTF
    charDraw_(0) = "+" + String(15, "-") + "+" + String(15, "-") + "+" + String(15, "-") + "+" + String(15, "-") + "+" + vbNewLine

    charDraw_(1) = "|" + PX(11, 7) + "|" + PX(0, 0) + "|" + PX(1, 0) + "|" + PX(2, 0) + "|" + vbNewLine
    charDraw_(2) = "|" + PX(11, 6) + "|" + PX(0, 1) + "|" + PX(1, 1) + "|" + PX(2, 1) + "|" + vbNewLine
    charDraw_(3) = "|" + PX(11, 5) + "|" + PX(0, 2) + "|" + PX(1, 2) + "|" + PX(2, 2) + "|" + vbNewLine
    charDraw_(4) = "|" + PX(11, 4) + "|" + PX(0, 3) + "|" + PX(1, 3) + "|" + PX(2, 3) + "|" + vbNewLine
    charDraw_(5) = "|" + PX(11, 3) + "|" + PX(0, 4) + "|" + PX(1, 4) + "|" + PX(2, 4) + "|" + vbNewLine
    charDraw_(6) = "|" + PX(11, 2) + "|" + PX(0, 5) + "|" + PX(1, 5) + "|" + PX(2, 5) + "|" + vbNewLine
    charDraw_(7) = "|" + PX(11, 1) + "|" + PX(0, 6) + "|" + PX(1, 6) + "|" + PX(2, 6) + "|" + vbNewLine
    charDraw_(8) = "|" + PX(11, 0) + "|" + PX(0, 7) + "|" + PX(1, 7) + "|" + PX(3, 7) + "|" + vbNewLine

    charDraw_(9) = "+" + String(15, "-") + "+" + String(15, "-") + "+" + String(15, "-") + "+" + String(15, "-") + "+" + vbNewLine

    charDraw_(10) = "|" + PX(10, 7) + "|" + Space(31) + "|" + PX(3, 0) + "|" + vbNewLine
    charDraw_(11) = "|" + PX(10, 6) + "|" + Space(31) + "|" + PX(3, 1) + "|" + vbNewLine
    charDraw_(12) = "|" + PX(10, 5) + "|" + Space(31) + "|" + PX(3, 2) + "|" + vbNewLine
    charDraw_(13) = "|" + PX(10, 4) + "|" + Space(1) + bAscM + Space(2) + "|" + PX(3, 3) + "|" + vbNewLine
    charDraw_(14) = "|" + PX(10, 3) + "|" + Space(1) + bDateM + Space(12) + "|" + PX(3, 4) + "|" + vbNewLine
    charDraw_(15) = "|" + PX(10, 2) + "|" + Space(1) + bTimeM + Space(12) + "|" + PX(3, 5) + "|" + vbNewLine
    charDraw_(16) = "|" + PX(10, 1) + "|" + Space(1) + bCorrM + Space(16) + "|" + PX(3, 6) + "|" + vbNewLine
    charDraw_(17) = "|" + PX(10, 0) + "|" + Space(1) + bPlaceM + Space(1) + "|" + PX(3, 7) + "|" + vbNewLine

    charDraw_(18) = "+" + String(15, "-") + "+" + Space(1) + bTzM + Space(12) + "+" + String(15, "-") + "+" + vbNewLine

    charDraw_(19) = "|" + PX(9, 7) + "|" + Space(1) + bAyanM + Space(7) + "|" + PX(4, 0) + "|" + vbNewLine
    charDraw_(20) = "|" + PX(9, 6) + "|" + Space(31) + "|" + PX(4, 1) + "|" + vbNewLine
    charDraw_(21) = "|" + PX(9, 5) + "|" + Space(31) + "|" + PX(4, 2) + "|" + vbNewLine
    charDraw_(22) = "|" + PX(9, 4) + "|" + Space(10) + bNmeM + Space(10) + "|" + PX(4, 3) + "|" + vbNewLine
    charDraw_(23) = "|" + PX(9, 3) + "|" + Space(31) + "|" + PX(4, 4) + "|" + vbNewLine
    charDraw_(24) = "|" + PX(9, 2) + "|" + Space(31) + "|" + PX(4, 5) + "|" + vbNewLine
    charDraw_(25) = "|" + PX(9, 1) + "|" + Space(31) + "|" + PX(4, 6) + "|" + vbNewLine
    charDraw_(26) = "|" + PX(9, 0) + "|" + Space(31) + "|" + PX(4, 7) + "|" + vbNewLine

    charDraw_(27) = "+" + String(15, "-") + "+" + String(15, "-") + "+" + String(15, "-") + "+" + String(15, "-") + "+" + vbNewLine

    charDraw_(28) = "|" + PX(8, 7) + "|" + PX(7, 7) + "|" + PX(6, 7) + "|" + PX(5, 0) + "|" + vbNewLine
    charDraw_(29) = "|" + PX(8, 6) + "|" + PX(7, 6) + "|" + PX(6, 6) + "|" + PX(5, 1) + "|" + vbNewLine
    charDraw_(30) = "|" + PX(8, 5) + "|" + PX(7, 5) + "|" + PX(6, 5) + "|" + PX(5, 2) + "|" + vbNewLine
    charDraw_(31) = "|" + PX(8, 4) + "|" + PX(7, 4) + "|" + PX(6, 4) + "|" + PX(5, 3) + "|" + vbNewLine
    charDraw_(32) = "|" + PX(8, 3) + "|" + PX(7, 3) + "|" + PX(6, 3) + "|" + PX(5, 4) + "|" + vbNewLine
    charDraw_(33) = "|" + PX(8, 2) + "|" + PX(7, 2) + "|" + PX(6, 2) + "|" + PX(5, 5) + "|" + vbNewLine
    charDraw_(34) = "|" + PX(8, 1) + "|" + PX(7, 1) + "|" + PX(6, 1) + "|" + PX(5, 6) + "|" + vbNewLine
    charDraw_(35) = "|" + PX(8, 0) + "|" + PX(7, 0) + "|" + PX(6, 0) + "|" + PX(5, 7) + "|" + vbNewLine

    charDraw_(36) = "+" + String(15, "-") + "+" + String(15, "-") + "+" + String(15, "-") + "+" + String(15, "-") + "+" + vbNewLine

    Dim mm As Byte  '#EXPORT VAR#
    For mm = 0 To 36    '#EXPORT#
        charDraw(mm) = charDraw_(mm)
    Next


    'PDF
    'Dim pdfPX(11, 7) As String
    Dim r1 As Integer, r2 As Integer

    For r1 = 0 To 11
        For r2 = 0 To 7
            pdfPX(r1, r2) = PX(r1, r2)
        Next
    Next

End Sub



Private Function kp_ChartPrint(posVal As Double, idInt As Integer) As String
'Set a String to print inside the KP Chart
    If idInt <= 11 Then
        If idInt = 0 Then
            kp_ChartPrint = "#ASC#" + " " + hp_FormalDegSh(posVal)
        Else
            kp_ChartPrint = " " + arToRom(idInt) + hp_FormalDegSh(posVal)
        End If
    Else
        idInt = idInt - 12
        kp_ChartPrint = " " + kp_PlanetName(idInt, isRe(idInt)) + "  " + hp_FormalDegSh(posVal)
    End If
End Function
Private Sub B01_posSort(posList() As Double, min As Integer, max As Integer)
' Sort the planet and house cusp

    Dim i As Integer
    Dim j As Integer
    Dim selVal As Double
    Dim selInt As Integer

    For i = min To max - 1
        selVal = posList(i)
        selInt = i
        For j = i + 1 To max
            If posList(j) < selVal Then
                selVal = posList(j)
                selInt = j
            End If
        Next j
        posList(selInt) = posList(i)
        posList(i) = selVal
    Next i

End Sub
'Private Sub B02_pVariableClear()
'    Dim i As Byte
'    For i = 0 To 6
'        aspVedic(i) = vbNullString
'        vedAspP(i) = vbNullString
'        smeRas(i) = vbNullString
'    Next
'End Sub

'Private Sub ZZ_DasaCalc()
'
'    Dim mPos As Double
'    Dim bJDate As Double
'
'    mPos = pPos(1)
'    bJDate = julDateLocal
'
'    'Fing initial position
'    Dim midVal(27) As Double
'    Dim i1 As Byte
'    For i1 = 0 To 27
'        midVal(i1) = 360# * (CDbl(i1) / 27#)
'    Next
'
'    Dim uBor As Double
'    Dim lBor As Double
'    Dim i2 As Byte
'
'    For i2 = 0 To 26
'        If (midVal(i2) < mPos And mPos < midVal(i2 + 1)) Then
'            lBor = midVal(i2)
'            uBor = midVal(i2 + 1)
'        End If
'    Next
'
'    Dim firstNakLordInt As Integer
'    firstNakLordInt = kp_StarLordInt(mPos, True)
'
'    Dim spndDasa As Double
'    spndDasa = ((mPos - lBor) / 13.3333333333333) * dasaYears(firstNakLordInt) * 365.2425
'
'    'M.Dasa
'    Dim mDasInt(8) As Integer
'    Dim mDasStart(8) As Double
'    mDasStart(0) = bJDate - spndDasa
'
'    Dim a1 As Byte, a2 As Byte
'
'    For a1 = 0 To 8
'        mDasInt(a1) = hp_Rnd0To8v(firstNakLordInt + a1)
'    Next
'
'    For a2 = 1 To 8
'        mDasStart(a2) = mDasStart(a2 - 1) + dasaYears(mDasInt(a2 - 1)) * 365.2425
'    Next
'
'    'A.Dasa
'    Dim a3 As Byte, a4 As Byte, a5 As Byte, a6 As Byte, a7 As Byte, a8 As Byte
'
'    Dim bDasInt(8, 8) As Integer
'    Dim bDasStart(8, 8) As Double
'
'    For a3 = 0 To 6
'        For a4 = 0 To 8
'            bDasInt(a3, a4) = hp_Rnd0To8v(mDasInt(a3) + a4)
'        Next
'    Next
'
'    Dim bDasaYears(8, 8) As Double
'    For a7 = 0 To 8
'        For a8 = 0 To 8
'            bDasaYears(a7, a8) = dasaYears(a7) * (dasaYears(hp_Rnd0To8v(a7 + a8)) / 120#)
'        Next
'    Next
'
'    For a5 = 0 To 8
'        bDasStart(a5, 0) = mDasStart(a5)
'        For a6 = 1 To 8
'            bDasStart(a5, a6) = bDasStart(a5, a6 - 1) + bDasaYears(mDasInt(a5), (a6 - 1)) * 365.2425
'        Next
'    Next
'
'End Sub

Private Sub cmb249_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub


Private Sub cmbDay_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub



Private Sub cmbHH_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub


Private Sub cmbLatDeg_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub


Private Sub cmbLatMin_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub


Private Sub cmbLatNS_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "NSns")
End Sub


Private Sub cmbLonDeg_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub


Private Sub cmbLonEW_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "EWew")
End Sub


Private Sub cmbLonMin_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub


Private Sub cmbMM_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub


Private Sub cmbMonth_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub


Private Sub cmbRot_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub


Private Sub cmbRotHor_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub










Private Sub cmbSS_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub


Private Sub cmbTcHH_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub


Private Sub cmbTcMM_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub


Private Sub cmbTcPM_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "+-")

End Sub


Private Sub cmbTzHH_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub


Private Sub cmbTzMM_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub


Private Sub cmbTzPM_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "+-")
End Sub


Private Sub cmbYear_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub


Private Sub cmdAtlas_Click()
    MsgBox "Sorry, currently not implemented. Please enter the required data manually.", vbOKOnly, "Sorry..."
End Sub



Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdGenerate_Click()

'Set Natal Mode
'isHorary = False

'Print All Data
    Call A99_doAll

End Sub

Private Sub cmdOK_Click()
'Set Horary Mode
    isHorary = True
    'Send Rotation to frmMain
    cmbRot.Text = cmbRotHor.Text
    sNo249Hor = CInt(cmb249.Text)    'Send the number via Global

    Call cmdGenerate_Click    'Call to frmMain's cmdGenerate button

End Sub

Private Sub Form_Load()
    
    'Posision
    Me.Height = 7970
    Me.Width = 10950

    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    'Caption
    Me.Caption = "KP New Astro Version : " + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision) + " beta"

    
    'PictureBox Positions
    picNatalHorary.Height = 6135
    picNatalHorary.Width = 10815
    picNatalHorary.Top = 480
    picNatalHorary.Left = 0

    picKPChart.Height = 6135
    picKPChart.Width = 10815
    picKPChart.Top = 480
    picKPChart.Left = 0

    '    picExtended.Height = 6135
    '    picExtended.Width = 10815
    '    picExtended.Top = 480
    '    picExtended.Left = 0

    picRulingPlanets.Height = 6135
    picRulingPlanets.Width = 10815
    picRulingPlanets.Top = 480
    picRulingPlanets.Left = 0

    picAbout.Height = 6135
    picAbout.Width = 10815
    picAbout.Top = 480
    picAbout.Left = 0

    Dim i As Integer
    For i = 1850 To CInt(Year(Now))
        cmbYear.AddItem CStr(i)
    Next
    
    '/------------------------------------------------------------------------------------/
    'Planet flag
    plntFlg = 256
    '/------------------------------------------------------------------------------------/
    
    'Load data from dat file
    'KP Horary
    Dim depLoaded_h() As String
    depLoaded_h() = hp_ReadDatFile(App.Path + "\Sett\depdet.dat", "_")

    Dim tzPM_h As String, tzHH_h As Integer, tzMM_h As Integer
    Dim lonDeg_h As Integer, lonMin_h As Integer, lonEW_h As String
    Dim latDeg_h As Integer, latMin_h As Integer, latNS_h As String

    'lblBirthPlace.Caption = "Default place : " + depLoaded_h(0)
    txtDefPlaceCur.Text = "Default place - " + depLoaded_h(0)
    
    tzPM_h = Trim(depLoaded_h(1))
    tzHH_h = CInt(depLoaded_h(2))
    tzMM_h = CInt(depLoaded_h(3))
    
    lonDeg_h = CInt(depLoaded_h(4))
    lonMin_h = CInt(depLoaded_h(5))
    lonEW_h = Trim(depLoaded_h(6))
    
    latDeg_h = CInt(depLoaded_h(7))
    latMin_h = CInt(depLoaded_h(8))
    latNS_h = Trim(depLoaded_h(9))

'    lblDetails1.Caption = "Lon : " + Format(lonDeg_h, "000") + lonEW_h + Format(lonMin_h, "00") + "   Lat : " + Format(latDeg_h, "00") + latNS_h + Format(latMin_h, "00")
'    lblDetails2.Caption = "Time Zone : " + tzPM_h + Format(tzHH_h, "00") + ":" + Format(tzMM_h, "00")

    If (CDbl(latDeg_h) + CDbl(latMin_h) / 60#) > 66# Then
        lblW.Visible = True
        cmdOK.Enabled = False
    End If

        
    'GeoTrue
    Dim geoTrue() As String
    geoTrue() = hp_ReadDatFile(App.Path + "\Sett\geotrue.dat", "_")
    isGeo = CBool(geoTrue(0))
    isTru = CBool(geoTrue(1))
    
    
    'Astrologer
    loadedPD() = hp_ReadDatFile(App.Path + "\Sett\asrdet.dat", "_")
    
    
    
    '/------------------------------------------------------------------------------------/
     'Load current (Horary)
    txtDayCur.Text = Format(Day(Now), "00")
    txtMonthCur.Text = Format(Month(Now), "00")
    txtYearCur.Text = Format(Year(Now), "0000")
    
    txtHHCur.Text = Format(Hour(Now), "00")
    txtMMCur.Text = Format(Minute(Now), "00")
    txtSSCur.Text = Format(Second(Now), "00")
    
    txtTZPMCur.Text = tzPM_h
    txtTZHHCur.Text = Format(tzHH_h, "00")
    txtTZMMCur.Text = Format(tzMM_h, "00")
    
    txtLonDDCur.Text = Format(lonDeg_h, "000")
    txtLonMMCur.Text = Format(lonMin_h, "00")
    txtLonEWCur.Text = lonEW_h
    
    txtLatDDCur.Text = Format(latDeg_h, "00")
    txtLatMMCur.Text = Format(latMin_h, "00")
    txtLatNSCur.Text = latNS_h
    '/------------------------------------------------------------------------------------/
    'Add a random number to combo box
    cmb249.Clear
    Dim ib As Integer
    For ib = 1 To 249
        cmb249.AddItem CStr(ib)
    Next

    Dim rndInt As Integer
    Call Randomize
    rndInt = Int(1 + Rnd(1) * 248)

    cmb249.Text = CStr(rndInt)

    '/------------------------------------------------------------------------------------/
    
    isHorary = False
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub


Private Sub lblEmail_Click()

    On Error GoTo errHnd
    Dim ret As Long
    ret = ShellExecute(Me.hwnd, "Open", "mailto:kpnewastro@gmail.com", vbNullString, vbNullString, 1)
    lblEmail.ForeColor = &H800080
    Exit Sub
errHnd:

    MsgBox "Please Use Your Default Application For Send The Email.", vbCritical, "Error"


End Sub

Private Sub lblEmail2_Click()

    On Error GoTo errHnd
    Dim ret As Long
    ret = ShellExecute(Me.hwnd, "Open", "mailto:kpnewastro@gmail.com", vbNullString, vbNullString, 1)
    lblEmail2.ForeColor = &H800080
    Exit Sub
errHnd:

    MsgBox "Please Use Your Default Application For Send The Email.", vbCritical, "Error"

End Sub


Private Sub lblWeb_Click()
    On Error GoTo errHnd
    Dim ret As Long
    ret = ShellExecute(Me.hwnd, "Open", "http://www.kpnewastro.blogspot.com", vbNullString, vbNullString, 1)
    lblWeb.ForeColor = &H800080
    Exit Sub
errHnd:

    MsgBox "Please Use Your Default Application For Visit Blog.", vbCritical, "Error"
End Sub


Private Sub lblWeb2_Click()
    On Error GoTo errHnd
    Dim ret As Long
    ret = ShellExecute(Me.hwnd, "Open", "http://www.kpnewastro.blogspot.com", vbNullString, vbNullString, 1)
    lblWeb2.ForeColor = &H800080
    Exit Sub
errHnd:

    MsgBox "Please Use Your Default Application For Visit Blog.", vbCritical, "Error"

End Sub


Private Sub mnuFileExit_Click()
    End
End Sub

Private Sub mnuFileNew_Click()
    TabStrip1.Tabs(1).Selected = True
End Sub



Private Sub mnuFileSaveAsPDF_Click()
    On Error GoTo FileError:

    cDlg1.InitDir = App.Path + "\Saved"

    cDlg1.CancelError = True
    cDlg1.DefaultExt = "pdf"
    cDlg1.Filter = "PDF Files|*.pdf"    '|Text File|*.txt|All Files|*.*"
    cDlg1.FileName = "KPNewAstro_" + Format(Now, "DDMMYYYYHHMMSS")
    cDlg1.ShowSave

    Call A98_printPDF(cDlg1.FileName)


    Exit Sub

FileError:
    If Err.Number = cdlCancel Then Exit Sub
    MsgBox "Unkown Error While Saving File " & cDlg1.FileName

End Sub

Private Sub mnuFileSaveAsRTF_Click()
On Error GoTo FileError:
            
    cDlg1.InitDir = App.Path + "\Saved"
    
    cDlg1.CancelError = True
    cDlg1.DefaultExt = "RTF"
    cDlg1.Filter = "RTF Files|*.RTF" '|Text File|*.txt|All Files|*.*"
    cDlg1.FileName = "KPNewAstro_" + Format(Now, "DDMMYYYYHHMMSS")
    cDlg1.ShowSave
    
    'If cDlg1.Filter = "RTF Files" Then
        RT1.SaveFile cDlg1.FileName, rtfRTF 'rtfText
    'Else
        'RTAll.SaveFile cDlg1.FileName, rtfText
    'End If
    
    Exit Sub
    
FileError:
    If Err.Number = cdlCancel Then Exit Sub
    MsgBox "Unkown Error While Saving File " & cDlg1.FileName
End Sub

Private Sub mnuFileSaveAsTXT_Click()
On Error GoTo FileError:
            
    cDlg1.InitDir = App.Path + "\Saved"
    
    cDlg1.CancelError = True
    cDlg1.DefaultExt = "txt"
    cDlg1.Filter = "Text Files|*.txt" '|Text File|*.txt|All Files|*.*"
    cDlg1.FileName = "KPNewAstro_" + Format(Now, "DDMMYYYYHHMMSS")
    cDlg1.ShowSave
    
    'If cDlg1.Filter = "RTF Files" Then
        RT1.SaveFile cDlg1.FileName, rtfText 'rtfText
    'Else
        'RTAll.SaveFile cDlg1.FileName, rtfText
    'End If
    
    Exit Sub
    
FileError:
    If Err.Number = cdlCancel Then Exit Sub
    MsgBox "Unkown Error While Saving File " & cDlg1.FileName
End Sub

Private Sub mnuHelpAbout_Click()
    Me.Enabled = False

    Load frmAbout
    frmAbout.Show
End Sub

Private Sub mnuHelpKP_Click()
    On Error GoTo errHnd
    Dim ret As Long
    ret = ShellExecute(Me.hwnd, "Open", "http://www.kpnewastro.blogspot.com", vbNullString, vbNullString, 1)
    'lblWeb.ForeColor = &H800080
    Exit Sub
errHnd:

    MsgBox "Please Use Your Default Application For Visit Blog.", vbCritical, "Error"

End Sub



Private Sub mnuHelpKPPDF_Click()
On Error GoTo errHnd
    Dim ret As Long
    ret = ShellExecute(Me.hwnd, "Open", App.Path + "\KPNA HELP.pdf", vbNullString, vbNullString, 1)
    lblWeb2.ForeColor = &H800080
    Exit Sub
errHnd:
    MsgBox "File cannot be found.", vbCritical, "Error"
End Sub

Private Sub mnuToolsSettings_Click()

    Me.Enabled = False

    Load frmSettings
    frmSettings.Show

End Sub



Private Sub TabStrip1_Click()

    If TabStrip1.SelectedItem.Index = 1 Then
        picNatalHorary.Visible = True
        picKPChart.Visible = False
        'picExtended.Visible = False
        picRulingPlanets.Visible = False
        picAbout.Visible = False
    End If

    If TabStrip1.SelectedItem.Index = 2 Then
        picNatalHorary.Visible = False
        picKPChart.Visible = True
        'picExtended.Visible = False
        picRulingPlanets.Visible = False
        picAbout.Visible = False
    End If

    If TabStrip1.SelectedItem.Index = 3 Then
        picNatalHorary.Visible = False
        picKPChart.Visible = False
        'picExtended.Visible = False
        picRulingPlanets.Visible = True
        picAbout.Visible = False
    End If

    If TabStrip1.SelectedItem.Index = 4 Then
        picNatalHorary.Visible = False
        picKPChart.Visible = False
        'picExtended.Visible = False
        picRulingPlanets.Visible = False
        picAbout.Visible = True
    End If

    
    If TabStrip1.SelectedItem.Index = 1 Then
        'Add a random number to combo box
        cmb249.Clear
        Dim ib As Integer
        For ib = 1 To 249
            cmb249.AddItem CStr(ib)
        Next

        Dim rndInt As Integer
        Call Randomize
        rndInt = Int(1 + Rnd(1) * 248)

        cmb249.Text = CStr(rndInt)

        'Load current (Horary)
        txtDayCur.Text = Format(Day(Now), "00")
        txtMonthCur.Text = Format(Month(Now), "00")
        txtYearCur.Text = Format(Year(Now), "0000")

        txtHHCur.Text = Format(Hour(Now), "00")
        txtMMCur.Text = Format(Minute(Now), "00")
        txtSSCur.Text = Format(Second(Now), "00")
    End If
End Sub


Private Sub txtBirthPlace_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890,.")
End Sub



Private Sub txtDayCur_KeyPress(KeyAscii As Integer)
KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub





Private Sub txtHHCur_KeyPress(KeyAscii As Integer)
KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub







Private Sub txtLatDDCur_KeyPress(KeyAscii As Integer)
KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub



Private Sub txtLatMMCur_KeyPress(KeyAscii As Integer)
KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub



Private Sub txtLatNSCur_KeyPress(KeyAscii As Integer)
KeyAscii = hp_KeyFilter(KeyAscii, "NSns")
End Sub

Private Sub txtLonDDCur_KeyPress(KeyAscii As Integer)
KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub

Private Sub txtLonEWCur_KeyPress(KeyAscii As Integer)
KeyAscii = hp_KeyFilter(KeyAscii, "EWew")
End Sub

Private Sub txtLonMMCur_KeyPress(KeyAscii As Integer)
KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub

Private Sub txtMMCur_KeyPress(KeyAscii As Integer)
KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub

Private Sub txtMonthCur_KeyPress(KeyAscii As Integer)
KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub

Private Sub txtName_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890")
End Sub



Public Sub A97_printRTF()

'Settings RTextBox
    frmMain.RT1.Text = vbNullString
    frmMain.RT1.SelIndent = 300

    'Basic Details
    'frmMain.RT1.SelColor = rgb(250, 0, 0)
    frmMain.RT1.SelText = vbNewLine + vbNewLine

    frmMain.RT1.SelText = ppName + vbNewLine
    frmMain.RT1.SelText = ppBDate + " ( " + weekDayName(stdWeekDay) + " )" + vbNewLine
    frmMain.RT1.SelText = ppBTime + vbNewLine

    frmMain.RT1.SelText = ppPlaceName + vbNewLine
    frmMain.RT1.SelText = ppPlaceDetails + vbNewLine

    frmMain.RT1.SelText = vbNewLine

    'Some Calculated Details
    frmMain.RT1.SelText = "ASCENDENT" + vbTab + vbTab + ": " + rashiNameTr(kp_RashiInt(hCus(0))) + " ( Rotated by " + cmbRot.Text + " )" + vbNewLine
    frmMain.RT1.SelText = "STAR" + vbTab + vbTab + ": " + starName(kp_StarInt(pPos(1))) + " (" + CStr(kp_StarPada(pPos(1))) + ")" + vbNewLine
    frmMain.RT1.SelText = "KP AYANAMSA" + vbTab + ": " + hp_FormalDeg(kpAyan) + vbNewLine
    frmMain.RT1.SelText = "DASA BALANCE" + vbTab + ": " + kp_DasaShesha(pPos(1)) + vbNewLine

    'House cusps details etc....
    frmMain.RT1.SelText = vbNewLine
    frmMain.RT1.SelText = vbNewLine

    frmMain.RT1.SelText = "DETAILS OF CUSPS AND PLANETS" + vbNewLine
    frmMain.RT1.SelText = vbNewLine

    frmMain.RT1.SelText = "CS " + "POSITION  " + "SGN     " + "SGL " + "STL " + "SBL " + "SSL" + " | " _
                          + "PLN " + "POSITION  " + "SGN     " + "PS " + "SGL " + "STL " + "SBL " + "SSL" + vbNewLine
    frmMain.RT1.SelText = String(79, "-") + vbNewLine

    Dim i As Byte
    For i = 0 To 11
        frmMain.RT1.SelText = Format(i + 1, "00") + " " + hp_FormalDegSh(hCus(i)) + " " + rashiNameTrSh(kp_RashiInt(hCus(i))) + " " + kp_PlanetName(kp_RashiLordInt(hCus(i), False), isRe(kp_RashiLordInt(hCus(i), False))) + " " + kp_PlanetName(kp_StarLordInt(hCus(i), False), isRe(kp_StarLordInt(hCus(i), False))) + " " + kp_PlanetName(kp_PID2N(kp_SubLord(hCus(i), False)), isRe(kp_PID2N(kp_SubLord(hCus(i), False)))) + " " + kp_PlanetName(kp_PID2N(kp_SubLord(hCus(i), True)), isRe(kp_PID2N(kp_SubLord(hCus(i), True)))) + " | " _
                              + UCase(kp_PlanetName(i, isRe(i))) + " " + hp_FormalDegSh(pPos(i)) + " " + rashiNameTrSh(kp_RashiInt(pPos(i))) + " " + Format(incHouse(i), "00") + " " + kp_PlanetName(kp_RashiLordInt(pPos(i), False), isRe(kp_RashiLordInt(pPos(i), False))) + " " + kp_PlanetName(kp_StarLordInt(pPos(i), False), isRe(kp_StarLordInt(pPos(i), False))) + " " + kp_PlanetName(kp_PID2N(kp_SubLord(pPos(i), False)), isRe(kp_PID2N(kp_SubLord(pPos(i), False)))) + " " + kp_PlanetName(kp_PID2N(kp_SubLord(pPos(i), True)), isRe(kp_PID2N(kp_SubLord(pPos(i), True)))) + vbNewLine
    Next

    'Draw KP Chart
    frmMain.RT1.SelText = vbNewLine
    frmMain.RT1.SelText = vbNewLine

    frmMain.RT1.SelText = "STANDARD KP CHART" + vbNewLine
    frmMain.RT1.SelText = vbNewLine

    Dim i1 As Byte
    For i1 = 0 To 36
        frmMain.RT1.SelText = charDraw(i1)
    Next

    'Vedic Aspects
    frmMain.RT1.SelText = vbNewLine
    frmMain.RT1.SelText = vbNewLine

    frmMain.RT1.SelText = "VEDIC ASPECTS" + vbNewLine
    frmMain.RT1.SelText = vbNewLine

    Dim i2 As Byte, i3 As Byte
    For i2 = 0 To 6
        frmMain.RT1.SelText = UCase(kp_PlanetName(i2, isRe(i2))) + ": " + aspVedic(i2) + vbNewLine
    Next

    frmMain.RT1.SelText = vbNewLine

    For i3 = 0 To 6
        frmMain.RT1.SelText = UCase(kp_PlanetName(i3, isRe(i3))) + ": " + vedAspP(i3) + " |" + smeRas(i3) + "|" + vbNewLine
    Next

    'Western Aspects
    frmMain.RT1.SelText = vbNewLine
    frmMain.RT1.SelText = vbNewLine

    frmMain.RT1.SelText = "WESTERN ASPECTS" + vbNewLine
    frmMain.RT1.SelText = vbNewLine

    Dim i6 As Byte
    Dim strPlArr As String
    strPlArr = "    "
    For i6 = 0 To 11
        strPlArr = strPlArr + UCase(kp_PlanetName(i6, isRe(i6))) + Space(2)
    Next
    frmMain.RT1.SelText = strPlArr + vbNewLine
    'frmmain.RT1.SelText = "    " + "RAV  " + "CHA  " + "KUJ  " + "BUD  " + "GUR  " + "SUK  " + "SAN  " + "RAH  " + "KET  " + "URA  " + "NEP  " + "FOR" + vbNewLine
    frmMain.RT1.SelText = String(62, "-") + vbNewLine

    Dim i4 As Byte, i5 As Byte

    For i4 = 0 To 11
        frmMain.RT1.SelText = Format(i4 + 1, "00") + "  " + kp_WAspFilter(pToHouse(0, i4), False) + "  " + kp_WAspFilter(pToHouse(1, i4), False) + "  " + kp_WAspFilter(pToHouse(2, i4), False) + "  " + kp_WAspFilter(pToHouse(3, i4), False) + "  " + kp_WAspFilter(pToHouse(4, i4), False) + "  " + kp_WAspFilter(pToHouse(5, i4), False) + "  " + kp_WAspFilter(pToHouse(6, i4), False) + "  " + kp_WAspFilter(pToHouse(7, i4), False) + "  " + kp_WAspFilter(pToHouse(8, i4), False) + "  " + kp_WAspFilter(pToHouse(9, i4), False) + "  " + kp_WAspFilter(pToHouse(10, i4), False) + "  " + kp_WAspFilter(pToHouse(11, i4), False) + vbNewLine
    Next

    frmMain.RT1.SelText = String(62, "-") + vbNewLine

    For i5 = 0 To 10
        frmMain.RT1.SelText = UCase(kp_PlanetName(i5, isRe(i5))) + " " + kp_WAspFilter(pToPlanet(0, i5), False) + "  " + kp_WAspFilter(pToPlanet(1, i5), False) + "  " + kp_WAspFilter(pToPlanet(2, i5), False) + "  " + kp_WAspFilter(pToPlanet(3, i5), False) + "  " + kp_WAspFilter(pToPlanet(4, i5), False) + "  " + kp_WAspFilter(pToPlanet(5, i5), False) + "  " + kp_WAspFilter(pToPlanet(6, i5), False) + "  " + kp_WAspFilter(pToPlanet(7, i5), False) + "  " + kp_WAspFilter(pToPlanet(8, i5), False) + "  " + kp_WAspFilter(pToPlanet(9, i5), False) + "  " + kp_WAspFilter(pToPlanet(10, i5), False) + "  " + kp_WAspFilter(pToPlanet(11, i5), False) + vbNewLine
    Next

    frmMain.RT1.SelText = vbNewLine
    'frmMain.RT1.SelText = "NOTE : Internally aspects are calculated upto 6 decimal points" + vbNewLine
    'frmMain.RT1.SelText = "       and shown above are nearest absolute values." + vbNewLine

    'Conventional Significators
    'plntWSig
    frmMain.RT1.SelText = vbNewLine

    frmMain.RT1.SelText = "CONVENTIONAL SIGNIFICATORS (WITHOUT CONSIDERING ASPECTS)" + vbNewLine
    frmMain.RT1.SelText = "ASCENDING ORDER OF STRENGTHS" + vbNewLine
    frmMain.RT1.SelText = vbNewLine

    Dim i7 As Byte
    For i7 = 0 To 11
        frmMain.RT1.SelText = Format(i7 + 1, "00") + " » " + conSig0(i7) + "|" + conSig1(i7) + "|" + conSig2(i7) + "|" + conSig3(i7) + vbNewLine
    Next

    'Conventional Significators planet wise
    frmMain.RT1.SelText = vbNewLine
    frmMain.RT1.SelText = "CONVENTIONAL SIGNIFICATORS PLANETWISE (WITHOUT CONSIDERING ASPECTS)" + vbNewLine
    frmMain.RT1.SelText = vbNewLine
    Dim i8 As Byte

    For i8 = 0 To 11
        frmMain.RT1.SelText = UCase(kp_PlanetName(i8, isRe(i8))) + " » " + Replace(plntWSig(i8), Space(1), ", ") + vbNewLine
    Next

    'Print 4-Step Helper
    frmMain.RT1.SelText = vbNewLine
    frmMain.RT1.SelText = vbNewLine
    frmMain.RT1.SelText = "FOUR STEP HELPER" + vbNewLine
    frmMain.RT1.SelText = vbNewLine
    frmMain.RT1.SelText = "PLANETS IN SELF CONSTELLATION : " + selfCons + vbNewLine
    frmMain.RT1.SelText = "NO PLANETS IN THE STARS OF    : " + Replace(noPInStar, Space(1), ", ") + vbNewLine
    frmMain.RT1.SelText = "EMPTY HOUSES                  : " + Replace(emptyBhava, Space(1), ", ") + vbNewLine
    frmMain.RT1.SelText = vbNewLine

    Dim i9 As Integer
    For i9 = 0 To 8
        'frmMain.RT1.SelText = UCase(planetName2Tr(i9)) + vbNewLine
        frmMain.RT1.SelText = UCase(kp_PlanetName(kp_PID2N(i9), isRe(kp_PID2N(i9)))) + "                     : " + planetName2Tr(i9) + "  : " + " DEPOSITED : " + hp_4StepFil(posBhava(i9)) + "   OWN : " + ownBhava(i9) + vbNewLine
        frmMain.RT1.SelText = "  » STAR LORD           : " + planetName2Tr(stLords(i9)) + "  : " + " DEPOSITED : " + hp_4StepFil(stLordsPos(i9)) + "   OWN : " + stLordsOwn(i9) + vbNewLine
        frmMain.RT1.SelText = "  » SUB LORD            : " + planetName2Tr(sbLords(i9)) + "  : " + " DEPOSITED : " + hp_4StepFil(sbLordsPos(i9)) + "   OWN : " + sbLordsOwn(i9) + vbNewLine
        frmMain.RT1.SelText = "  » ST.LORD OF SUB LORD : " + planetName2Tr(stSbLords(i9)) + "  : " + " DEPOSITED : " + hp_4StepFil(stSbLordsPos(i9)) + "   OWN : " + stSbLordsOwn(i9) + vbNewLine
        frmMain.RT1.SelText = vbNewLine
    Next

    'Dasas And Bukthis
    frmMain.RT1.SelText = vbNewLine
    frmMain.RT1.SelText = "VIMSHOTTARI DASAS" + vbNewLine
    frmMain.RT1.SelText = vbNewLine

    Dim k1 As Byte, k2 As Byte, k3 As Byte
    '0-5    'frmmain.RT1.SelText = "Age - " + hp_AgeCal()+vbNewLine

    frmMain.RT1.SelText = String(19, "-") + Space(6) + String(19, "-") + Space(6) + String(19, "-") + vbNewLine
    frmMain.RT1.SelText = planetName2Tr(mDasL(0)) + " M.D" + Space(2) + kp_RevJday(mDas(0)) + Space(6) + planetName2Tr(mDasL(1)) + " M.D" + Space(2) + kp_RevJday(mDas(1)) + Space(6) + planetName2Tr(mDasL(2)) + " M.D" + Space(2) + kp_RevJday(mDas(2)) + vbNewLine
    frmMain.RT1.SelText = String(19, "-") + Space(6) + String(19, "-") + Space(6) + String(19, "-") + vbNewLine
    For k1 = 0 To 8
        frmMain.RT1.SelText = planetName2Tr(bDasL(0, k1)) + " B.D" + Space(2) + kp_RevJday(bDas(0, k1)) + Space(6) + planetName2Tr(bDasL(1, k1)) + " B.D" + Space(2) + kp_RevJday(bDas(1, k1)) + Space(6) + planetName2Tr(bDasL(2, k1)) + " B.D" + Space(2) + kp_RevJday(bDas(2, k1)) + vbNewLine
    Next
    frmMain.RT1.SelText = vbNewLine

    frmMain.RT1.SelText = String(19, "-") + Space(6) + String(19, "-") + Space(6) + String(19, "-") + vbNewLine
    frmMain.RT1.SelText = planetName2Tr(mDasL(3)) + " M.D" + Space(2) + kp_RevJday(mDas(3)) + Space(6) + planetName2Tr(mDasL(4)) + " M.D" + Space(2) + kp_RevJday(mDas(4)) + Space(6) + planetName2Tr(mDasL(5)) + " M.D" + Space(2) + kp_RevJday(mDas(5)) + vbNewLine
    frmMain.RT1.SelText = String(19, "-") + Space(6) + String(19, "-") + Space(6) + String(19, "-") + vbNewLine
    For k2 = 0 To 8
        frmMain.RT1.SelText = planetName2Tr(bDasL(3, k2)) + " B.D" + Space(2) + kp_RevJday(bDas(3, k2)) + Space(6) + planetName2Tr(bDasL(4, k2)) + " B.D" + Space(2) + kp_RevJday(bDas(4, k2)) + Space(6) + planetName2Tr(bDasL(5, k2)) + " B.D" + Space(2) + kp_RevJday(bDas(5, k2)) + vbNewLine
    Next
    frmMain.RT1.SelText = vbNewLine

    frmMain.RT1.SelText = String(19, "-") + Space(6) + String(19, "-") + Space(6) + String(19, "-") + vbNewLine
    frmMain.RT1.SelText = planetName2Tr(mDasL(6)) + " M.D" + Space(2) + kp_RevJday(mDas(6)) + Space(6) + planetName2Tr(mDasL(7)) + " M.D" + Space(2) + kp_RevJday(mDas(7)) + Space(6) + planetName2Tr(mDasL(8)) + " M.D" + Space(2) + kp_RevJday(mDas(8)) + vbNewLine
    frmMain.RT1.SelText = String(19, "-") + Space(6) + String(19, "-") + Space(6) + String(19, "-") + vbNewLine
    For k3 = 0 To 8
        frmMain.RT1.SelText = planetName2Tr(bDasL(0, k3)) + " B.D" + Space(2) + kp_RevJday(bDas(6, k3)) + Space(6) + planetName2Tr(bDasL(7, k3)) + " B.D" + Space(2) + kp_RevJday(bDas(7, k3)) + Space(6) + planetName2Tr(bDasL(8, k3)) + " B.D" + Space(2) + kp_RevJday(bDas(8, k3)) + vbNewLine
    Next
    'frmmain.RT1.SelText = String(19, "-") + Space(6) + String(19, "-") + Space(6) + String(19, "-") + vbNewLine

    'Astrologer Personal
    frmMain.RT1.SelText = vbNewLine
    frmMain.RT1.SelText = vbNewLine

    frmMain.RT1.SelText = "ASTROLOGER" + vbNewLine
    frmMain.RT1.SelText = "Name      : " + loadedPD(0) + vbNewLine
    frmMain.RT1.SelText = "Address   : " + loadedPD(1) + vbNewLine
    frmMain.RT1.SelText = "Telephone : " + loadedPD(2) + vbNewLine
    frmMain.RT1.SelText = "Email     : " + loadedPD(3) + vbNewLine

    frmMain.RT1.SelText = vbNewLine
    frmMain.RT1.SelText = "KP New Astro Version : " + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision) + " beta" + vbNewLine
    frmMain.RT1.SelText = "kpnewastro@gmail.com" + vbNewLine

    frmMain.RT1.SelText = vbNewLine
    '    frmmain.RT1.SelText = "" + vbNewLine
    '    frmmain.RT1.SelText = "" + vbNewLine
    '    frmmain.RT1.SelText = "" + vbNewLine
    '    frmmain.RT1.SelText = "" + vbNewLine
End Sub

Private Sub A98_printPDF(savePath As String)

    Dim sPDF As New clsPDFCreator

    With sPDF
        .Title = "KP New Astro"             ' Title
        .ScaleMode = pdfCentimeter          ' Units
        .PaperSize = pdfA4                  ' Paper Size
        .Margin = 0                         ' Margin
        .Orientation = pdfPortrait          ' Orientation

        .InitPDFFile (savePath) '(App.Path + "/KP NEW ASTRO.pdf")     'Initialize the pdf

        ' Define the Fonts
        .LoadFont "TNRR", "Times New Roman"
        .LoadFont "TNRB", "Times New Roman", pdfBold
        .LoadFont "TNRI", "Times New Roman", pdfItalic
        .LoadFont "TNRBI", "Times New Roman", pdfBoldItalic

        .LoadFont "ARR", "Arial"
        .LoadFont "ARB", "Arial", pdfBold
        .LoadFont "ARI", "Arial", pdfItalic
        .LoadFont "ARBI", "Arial", pdfBoldItalic

        .LoadFont "CNR", "Courier New"
        .LoadFont "CNB", "Courier New", pdfBold
        .LoadFont "CNI", "Courier New", pdfItalic
        .LoadFont "CNBI", "Courier New", pdfBoldItalic

        'A4 = 21.0 x 29.7 cm
        .StartPage    'Start Page 01

        .DrawText 19, 1, Trim(CStr(.Pages)), "TNRI", 10, pdfAlignRight
        '.DrawObject "Footers"

        'WATER MARK
        '.SetColorFill rgb(252, 252, 252)
        '.DrawText 5, 3, "KP New Astro", "ARI", 120, , 60

        'Heding
        .SetColorFill rgb(0, 0, 0)
        .DrawText 19.9, 28, "KP New Astro Version " + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision) + " beta", "ARBI", 12, pdfAlignRight

        'A line
        .MoveTo 2.5, 27.9
        .LineTo 19.9, 27.9, Filled

        'Basic Details 01
        .DrawText 2.5, 27.2, "NAME", "TNRB", 12, pdfAlignLeft
        .DrawText 6, 27.2, ": " + ppNamePdf, "TNRR", 12, pdfAlignLeft

        .DrawText 2.5, 26.7, "DATE", "TNRB", 12, pdfAlignLeft
        .DrawText 6, 26.7, ": " + ppBDatePdf, "TNRR", 12, pdfAlignLeft

        .DrawText 2.5, 26.2, "TIME", "TNRB", 12, pdfAlignLeft
        .DrawText 6, 26.2, ": " + ppBTimePdf, "TNRR", 12, pdfAlignLeft

        .DrawText 2.5, 25.7, "PLACE", "TNRB", 12, pdfAlignLeft
        .DrawText 6, 25.7, ": " + ppPlaceNamePdf, "TNRR", 12, pdfAlignLeft

        .DrawText 2.5, 25.2, "PLACE DETAILS", "TNRB", 12, pdfAlignLeft
        .DrawText 6, 25.2, ": " + ppPlaceDetailsPdf, "TNRR", 12, pdfAlignLeft

        'Basic Details 02
        .DrawText 2.5, 24.3, "ASCENDENT", "TNRB", 12, pdfAlignLeft
        .DrawText 6, 24.3, ": " + rashiNameTr(kp_RashiInt(hCus(0))) + " ( Rotated by " + cmbRot.Text + " )", "TNRR", 12, pdfAlignLeft

        .DrawText 2.5, 23.8, "STAR", "TNRB", 12, pdfAlignLeft
        .DrawText 6, 23.8, ": " + starName(kp_StarInt(pPos(1))) + " (" + CStr(kp_StarPada(pPos(1))) + ")", "TNRR", 12, pdfAlignLeft

        .DrawText 2.5, 23.3, "KP AYANAMSA", "TNRB", 12, pdfAlignLeft
        .DrawText 6, 23.3, ": " + hp_FormalDeg(kpAyan), "TNRR", 12, pdfAlignLeft

        .DrawText 2.5, 22.8, "DASA BALANCE", "TNRB", 12, pdfAlignLeft
        .DrawText 6, 22.8, ": " + kp_DasaShesha(pPos(1)), "TNRR", 12, pdfAlignLeft

        'Cusps & Planets
        .DrawText 2.5, 21.8, "DETAILS OF CUSPS AND PLANETS", "TNRB", 12, pdfAlignLeft

        'CS
        .DrawText 2.5, 21.1, "CS", "TNRB", 12, pdfAlignLeft
        Dim i1 As Byte
        For i1 = 0 To 11
            .DrawText 2.5, (21 - (i1 + 1) / 2), Format(i1 + 1, "00"), "TNRB", 12, pdfAlignLeft
        Next

        'POSN
        .DrawText 3.2, 21.1, "POSN", "TNRB", 12, pdfAlignLeft
        Dim i2 As Byte
        For i2 = 0 To 11
            .DrawText 3.2, (21 - (i2 + 1) / 2), hp_FormalDegSh(hCus(i2)), "TNRR", 12, pdfAlignLeft
        Next

        'SGN
        .DrawText 5.15, 21.1, "SGN", "TNRB", 12, pdfAlignLeft
        Dim i3 As Byte
        For i3 = 0 To 11
            .DrawText 5.15, (21 - (i3 + 1) / 2), rashiNameTrSh(kp_RashiInt(hCus(i3))), "TNRR", 12, pdfAlignLeft
        Next

        'SGL
        .DrawText 7.05, 21.1, "SGL", "TNRB", 12, pdfAlignLeft
        Dim i4 As Byte
        For i4 = 0 To 11
            .DrawText 7.05, (21 - (i4 + 1) / 2), kp_PlanetName(kp_RashiLordInt(hCus(i4), False), isRe(kp_RashiLordInt(hCus(i4), False))), "TNRR", 12, pdfAlignLeft
        Next

        'STL
        .DrawText 7.95, 21.1, "STL", "TNRB", 12, pdfAlignLeft
        Dim i5 As Byte
        For i5 = 0 To 11
            .DrawText 7.95, (21 - (i5 + 1) / 2), kp_PlanetName(kp_StarLordInt(hCus(i5), False), isRe(kp_StarLordInt(hCus(i5), False))), "TNRR", 12, pdfAlignLeft
        Next

        'SBL
        .DrawText 8.85, 21.1, "SBL", "TNRB", 12, pdfAlignLeft
        Dim i6 As Byte
        For i6 = 0 To 11
            .DrawText 8.85, (21 - (i6 + 1) / 2), kp_PlanetName(kp_PID2N(kp_SubLord(hCus(i6), False)), isRe(kp_PID2N(kp_SubLord(hCus(i6), False)))), "TNRR", 12, pdfAlignLeft
        Next

        'SSL
        .DrawText 9.75, 21.1, "SSL", "TNRB", 12, pdfAlignLeft
        Dim i7 As Byte
        For i7 = 0 To 11
            .DrawText 9.75, (21 - (i7 + 1) / 2), kp_PlanetName(kp_PID2N(kp_SubLord(hCus(i7), True)), isRe(kp_PID2N(kp_SubLord(hCus(i7), True)))), "TNRR", 12, pdfAlignLeft
        Next

        'PLANETS
        'PLN
        .DrawText 10.9, 21.1, "PLN", "TNRB", 12, pdfAlignLeft
        Dim j1 As Byte
        For j1 = 0 To 11
            .DrawText 10.9, (21 - (j1 + 1) / 2), UCase(kp_PlanetName(j1, isRe(j1))), "TNRB", 12, pdfAlignLeft
        Next

        'POSN
        .DrawText 12.1, 21.1, "POSN", "TNRB", 12, pdfAlignLeft
        Dim j2 As Byte
        For j2 = 0 To 11
            .DrawText 12.1, (21 - (j2 + 1) / 2), hp_FormalDegSh(pPos(j2)), "TNRR", 12, pdfAlignLeft
        Next

        'SGN
        .DrawText 14.05, 21.1, "SGN", "TNRB", 12, pdfAlignLeft
        Dim j3 As Byte
        For j3 = 0 To 11
            .DrawText 14.05, (21 - (j3 + 1) / 2), rashiNameTrSh(kp_RashiInt(pPos(j3))), "TNRR", 12, pdfAlignLeft
        Next

        'PS
        .DrawText 15.9, 21.1, "PS", "TNRB", 12, pdfAlignLeft
        Dim j4 As Byte
        For j4 = 0 To 11
            .DrawText 15.9, (21 - (j4 + 1) / 2), Format(incHouse(j4), "00"), "TNRR", 12, pdfAlignLeft
        Next

        'SGL
        .DrawText 16.55, 21.1, "SGL", "TNRB", 12, pdfAlignLeft
        Dim j5 As Byte
        For j5 = 0 To 11
            .DrawText 16.55, (21 - (j5 + 1) / 2), kp_PlanetName(kp_RashiLordInt(pPos(j5), False), isRe(kp_RashiLordInt(pPos(j5), False))), "TNRR", 12, pdfAlignLeft
        Next

        'STL
        .DrawText 17.45, 21.1, "STL", "TNRB", 12, pdfAlignLeft
        Dim j6 As Byte
        For j6 = 0 To 11
            .DrawText 17.45, (21 - (j6 + 1) / 2), kp_PlanetName(kp_StarLordInt(pPos(j6), False), isRe(kp_StarLordInt(pPos(j6), False))), "TNRR", 12, pdfAlignLeft
        Next

        'SBL
        .DrawText 18.35, 21.1, "SBL", "TNRB", 12, pdfAlignLeft
        Dim j7 As Byte
        For j7 = 0 To 11
            .DrawText 18.35, (21 - (j7 + 1) / 2), kp_PlanetName(kp_PID2N(kp_SubLord(pPos(j7), False)), isRe(kp_PID2N(kp_SubLord(pPos(j7), False)))), "TNRR", 12, pdfAlignLeft
        Next

        'SSL
        .DrawText 19.25, 21.1, "SSL", "TNRB", 12, pdfAlignLeft
        Dim j8 As Byte
        For j8 = 0 To 11
            .DrawText 19.25, (21 - (j8 + 1) / 2), kp_PlanetName(kp_PID2N(kp_SubLord(pPos(j8), True)), isRe(kp_PID2N(kp_SubLord(pPos(j8), True)))), "TNRR", 12, pdfAlignLeft
        Next


        'KP Chart
        .DrawText 2.5, 14, "STANDARD KP CHART", "TNRB", 12, pdfAlignLeft

        'Cage
        '.SetColorFill rgb(0, 0, 0)
        '.SetLineWidth 10
        .Rectangle 2.5, 13.5, 16, -12
        .MoveTo 6.5, 13.5
        .LineTo 6.5, 1.5
        .MoveTo 14.5, 13.5
        .LineTo 14.5, 1.5
        .MoveTo 10.5, 13.5
        .LineTo 10.5, 10.5
        .MoveTo 10.5, 4.5
        .LineTo 10.5, 1.5
        .MoveTo 2.5, 10.5
        .LineTo 18.5, 10.5
        .MoveTo 2.5, 4.5
        .LineTo 18.5, 4.5
        .MoveTo 2.5, 7.5
        .LineTo 6.5, 7.5
        .MoveTo 14.5, 7.5
        .LineTo 18.5, 7.5

        'Vals
        'Cell 01
        Dim r1 As Byte
        For r1 = 0 To 7
            .DrawText 6.8, (13.2 - r1 * 0.35), pdfPX(0, r1), "CNR", 10, pdfAlignLeft
        Next

        'Cell 02
        Dim r2 As Byte
        For r2 = 0 To 7
            .DrawText 10.8, (13.2 - r2 * 0.35), pdfPX(1, r2), "CNR", 10, pdfAlignLeft
        Next

        'Cell 03
        Dim r3 As Byte
        For r3 = 0 To 7
            .DrawText 14.8, (13.2 - r3 * 0.35), pdfPX(2, r3), "CNR", 10, pdfAlignLeft
        Next

        'Cell 04
        Dim r4 As Byte
        For r4 = 0 To 7
            .DrawText 14.8, (10.2 - r4 * 0.35), pdfPX(3, r4), "CNR", 10, pdfAlignLeft
        Next

        'Cell 05
        Dim r5 As Byte
        For r5 = 0 To 7
            .DrawText 14.8, (7.2 - r5 * 0.35), pdfPX(4, r5), "CNR", 10, pdfAlignLeft
        Next

        'Cell 06
        Dim r6 As Byte
        For r6 = 0 To 7
            .DrawText 14.8, (4.2 - r6 * 0.35), pdfPX(5, r6), "CNR", 10, pdfAlignLeft
        Next

        'Cell 07
        Dim r7 As Byte
        For r7 = 0 To 7
            .DrawText 10.8, (4.2 - r7 * 0.35), pdfPX(6, 7 - r7), "CNR", 10, pdfAlignLeft
        Next

        'Cell 08
        Dim r8 As Byte
        For r8 = 0 To 7
            .DrawText 6.8, (4.2 - r8 * 0.35), pdfPX(7, 7 - r8), "CNR", 10, pdfAlignLeft
        Next

        'Cell 09
        Dim r9 As Byte
        For r9 = 0 To 7
            .DrawText 2.8, (4.2 - r9 * 0.35), pdfPX(8, 7 - r9), "CNR", 10, pdfAlignLeft
        Next

        'Cell 10
        Dim r10 As Byte
        For r10 = 0 To 7
            .DrawText 2.8, (7.2 - r10 * 0.35), pdfPX(9, 7 - r10), "CNR", 10, pdfAlignLeft
        Next

        'Cell 11
        Dim r11 As Byte
        For r11 = 0 To 7
            .DrawText 2.8, (10.2 - r11 * 0.35), pdfPX(10, 7 - r11), "CNR", 10, pdfAlignLeft
        Next

        'Cell 12
        Dim r12 As Byte
        For r12 = 0 To 7
            .DrawText 2.8, (13.2 - r12 * 0.35), pdfPX(11, 7 - r12), "CNR", 10, pdfAlignLeft
        Next

        'Cage MID
        .DrawText 6.8, 10, bAscM, "CNR", 10, pdfAlignLeft
        .DrawText 6.8, 9.5, bDateM, "CNR", 10, pdfAlignLeft
        .DrawText 6.8, 9, bTimeM, "CNR", 10, pdfAlignLeft
        .DrawText 6.8, 8.5, bCorrM, "CNR", 10, pdfAlignLeft
        .DrawText 6.8, 8, bPlaceM, "CNR", 10, pdfAlignLeft
        .DrawText 6.8, 7.5, bTzM, "CNR", 10, pdfAlignLeft
        .DrawText 6.8, 7, bAyanM, "CNR", 10, pdfAlignLeft
        .DrawText 10.5, 5.5, bNmeM, "CNR", 12, pdfCenter

        .EndPage    'End Page 01

        .StartPage    'Start Page 02
        .DrawText 19, 1, Trim(CStr(.Pages)), "TNRI", 10, pdfAlignRight
        'Heding
        .SetColorFill rgb(0, 0, 0)
        .DrawText 19.9, 28, "KP New Astro Version " + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision) + " beta", "ARBI", 12, pdfAlignRight

        'A line
        .MoveTo 2.5, 27.9
        .LineTo 19.9, 27.9, Filled

        'Vedic Aspects
        .DrawText 2.5, 26.5, "VEDIC ASPECTS", "TNRB", 12, pdfAlignLeft
        Dim k1 As Byte
        For k1 = 0 To 6
            .DrawText 2.5, (25.8 - k1 / 2), UCase(kp_PlanetName(k1, isRe(k1))), "TNRB", 12, pdfAlignLeft
        Next

        Dim k2 As Byte
        For k2 = 0 To 6
            .DrawText 3.7, (25.8 - k2 / 2), ":  " + aspVedic(k2), "TNRR", 12, pdfAlignLeft
        Next

        Dim k3 As Byte
        For k3 = 0 To 6
            .DrawText 2.5, (22 - k3 / 2), UCase(kp_PlanetName(k3, isRe(k3))), "TNRB", 12, pdfAlignLeft
        Next

        Dim k4 As Byte
        For k4 = 0 To 6
            .DrawText 3.7, (22 - k4 / 2), ":  " + vedAspP(k4) + " [" + smeRas(k4) + "]", "TNRR", 12, pdfAlignLeft
        Next

        'Western Aspects
        .DrawText 2.5, 17.5, "WESTERN ASPECTS", "TNRB", 12, pdfAlignLeft

        Dim k5 As Byte
        For k5 = 0 To 11
            .DrawText 2.5, (16.2 - k5 / 2), Format(k5 + 1, "00"), "TNRB", 12, pdfAlignLeft
        Next

        Dim k6 As Byte
        For k6 = 0 To 10
            .DrawText 2.5, (9.8 - k6 / 2), UCase(kp_PlanetName(k6, isRe(k6))), "TNRB", 12, pdfAlignLeft
        Next

        Dim m1 As Byte, m2 As Byte
        For m2 = 0 To 11
            .DrawText (3.9 + m2 * 1.25), 16.8, UCase(kp_PlanetName(m2, isRe(m2))), "TNRB", 12, pdfAlignLeft
            For m1 = 0 To 11
                .DrawText (3.9 + m2 * 1.25), (16.2 - m1 / 2), kp_WAspFilter(pToHouse(m2, m1), True), "TNRR", 12, pdfAlignLeft
            Next
        Next

        Dim m3 As Byte, m4 As Byte
        For m4 = 0 To 11
            For m3 = 0 To 10
                .DrawText (3.9 + m4 * 1.25), (9.8 - m3 / 2), kp_WAspFilter(pToPlanet(m4, m3), True), "TNRR", 12, pdfAlignLeft
            Next
        Next

        .EndPage    'End Page 02

        .StartPage  'Start Page 03
        .DrawText 19, 1, Trim(CStr(.Pages)), "TNRI", 10, pdfAlignRight

        'Heding
        .SetColorFill rgb(0, 0, 0)
        .DrawText 19.9, 28, "KP New Astro Version " + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision) + " beta", "ARBI", 12, pdfAlignRight

        'A line
        .MoveTo 2.5, 27.9
        .LineTo 19.9, 27.9, Filled

        'Significators
        .DrawText 2.5, 26.5, "CONVENTIONAL SIGNIFICATORS (WITHOUT CONSIDERING ASPECTS)", "TNRB", 12, pdfAlignLeft
        .DrawText 2.5, 26, "ASCENDING ORDER OF STRENGTHS", "TNRB", 12, pdfAlignLeft

        Dim a1 As Byte, a2 As Byte, a3 As Byte, a4 As Byte

        For a1 = 0 To 11
            .DrawText 2.5, (25.2 - a1 / 2), Format(a1 + 1, "00") + " -: ", "TNRB", 12, pdfAlignLeft
        Next

        For a2 = 0 To 11
            .DrawText 3.8, (25.2 - a2 / 2), conSig0(a2) + " -  " + conSig1(a2) + " -  " + conSig2(a2) + " -  " + conSig3(a2), "TNRR", 12, pdfAlignLeft
        Next

        'Planetwise
        .DrawText 2.5, 18.5, "CONVENTIONAL SIGNIFICATORS PLANETWISE", "TNRB", 12, pdfAlignLeft
        .DrawText 2.5, 18, "(WITHOUT CONSIDERING ASPECTS)", "TNRB", 12, pdfAlignLeft

        For a3 = 0 To 11
            .DrawText 2.5, (17.2 - a3 / 2), UCase(kp_PlanetName(a3, isRe(a3))), "TNRB", 12, pdfAlignLeft
        Next

        For a4 = 0 To 11
            .DrawText 3.8, (17.2 - a4 / 2), " -: " + Replace(Trim$(plntWSig(a4)), Space(1), ", "), "TNRR", 12, pdfAlignLeft
        Next

        'Extra Details

        .DrawText 2.5, 9, "Warning :", "ARB", 12, pdfAlignLeft
        .DrawText 2.5, 8.5, "Reports are generated by using " + "KP New Astro Version : " + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision) + " beta" + vbNewLine, "TNRR", 12, pdfAlignLeft
        .DrawText 2.5, 8, "Please send me bugs and comments for further development of this software.", "TNRR", 12, pdfAlignLeft
        .DrawText 2.5, 7.5, "kpnewastro@gmail.com", "TNRR", 12, pdfAlignLeft
        '        .DrawText 2.5, 7, "JJJJ", "TNRR", 12, pdfAlignLeft


        .EndPage    'End Page 03


        .StartPage    'Start Page 04
        .DrawText 19, 1, Trim(CStr(.Pages)), "TNRI", 10, pdfAlignRight
        'Heding
        .SetColorFill rgb(0, 0, 0)
        .DrawText 19.9, 28, "KP New Astro Version " + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision) + " beta", "ARBI", 12, pdfAlignRight

        'A line
        .MoveTo 2.5, 27.9
        .LineTo 19.9, 27.9, Filled

        '4-Step
        .DrawText 2.5, 26.5, "FOUR STEP HELPER", "TNRB", 12, pdfAlignLeft

        .DrawText 2.5, 25.8, "Planets in self constellation", "TNRB", 12, pdfAlignLeft
        .DrawText 7.5, 25.8, " :-  " + Replace(selfCons, Space(1), ", "), "TNRR", 12, pdfAlignLeft

        .DrawText 2.5, 25.3, "No planets in the starts of", "TNRB", 12, pdfAlignLeft
        .DrawText 7.5, 25.3, " :-  " + Replace(noPInStar, Space(1), ", "), "TNRR", 12, pdfAlignLeft

        .DrawText 2.5, 24.8, "Empty Houses", "TNRB", 12, pdfAlignLeft
        .DrawText 7.5, 24.8, " :-  " + Replace(emptyBhava, Space(1), ", "), "TNRR", 12, pdfAlignLeft


        Dim a5 As Byte, a6 As Byte, a7 As Byte, a8 As Byte

        For a5 = 0 To 8
            .DrawText 2.5, (24 - a5 * 2.4), UCase(kp_PlanetName(kp_PID2N(CInt(a5)), isRe(kp_PID2N(CInt(a5))))), "TNRB", 12, pdfAlignLeft
            .DrawText 2.5, (23.5 - a5 * 2.4), "  » STAR LORD", "TNRR", 12, pdfAlignLeft
            .DrawText 2.5, (23 - a5 * 2.4), "  » SUB LORD", "TNRR", 12, pdfAlignLeft
            .DrawText 2.5, (22.5 - a5 * 2.4), "  » ST.LORD OF SUB LORD", "TNRR", 12, pdfAlignLeft
        Next

        For a6 = 0 To 8
            .DrawText 8, (24 - a6 * 2.4), ": " + planetName2Tr(a6), "TNRR", 12, pdfAlignLeft
            .DrawText 8, (23.5 - a6 * 2.4), ": " + planetName2Tr(stLords(a6)), "TNRR", 12, pdfAlignLeft
            .DrawText 8, (23 - a6 * 2.4), ": " + planetName2Tr(sbLords(a6)), "TNRR", 12, pdfAlignLeft
            .DrawText 8, (22.5 - a6 * 2.4), ": " + planetName2Tr(stSbLords(a6)), "TNRR", 12, pdfAlignLeft
        Next

        For a7 = 0 To 8
            .DrawText 9.1, (24 - a7 * 2.4), ": DEPOSITED : " + Trim$(hp_4StepFil(posBhava(a7))), "TNRR", 12, pdfAlignLeft
            .DrawText 9.1, (23.5 - a7 * 2.4), ": DEPOSITED : " + Trim$(hp_4StepFil(stLordsPos(a7))), "TNRR", 12, pdfAlignLeft
            .DrawText 9.1, (23 - a7 * 2.4), ": DEPOSITED : " + Trim$(hp_4StepFil(sbLordsPos(a7))), "TNRR", 12, pdfAlignLeft
            .DrawText 9.1, (22.5 - a7 * 2.4), ": DEPOSITED : " + Trim$(hp_4StepFil(stSbLordsPos(a7))), "TNRR", 12, pdfAlignLeft
        Next

        For a8 = 0 To 8
            .DrawText 12.5, (24 - a8 * 2.4), ": OWN : " + Replace(Trim$(ownBhava(a8)), Space(1), ", "), "TNRR", 12, pdfAlignLeft
            .DrawText 12.5, (23.5 - a8 * 2.4), ": OWN : " + Replace(Trim$(stLordsOwn(a8)), Space(1), ", "), "TNRR", 12, pdfAlignLeft
            .DrawText 12.5, (23 - a8 * 2.4), ": OWN : " + Replace(Trim$(sbLordsOwn(a8)), Space(1), ", "), "TNRR", 12, pdfAlignLeft
            .DrawText 12.5, (22.5 - a8 * 2.4), ": OWN : " + Replace(Trim$(stSbLordsOwn(a8)), Space(1), ", "), "TNRR", 12, pdfAlignLeft
        Next

        .EndPage    'End Page 04

        .StartPage  'Start Page 05
        .DrawText 19, 1, Trim(CStr(.Pages)), "TNRI", 10, pdfAlignRight
        'Heding
        .SetColorFill rgb(0, 0, 0)
        .DrawText 19.9, 28, "KP New Astro Version " + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision) + " beta", "ARBI", 12, pdfAlignRight

        'A line
        .MoveTo 2.5, 27.9
        .LineTo 19.9, 27.9, Filled

        Dim b1 As Byte, b2 As Byte
        .DrawText 2.5, 26.5, "VIMSHOTTARI DASAS", "TNRB", 12, pdfAlignLeft
        b1 = 0
        Do While b1 <= 6
            .DrawText 2.5, 25.7 - b1 * 2, planetName2Tr(mDasL(b1)) + " M.D" + Space(2) + kp_RevJday(mDas(b1)) + Space(4) + planetName2Tr(mDasL(b1 + 1)) + " M.D" + Space(2) + kp_RevJday(mDas(b1 + 1)) + Space(4) + planetName2Tr(mDasL(b1 + 2)) + " M.D" + Space(2) + kp_RevJday(mDas(b1 + 2)), "CNB", 12, pdfAlignLeft
            For b2 = 0 To 8
                .DrawText 2.5, 25 - b2 / 2 - b1 * 2, planetName2Tr(bDasL(b1, b2)) + " B.D" + Space(2) + kp_RevJday(bDas(b1, b2)) + Space(4) + planetName2Tr(bDasL(b1 + 1, b2)) + " B.D" + Space(2) + kp_RevJday(bDas(b1 + 1, b2)) + Space(4) + planetName2Tr(bDasL(b1 + 2, b2)) + " B.D" + Space(2) + kp_RevJday(bDas(b1 + 2, b2)), "CNR", 12, pdfAlignLeft
            Next
            b1 = b1 + 3
        Loop

        .DrawText 2.5, 7, "ASTROLOGER", "TNRB", 12, pdfAlignLeft
        .DrawText 2.5, 6.2, "Name", "TNRR", 12, pdfAlignLeft
        .DrawText 2.5, 5.7, "Address", "TNRR", 12, pdfAlignLeft
        .DrawText 2.5, 5.2, "Telephone", "TNRR", 12, pdfAlignLeft
        .DrawText 2.5, 4.7, "Email", "TNRR", 12, pdfAlignLeft

        Dim b6 As Byte
        For b6 = 0 To 3
            .DrawText 4.5, (6.2 - b6 / 2), ":  " + loadedPD(b6), "TNRR", 12, pdfAlignLeft
        Next

        .EndPage    'End Page 05

        .ClosePDFFile    'Close the pdf


    End With



End Sub





Private Sub txtSSCur_KeyPress(KeyAscii As Integer)
KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub



Private Sub txtTZHHCur_KeyPress(KeyAscii As Integer)
KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub



Private Sub txtTZMMCur_KeyPress(KeyAscii As Integer)
KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub



Private Sub txtTZPMCur_KeyPress(KeyAscii As Integer)
KeyAscii = hp_KeyFilter(KeyAscii, "+-")
End Sub

Private Sub txtYearCur_KeyPress(KeyAscii As Integer)
KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub
