VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmSettings 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Settings"
   ClientHeight    =   5535
   ClientLeft      =   2100
   ClientTop       =   2325
   ClientWidth     =   6645
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5535
   ScaleWidth      =   6645
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picOther 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   3960
      ScaleHeight     =   855
      ScaleWidth      =   1095
      TabIndex        =   95
      Top             =   2160
      Visible         =   0   'False
      Width           =   1095
      Begin VB.CheckBox chkGeo 
         Alignment       =   1  'Right Justify
         Caption         =   "Use Geocentric Correction ------"
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
         Left            =   1680
         TabIndex        =   138
         Top             =   600
         Width           =   3015
      End
      Begin VB.CheckBox chkTrue 
         Alignment       =   1  'Right Justify
         Caption         =   "Use True Rahu/Ketu Positions --"
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
         Left            =   1680
         TabIndex        =   137
         Top             =   1200
         Width           =   3015
      End
      Begin VB.TextBox txtPD 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   2160
         MaxLength       =   70
         TabIndex        =   89
         Text            =   "kpnewastro@gmail.com"
         Top             =   3960
         Width           =   3975
      End
      Begin VB.TextBox txtPD 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   2160
         MaxLength       =   70
         TabIndex        =   88
         Text            =   "0000000000"
         Top             =   3600
         Width           =   3975
      End
      Begin VB.TextBox txtPD 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2160
         MaxLength       =   70
         TabIndex        =   87
         Text            =   "Kandy, Sri Lanka"
         Top             =   3240
         Width           =   3975
      End
      Begin VB.TextBox txtPD 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2160
         MaxLength       =   70
         TabIndex        =   86
         Text            =   "KP New Astro"
         Top             =   2880
         Width           =   3975
      End
      Begin VB.Line Line8 
         X1              =   240
         X2              =   6120
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label39 
         Caption         =   "Personal Details :"
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
         Left            =   240
         TabIndex        =   136
         Top             =   2520
         Width           =   1455
      End
      Begin VB.Label Label38 
         Caption         =   "Email :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   135
         Top             =   3975
         Width           =   1935
      End
      Begin VB.Label Label36 
         Caption         =   "Telephone : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   134
         Top             =   3615
         Width           =   1935
      End
      Begin VB.Label Label35 
         Caption         =   "Address : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   133
         Top             =   3255
         Width           =   1935
      End
      Begin VB.Label Label34 
         Caption         =   "Name Of The Astrologer :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   132
         Top             =   2895
         Width           =   1935
      End
   End
   Begin VB.PictureBox picPlanetNames 
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   480
      ScaleHeight     =   975
      ScaleWidth      =   975
      TabIndex        =   92
      Top             =   1680
      Width           =   975
      Begin VB.TextBox txtP 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   0
         Left            =   960
         MaxLength       =   3
         TabIndex        =   1
         Text            =   "Rav"
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txtP 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   1
         Left            =   960
         MaxLength       =   3
         TabIndex        =   2
         Text            =   "Cha"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox txtP 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   2
         Left            =   960
         MaxLength       =   3
         TabIndex        =   3
         Text            =   "Kuj"
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox txtP 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   3
         Left            =   960
         MaxLength       =   3
         TabIndex        =   4
         Text            =   "Bud"
         Top             =   3000
         Width           =   495
      End
      Begin VB.TextBox txtP 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   4
         Left            =   960
         MaxLength       =   3
         TabIndex        =   5
         Text            =   "Gur"
         Top             =   3480
         Width           =   495
      End
      Begin VB.TextBox txtP 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   5
         Left            =   960
         MaxLength       =   3
         TabIndex        =   6
         Text            =   "Suk"
         Top             =   3960
         Width           =   495
      End
      Begin VB.TextBox txtP 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   18
         Left            =   5760
         MaxLength       =   3
         TabIndex        =   19
         Text            =   "Ne®"
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox txtP 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   17
         Left            =   5760
         MaxLength       =   3
         TabIndex        =   18
         Text            =   "Ur®"
         Top             =   2640
         Width           =   495
      End
      Begin VB.TextBox txtP 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   16
         Left            =   5760
         MaxLength       =   3
         TabIndex        =   17
         Text            =   "Sa®"
         Top             =   2160
         Width           =   495
      End
      Begin VB.TextBox txtP 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   15
         Left            =   4200
         MaxLength       =   3
         TabIndex        =   16
         Text            =   "Su®"
         Top             =   3600
         Width           =   495
      End
      Begin VB.TextBox txtP 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   14
         Left            =   4200
         MaxLength       =   3
         TabIndex        =   15
         Text            =   "Gu®"
         Top             =   3120
         Width           =   495
      End
      Begin VB.TextBox txtP 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   13
         Left            =   4200
         MaxLength       =   3
         TabIndex        =   14
         Text            =   "Bu®"
         Top             =   2640
         Width           =   495
      End
      Begin VB.TextBox txtP 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   12
         Left            =   4200
         MaxLength       =   3
         TabIndex        =   13
         Text            =   "Ku®"
         Top             =   2160
         Width           =   495
      End
      Begin VB.CommandButton cmdOptGen 
         Caption         =   "<< Options"
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
         Left            =   5040
         TabIndex        =   0
         Top             =   120
         Width           =   1215
      End
      Begin VB.TextBox txtP 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   11
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   12
         Text            =   "For"
         Top             =   3960
         Width           =   495
      End
      Begin VB.TextBox txtP 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   10
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   11
         Text            =   "Nep"
         Top             =   3480
         Width           =   495
      End
      Begin VB.TextBox txtP 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   9
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   10
         Text            =   "Ura"
         Top             =   3000
         Width           =   495
      End
      Begin VB.TextBox txtP 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   8
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   9
         Text            =   "Ket"
         Top             =   2520
         Width           =   495
      End
      Begin VB.TextBox txtP 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   7
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   8
         Text            =   "Rah"
         Top             =   2040
         Width           =   495
      End
      Begin VB.TextBox txtP 
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Index           =   6
         Left            =   2520
         MaxLength       =   3
         TabIndex        =   7
         Text            =   "San"
         Top             =   1560
         Width           =   495
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "Note : Name of a planet must be a short name of three (3) characters."
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
         Left            =   120
         TabIndex        =   119
         Top             =   720
         Width           =   5880
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "Mars :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3360
         TabIndex        =   118
         Top             =   2205
         Width           =   450
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "Mercury :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3360
         TabIndex        =   117
         Top             =   2685
         Width           =   690
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "Jupiter :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3360
         TabIndex        =   116
         Top             =   3165
         Width           =   600
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "Saturn :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4920
         TabIndex        =   115
         Top             =   2205
         Width           =   585
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "Venus :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   3360
         TabIndex        =   114
         Top             =   3645
         Width           =   540
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "Uranus :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4920
         TabIndex        =   113
         Top             =   2685
         Width           =   615
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "Neptune :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   4920
         TabIndex        =   112
         Top             =   3165
         Width           =   720
      End
      Begin VB.Line Line6 
         X1              =   3120
         X2              =   3120
         Y1              =   1200
         Y2              =   4320
      End
      Begin VB.Line Line5 
         X1              =   120
         X2              =   6240
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line4 
         X1              =   120
         X2              =   6240
         Y1              =   600
         Y2              =   600
      End
      Begin VB.Label Label37 
         Caption         =   "Retrograde Planet Name :"
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
         Left            =   3360
         TabIndex        =   111
         Top             =   1200
         Width           =   2820
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "Planet Names :"
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
         Left            =   120
         TabIndex        =   110
         Top             =   1200
         Width           =   1245
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "Neptune :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   109
         Top             =   3525
         Width           =   720
      End
      Begin VB.Label Label27 
         AutoSize        =   -1  'True
         Caption         =   "Uranus :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   108
         Top             =   3045
         Width           =   615
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "Dec.Node"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   107
         Top             =   2565
         Width           =   705
      End
      Begin VB.Label Label25 
         Caption         =   "Note : Use special 3rd character for display retrograde planets."
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   435
         Left            =   3360
         TabIndex        =   106
         Top             =   1560
         Width           =   3000
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "Venus :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   105
         Top             =   4005
         Width           =   540
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "Fortune :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   104
         Top             =   4005
         Width           =   675
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "Asc.Node :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   103
         Top             =   2085
         Width           =   795
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "Saturn :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1680
         TabIndex        =   102
         Top             =   1605
         Width           =   585
      End
      Begin VB.Label Label15 
         AutoSize        =   -1  'True
         Caption         =   "Jupiter :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   101
         Top             =   3525
         Width           =   600
      End
      Begin VB.Label Label14 
         AutoSize        =   -1  'True
         Caption         =   "Mercury :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   100
         Top             =   3045
         Width           =   690
      End
      Begin VB.Label Label13 
         AutoSize        =   -1  'True
         Caption         =   "Mars :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   99
         Top             =   2565
         Width           =   450
      End
      Begin VB.Label Label12 
         AutoSize        =   -1  'True
         Caption         =   "Moon :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   98
         Top             =   2085
         Width           =   495
      End
      Begin VB.Label Label11 
         AutoSize        =   -1  'True
         Caption         =   "Sun :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   97
         Top             =   1605
         Width           =   375
      End
   End
   Begin VB.PictureBox picDefaultPlace 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   2040
      ScaleHeight     =   1095
      ScaleWidth      =   1335
      TabIndex        =   94
      Top             =   720
      Visible         =   0   'False
      Width           =   1335
      Begin VB.TextBox txtBirthPlace 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1740
         MaxLength       =   50
         TabIndex        =   75
         Text            =   "Panadura, Sri Lanka"
         Top             =   480
         Width           =   3375
      End
      Begin VB.CommandButton cmdAtlas 
         Caption         =   "Atlas"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5280
         TabIndex        =   76
         Top             =   480
         Width           =   975
      End
      Begin VB.ComboBox cmbTzPM 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmSettings.frx":1CFA
         Left            =   1740
         List            =   "frmSettings.frx":1D04
         TabIndex        =   77
         Text            =   "+"
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cmbTzHH 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmSettings.frx":1D0E
         Left            =   2580
         List            =   "frmSettings.frx":1D36
         TabIndex        =   78
         Text            =   "05"
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cmbTzMM 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmSettings.frx":1D60
         Left            =   3420
         List            =   "frmSettings.frx":1E18
         TabIndex        =   79
         Text            =   "30"
         Top             =   960
         Width           =   735
      End
      Begin VB.ComboBox cmbLonDeg 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmSettings.frx":1F0C
         Left            =   1740
         List            =   "frmSettings.frx":212C
         TabIndex        =   80
         Text            =   "079"
         Top             =   1320
         Width           =   735
      End
      Begin VB.ComboBox cmbLonMin 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmSettings.frx":24B4
         Left            =   2580
         List            =   "frmSettings.frx":256C
         TabIndex        =   81
         Text            =   "54"
         Top             =   1320
         Width           =   735
      End
      Begin VB.ComboBox cmbLonEW 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmSettings.frx":2660
         Left            =   3420
         List            =   "frmSettings.frx":266A
         TabIndex        =   82
         Text            =   "E"
         Top             =   1320
         Width           =   735
      End
      Begin VB.ComboBox cmbLatDeg 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmSettings.frx":2674
         Left            =   1740
         List            =   "frmSettings.frx":2786
         TabIndex        =   83
         Text            =   "06"
         Top             =   1680
         Width           =   735
      End
      Begin VB.ComboBox cmbLatMin 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmSettings.frx":28F2
         Left            =   2580
         List            =   "frmSettings.frx":29AA
         TabIndex        =   84
         Text            =   "39"
         Top             =   1680
         Width           =   735
      End
      Begin VB.ComboBox cmbLatNS 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         ItemData        =   "frmSettings.frx":2A9E
         Left            =   3420
         List            =   "frmSettings.frx":2AA8
         TabIndex        =   85
         Text            =   "N"
         Top             =   1680
         Width           =   735
      End
      Begin VB.Label Label18 
         Caption         =   "Note : No time correction (DST/WT) for default location."
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
         Index           =   0
         Left            =   120
         TabIndex        =   131
         Top             =   120
         Width           =   4815
      End
      Begin VB.Line Line3 
         X1              =   120
         X2              =   6240
         Y1              =   2160
         Y2              =   2160
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "Default Location :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   130
         Top             =   540
         Width           =   1275
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "Time Zone (HH:MM) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   129
         Top             =   1020
         Width           =   1515
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         Caption         =   "Longitude (DDD:MM) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   128
         Top             =   1380
         Width           =   1590
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "Latitude (DD:MM) :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   120
         TabIndex        =   127
         Top             =   1740
         Width           =   1365
      End
      Begin VB.Label Label18 
         Caption         =   $"frmSettings.frx":2AB2
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Index           =   1
         Left            =   120
         TabIndex        =   126
         Top             =   2280
         Width           =   6135
      End
   End
   Begin VB.PictureBox picAspects 
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   1560
      ScaleHeight     =   855
      ScaleWidth      =   975
      TabIndex        =   93
      Top             =   2880
      Visible         =   0   'False
      Width           =   975
      Begin VB.TextBox txApp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   17
         Left            =   5280
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   73
         Text            =   "0"
         ToolTipText     =   "This value Cannot be changed"
         Top             =   4080
         Width           =   495
      End
      Begin VB.TextBox txApp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   5280
         MaxLength       =   4
         TabIndex        =   70
         Text            =   "1"
         Top             =   3720
         Width           =   495
      End
      Begin VB.TextBox txApp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   5280
         MaxLength       =   4
         TabIndex        =   67
         Text            =   "1"
         Top             =   3360
         Width           =   495
      End
      Begin VB.TextBox txApp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   5280
         MaxLength       =   4
         TabIndex        =   64
         Text            =   "2"
         Top             =   3000
         Width           =   495
      End
      Begin VB.TextBox txApp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   5280
         MaxLength       =   4
         TabIndex        =   61
         Text            =   "2"
         Top             =   2640
         Width           =   495
      End
      Begin VB.TextBox txApp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   5280
         MaxLength       =   4
         TabIndex        =   58
         Text            =   "1"
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox txApp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   5280
         MaxLength       =   4
         TabIndex        =   55
         Text            =   "2"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox txApp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   5280
         MaxLength       =   4
         TabIndex        =   52
         Text            =   "2"
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txApp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   5280
         MaxLength       =   4
         TabIndex        =   49
         Text            =   "2"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txApp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   43
         Text            =   "3"
         Top             =   3720
         Width           =   495
      End
      Begin VB.TextBox txApp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   46
         Text            =   "2"
         Top             =   4080
         Width           =   495
      End
      Begin VB.TextBox txApp 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   40
         Text            =   "3"
         Top             =   3360
         Width           =   495
      End
      Begin VB.TextBox txApp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   37
         Text            =   "6"
         Top             =   3000
         Width           =   495
      End
      Begin VB.TextBox txApp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   34
         Text            =   "6"
         Top             =   2640
         Width           =   495
      End
      Begin VB.TextBox txApp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   31
         Text            =   "2"
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox txApp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   28
         Text            =   "6"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox txApp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   22
         Text            =   "10"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txApp 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2040
         MaxLength       =   4
         TabIndex        =   25
         Text            =   "8"
         Top             =   1560
         Width           =   495
      End
      Begin VB.CommandButton cmdOptions 
         Caption         =   "<< Options"
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
         Left            =   5040
         TabIndex        =   20
         Top             =   120
         Width           =   1215
      End
      Begin VB.CheckBox chkAsp 
         Caption         =   "126 Degrees (126)------"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   17
         Left            =   3240
         TabIndex        =   72
         Top             =   4080
         Width           =   2055
      End
      Begin VB.CheckBox chkAsp 
         Caption         =   "162 Degrees (162)------"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   16
         Left            =   3240
         TabIndex        =   69
         Top             =   3720
         Width           =   2055
      End
      Begin VB.CheckBox chkAsp 
         Caption         =   "54 Degrees (54)---------"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   15
         Left            =   3240
         TabIndex        =   66
         Top             =   3360
         Width           =   2055
      End
      Begin VB.CheckBox chkAsp 
         Caption         =   "Tredecile (108)-----------"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   14
         Left            =   3240
         TabIndex        =   63
         Top             =   3000
         Width           =   2055
      End
      Begin VB.CheckBox chkAsp 
         Caption         =   "Quindecile (24)-----------"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   12
         Left            =   3240
         TabIndex        =   57
         Top             =   2280
         Width           =   2055
      End
      Begin VB.CheckBox chkAsp 
         Caption         =   "Vigintile (18)--------------"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   11
         Left            =   3240
         TabIndex        =   54
         Top             =   1920
         Width           =   2055
      End
      Begin VB.CheckBox chkAsp 
         Caption         =   "Semi-Sextile (30)---------"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   10
         Left            =   3240
         TabIndex        =   51
         Top             =   1560
         Width           =   2055
      End
      Begin VB.CheckBox chkAsp 
         Caption         =   "Semi-Square (45)--------"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   9
         Left            =   3240
         TabIndex        =   48
         Top             =   1200
         Width           =   2055
      End
      Begin VB.CheckBox chkAsp 
         Caption         =   "Decile Semi-Quintile (36)"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   3240
         TabIndex        =   60
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox txSep 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   7
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   44
         Text            =   "3"
         Top             =   3720
         Width           =   495
      End
      Begin VB.TextBox txSep 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   8
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   47
         Text            =   "2"
         Top             =   4080
         Width           =   495
      End
      Begin VB.TextBox txSep 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   6
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   41
         Text            =   "3"
         Top             =   3360
         Width           =   495
      End
      Begin VB.TextBox txSep 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   5
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   38
         Text            =   "6"
         Top             =   3000
         Width           =   495
      End
      Begin VB.TextBox txSep 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   4
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   35
         Text            =   "6"
         Top             =   2640
         Width           =   495
      End
      Begin VB.TextBox txSep 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   32
         Text            =   "3"
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox txSep 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   29
         Text            =   "6"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox txSep 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   26
         Text            =   "8"
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txSep 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   2520
         MaxLength       =   4
         TabIndex        =   23
         Text            =   "10"
         Top             =   1200
         Width           =   495
      End
      Begin VB.TextBox txSep 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   17
         Left            =   5760
         MaxLength       =   4
         TabIndex        =   74
         Text            =   "4"
         Top             =   4080
         Width           =   495
      End
      Begin VB.TextBox txSep 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   16
         Left            =   5760
         MaxLength       =   4
         TabIndex        =   71
         Text            =   "1"
         Top             =   3720
         Width           =   495
      End
      Begin VB.TextBox txSep 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   15
         Left            =   5760
         MaxLength       =   4
         TabIndex        =   68
         Text            =   "1"
         Top             =   3360
         Width           =   495
      End
      Begin VB.TextBox txSep 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   14
         Left            =   5760
         MaxLength       =   4
         TabIndex        =   65
         Text            =   "2"
         Top             =   3000
         Width           =   495
      End
      Begin VB.TextBox txSep 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   13
         Left            =   5760
         MaxLength       =   4
         TabIndex        =   62
         Text            =   "2"
         Top             =   2640
         Width           =   495
      End
      Begin VB.TextBox txSep 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   12
         Left            =   5760
         MaxLength       =   4
         TabIndex        =   59
         Text            =   "1"
         Top             =   2280
         Width           =   495
      End
      Begin VB.TextBox txSep 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   11
         Left            =   5760
         MaxLength       =   4
         TabIndex        =   56
         Text            =   "2"
         Top             =   1920
         Width           =   495
      End
      Begin VB.TextBox txSep 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   10
         Left            =   5760
         MaxLength       =   4
         TabIndex        =   53
         Text            =   "2"
         Top             =   1560
         Width           =   495
      End
      Begin VB.TextBox txSep 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   9
         Left            =   5760
         MaxLength       =   4
         TabIndex        =   50
         Text            =   "2"
         Top             =   1200
         Width           =   495
      End
      Begin VB.CheckBox chkAsp 
         Caption         =   "Conjuction (0)----------"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   21
         Top             =   1200
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkAsp 
         Caption         =   "Opposition (180)-------"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   24
         Top             =   1560
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkAsp 
         Caption         =   "Trine (120)--------------"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   27
         Top             =   1920
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkAsp 
         Caption         =   "Quincunx (150)--------"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   30
         Top             =   2280
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkAsp 
         Caption         =   "Square (90)-------------"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   33
         Top             =   2640
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkAsp 
         Caption         =   "Sextile (60)-------------"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   255
         Index           =   5
         Left            =   120
         TabIndex        =   36
         Top             =   3000
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.CheckBox chkAsp 
         Caption         =   "Biquintile (144)---------"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   39
         Top             =   3360
         Width           =   1935
      End
      Begin VB.CheckBox chkAsp 
         Caption         =   "Sesquiquadrate (135)-"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   42
         Top             =   3720
         Width           =   1935
      End
      Begin VB.CheckBox chkAsp 
         Caption         =   "Quintile (72)------------"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   45
         Top             =   4080
         Width           =   1935
      End
      Begin VB.Line Line2 
         X1              =   3120
         X2              =   3120
         Y1              =   1200
         Y2              =   4320
      End
      Begin VB.Label Label6 
         AutoSize        =   -1  'True
         Caption         =   "Separ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5760
         TabIndex        =   125
         Top             =   960
         Width           =   420
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "Apply"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   5280
         TabIndex        =   124
         Top             =   960
         Width           =   405
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "Separ"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2520
         TabIndex        =   123
         Top             =   960
         Width           =   420
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Apply"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2040
         TabIndex        =   122
         Top             =   960
         Width           =   405
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Orb"
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
         Left            =   5640
         TabIndex        =   121
         Top             =   720
         Width           =   300
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Orb"
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
         Left            =   2400
         TabIndex        =   120
         Top             =   720
         Width           =   300
      End
      Begin VB.Line Line1 
         X1              =   120
         X2              =   6240
         Y1              =   600
         Y2              =   600
      End
   End
   Begin ComctlLib.TabStrip tStrip1 
      Height          =   375
      Left            =   0
      TabIndex        =   96
      Top             =   120
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   661
      TabWidthStyle   =   2
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   4
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Planet Names"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Aspects"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Default Location"
            Key             =   ""
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Other"
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
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
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
      Left            =   5040
      TabIndex        =   91
      Top             =   5040
      Width           =   1215
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "&Save"
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
      Left            =   3720
      TabIndex        =   90
      Top             =   5040
      Width           =   1215
   End
   Begin VB.Line Line7 
      BorderWidth     =   3
      X1              =   120
      X2              =   6240
      Y1              =   4920
      Y2              =   4920
   End
   Begin VB.Menu mnuOptAsp 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu mnuVeryStrong 
         Caption         =   "Select Very Strong Aspects"
      End
      Begin VB.Menu mnuVSAspects 
         Caption         =   "Select Very Strong And Strong Aspects"
      End
      Begin VB.Menu mnuSelectAll 
         Caption         =   "Select All"
      End
      Begin VB.Menu sep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuSet8 
         Caption         =   "Set Orb = 8"
      End
      Begin VB.Menu mnuSetOrb333 
         Caption         =   "Set Orb = 3.33 (4 Step Theory)"
      End
      Begin VB.Menu mnuDefaultOrb 
         Caption         =   "Set Default Orb"
      End
   End
   Begin VB.Menu mnuOptGeneral 
      Caption         =   "Options"
      Visible         =   0   'False
      Begin VB.Menu optDefEng 
         Caption         =   "Set Default English Names"
      End
      Begin VB.Menu optDefTra 
         Caption         =   "Set Default Traditional Names"
      End
      Begin VB.Menu sep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuClearAll 
         Caption         =   "Clear All"
      End
   End
End
Attribute VB_Name = "frmSettings"
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

Private Sub chkAsp_Click(Index As Integer)
    Dim i As Byte
    For i = 0 To 17
        If chkAsp(i).Value = 0 Then
            txApp(i).Enabled = False
            txSep(i).Enabled = False
        Else
            txApp(i).Enabled = True
            txSep(i).Enabled = True
        End If
    Next

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


Private Sub cmbTzHH_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub


Private Sub cmbTzMM_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "1234567890")
End Sub


Private Sub cmbTzPM_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "+-")
End Sub

Private Sub cmdAtlas_Click()
    MsgBox "Sorry, currently not implemented. Please enter the required data manually.", vbOKOnly, "Sorry..."
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOptGen_Click()
    PopupMenu mnuOptGeneral
End Sub

Private Sub cmdOptions_Click()
    PopupMenu mnuOptAsp
End Sub

Private Sub cmdSave_Click()

    '/----------------------------------------------------------------------------/
    'Aspects Settings
    Dim exSet As String, i As Byte
    For i = 0 To 17
        exSet = exSet + CStr(chkAsp(i).Value) + " " + CStr(Val(txApp(i).Text)) + " " + CStr(Val(txSep(i).Text)) + "_"
    Next
    Call hp_CreateDatFile(exSet, "aspdet", App.Path + "\Sett")

    '/----------------------------------------------------------------------------/
    'Default Place Settings
    Dim exPl As String
    exPl = txtBirthPlace.Text + "_" + cmbTzPM.Text + "_" + cmbTzHH.Text + "_" + cmbTzMM.Text + "_" + cmbLonDeg.Text + "_" + cmbLonMin.Text + "_" + cmbLonEW.Text + "_" + cmbLatDeg.Text + "_" + cmbLatMin.Text + "_" + cmbLatNS.Text
    Call hp_CreateDatFile(exPl, "depdet", App.Path + "\Sett")
    '/----------------------------------------------------------------------------/
    'Planet Names Settings

    Dim plNme As String
    Dim i1 As Byte
    For i1 = 0 To 18
        plNme = plNme + txtP(i1).Text + "_"
    Next
    Call hp_CreateDatFile(plNme, "nmedet", App.Path + "\Sett")
    
    '/----------------------------------------------------------------------------/
    'Geocentric and True
    Dim geoTrue As String
    geoTrue = CStr(chkGeo.Value) + "_" + CStr(chkTrue.Value)
    Call hp_CreateDatFile(geoTrue, "geotrue", App.Path + "\Sett")
    
    '/----------------------------------------------------------------------------/
    'Personal Details
    Dim astrPD As String
    Dim i2 As Byte
    For i2 = 0 To 3
        astrPD = astrPD + txtPD(i2).Text + "_"
    Next
    Call hp_CreateDatFile(astrPD, "asrdet", App.Path + "\Sett")

    '/----------------------------------------------------------------------------/

    If MsgBox("All Changes Are Saved. Application Must Be Restarted To Apply Settings. Restart Now ?", vbYesNo, "Settings") = vbYes Then
        End
    Else
        Unload Me
    End If

End Sub

Private Sub Form_Load()

'Size
    Me.Height = 5985
    Me.Width = 6475

    'Center Screen
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2

    'PictureBox Positions
    picPlanetNames.Height = 4455
    picPlanetNames.Width = 6375
    picPlanetNames.Top = 480
    picPlanetNames.Left = 0

    picAspects.Height = 4455
    picAspects.Width = 6375
    picAspects.Top = 480
    picAspects.Left = 0

    picDefaultPlace.Height = 4455
    picDefaultPlace.Width = 6375
    picDefaultPlace.Top = 480
    picDefaultPlace.Left = 0

    picOther.Height = 4455
    picOther.Width = 6375
    picOther.Top = 480
    picOther.Left = 0



    'Load Aspect Settings
    Dim aspLoaded() As String
    aspLoaded() = hp_ReadDatFile(App.Path + "\Sett\aspdet.dat", "_")


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

    '0
    chkAsp(0).Value = CInt(dataToAdd0(0))
    txApp(0).Text = dataToAdd0(1)
    txSep(0).Text = dataToAdd0(2)

    '1
    chkAsp(1).Value = CInt(dataToAdd1(0))
    txApp(1).Text = dataToAdd1(1)
    txSep(1).Text = dataToAdd1(2)

    '2
    chkAsp(2).Value = CInt(dataToAdd2(0))
    txApp(2).Text = dataToAdd2(1)
    txSep(2).Text = dataToAdd2(2)

    '3
    chkAsp(3).Value = CInt(dataToAdd3(0))
    txApp(3).Text = dataToAdd3(1)
    txSep(3).Text = dataToAdd3(2)

    '4
    chkAsp(4).Value = CInt(dataToAdd4(0))
    txApp(4).Text = dataToAdd4(1)
    txSep(4).Text = dataToAdd4(2)

    '5
    chkAsp(5).Value = CInt(dataToAdd5(0))
    txApp(5).Text = dataToAdd5(1)
    txSep(5).Text = dataToAdd5(2)

    '6
    chkAsp(6).Value = CInt(dataToAdd6(0))
    txApp(6).Text = dataToAdd6(1)
    txSep(6).Text = dataToAdd6(2)

    '7
    chkAsp(7).Value = CInt(dataToAdd7(0))
    txApp(7).Text = dataToAdd7(1)
    txSep(7).Text = dataToAdd7(2)

    '8
    chkAsp(8).Value = CInt(dataToAdd8(0))
    txApp(8).Text = dataToAdd8(1)
    txSep(8).Text = dataToAdd8(2)

    '9
    chkAsp(9).Value = CInt(dataToAdd9(0))
    txApp(9).Text = dataToAdd9(1)
    txSep(9).Text = dataToAdd9(2)

    '10
    chkAsp(10).Value = CInt(dataToAdd10(0))
    txApp(10).Text = dataToAdd10(1)
    txSep(10).Text = dataToAdd10(2)

    '11
    chkAsp(11).Value = CInt(dataToAdd11(0))
    txApp(11).Text = dataToAdd11(1)
    txSep(11).Text = dataToAdd11(2)

    '12
    chkAsp(12).Value = CInt(dataToAdd12(0))
    txApp(12).Text = dataToAdd12(1)
    txSep(12).Text = dataToAdd12(2)

    '13
    chkAsp(13).Value = CInt(dataToAdd13(0))
    txApp(13).Text = dataToAdd13(1)
    txSep(13).Text = dataToAdd13(2)

    '14
    chkAsp(14).Value = CInt(dataToAdd14(0))
    txApp(14).Text = dataToAdd14(1)
    txSep(14).Text = dataToAdd14(2)

    '15
    chkAsp(15).Value = CInt(dataToAdd15(0))
    txApp(15).Text = dataToAdd15(1)
    txSep(15).Text = dataToAdd15(2)

    '16
    chkAsp(16).Value = CInt(dataToAdd16(0))
    txApp(16).Text = dataToAdd16(1)
    txSep(16).Text = dataToAdd16(2)

    '17
    chkAsp(17).Value = CInt(dataToAdd17(0))
    txApp(17).Text = dataToAdd17(1)
    txSep(17).Text = dataToAdd17(2)


    '/----------------------------------------------------------------------------/
    'Load Default place settings

    Dim depLoaded() As String
    depLoaded() = hp_ReadDatFile(App.Path + "\Sett\depdet.dat", "_")

    txtBirthPlace.Text = depLoaded(0)
    cmbTzPM.Text = depLoaded(1)
    cmbTzHH.Text = depLoaded(2)
    cmbTzMM.Text = depLoaded(3)
    cmbLonDeg.Text = depLoaded(4)
    cmbLonMin.Text = depLoaded(5)
    cmbLonEW.Text = depLoaded(6)
    cmbLatDeg.Text = depLoaded(7)
    cmbLatMin.Text = depLoaded(8)
    cmbLatNS.Text = depLoaded(9)

    '/----------------------------------------------------------------------------/
    'Load Name settings
    Dim loadedPlntNames() As String
    loadedPlntNames() = hp_ReadDatFile(App.Path + "\Sett\nmedet.dat", "_")
    Dim i1 As Byte
    For i1 = 0 To 18
        txtP(i1).Text = loadedPlntNames(i1)
    Next
    '/----------------------------------------------------------------------------/
    'Load Geo True
    Dim geoTrue() As String
    geoTrue() = hp_ReadDatFile(App.Path + "\Sett\geotrue.dat", "_")
    chkGeo.Value = CByte(geoTrue(0))
    chkTrue.Value = CByte(geoTrue(1))
    
    
    
    
    '/----------------------------------------------------------------------------/
    'Load Pesonal details
    Dim loadedPD() As String
    loadedPD() = hp_ReadDatFile(App.Path + "\Sett\asrdet.dat", "_")
    Dim i2 As Byte
    For i2 = 0 To 3
        txtPD(i2).Text = loadedPD(i2)
    Next

    '/----------------------------------------------------------------------------/
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
End Sub

Private Sub mnuClearAll_Click()
    txtP(0).Text = vbNullString
    txtP(1).Text = vbNullString
    txtP(2).Text = vbNullString    '
    txtP(3).Text = vbNullString    '
    txtP(4).Text = vbNullString    '
    txtP(5).Text = vbNullString    '
    txtP(6).Text = vbNullString    '
    txtP(7).Text = vbNullString
    txtP(8).Text = vbNullString
    txtP(9).Text = vbNullString    '
    txtP(10).Text = vbNullString    '
    txtP(11).Text = vbNullString

    txtP(12).Text = vbNullString
    txtP(13).Text = vbNullString
    txtP(14).Text = vbNullString
    txtP(15).Text = vbNullString
    txtP(16).Text = vbNullString
    txtP(17).Text = vbNullString
    txtP(18).Text = vbNullString

End Sub

Private Sub mnuDefaultOrb_Click()

    txApp(0).Text = "10": txSep(0).Text = "10"
    txApp(1).Text = "8": txSep(1).Text = "8"
    txApp(2).Text = "6": txSep(2).Text = "6"
    txApp(3).Text = "2": txSep(3).Text = "3"
    txApp(4).Text = "6": txSep(4).Text = "6"
    txApp(5).Text = "6": txSep(5).Text = "6"
    txApp(6).Text = "3": txSep(6).Text = "3"
    txApp(7).Text = "3": txSep(7).Text = "3"
    txApp(8).Text = "2": txSep(8).Text = "2"

    txApp(9).Text = "2": txSep(9).Text = "2"
    txApp(10).Text = "2": txSep(10).Text = "2"
    txApp(11).Text = "2": txSep(11).Text = "2"
    txApp(12).Text = "1": txSep(12).Text = "1"
    txApp(13).Text = "2": txSep(13).Text = "2"
    txApp(14).Text = "2": txSep(14).Text = "2"
    txApp(15).Text = "1": txSep(15).Text = "1"
    txApp(16).Text = "1": txSep(16).Text = "1"
    'txApp(17).Text = "10"
    txSep(17).Text = "4"

End Sub

Private Sub mnuSelectAll_Click()

    Dim i As Byte
    For i = 0 To 17
        chkAsp(i).Value = 1
    Next

End Sub

Private Sub mnuSet8_Click()
    Dim i As Byte
    For i = 0 To 17
        txApp(i).Text = "8"
        txSep(i).Text = "8"
    Next
    txApp(17).Text = "0"
End Sub

Private Sub mnuSetOrb333_Click()

    Dim i As Byte
    For i = 0 To 17
        txApp(i).Text = "3.33333333333333"
        txSep(i).Text = "3.33333333333333"
    Next
    txApp(17).Text = "0"
End Sub

Private Sub mnuVeryStrong_Click()

    Dim i As Byte, j As Byte
    For i = 0 To 5
        chkAsp(i).Value = 1
    Next

    For j = 6 To 17
        chkAsp(j).Value = 0
    Next

End Sub

Private Sub mnuVSAspects_Click()

    Dim i As Byte, j As Byte
    For i = 0 To 8
        chkAsp(i).Value = 1
    Next

    For j = 9 To 17
        chkAsp(j).Value = 0
    Next

End Sub

Private Sub optDefEng_Click()

    txtP(0).Text = "Sun"
    txtP(1).Text = "Moo"
    txtP(2).Text = "Mar"    '
    txtP(3).Text = "Mer"    '
    txtP(4).Text = "Jup"    '
    txtP(5).Text = "Ven"    '
    txtP(6).Text = "Sat"    '
    txtP(7).Text = "Rah"
    txtP(8).Text = "Ket"
    txtP(9).Text = "Ura"    '
    txtP(10).Text = "Nep"    '
    txtP(11).Text = "For"

    txtP(12).Text = "Ma®"
    txtP(13).Text = "Me®"
    txtP(14).Text = "Ju®"
    txtP(15).Text = "Ve®"
    txtP(16).Text = "Sa®"
    txtP(17).Text = "Ur®"
    txtP(18).Text = "Ne®"

End Sub

Private Sub optDefTra_Click()

    txtP(0).Text = "Rav"
    txtP(1).Text = "Cha"
    txtP(2).Text = "Kuj"    '
    txtP(3).Text = "Bud"    '
    txtP(4).Text = "Gur"    '
    txtP(5).Text = "Suk"    '
    txtP(6).Text = "San"    '
    txtP(7).Text = "Rah"
    txtP(8).Text = "Ket"
    txtP(9).Text = "Ura"    '
    txtP(10).Text = "Nep"    '
    txtP(11).Text = "For"

    txtP(12).Text = "Ku®"
    txtP(13).Text = "Bu®"
    txtP(14).Text = "Gu®"
    txtP(15).Text = "Su®"
    txtP(16).Text = "Sa®"
    txtP(17).Text = "Ur®"
    txtP(18).Text = "Ne®"

End Sub


Private Sub tStrip1_Click()

    If tStrip1.SelectedItem.Index = 1 Then
        picPlanetNames.Visible = True
        picAspects.Visible = False
        picDefaultPlace.Visible = False
        picOther.Visible = False
    End If

    If tStrip1.SelectedItem.Index = 2 Then
        picPlanetNames.Visible = False
        picAspects.Visible = True
        picDefaultPlace.Visible = False
        picOther.Visible = False
    End If

    If tStrip1.SelectedItem.Index = 3 Then
        picPlanetNames.Visible = False
        picAspects.Visible = False
        picDefaultPlace.Visible = True
        picOther.Visible = False
    End If

    If tStrip1.SelectedItem.Index = 4 Then
        picPlanetNames.Visible = False
        picAspects.Visible = False
        picDefaultPlace.Visible = False
        picOther.Visible = True
    End If

End Sub


Private Sub txApp_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "1234567890.")
End Sub

Private Sub txSep_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "1234567890.")
End Sub

Private Sub txtBirthPlace_KeyPress(KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890,.")
End Sub

Private Sub txtP_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890®*@#$%^&*")
End Sub

Private Sub txtPD_KeyPress(Index As Integer, KeyAscii As Integer)
    KeyAscii = hp_KeyFilter(KeyAscii, "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz1234567890,.")
End Sub


