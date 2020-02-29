VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About KP New Astro"
   ClientHeight    =   2625
   ClientLeft      =   3120
   ClientTop       =   3915
   ClientWidth     =   4800
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2625
   ScaleWidth      =   4800
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
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
      Left            =   3360
      TabIndex        =   0
      Top             =   2160
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "(version 2 or later)"
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
      Left            =   120
      TabIndex        =   8
      Top             =   1680
      Width           =   1455
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   4080
      Picture         =   "frmAbout.frx":1CFA
      Stretch         =   -1  'True
      Top             =   120
      Width           =   600
   End
   Begin VB.Label Label9 
      Caption         =   "Email : kpnewastro@gmail.com"
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
      Left            =   120
      TabIndex        =   7
      Top             =   2280
      Width           =   2415
   End
   Begin VB.Label Label7 
      Caption         =   "Web : kpnewastro.blogspot.com"
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
      Left            =   120
      TabIndex        =   6
      Top             =   2040
      Width           =   2415
   End
   Begin VB.Label Label8 
      Caption         =   "KP New Astro is released under the  GNU General Public License."
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
      Left            =   120
      TabIndex        =   5
      Top             =   1440
      Width           =   4695
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "KP New Astro. Copyright (C) 2009-2010 JSW"
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
      TabIndex        =   4
      Top             =   600
      Width           =   3240
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "Version"
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
      TabIndex        =   3
      Top             =   840
      Width           =   525
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "KP New Astro"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1403
      TabIndex        =   2
      Top             =   120
      Width           =   1965
   End
   Begin VB.Label lblRelDate 
      AutoSize        =   -1  'True
      Caption         =   "Released Date : 05/12/2010"
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
      TabIndex        =   1
      Top             =   1080
      Width           =   2040
   End
End
Attribute VB_Name = "frmAbout"
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

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Me.Top = (Screen.Height - Me.Height) / 2
    Me.Left = (Screen.Width - Me.Width) / 2
    
    lblVersion.Caption = "Version : " + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision) + " beta"
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    frmMain.Enabled = True
End Sub
