VERSION 5.00
Begin VB.Form frmHelp 
   BackColor       =   &H00800000&
   Caption         =   "Form1"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   14760
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "HELP GUIDE"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   8900
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   14775
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   480
         TabIndex        =   2
         Top             =   6000
         Width           =   1020
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "Call 07036883184 for help or support "
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   975
         Left            =   600
         TabIndex        =   3
         Top             =   3720
         Width           =   5295
      End
      Begin VB.Image Image1 
         Height          =   1350
         Left            =   7320
         Picture         =   "frmHelp.frx":0000
         Top             =   2160
         Width           =   1350
      End
      Begin VB.Label lblHelp 
         BackColor       =   &H00800000&
         Caption         =   $"frmHelp.frx":0A90
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         Top             =   720
         Width           =   14055
      End
   End
End
Attribute VB_Name = "frmHelp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBack_Click()
Me.Hide
frmMenu.Show
End Sub

