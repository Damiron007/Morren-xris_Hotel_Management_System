VERSION 5.00
Begin VB.Form frmSplash 
   BackColor       =   &H00800000&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   8910
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   15210
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmSplash.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8910
   ScaleWidth      =   15210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Height          =   8835
      Left            =   -480
      TabIndex        =   0
      Top             =   0
      Width           =   15570
      Begin VB.Timer Timer2 
         Left            =   360
         Top             =   8640
      End
      Begin VB.Timer Timer1 
         Interval        =   3000
         Left            =   120
         Top             =   7200
      End
      Begin VB.Label dtt 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   375
         Left            =   840
         TabIndex        =   5
         Top             =   6120
         Width           =   4575
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "www.morris-xrishotels.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   3720
         TabIndex        =   4
         Top             =   7920
         Width           =   3615
      End
      Begin VB.Image Image1 
         Height          =   3795
         Left            =   240
         Picture         =   "frmSplash.frx":000C
         Top             =   120
         Width           =   7740
      End
      Begin VB.Image Image3 
         Height          =   1350
         Left            =   7080
         Picture         =   "frmSplash.frx":881B
         Top             =   5520
         Width           =   1350
      End
      Begin VB.Image Image2 
         Height          =   3270
         Left            =   9120
         Picture         =   "frmSplash.frx":92AB
         Top             =   120
         Width           =   5685
      End
      Begin VB.Label lblHotelName 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "WELCOME TO MORREN-XRIS SUITES AND GARDEN"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   26.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   1800
         TabIndex        =   3
         Top             =   3840
         Width           =   13005
      End
      Begin VB.Label lblVersion 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "Version 1.0"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0000FFFF&
         Height          =   285
         Left            =   10680
         TabIndex        =   2
         Top             =   8040
         Width           =   1275
      End
      Begin VB.Label LblProjectname 
         BackColor       =   &H00800000&
         Caption         =   "                      HOTEL MANAGEMENT SYSTEM"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   615
         Left            =   720
         TabIndex        =   1
         Top             =   4680
         Width           =   12255
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Timer1_Timer()
Unload Me
frmLogin.Show
End Sub

Private Sub Timer2_Timer()
dtt = Now
Frame1.BackColor = vbWhite
End Sub
