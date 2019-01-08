VERSION 5.00
Begin VB.Form frmMenu 
   Caption         =   "Welcome to Morren-xris suites and Garden"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14760
   LinkTopic       =   "Form4"
   ScaleHeight     =   8430
   ScaleWidth      =   14760
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "Main Menu"
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
      Height          =   8890
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   14775
      Begin VB.OptionButton optHelp 
         BackColor       =   &H00800000&
         Caption         =   "Help"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   500
         Left            =   360
         TabIndex        =   3
         Top             =   3960
         Width           =   2500
      End
      Begin VB.OptionButton optBar 
         BackColor       =   &H00800000&
         Caption         =   "Bar kitchen Order"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   500
         Left            =   360
         TabIndex        =   2
         Top             =   3000
         Width           =   2865
      End
      Begin VB.OptionButton optCaptin_Order 
         BackColor       =   &H00800000&
         Caption         =   "Make Captain Order"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   500
         Left            =   360
         TabIndex        =   1
         Top             =   2040
         Width           =   3345
      End
      Begin VB.OptionButton optExit 
         BackColor       =   &H00800000&
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   500
         Left            =   360
         TabIndex        =   5
         Top             =   4920
         Width           =   2500
      End
      Begin VB.OptionButton optReservation 
         BackColor       =   &H00800000&
         Caption         =   "Make reservation"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   15.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   500
         Left            =   360
         TabIndex        =   0
         Top             =   960
         Width           =   3225
      End
      Begin VB.Image Image4 
         Height          =   1350
         Left            =   5040
         Picture         =   "frmMenu.frx":0000
         Top             =   600
         Width           =   1350
      End
      Begin VB.Image Image3 
         Height          =   2955
         Left            =   9240
         Picture         =   "frmMenu.frx":0A90
         Top             =   5280
         Width           =   4605
      End
      Begin VB.Image Image2 
         Height          =   2610
         Left            =   4920
         Picture         =   "frmMenu.frx":4FB4
         Top             =   3720
         Width           =   4350
      End
      Begin VB.Image Image1 
         Height          =   3600
         Left            =   7080
         Picture         =   "frmMenu.frx":7413
         Top             =   480
         Width           =   6180
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub optBar_Click()
Me.Hide
frmBar_kitchen.Show
End Sub

Private Sub optCaptin_Order_Click()
Me.Hide
frmCaptinOrder.Show
End Sub

Private Sub optExit_Click()
Dim intresponse As Integer
intresponse = MsgBox("Do you want to Quit", vbYesNo + vbInformation, "You Quit")
If intresponse = vbYes Then
End
End If
End Sub

Private Sub optHelp_Click()
Me.Hide
frmHelp.Show
End Sub

Private Sub optHome_Click()
Me.Hide
frmLogin.Show
End Sub

Private Sub optReservation_Click()
Me.Hide
frmReservation.Show
End Sub

Private Sub optSearch_Click()
Me.Hide
frmSearch.Show
End Sub

Private Sub optView_List_Click()
Me.Hide
frmRecrd_list.Show
End Sub


Private Sub tmrScroll_Timer()
lblWelcome = "  " & lblWelcome
End Sub
