VERSION 5.00
Begin VB.Form frmReceipt2 
   Caption         =   "Form1"
   ClientHeight    =   5835
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   6255
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Morren xris suites and Garden "
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   6400
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   6500
      Begin VB.CommandButton cmdPrint 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   5160
         TabIndex        =   1
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label lblClass_of_Room 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Class of Room"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   2640
         Width           =   1500
      End
      Begin VB.Label lblDuration 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Duration"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   4080
         Width           =   1500
      End
      Begin VB.Label lblDate 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1080
         Width           =   1500
      End
      Begin VB.Label lblCheckin 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Check in Time"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   3720
         Width           =   1500
      End
      Begin VB.Label lblCustomer_Name 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Customer Name"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1440
         Width           =   1500
      End
      Begin VB.Label lblCustomer_Address 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Customer Address"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   1920
         Width           =   1380
      End
      Begin VB.Label lblAmount 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Amount paid"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   20
         Top             =   3360
         Width           =   1500
      End
      Begin VB.Label lblPhone_number 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Phone Number"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   2280
         Width           =   1500
      End
      Begin VB.Label lblReceptionist 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Name of Receptionist"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   4440
         Width           =   1500
      End
      Begin VB.Label lblWelcome 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Thank you for patronizing us. Please visit us again"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   5520
         Width           =   3645
      End
      Begin VB.Label roomclass 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   1800
         TabIndex        =   16
         Top             =   2640
         Width           =   1500
      End
      Begin VB.Label duration 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   15
         Top             =   4080
         Width           =   1500
      End
      Begin VB.Label date1 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   14
         Top             =   1080
         Width           =   1500
      End
      Begin VB.Label checkin 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   13
         Top             =   3720
         Width           =   1500
      End
      Begin VB.Label customername 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   255
         Left            =   1800
         TabIndex        =   12
         Top             =   1440
         Width           =   2820
         WordWrap        =   -1  'True
      End
      Begin VB.Label customeraddress 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000009&
         Height          =   375
         Left            =   1800
         TabIndex        =   11
         Top             =   1800
         Width           =   2865
         WordWrap        =   -1  'True
      End
      Begin VB.Label amountpaid 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   10
         Top             =   3360
         Width           =   1500
      End
      Begin VB.Label phonenumber 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   9
         Top             =   2280
         Width           =   1500
      End
      Begin VB.Label receptionist 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   1800
         TabIndex        =   8
         Top             =   4440
         Width           =   2820
      End
      Begin VB.Label dtt 
         BackStyle       =   0  'Transparent
         Height          =   375
         Left            =   5160
         TabIndex        =   7
         Top             =   360
         Width           =   1125
      End
      Begin VB.Image Image1 
         Height          =   540
         Left            =   3600
         Picture         =   "frmReceipt2.frx":0000
         Top             =   360
         Width           =   540
      End
      Begin VB.Label lblHotelSign 
         BackColor       =   &H8000000E&
         Caption         =   "Staff Sign"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   5280
         Width           =   1815
      End
      Begin VB.Line Line1 
         X1              =   240
         X2              =   1560
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Label Label5 
         BackColor       =   &H8000000E&
         Caption         =   "Customer Sign"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   2400
         TabIndex        =   5
         Top             =   5280
         Width           =   1575
      End
      Begin VB.Line Line4 
         X1              =   2280
         X2              =   3840
         Y1              =   5160
         Y2              =   5160
      End
      Begin VB.Label lblRoomNo 
         BackColor       =   &H8000000E&
         Caption         =   "Room Number"
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label RoomNo 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   1800
         TabIndex        =   3
         Top             =   3000
         Width           =   1575
      End
      Begin VB.Label lblHotelAddress 
         BackColor       =   &H8000000E&
         Caption         =   "Blessed Anyi Lane, Amaokpala  08124733878, 08033099483            Customer receipt"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   600
         TabIndex        =   2
         Top             =   360
         Width           =   2775
         WordWrap        =   -1  'True
      End
   End
End
Attribute VB_Name = "frmReceipt2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdPrint_Click()
With Me
    .WindowState = 0
    .PrintForm
    Unload Me
End With

End Sub

