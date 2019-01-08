VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmCaptinOrder 
   BackColor       =   &H00FF0000&
   Caption         =   "Make Captain Order"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   14760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdReceipt 
      Caption         =   "Display Reciept"
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
      Left            =   7560
      TabIndex        =   21
      Top             =   2280
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.TextBox txtName_of_waiter 
      DataField       =   "Name_of_waitress"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4080
      TabIndex        =   4
      Top             =   6120
      Width           =   2535
   End
   Begin VB.TextBox txtAmount2 
      DataField       =   "Amount"
      DataSource      =   "Adodc1"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   4080
      TabIndex        =   3
      Top             =   5280
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "Make Captain Order"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   9120
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14655
      Begin VB.TextBox txtSerial_Number 
         DataField       =   "SerialNumber"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "0000"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   465
         Left            =   4080
         Locked          =   -1  'True
         TabIndex        =   32
         Top             =   3000
         Width           =   1545
      End
      Begin VB.CommandButton Command2 
         Caption         =   "View Record"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7920
         TabIndex        =   10
         Top             =   6720
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Print"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   7920
         TabIndex        =   5
         Top             =   3000
         Width           =   1095
      End
      Begin VB.TextBox txtOrderedFrom 
         DataField       =   "Ordered_From"
         DataSource      =   "Adodc1"
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
         Left            =   4080
         TabIndex        =   1
         Top             =   2160
         Width           =   2415
      End
      Begin VB.Timer tmrDate 
         Interval        =   1000
         Left            =   0
         Top             =   8160
      End
      Begin VB.CommandButton cmdSave 
         Caption         =   "Save"
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
         Left            =   7920
         TabIndex        =   6
         Top             =   3720
         Width           =   1215
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add New"
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
         Left            =   7920
         TabIndex        =   7
         Top             =   4440
         Width           =   1335
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   615
         Left            =   1680
         Top             =   6960
         Width           =   5175
         _ExtentX        =   9128
         _ExtentY        =   1085
         ConnectMode     =   0
         CursorLocation  =   3
         IsolationLevel  =   -1
         ConnectionTimeout=   15
         CommandTimeout  =   30
         CursorType      =   3
         LockType        =   3
         CommandType     =   2
         CursorOptions   =   0
         CacheSize       =   50
         MaxRecords      =   0
         BOFAction       =   0
         EOFAction       =   0
         ConnectStringType=   1
         Appearance      =   1
         BackColor       =   -2147483643
         ForeColor       =   -2147483640
         Orientation     =   0
         Enabled         =   -1
         Connect         =   $"frmCaptinOrder.frx":0000
         OLEDBString     =   $"frmCaptinOrder.frx":008B
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Captin_Order"
         Caption         =   ""
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _Version        =   393216
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
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
         Left            =   7920
         TabIndex        =   9
         Top             =   5880
         Width           =   1140
      End
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
         Left            =   7920
         TabIndex        =   8
         Top             =   5160
         Width           =   1140
      End
      Begin VB.TextBox txtDescription 
         DataField       =   "Description"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1140
         Left            =   4080
         TabIndex        =   2
         Top             =   3720
         Width           =   2775
      End
      Begin VB.Label lblOrderedFrom 
         BackColor       =   &H00800000&
         Caption         =   "Ordered From"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   495
         Left            =   600
         TabIndex        =   31
         Top             =   2160
         Width           =   2775
      End
      Begin VB.Label lblMonth2 
         BackColor       =   &H00800000&
         Caption         =   "March"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   4080
         TabIndex        =   30
         Top             =   840
         Width           =   735
      End
      Begin VB.Label lblNumber2 
         BackColor       =   &H00800000&
         Caption         =   "31"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   4800
         TabIndex        =   29
         Top             =   840
         Width           =   255
      End
      Begin VB.Label lblYear2 
         BackColor       =   &H00800000&
         Caption         =   "1998"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   375
         Left            =   5040
         TabIndex        =   28
         Top             =   840
         Width           =   615
      End
      Begin VB.Label lblTime2 
         BackColor       =   &H00800000&
         Caption         =   "00:00:00 PM"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   4080
         TabIndex        =   27
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Image Image1 
         Height          =   1350
         Left            =   7680
         Picture         =   "frmCaptinOrder.frx":0116
         Top             =   600
         Width           =   1350
      End
      Begin VB.Label Label9 
         BackColor       =   &H00800000&
         Caption         =   "Name of Waiter/Waitress"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   480
         TabIndex        =   20
         Top             =   6120
         Width           =   3135
      End
      Begin VB.Label Label7 
         BackColor       =   &H00800000&
         Caption         =   "Amount"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   405
         Left            =   480
         TabIndex        =   18
         Top             =   5280
         Width           =   2895
      End
      Begin VB.Label lblDateofOrder 
         BackColor       =   &H00800000&
         Caption         =   "Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   405
         Left            =   600
         TabIndex        =   14
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label lblAmount 
         BackColor       =   &H00800000&
         Caption         =   "Description"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   405
         Left            =   480
         TabIndex        =   13
         Top             =   3720
         Width           =   3135
      End
      Begin VB.Label lblTim 
         BackColor       =   &H00800000&
         Caption         =   "Time"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   405
         Left            =   600
         TabIndex        =   12
         Top             =   1440
         Width           =   3135
      End
      Begin VB.Label lblSerialNumber 
         BackColor       =   &H00800000&
         Caption         =   "SerialNumber"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   405
         Left            =   480
         TabIndex        =   11
         Top             =   3000
         Width           =   3015
      End
   End
   Begin VB.Label lblDay 
      BackColor       =   &H00C00000&
      Caption         =   "Sunday"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   240
      TabIndex        =   26
      Top             =   7320
      Width           =   975
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C00000&
      Caption         =   "00:00:00 PM"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   1440
      TabIndex        =   25
      Top             =   7320
      Width           =   1815
   End
   Begin VB.Label lblMonth 
      BackColor       =   &H00C00000&
      Caption         =   "March"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   24
      Top             =   7800
      Width           =   855
   End
   Begin VB.Label lblYear 
      BackColor       =   &H00FFFFFF&
      Caption         =   "1998"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   1800
      TabIndex        =   23
      Top             =   7800
      Width           =   855
   End
   Begin VB.Label lblTime 
      BackColor       =   &H00C00000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   0
      TabIndex        =   22
      Top             =   0
      Width           =   2295
   End
   Begin VB.Label Label8 
      Caption         =   "Label6"
      Height          =   405
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C00000&
      Caption         =   "Next of  Kin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   405
      Left            =   0
      TabIndex        =   17
      Top             =   0
      Width           =   2505
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C00000&
      Caption         =   "Next of  Kin"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   405
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   2505
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C00000&
      Caption         =   "Amount paid"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   405
      Left            =   0
      TabIndex        =   15
      Top             =   720
      Width           =   2505
   End
End
Attribute VB_Name = "frmCaptinOrder"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset
Option Explicit

Private Sub cmdAdd_Click()
Dim res As Integer
On Error GoTo trap:
Adodc1.Recordset.AddNew
Exit Sub
trap:
End Sub

Private Sub cmdBack_Click()
Me.Hide
frmMenu.Show
End Sub

Private Sub cmdExit_Click()
Dim intresponse As Integer
intresponse = MsgBox("Do you want to Quit", vbYesNo + vbInformation, "You Quit")
If intresponse = vbYes Then
End
End If
End Sub


Private Sub cmdPrint_Click()
With Me
    .WindowState = 0
    .PrintForm
    Unload Me
End With
End Sub

Private Sub cmdReceipt_Click()
Dim Today As Date
Today = Now = #8/2/2007#
Dim OrderedFrom As Object
Dim SerialNumber As Object
Dim Tim As Date
Dim Description As Object
Dim Amount2 As Object
Dim Waiter As Object
cmdReceipt.Enabled = False
Set OrderedFrom = txtOrderedFrom
Set SerialNumber = txtSerial_Number
Set Description = txtDescription
Set Amount2 = txtAmount2
Set Waiter = txtName_of_waiter
 
With frmOrderReceipt
    .Show
    .OrderedFrom = txtOrderedFrom
    .SerialNumber = txtSerial_Number
    .Tim = Now
    .Description = txtDescription
    .Amount2 = txtAmount2
    .Waiter = txtName_of_waiter
End With

End Sub

Private Sub cmdSave_Click()
Dim res As Integer
On Error GoTo trap:
Adodc1.Recordset.Update
res = MsgBox("Record Saved", vbInformation, "Record was successfully Saved")
Adodc1.Refresh
Exit Sub
trap: res = MsgBox("Empty field cann't Saved", vbInformation, "Saved")
Adodc1.Refresh
End Sub

Private Sub Command1_Click()
cmdReceipt = True
End Sub

Private Sub Command2_Click()
Me.Hide
frmCaptainRecordList.Show
End Sub


Private Sub Form_Load()
' open the connection
    con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Morrenxis_Hotel_DB1.mdb;Persist Security Info=False"

    
  'create a recordset
   rs.Open "Select * from Captin_Order", con, adOpenDynamic, adLockPessimistic
End Sub

Private Sub tmrDate_Timer()
Dim Today As Variant
Today = Now
lblMonth2.Caption = Format(Today, "mmmm")
lblYear2.Caption = Format(Today, "yyyy")
lblNumber2.Caption = Format(Today, "d")
lblTime2.Caption = Format(Today, "h:mm:ss ampm")
End Sub
