VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmReservation 
   Caption         =   "Reservation Billing"
   ClientHeight    =   8430
   ClientLeft      =   225
   ClientTop       =   450
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   14760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdBillDisplay2 
      Caption         =   "Display Customer copy"
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
      Left            =   240
      TabIndex        =   32
      Top             =   7560
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox txtReceptionist 
      DataField       =   "Name of Receptionist"
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
      Left            =   3960
      TabIndex        =   10
      Top             =   6840
      Width           =   2535
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "Customer Reservation Billing"
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
      Height          =   8415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   15135
      Begin VB.CommandButton cmdRecord 
         Caption         =   "Reservation Record"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   8160
         TabIndex        =   35
         Top             =   5880
         Width           =   1575
      End
      Begin VB.CommandButton Command2 
         Caption         =   "Print "
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
         Left            =   8160
         TabIndex        =   34
         Top             =   5040
         Width           =   1380
      End
      Begin VB.TextBox txtRoomNo 
         DataField       =   "Room Number"
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
         Height          =   375
         Left            =   3960
         TabIndex        =   2
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Timer trmDisplay 
         Interval        =   1000
         Left            =   0
         Top             =   8280
      End
      Begin VB.CommandButton cmdBillDisplay 
         Caption         =   "Display Hotel copy"
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
         Left            =   2040
         TabIndex        =   27
         Top             =   7560
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.TextBox txtDuration 
         DataField       =   "Duration (Days)"
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
         Left            =   3960
         TabIndex        =   9
         Top             =   6360
         Width           =   2535
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
         Height          =   490
         Left            =   8160
         TabIndex        =   11
         Top             =   2040
         Width           =   1335
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
         Left            =   8160
         TabIndex        =   12
         Top             =   2760
         Width           =   1335
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   615
         Left            =   4200
         Top             =   7440
         Width           =   3255
         _ExtentX        =   5741
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
         Connect         =   $"frmReservation (2).frx":0000
         OLEDBString     =   $"frmReservation (2).frx":008B
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Reservation_Billing"
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
      Begin VB.CommandButton cmdBack 
         Caption         =   "Back"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   8160
         TabIndex        =   13
         Top             =   3480
         Width           =   1380
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "Exit"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   8160
         TabIndex        =   14
         Top             =   4200
         Width           =   1380
      End
      Begin VB.TextBox txtCheckout 
         DataField       =   "Check out TIme"
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
         Left            =   3960
         TabIndex        =   8
         Top             =   5280
         Width           =   2535
      End
      Begin VB.TextBox txtCustomer_Name 
         DataField       =   "Customer Name"
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
         Left            =   3960
         TabIndex        =   3
         Top             =   1680
         Width           =   2500
      End
      Begin VB.TextBox txtCustomer_Address 
         DataField       =   "Customer Address"
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
         Left            =   3960
         TabIndex        =   4
         Top             =   2280
         Width           =   2500
      End
      Begin VB.TextBox txtPhone_number 
         DataField       =   "Phone Number"
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
         Left            =   3960
         TabIndex        =   5
         Top             =   2880
         Width           =   2500
      End
      Begin VB.TextBox txtOccupation 
         DataField       =   "Occupation"
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
         Left            =   3960
         TabIndex        =   6
         Top             =   3480
         Width           =   2500
      End
      Begin VB.ComboBox cboClassofRoom 
         DataField       =   "Class of Room"
         DataSource      =   "Adodc1"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         ItemData        =   "frmReservation (2).frx":0116
         Left            =   3960
         List            =   "frmReservation (2).frx":0126
         TabIndex        =   1
         Text            =   "Choose Class of Room"
         Top             =   720
         Width           =   2535
      End
      Begin VB.TextBox txtAmountPaid 
         DataField       =   "Amount paid"
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
         Left            =   3960
         TabIndex        =   7
         Top             =   4080
         Width           =   2535
      End
      Begin VB.Label Label1 
         BackColor       =   &H00800000&
         Caption         =   "Room Number"
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
         Height          =   375
         Left            =   960
         TabIndex        =   33
         Top             =   1200
         Width           =   1695
      End
      Begin VB.Label lblYear 
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
         Left            =   4920
         TabIndex        =   31
         Top             =   5880
         Width           =   735
      End
      Begin VB.Label lblNumber 
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
         Left            =   4680
         TabIndex        =   30
         Top             =   5880
         Width           =   375
      End
      Begin VB.Label lblMonth 
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
         Left            =   3960
         TabIndex        =   29
         Top             =   5880
         Width           =   735
      End
      Begin VB.Label lblTime 
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
         Left            =   3960
         TabIndex        =   28
         Top             =   4680
         Width           =   1815
      End
      Begin VB.Label lblDuration 
         BackColor       =   &H00800000&
         Caption         =   "Duration (days)"
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
         Left            =   960
         TabIndex        =   26
         Top             =   6360
         Width           =   2505
      End
      Begin VB.Image Image1 
         Height          =   1350
         Left            =   7200
         Picture         =   "frmReservation (2).frx":014F
         Top             =   480
         Width           =   1350
      End
      Begin VB.Label lblReceptionist 
         BackColor       =   &H00800000&
         Caption         =   "Name of Receptionist"
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
         Left            =   960
         TabIndex        =   25
         Top             =   6840
         Width           =   2505
      End
      Begin VB.Label lblDate 
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
         Left            =   960
         TabIndex        =   23
         Top             =   5880
         Width           =   2505
      End
      Begin VB.Label lblCheckout 
         BackColor       =   &H00800000&
         Caption         =   "Check out time"
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
         Left            =   960
         TabIndex        =   22
         Top             =   5400
         Width           =   2505
      End
      Begin VB.Label lblCheckin 
         BackColor       =   &H00800000&
         Caption         =   "Check in Time"
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
         Left            =   960
         TabIndex        =   21
         Top             =   4800
         Width           =   2505
      End
      Begin VB.Label lblCustomer_Name 
         BackColor       =   &H00800000&
         Caption         =   "Customer Name"
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
         Left            =   960
         TabIndex        =   20
         Top             =   1680
         Width           =   2505
      End
      Begin VB.Label lblCustomer_Address 
         BackColor       =   &H00800000&
         Caption         =   "Customer Address"
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
         Left            =   960
         TabIndex        =   19
         Top             =   2280
         Width           =   2505
      End
      Begin VB.Label lblAmount 
         BackColor       =   &H00800000&
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
         Left            =   960
         TabIndex        =   18
         Top             =   4080
         Width           =   2505
      End
      Begin VB.Label lblPhone_number 
         BackColor       =   &H00800000&
         Caption         =   "Phone Number"
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
         Left            =   960
         TabIndex        =   17
         Top             =   2880
         Width           =   2505
      End
      Begin VB.Label lblOccupation 
         BackColor       =   &H00800000&
         Caption         =   "Occupation"
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
         Left            =   960
         TabIndex        =   16
         Top             =   3480
         Width           =   2505
      End
      Begin VB.Label lblClass_of_room 
         BackColor       =   &H00800000&
         Caption         =   "Class of Room"
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
         Left            =   960
         TabIndex        =   15
         Top             =   720
         Width           =   2535
      End
   End
   Begin VB.Label Label5 
      BackColor       =   &H00C00000&
      Caption         =   "Date"
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
      Height          =   495
      Left            =   480
      TabIndex        =   24
      Top             =   7800
      Width           =   2505
   End
End
Attribute VB_Name = "frmReservation"
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

Private Sub cmdBillDisplay_Click()
Dim customername As Object
Dim roomclass As Object
Dim RoomNo As Object
Dim customeraddress As Object
Dim phonenumber As Object
Dim amountpaid As Object
Dim checkin As Object
Dim duration As Object
Dim date1 As Object
Dim receptionist As Object
cmdBillDisplay.Enabled = False
Set customername = txtCustomer_Name
Set roomclass = cboClassofRoom
Set RoomNo = txtRoomNo
Set customeraddress = txtCustomer_Address
Set phonenumber = txtPhone_number
Set amountpaid = txtAmountPaid
Set checkin = txtCheckin
Set duration = txtDuration
Set date1 = txtDate
Set receptionist = txtReceptionist

With frmReceipt
    .Show
    .roomclass = cboClassofRoom
    .RoomNo = txtRoomNo
    .customername = txtCustomer_Name
    .customeraddress = txtCustomer_Address
    .phonenumber = txtPhone_number
    .amountpaid = txtAmountPaid
    .checkin = txtCheckin
    .duration = txtDuration
    .date1 = txtDate
    .receptionist = txtReceptionist
End With
    

End Sub

Private Sub cmdBillDisplay2_Click()
Dim Today As Variant
Today = Now
Dim customername As Object
Dim roomclass As Object
Dim RoomNo As Object
Dim customeraddress As Object
Dim phonenumber As Object
Dim amountpaid As Object
Dim duration As Object
Dim checkin As Date
Dim date1 As Date
Dim receptionist As Object
cmdBillDisplay2.Enabled = False
Set customername = txtCustomer_Name
Set roomclass = cboClassofRoom
Set RoomNo = txtRoomNo
Set customeraddress = txtCustomer_Address
Set phonenumber = txtPhone_number
Set amountpaid = txtAmountPaid
Set duration = txtDuration
Set receptionist = txtReceptionist

With frmReceipt2
    .Show
    .roomclass = cboClassofRoom
    .RoomNo = txtRoomNo
    .customername = txtCustomer_Name
    .customeraddress = txtCustomer_Address
    .phonenumber = txtPhone_number
    .amountpaid = txtAmountPaid
    .checkin = Now
    .duration = txtDuration
    .date1 = Now
    .receptionist = txtReceptionist
End With
    
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

Private Sub cmdRecord_Click()
Me.Hide
frmRecrd_list.Show
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
cmdBillDisplay = True
End Sub

Private Sub Command2_Click()
cmdBillDisplay2 = True
End Sub

Private Sub txtCustomer_Name_Change()
cmdBillDisplay.Enabled = True
End Sub

Private Sub DTPicker1_Change()
txtDate = DTPicker1.Value
End Sub

Private Sub Form_Load()
' open the connection
   con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Morrenxis_Hotel_DB1.mdb;Persist Security Info=False"

  
  'create a recordset
   rs.Open "Select * from Reservation_Billing", con, adOpenDynamic, adLockPessimistic
End Sub

Private Sub lblTime_Click()
lblTime = Time
End Sub

Private Sub trmDisplay_Timer()
Dim Today As Variant
Today = Now
lblMonth.Caption = Format(Today, "mmmm")
lblYear.Caption = Format(Today, "yyyy")
lblNumber.Caption = Format(Today, "d")
lblTime.Caption = Format(Today, "h:mm:ss ampm")
End Sub
