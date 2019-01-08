VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmBar_kitchen 
   BackColor       =   &H00800000&
   Caption         =   "Make Bar Kitchen order"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14760
   LinkTopic       =   "Form1"
   ScaleHeight     =   8430
   ScaleWidth      =   14760
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCashier 
      DataField       =   "Cashier"
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
      Left            =   3360
      TabIndex        =   26
      Top             =   5760
      Width           =   2500
   End
   Begin VB.TextBox txtCustomername 
      DataField       =   "Customer_name"
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
      Left            =   3360
      TabIndex        =   4
      Top             =   3720
      Width           =   2500
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "Bar/Kitchen Booklet"
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
      Height          =   9000
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   14700
      Begin VB.Timer tmrDateBar 
         Left            =   600
         Top             =   8160
      End
      Begin VB.TextBox txtDate3 
         DataField       =   "Date_of_entry"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3240
         TabIndex        =   44
         Top             =   6360
         Width           =   2655
      End
      Begin VB.CommandButton cmdViewRecord 
         Caption         =   "View Record "
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7440
         TabIndex        =   32
         Top             =   6600
         Width           =   1455
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
         Left            =   7560
         TabIndex        =   43
         Top             =   2880
         Width           =   1095
      End
      Begin VB.Timer tmrTimeDisplay 
         Interval        =   1000
         Left            =   0
         Top             =   8160
      End
      Begin VB.CommandButton cmdBarReceipt 
         Caption         =   "View Reciept"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   7440
         TabIndex        =   31
         Top             =   2160
         Visible         =   0   'False
         Width           =   1455
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
         Left            =   7560
         TabIndex        =   30
         Top             =   3480
         Width           =   1215
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
         Left            =   7560
         TabIndex        =   29
         Top             =   4200
         Width           =   1095
      End
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   615
         Left            =   3120
         Top             =   7560
         Width           =   2895
         _ExtentX        =   5106
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
         Connect         =   $"frmBar_kitchen.frx":0000
         OLEDBString     =   $"frmBar_kitchen.frx":008B
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "Bar_kitchen"
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
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   7560
         TabIndex        =   28
         Top             =   5400
         Width           =   1020
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
         Left            =   7560
         TabIndex        =   27
         Top             =   4800
         Width           =   1020
      End
      Begin VB.TextBox txtAmountPaid 
         DataField       =   "Amount_Paid"
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
         Left            =   3360
         TabIndex        =   6
         Top             =   5160
         Width           =   2500
      End
      Begin VB.TextBox txtRefNo 
         DataField       =   "Reference_number"
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
         Left            =   3360
         TabIndex        =   3
         Top             =   2280
         Width           =   2500
      End
      Begin VB.TextBox txtRoomNumber 
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
         Height          =   420
         Left            =   3360
         TabIndex        =   2
         Top             =   1560
         Width           =   2500
      End
      Begin VB.TextBox txtDescription2 
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
         Height          =   420
         Left            =   3360
         TabIndex        =   5
         Top             =   4560
         Width           =   2500
      End
      Begin VB.ComboBox cboClassofRoom 
         DataField       =   "Class_of_Room"
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
         ItemData        =   "frmBar_kitchen.frx":0116
         Left            =   3360
         List            =   "frmBar_kitchen.frx":0126
         TabIndex        =   1
         Text            =   "Choose Class of Room"
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label lblDate3 
         Caption         =   "March 31 1998"
         DataField       =   "Date_of_entry"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3240
         TabIndex        =   45
         Top             =   6960
         Width           =   2655
      End
      Begin VB.Label lblDay3 
         BackColor       =   &H00800000&
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
         Left            =   9120
         TabIndex        =   42
         Top             =   600
         Width           =   975
      End
      Begin VB.Label lblTime3 
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
         Left            =   10440
         TabIndex        =   41
         Top             =   600
         Width           =   1815
      End
      Begin VB.Label lblYear3 
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
         Left            =   4320
         TabIndex        =   40
         Top             =   3000
         Width           =   735
      End
      Begin VB.Label lblNumber3 
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
         Left            =   4080
         TabIndex        =   39
         Top             =   3000
         Width           =   315
      End
      Begin VB.Label lblMonth3 
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
         Left            =   3360
         TabIndex        =   38
         Top             =   3000
         Width           =   750
      End
      Begin VB.Image Image2 
         Height          =   1350
         Left            =   6480
         Picture         =   "frmBar_kitchen.frx":014F
         Top             =   480
         Width           =   1350
      End
      Begin VB.Label Label12 
         Alignment       =   2  'Center
         BackColor       =   &H00800000&
         Caption         =   "Customer  name"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   360
         TabIndex        =   25
         Top             =   3600
         Width           =   2265
      End
      Begin VB.Label Label9 
         BackColor       =   &H00800000&
         Caption         =   "Cashier Name"
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
         TabIndex        =   22
         Top             =   5760
         Width           =   2505
      End
      Begin VB.Label Label8 
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
         Left            =   480
         TabIndex        =   21
         Top             =   5160
         Width           =   2145
      End
      Begin VB.Label Label7 
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
         Left            =   600
         TabIndex        =   20
         Top             =   4560
         Width           =   2265
      End
      Begin VB.Label Label5 
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
         Left            =   600
         TabIndex        =   19
         Top             =   960
         Width           =   2535
      End
      Begin VB.Label Label4 
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
         Height          =   405
         Left            =   600
         TabIndex        =   18
         Top             =   1560
         Width           =   2265
      End
      Begin VB.Label Label3 
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
         TabIndex        =   17
         Top             =   3000
         Width           =   2505
      End
      Begin VB.Label Label2 
         BackColor       =   &H00800000&
         Caption         =   "Ref.  Number"
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
         TabIndex        =   16
         Top             =   2280
         Width           =   2505
      End
   End
   Begin VB.Label lblDay2 
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
      Left            =   0
      TabIndex        =   37
      Top             =   0
      Width           =   975
   End
   Begin VB.Label lblTime2 
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
      Left            =   1320
      TabIndex        =   36
      Top             =   0
      Width           =   1815
   End
   Begin VB.Label lblYear2 
      BackColor       =   &H00C00000&
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
      TabIndex        =   35
      Top             =   720
      Width           =   855
   End
   Begin VB.Label lblNumber2 
      BackColor       =   &H00C00000&
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
      Left            =   1080
      TabIndex        =   34
      Top             =   720
      Width           =   495
   End
   Begin VB.Label lblMonth2 
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
      Left            =   0
      TabIndex        =   33
      Top             =   720
      Width           =   855
   End
   Begin VB.Label Label11 
      BackColor       =   &H00C00000&
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
      Left            =   360
      TabIndex        =   24
      Top             =   5400
      Width           =   2505
   End
   Begin VB.Label Label10 
      BackColor       =   &H00C00000&
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
      Left            =   3120
      TabIndex        =   23
      Top             =   5640
      Width           =   2505
   End
   Begin VB.Label Label6 
      BackColor       =   &H00FF0000&
      Caption         =   "Unit Price"
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
      Left            =   2280
      TabIndex        =   15
      Top             =   5520
      Width           =   3135
   End
   Begin VB.Label Label1 
      BackColor       =   &H00C00000&
      Caption         =   "Cover"
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
      Left            =   2280
      TabIndex        =   14
      Top             =   1800
      Width           =   3135
   End
   Begin VB.Label lblNextofKin 
      BackColor       =   &H00C00000&
      Caption         =   "Quantity"
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
      Left            =   2280
      TabIndex        =   13
      Top             =   4200
      Width           =   3135
   End
   Begin VB.Label lblCustomer_Name 
      BackColor       =   &H00C00000&
      Caption         =   "Table Number"
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
      Left            =   2280
      TabIndex        =   12
      Top             =   1320
      Width           =   3135
   End
   Begin VB.Label lblCustomer_Address 
      BackColor       =   &H000000FF&
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
      ForeColor       =   &H000000C0&
      Height          =   405
      Left            =   2280
      TabIndex        =   11
      Top             =   2400
      Width           =   3135
   End
   Begin VB.Label lblAmount 
      BackColor       =   &H00C00000&
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
      Left            =   2280
      TabIndex        =   10
      Top             =   4920
      Width           =   3135
   End
   Begin VB.Label lblPhone_number 
      BackColor       =   &H00C00000&
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
      Left            =   2280
      TabIndex        =   9
      Top             =   3000
      Width           =   3135
   End
   Begin VB.Label lblOccupation 
      BackColor       =   &H00C00000&
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
      Left            =   2280
      TabIndex        =   8
      Top             =   3600
      Width           =   3135
   End
   Begin VB.Label lblMarketer_Type 
      BackColor       =   &H00C00000&
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
      Left            =   2640
      TabIndex        =   7
      Top             =   840
      Width           =   3135
   End
End
Attribute VB_Name = "frmBar_kitchen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

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

Private Sub cmdBarReceipt_Click()
Dim Today As Date
Today = Now = #8/2/2007#
Dim RoomNo As Object
Dim roomclass As Object
Dim ReferenceNo As Object
Dim customername As Object
Dim Description As Object
Dim amountpaid As Object
Dim CashierName As Object
Dim Date3 As Date
cmdBarReceipt.Enabled = False
Set roomclass = cboClassofRoom
Set RoomNo = txtRoomNumber
Set ReferenceNo = txtRefNo
Set customername = txtCustomername
Set Description = txtDescription2
Set amountpaid = txtAmountPaid
Set CashierName = txtCashier

With frmbarReciept
     .Show
     .roomclass = cboClassofRoom
     .RoomNo = txtRoomNumber
     .ReferenceNo = txtRefNo
     .customername = txtCustomername
     .Description = txtDescription2
     .amountpaid = txtAmountPaid
     .CashierName = txtCashier
     .Date3 = Now
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

Private Sub cmdSave_Click()
  Dim res As Integer
On Error GoTo trap:
Adodc1.Recordset.Update
res = MsgBox("Record Saved", vbInformation, "Record was successfully Saved")
Adodc1.Refresh
Exit Sub
trap: res = MsgBox("Empty field cann't Saved", vbInformation, "Saved")
Adodc1.Refresh
  'Disable text boxes
  If cmdSave_Click = True Then
   cboClassofRoom.Locked = True
   txtCustomer_Name.Locked = True
   txtRoomNo.Locked = True
   txtCustomer_Address.Locked = True
   txtPhone_number.Locked = True
   txtAmountPaid.Locked = True
   txtCheckin.Locked = True
   txtDuration.Locked = True
   txtDate.Locked = True
   txtReceptionist.Locked = True
   MsgBox "Cannot modify record", vbRetryCancel
 End If
End Sub

Private Sub cmdViewRecord_Click()
Me.Hide
frmBarRecordList.Show
End Sub


Private Sub Command1_Click()
cmdBarReceipt = True
End Sub

Private Sub Form_Load()
' open the connection
   con.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & "\Morrenxis_Hotel_DB1.mdb;Persist Security Info=False"
 
  'create a recordset
   rs.Open "Select * from Bar_kitchen", con, adOpenDynamic, adLockPessimistic
End Sub

Private Sub lblDate3_Change()
Dim Today As Variant
Today = Now
End Sub

Private Sub tmrDateBar_Timer()
Dim Today As Variant
Today = Now
lblDate3 = Format(Today, "mmmm", "yyyy", "d")
End Sub

Private Sub tmrTimeDisplay_Timer()
Dim Today As Variant
Today = Now
lblDay3.Caption = Format(Today, "dddd")
lblMonth3.Caption = Format(Today, "mmmm")
lblYear3.Caption = Format(Today, "yyyy")
lblNumber3.Caption = Format(Today, "d")
lblTime3.Caption = Format(Today, "h:mm:ss ampm")
End Sub

Private Sub txtDate3_Change()
'Call date function
Call tmrTimeDisplay_Timer
End Sub
