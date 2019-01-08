VERSION 5.00
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Begin VB.Form frmLogin 
   BackColor       =   &H8000000D&
   Caption         =   "UserLogin"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14760
   LinkTopic       =   "Form3"
   ScaleHeight     =   8430
   ScaleWidth      =   14760
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Caption         =   "Kindly enter username and password"
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
      Width           =   14715
      Begin MSAdodcLib.Adodc Adodc1 
         Height          =   615
         Left            =   2640
         Top             =   7680
         Visible         =   0   'False
         Width           =   4935
         _ExtentX        =   8705
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
         Connect         =   $"frmLogin.frx":0000
         OLEDBString     =   $"frmLogin.frx":008B
         OLEDBFile       =   ""
         DataSourceName  =   ""
         OtherAttributes =   ""
         UserName        =   ""
         Password        =   ""
         RecordSource    =   "LoginTB"
         Caption         =   "Login Form"
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
         BackColor       =   &H00FFFFFF&
         Caption         =   "Cancel"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   5160
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   5040
         Width           =   2100
      End
      Begin VB.CommandButton cmdLogin 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Login"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2520
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5040
         Width           =   2100
      End
      Begin VB.TextBox txtUserName 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         Left            =   4200
         TabIndex        =   1
         Top             =   2280
         Width           =   3075
      End
      Begin VB.TextBox txtPassword 
         BeginProperty Font 
            Name            =   "Calibri"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   525
         IMEMode         =   3  'DISABLE
         Left            =   4200
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   3360
         Width           =   3075
      End
      Begin VB.Image Image1 
         Height          =   1350
         Left            =   10080
         Picture         =   "frmLogin.frx":0116
         Top             =   360
         Width           =   1350
      End
      Begin VB.Label Label3 
         BackColor       =   &H00800000&
         Caption         =   "USER AUTHENTICATION"
         BeginProperty Font 
            Name            =   "Arial Narrow"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   855
         Left            =   3240
         TabIndex        =   7
         Top             =   840
         Width           =   4455
      End
      Begin VB.Image Image2 
         Height          =   495
         Left            =   8640
         Picture         =   "frmLogin.frx":0BA6
         Stretch         =   -1  'True
         Top             =   5040
         Width           =   615
      End
      Begin VB.Shape Shape2 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   1575
         Left            =   1800
         Shape           =   4  'Rounded Rectangle
         Top             =   4560
         Width           =   7695
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Username"
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
         Left            =   2040
         TabIndex        =   6
         Top             =   2400
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Password"
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
         Left            =   2040
         TabIndex        =   5
         Top             =   3360
         Width           =   1815
      End
      Begin VB.Shape Shape1 
         BackStyle       =   1  'Opaque
         BorderColor     =   &H00C0C0C0&
         BorderWidth     =   3
         Height          =   2655
         Left            =   1800
         Shape           =   4  'Rounded Rectangle
         Top             =   1800
         Width           =   7695
      End
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public LoginSucceeded As Boolean

Private Sub cmdExit_Click()
Dim res As Integer
res = MsgBox("Do you want to cancel", vbInformation + vbYesNo, "Exit")
If res = vbYes Then
End
Else
End If
End Sub

Private Sub cmdLogin_Click()
'check for correct password
Static y As Integer
Adodc1.Refresh
    With Adodc1.Recordset
        Dim query As String
        query = "username = '" & txtUserName.Text & "'"
        query = query & " and password = '" & txtPassword.Text & "'"
             
        If .RecordCount > 0 Then
            MsgBox "Welcome, Your Login was Successful!"
            frmMenu.Show
            Unload Me
        Else
            MsgBox "Invalid username or password!", vbOKOnly + vbCritical, "Access denied!"
            txtUserName.Text = ""
            txtPassword.Text = ""
            txtUserName.SetFocus
        End If
    End With

   End Sub
