VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form5 
   BackColor       =   &H000080FF&
   Caption         =   "Form5"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form5"
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Bindings        =   "vehicel.frx":0000
      Height          =   1695
      Left            =   120
      TabIndex        =   38
      Top             =   9360
      Width           =   15975
      _ExtentX        =   28178
      _ExtentY        =   2990
      _Version        =   393216
      BackColor       =   33023
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13560
      TabIndex        =   35
      Top             =   6240
      Width           =   1455
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13560
      TabIndex        =   34
      Top             =   3840
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   13680
      Top             =   1200
      Visible         =   0   'False
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
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
      Connect         =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=rto;Data Source=RANJITHA-5F079E"
      OLEDBString     =   "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=rto;Data Source=RANJITHA-5F079E"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   "Vehicle1"
      Caption         =   "Adodc1"
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
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13560
      TabIndex        =   30
      Top             =   7440
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13560
      TabIndex        =   29
      Top             =   8640
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13560
      TabIndex        =   28
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Vehicle Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6975
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   13215
      Begin VB.ComboBox Combo4 
         DataField       =   "FuelUsed"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "vehicel.frx":0015
         Left            =   3600
         List            =   "vehicel.frx":0025
         TabIndex        =   33
         Top             =   6240
         Width           =   2415
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         DataField       =   "YearOfManufacture"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3600
         TabIndex        =   32
         Top             =   5400
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   42270
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         DataField       =   "RegisterDate"
         BeginProperty DataFormat 
            Type            =   1
            Format          =   "dd-MMM-yyyy"
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   3
         EndProperty
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3600
         TabIndex        =   31
         Top             =   3720
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   873
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   42220
      End
      Begin VB.TextBox Text9 
         DataField       =   "ValidityOfIssuingAuthorithy"
         DataSource      =   "Adodc1"
         Height          =   615
         Left            =   10080
         TabIndex        =   27
         Top             =   5880
         Width           =   2055
      End
      Begin VB.TextBox Text8 
         DataField       =   "Color"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   10080
         TabIndex        =   25
         Top             =   5040
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         DataField       =   "ModelNumber"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   10080
         TabIndex        =   23
         Top             =   3960
         Width           =   1935
      End
      Begin VB.TextBox Text6 
         DataField       =   "EngineNumber"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   10080
         TabIndex        =   21
         Top             =   2880
         Width           =   1935
      End
      Begin VB.TextBox Text5 
         DataField       =   "ChassiNumber"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   10080
         TabIndex        =   19
         Top             =   2040
         Width           =   1935
      End
      Begin VB.TextBox Text4 
         DataField       =   "MakersName"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   7920
         TabIndex        =   17
         Top             =   1200
         Width           =   3615
      End
      Begin VB.TextBox Text2 
         DataField       =   "ClassOfVehicle"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3600
         TabIndex        =   13
         Top             =   4560
         Width           =   2415
      End
      Begin VB.TextBox Text1 
         DataField       =   "RegisterNumber"
         DataSource      =   "Adodc1"
         Height          =   495
         Left            =   3600
         TabIndex        =   10
         Top             =   2880
         Width           =   2415
      End
      Begin VB.ComboBox Combo3 
         DataField       =   "CellUsed"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "vehicel.frx":0047
         Left            =   3600
         List            =   "vehicel.frx":0051
         TabIndex        =   8
         Top             =   2160
         Width           =   2415
      End
      Begin VB.ComboBox Combo2 
         DataField       =   "VehicleVariety"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "vehicel.frx":0064
         Left            =   3600
         List            =   "vehicel.frx":006E
         TabIndex        =   6
         Top             =   1320
         Width           =   2415
      End
      Begin VB.ComboBox Combo1 
         DataField       =   "VehicleType"
         DataSource      =   "Adodc1"
         Height          =   315
         ItemData        =   "vehicel.frx":008B
         Left            =   3600
         List            =   "vehicel.frx":009B
         TabIndex        =   4
         Top             =   480
         Width           =   2415
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "Validity of issuing           Authority"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   735
         Left            =   6960
         TabIndex        =   26
         Top             =   5880
         Width           =   2415
      End
      Begin VB.Label Label15 
         BackColor       =   &H00000000&
         Caption         =   "          color     "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   6960
         TabIndex        =   24
         Top             =   5040
         Width           =   2415
      End
      Begin VB.Label Label14 
         BackColor       =   &H00000000&
         Caption         =   "   Model number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   6960
         TabIndex        =   22
         Top             =   3960
         Width           =   2415
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Caption         =   "    Engine number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   6960
         TabIndex        =   20
         Top             =   2880
         Width           =   2415
      End
      Begin VB.Label Label12 
         BackColor       =   &H00000000&
         Caption         =   "   Chassi number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   6960
         TabIndex        =   18
         Top             =   2040
         Width           =   2415
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         Caption         =   "    Maker's Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   8520
         TabIndex        =   16
         Top             =   360
         Width           =   2415
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Caption         =   "        Fuel used"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   120
         TabIndex        =   15
         Top             =   6240
         Width           =   3255
      End
      Begin VB.Line Line1 
         X1              =   6840
         X2              =   6840
         Y1              =   240
         Y2              =   6960
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "Year of Manufacture"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   120
         TabIndex        =   14
         Top             =   5400
         Width           =   3255
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "      Class of Vehicle"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   4560
         Width           =   3255
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "     Register Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   120
         TabIndex        =   11
         Top             =   3720
         Width           =   3255
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "     Register number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   120
         TabIndex        =   9
         Top             =   2880
         Width           =   3255
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Cell used to run Engine"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   3255
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "        Vehicle Variety"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1200
         Width           =   3255
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "         Vehicle Type "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   3255
      End
   End
   Begin VB.Label Label18 
      BackColor       =   &H00000000&
      DataField       =   "ApplicationNo"
      DataSource      =   "Adodc1"
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   1920
      TabIndex        =   37
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label17 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Application No:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   120
      TabIndex        =   36
      Top             =   1200
      Width           =   1815
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "         VEHICLE REGISTRATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   4560
      TabIndex        =   1
      Top             =   1320
      Width           =   6375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "                             REGIONAL TRANSPORT OFFICE"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   975
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16215
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
Combo1.Text = ""
Combo2.Text = ""
Combo3.Text = ""
Combo4.Text = ""
DTPicker1.Day = ""
Text1.Text = ""
Text2.Text = ""
Text4.Text = ""
Text6.Text = ""
Text7.Text = ""
Text8.Text = ""
Text9.Text = ""
End Sub

Private Sub Command2_Click()
Form5.Hide
Form6.Show
End Sub

Private Sub Command3_Click()
Form5.Hide
Form4.Show
End Sub

Private Sub Command4_Click()
Dim id As Integer
On Error GoTo errormsg
Adodc1.Refresh
Adodc1.Recordset.MoveLast
id = Adodc1.Recordset.Fields(0) + 1
Adodc1.Recordset.AddNew
Label18.Caption = id
Text1.SetFocus
Exit Sub
errormsg:
Adodc1.Recordset.AddNew
Label18.Caption = 1001
Text1.SetFocus
End Sub

Private Sub Command5_Click()
Dim a As Integer
Adodc1.Recordset.Update
MsgBox "Current details are saved"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or KeyAscii = 45 Or KeyAscii = 32 Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
ElseIf Text1.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text1.Text = ""
MsgBox "Enter valid number"
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
ElseIf Text2.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text2.Text = ""
MsgBox "Enter valid name"
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Then
ElseIf Text4.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text4.Text = ""
MsgBox "Enter valid name"
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
ElseIf Text6.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text6.Text = ""
MsgBox "Enter valid name"
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
ElseIf Text7.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text7.Text = ""
MsgBox "Enter valid number"
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Then
ElseIf Text8.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text8.Text = ""
MsgBox "Enter valid number"
End If

End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
ElseIf Text9.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text9.Text = ""
MsgBox "Enter valid number"
End If
End Sub
