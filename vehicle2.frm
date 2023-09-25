VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form6 
   BackColor       =   &H000080FF&
   Caption         =   "Form6"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form6"
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2655
      Left            =   120
      TabIndex        =   24
      Top             =   8280
      Width           =   15855
      _ExtentX        =   27966
      _ExtentY        =   4683
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
      TabIndex        =   23
      Top             =   2640
      Width           =   1455
   End
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   495
      Left            =   11760
      Top             =   1320
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
      RecordSource    =   "Vehicle2"
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
   Begin VB.CommandButton Command4 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13560
      TabIndex        =   22
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Cancel"
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
      TabIndex        =   21
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
      TabIndex        =   19
      Top             =   6240
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
      TabIndex        =   18
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "OWNER DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   13215
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   615
         Left            =   3000
         TabIndex        =   20
         Top             =   1560
         Width           =   2775
         _ExtentX        =   4895
         _ExtentY        =   1085
         _Version        =   393216
         Format          =   16515073
         CurrentDate     =   42220
      End
      Begin VB.TextBox Text7 
         Height          =   1335
         Left            =   7080
         TabIndex        =   17
         Top             =   3240
         Width           =   5655
      End
      Begin VB.TextBox Text6 
         Height          =   1335
         Left            =   7080
         TabIndex        =   15
         Top             =   1200
         Width           =   5535
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   3000
         TabIndex        =   13
         Top             =   5280
         Width           =   3255
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   3000
         TabIndex        =   11
         Top             =   4440
         Width           =   3255
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   3000
         TabIndex        =   10
         Top             =   3480
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   3000
         TabIndex        =   7
         Top             =   2520
         Width           =   3255
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   3000
         TabIndex        =   4
         Top             =   480
         Width           =   3255
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Caption         =   "   Temporary Address"
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
         Left            =   8280
         TabIndex        =   16
         Top             =   2640
         Width           =   3015
      End
      Begin VB.Line Line1 
         X1              =   6600
         X2              =   6600
         Y1              =   240
         Y2              =   6120
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "   Permanent Address"
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
         Left            =   8280
         TabIndex        =   14
         Top             =   480
         Width           =   3015
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "FATHER/HUSBAND              NAME"
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
         Left            =   120
         TabIndex        =   12
         Top             =   5160
         Width           =   2655
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "          E-mail"
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
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   4440
         Width           =   2655
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "  MOBILE NUMBER"
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
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   3480
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "  PLACE OF BIRTH"
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
         TabIndex        =   6
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "  DATE OF BIRTH"
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
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   1560
         Width           =   2655
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "          NAME"
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
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   480
         Width           =   2655
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "        VEHICLE   REGISTRATION"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   615
      Left            =   5400
      TabIndex        =   1
      Top             =   1200
      Width           =   5295
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "                             REGIONALTRANSPORT OFFICE"
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
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Text1.Text = ""
Text2.Text = ""
Text3.Text = ""
Text4.Text = ""
Text5.Text = ""
Text6.Text = ""
Text7.Text = ""
End Sub

Private Sub Command2_Click()
Form6.Hide
Form13.Show
End Sub

Private Sub Command3_Click()
Form6.Hide
Form4.Show
End Sub

Private Sub Command5_Click()
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

Private Sub Command4_Click()
Dim a As Integer
Adodc1.Recordset.Update
MsgBox "Current details are saved"

End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Then
ElseIf Text1.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text1.Text = ""
MsgBox "Enter valid name"
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Then
ElseIf Text2.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text2.Text = ""
MsgBox "Enter valid name"
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
ElseIf Text3.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text3.Text = ""
MsgBox "Enter valid name"
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 64 Or KeyAscii = 95 Or (KeyAscii > 48 And KeyAscii < 57) Then
ElseIf Text4.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
KeyAscii = 0
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Then
ElseIf Text5.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text5.Text = ""
MsgBox "Enter valid name"
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 64 Or KeyAscii = 95 Or (KeyAscii > 48 And KeyAscii < 57) Then
ElseIf Text6.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text6.Text = ""
MsgBox "Enter valid address"
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 35 Or KeyAscii = 8 Or (KeyAscii > 48 And KeyAscii < 57) Then
ElseIf Text7.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text7.Text = ""
MsgBox "Enter valid address"
End If
End Sub
