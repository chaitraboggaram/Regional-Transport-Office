VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form13 
   BackColor       =   &H000080FF&
   Caption         =   "Form13"
   ClientHeight    =   10095
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   19185
   LinkTopic       =   "Form13"
   ScaleHeight     =   10095
   ScaleWidth      =   19185
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3255
      Left            =   120
      TabIndex        =   19
      Top             =   7680
      Width           =   15975
      _ExtentX        =   28178
      _ExtentY        =   5741
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
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   13200
      Top             =   1200
      Visible         =   0   'False
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   661
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   0
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
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
      Height          =   735
      Left            =   12000
      TabIndex        =   18
      Top             =   1680
      Width           =   1695
   End
   Begin VB.CommandButton Command4 
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
      Height          =   735
      Left            =   12000
      TabIndex        =   17
      Top             =   4080
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Next"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12000
      TabIndex        =   16
      Top             =   5400
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12000
      TabIndex        =   15
      Top             =   6720
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   12000
      TabIndex        =   14
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "AGENT DETAILS"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6255
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   11655
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   615
         Left            =   6120
         TabIndex        =   13
         Top             =   5280
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   1085
         _Version        =   393216
         Format          =   60030977
         CurrentDate     =   42276
      End
      Begin VB.TextBox Text5 
         Height          =   1455
         Left            =   6120
         TabIndex        =   11
         Top             =   3600
         Width           =   4455
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   6120
         TabIndex        =   9
         Top             =   2760
         Width           =   3255
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   6120
         TabIndex        =   7
         Top             =   1920
         Width           =   3255
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   6120
         TabIndex        =   5
         Top             =   1080
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   6120
         TabIndex        =   3
         Top             =   360
         Width           =   2655
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "             Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   615
         Left            =   1200
         TabIndex        =   12
         Top             =   5280
         Width           =   2895
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Showroom address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   1200
         TabIndex        =   10
         Top             =   3600
         Width           =   2775
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "           E-mail  "
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   375
         Left            =   1200
         TabIndex        =   8
         Top             =   2040
         Width           =   2775
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   " Showroom name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   1200
         TabIndex        =   6
         Top             =   2760
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Mobile number"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   1200
         TabIndex        =   4
         Top             =   1200
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H00000000&
         Caption         =   "  Agent Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000080FF&
         Height          =   495
         Left            =   1200
         TabIndex        =   2
         Top             =   360
         Width           =   2775
      End
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "                            REGIONAL TRANSPORT OFFICE"
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
Attribute VB_Name = "Form13"
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
End Sub

Private Sub Command2_Click()
Form13.Hide
Form6.Show
End Sub

Private Sub Command3_Click()
Form13.Hide
Form11.Show
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
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 96 And KeyAscii <= 122) Or kayacsii = 32 Or KeyAscii = 8 Then
ElseIf Text1.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text1.Text = ""
MsgBox "Enter valid name"
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = 8 Then
ElseIf Text3.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text3.Text = ""
MsgBox "Enter valid name"
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or kayacsii = 32 Or KeyAscii = 8 Or KeyAscii = 64 Or KeyAscii = 95 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
ElseIf Text4.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
KeyAscii = 0
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or kayacsii = 32 Or KeyAscii = 8 Or KeyAscii = 64 Or KeyAscii = 95 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
ElseIf Text4.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
KeyAscii = 0
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or kayacsii = 32 Or KeyAscii = 8 Then
ElseIf Text5.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text5.Text = ""
MsgBox "Enter valid name"
End If
End Sub
