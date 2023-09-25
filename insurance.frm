VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form7 
   BackColor       =   &H000080FF&
   Caption         =   "Form7"
   ClientHeight    =   10545
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10875
   LinkTopic       =   "Form7"
   ScaleHeight     =   10545
   ScaleWidth      =   10875
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2175
      Left            =   120
      TabIndex        =   40
      Top             =   8880
      Width           =   15975
      _ExtentX        =   28178
      _ExtentY        =   3836
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
      Height          =   615
      Left            =   11400
      Top             =   1200
      Visible         =   0   'False
      Width           =   2295
      _ExtentX        =   4048
      _ExtentY        =   1085
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
      BackColor       =   -2147483643
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
      Height          =   615
      Left            =   14280
      TabIndex        =   39
      Top             =   3360
      Width           =   1455
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
      Left            =   14280
      TabIndex        =   38
      Top             =   5760
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
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
      Left            =   14280
      TabIndex        =   36
      Top             =   6960
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
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
      Left            =   14280
      TabIndex        =   35
      Top             =   8160
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
      Left            =   14280
      TabIndex        =   34
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Height          =   6975
      Left            =   120
      TabIndex        =   2
      Top             =   1800
      Width           =   13935
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   735
         Left            =   10680
         TabIndex        =   37
         Top             =   2880
         Width           =   2895
         _ExtentX        =   5106
         _ExtentY        =   1296
         _Version        =   393216
         Format          =   60030977
         CurrentDate     =   42220
      End
      Begin VB.TextBox Text15 
         Height          =   615
         Left            =   10680
         TabIndex        =   33
         ToolTipText     =   "enter in words"
         Top             =   4800
         Width           =   2895
      End
      Begin VB.TextBox Text14 
         Height          =   735
         Left            =   10680
         TabIndex        =   31
         ToolTipText     =   "DD in amount"
         Top             =   3840
         Width           =   2895
      End
      Begin VB.TextBox Text13 
         Height          =   405
         Left            =   3240
         TabIndex        =   28
         Top             =   1320
         Width           =   2655
      End
      Begin VB.TextBox Text12 
         Height          =   375
         Left            =   3240
         TabIndex        =   27
         Top             =   720
         Width           =   2655
      End
      Begin VB.TextBox Text11 
         Height          =   615
         Left            =   10680
         TabIndex        =   24
         Top             =   2040
         Width           =   2895
      End
      Begin VB.TextBox Text10 
         Height          =   615
         Left            =   10680
         TabIndex        =   22
         ToolTipText     =   "dd issued bank"
         Top             =   1080
         Width           =   2895
      End
      Begin VB.TextBox Text9 
         Height          =   495
         Left            =   10680
         TabIndex        =   20
         ToolTipText     =   "respective branch code"
         Top             =   240
         Width           =   2895
      End
      Begin VB.TextBox Text8 
         Height          =   375
         Left            =   3240
         TabIndex        =   18
         Top             =   3720
         Width           =   2655
      End
      Begin VB.TextBox Text7 
         Height          =   375
         Left            =   3240
         TabIndex        =   16
         Top             =   4320
         Width           =   2655
      End
      Begin VB.TextBox Text6 
         Height          =   375
         Left            =   3240
         TabIndex        =   14
         Top             =   3120
         Width           =   2655
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   3240
         TabIndex        =   12
         Top             =   6120
         Width           =   2655
      End
      Begin VB.TextBox Text4 
         Height          =   975
         Left            =   3240
         TabIndex        =   11
         Top             =   4920
         Width           =   2655
      End
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   3240
         TabIndex        =   10
         Top             =   2520
         Width           =   2655
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   3240
         TabIndex        =   9
         Top             =   1920
         Width           =   2655
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   3240
         TabIndex        =   8
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label Label18 
         BackColor       =   &H00000000&
         Caption         =   "Amount in words"
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
         Left            =   7560
         TabIndex        =   32
         Top             =   4800
         Width           =   2655
      End
      Begin VB.Label Label17 
         BackColor       =   &H00000000&
         Caption         =   "        Amount"
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
         Left            =   7560
         TabIndex        =   30
         Top             =   3840
         Width           =   2655
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "DATE OF REGISTER"
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
         Left            =   7560
         TabIndex        =   29
         Top             =   2880
         Width           =   2655
      End
      Begin VB.Label Label15 
         BackColor       =   &H00000000&
         Caption         =   "Engine number"
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
         TabIndex        =   26
         Top             =   1320
         Width           =   2775
      End
      Begin VB.Label Label14 
         BackColor       =   &H00000000&
         Caption         =   "Chassi number"
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
         TabIndex        =   25
         Top             =   720
         Width           =   2775
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Caption         =   "DD NUMBER"
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
         Left            =   7560
         TabIndex        =   23
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Label Label12 
         BackColor       =   &H00000000&
         Caption         =   "    BANK NAME"
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
         Left            =   7560
         TabIndex        =   21
         Top             =   1080
         Width           =   2655
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         Caption         =   "RTO OFFICE CODE"
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
         Left            =   7560
         TabIndex        =   19
         Top             =   240
         Width           =   2655
      End
      Begin VB.Line Line2 
         X1              =   6720
         X2              =   6720
         Y1              =   3120
         Y2              =   6840
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Caption         =   "       E-mail "
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
         TabIndex        =   17
         Top             =   3720
         Width           =   2775
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "Certificate number"
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
         TabIndex        =   15
         Top             =   4320
         Width           =   2775
      End
      Begin VB.Label Label8 
         BackColor       =   &H00000000&
         Caption         =   "Mobile number"
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
         TabIndex        =   13
         Top             =   3120
         Width           =   2775
      End
      Begin VB.Line Line1 
         X1              =   6720
         X2              =   6720
         Y1              =   120
         Y2              =   3600
      End
      Begin VB.Label Label7 
         BackColor       =   &H00000000&
         Caption         =   "     Occupation"
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
         Top             =   6120
         Width           =   2775
      End
      Begin VB.Label Label6 
         BackColor       =   &H00000000&
         Caption         =   "Address of the                     Insurer"
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
         Height          =   975
         Left            =   120
         TabIndex        =   6
         Top             =   4920
         Width           =   2775
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "Name of Insurer"
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
         TabIndex        =   5
         Top             =   2520
         Width           =   2775
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "Customer ID"
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
         TabIndex        =   4
         Top             =   1920
         Width           =   2775
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "Vehicle number"
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
         TabIndex        =   3
         Top             =   240
         Width           =   2775
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "        INSURANCE REGISTRATION"
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
      Left            =   5040
      TabIndex        =   1
      Top             =   1200
      Width           =   5535
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "                           REGIONAL TRANSPORT OFFICE"
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
Attribute VB_Name = "Form7"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Form7.Hide
Form4.Show

End Sub



Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 64 Or KeyAscii = 95 Or (KeyAscii > 48 And KeyAscii < 57) Then
ElseIf Text1.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text1.Text = ""
MsgBox "Enter valid number"
End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 96 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 46 Then
ElseIf Text10.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text10.Text = ""
MsgBox "Enter valid name"
End If
End Sub

Private Sub Text11_Change()
If (KeyAscii >= 48 And KeyAscii <= 57) Then
ElseIf Text11.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text11.Text = ""
MsgBox "Enter valid number"
End If
End Sub

Private Sub Text12_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 96 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
ElseIf Text12.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text12.Text = ""
MsgBox "Enter valid number"
End If
End Sub

Private Sub Text13_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 96 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
ElseIf Text13.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text13.Text = ""
MsgBox "Enter valid number"
End If
End Sub

Private Sub Text14_Change()
If (KeyAscii >= 48 And KeyAscii <= 57) Then
ElseIf Text14.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text14.Text = ""
MsgBox "Enter valid number"
End If
End Sub

Private Sub Text15_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 96 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Then
ElseIf Text15.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text15.Text = ""
MsgBox "Enter valid name"
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 96 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
ElseIf Text2.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text2.Text = ""
MsgBox "Enter valid number"
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 96 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 46 Then
ElseIf Text3.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text3.Text = ""
MsgBox "Enter valid name"
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 35 Or KeyAscii = 32 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
ElseIf Text4.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text4.Text = ""
MsgBox "Enter valid address"
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 96 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 46 Then
ElseIf Text5.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text5.Text = ""
MsgBox "Enter valid name"
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 96 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
ElseIf Text6.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text6.Text = ""
MsgBox "Enter valid number"
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 96 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
ElseIf Text7.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text7.Text = ""
MsgBox "Enter valid number"
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Or KeyAscii = 64 Or KeyAscii = 95 Or (KeyAscii >= 48 And KeyAscii <= 57) Then
ElseIf Text8.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
MsgBox "Enter valid mail id"
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If (KeyAscii >= 65 And KeyAscii <= 90) Or (KeyAscii >= 97 And KeyAscii <= 122) Or KeyAscii = 32 Or KeyAscii = 8 Or (KeyAscii >= 48 Or KeyAscii <= 57) Then
ElseIf Text9.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text9.Text = ""
MsgBox "Enter valid number"
End If
End Sub
