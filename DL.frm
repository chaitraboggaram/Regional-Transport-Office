VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form9 
   BackColor       =   &H000080FF&
   Caption         =   "Form9"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form9"
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   Begin MSAdodcLib.Adodc Adodc1 
      Height          =   375
      Left            =   11880
      Top             =   1320
      Visible         =   0   'False
      Width           =   3135
      _ExtentX        =   5530
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
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2055
      Left            =   0
      TabIndex        =   41
      Top             =   8880
      Width           =   15615
      _ExtentX        =   27543
      _ExtentY        =   3625
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
      Left            =   15840
      TabIndex        =   40
      Top             =   3960
      Width           =   1695
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
      Left            =   15840
      TabIndex        =   39
      Top             =   6000
      Width           =   1695
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
      Left            =   15840
      TabIndex        =   38
      Top             =   8160
      Width           =   1695
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
      Left            =   15840
      TabIndex        =   37
      Top             =   7080
      Width           =   1695
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
      Left            =   15840
      TabIndex        =   36
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      Height          =   495
      Left            =   13680
      TabIndex        =   25
      ToolTipText     =   "select the specified option."
      Top             =   4080
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Caption         =   "Applicant Details"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6855
      Left            =   0
      TabIndex        =   2
      Top             =   1920
      Width           =   15615
      Begin VB.TextBox Text11 
         Height          =   855
         Left            =   10320
         TabIndex        =   35
         Top             =   5640
         Width           =   4935
      End
      Begin MSComCtl2.DTPicker DTPicker3 
         Height          =   735
         Left            =   13080
         TabIndex        =   33
         Top             =   4680
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1296
         _Version        =   393216
         Format          =   16580609
         CurrentDate     =   42221
      End
      Begin VB.TextBox Text10 
         Height          =   735
         Left            =   10320
         TabIndex        =   32
         Top             =   4680
         Width           =   2535
      End
      Begin MSComCtl2.DTPicker DTPicker2 
         Height          =   735
         Left            =   13080
         TabIndex        =   30
         Top             =   3600
         Width           =   2175
         _ExtentX        =   3836
         _ExtentY        =   1296
         _Version        =   393216
         Format          =   16580609
         CurrentDate     =   42221
      End
      Begin VB.TextBox Text9 
         Height          =   735
         Left            =   10320
         TabIndex        =   29
         Top             =   3600
         Width           =   2535
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "NON-INDIAN"
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
         Left            =   10320
         TabIndex        =   27
         Top             =   2880
         Width           =   2055
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "INDIAN"
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
         TabIndex        =   26
         Top             =   2880
         Width           =   2055
      End
      Begin VB.ComboBox Combo2 
         Height          =   315
         ItemData        =   "DL.frx":0000
         Left            =   10320
         List            =   "DL.frx":001C
         TabIndex        =   23
         Text            =   "SELECT"
         Top             =   1440
         Width           =   1935
      End
      Begin VB.TextBox Text7 
         Height          =   975
         Left            =   10320
         TabIndex        =   21
         Top             =   240
         Width           =   4935
      End
      Begin VB.TextBox Text6 
         Height          =   1335
         Left            =   3000
         TabIndex        =   19
         Top             =   5400
         Width           =   3375
      End
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "DL.frx":0042
         Left            =   3000
         List            =   "DL.frx":0058
         TabIndex        =   17
         Text            =   "SELECT"
         Top             =   4680
         Width           =   2655
      End
      Begin VB.TextBox Text5 
         Height          =   495
         Left            =   3000
         TabIndex        =   15
         Top             =   3960
         Width           =   3495
      End
      Begin VB.TextBox Text4 
         Height          =   495
         Left            =   3000
         TabIndex        =   14
         Top             =   3240
         Width           =   3495
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   3000
         TabIndex        =   10
         Top             =   2520
         Width           =   3495
      End
      Begin MSComCtl2.DTPicker DTPicker1 
         Height          =   495
         Left            =   3000
         TabIndex        =   8
         Top             =   1800
         Width           =   2655
         _ExtentX        =   4683
         _ExtentY        =   873
         _Version        =   393216
         Format          =   16580609
         CurrentDate     =   42221
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   3000
         TabIndex        =   6
         Top             =   1080
         Width           =   3495
      End
      Begin VB.TextBox Text1 
         Height          =   495
         Left            =   3000
         TabIndex        =   4
         Top             =   360
         Width           =   3495
      End
      Begin VB.Label Label17 
         BackColor       =   &H00000000&
         Caption         =   "          APPLICANT                         SIGNATURE"
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
         Height          =   855
         Left            =   6840
         TabIndex        =   34
         Top             =   5640
         Width           =   3255
      End
      Begin VB.Label Label16 
         BackColor       =   &H00000000&
         Caption         =   "   Driving licence number              and date"
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
         Left            =   6840
         TabIndex        =   31
         Top             =   4680
         Width           =   3255
      End
      Begin VB.Label Label15 
         BackColor       =   &H00000000&
         Caption         =   "Leaner's Licence number             and date"
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
         Left            =   6840
         TabIndex        =   28
         Top             =   3600
         Width           =   3255
      End
      Begin VB.Label Label14 
         BackColor       =   &H00000000&
         Caption         =   "SELECT DECLARATIONOF CITIZENSHIP STATUS ?"
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
         Left            =   6840
         TabIndex        =   24
         Top             =   2160
         Width           =   6735
      End
      Begin VB.Label Label13 
         BackColor       =   &H00000000&
         Caption         =   "   BLOOD GROUP"
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
         Left            =   6840
         TabIndex        =   22
         Top             =   1440
         Width           =   2535
      End
      Begin VB.Label Label12 
         BackColor       =   &H00000000&
         Caption         =   "    TEMPORARY              ADDRESS"
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
         Left            =   6840
         TabIndex        =   20
         Top             =   360
         Width           =   2535
      End
      Begin VB.Line Line1 
         X1              =   6720
         X2              =   6720
         Y1              =   120
         Y2              =   6840
      End
      Begin VB.Label Label11 
         BackColor       =   &H00000000&
         Caption         =   "      PERMANENT                  ADDRESS"
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
         TabIndex        =   18
         Top             =   5400
         Width           =   2655
      End
      Begin VB.Label Label10 
         BackColor       =   &H00000000&
         Caption         =   "   QUALIFICATION"
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
         TabIndex        =   16
         Top             =   4680
         Width           =   2655
      End
      Begin VB.Label Label9 
         BackColor       =   &H00000000&
         Caption         =   "       E-MAIL ID"
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
         TabIndex        =   13
         Top             =   3960
         Width           =   2655
      End
      Begin VB.Label Label8 
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
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   3240
         Width           =   2655
      End
      Begin VB.Label Label6 
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
         TabIndex        =   9
         Top             =   2520
         Width           =   2655
      End
      Begin VB.Label Label5 
         BackColor       =   &H00000000&
         Caption         =   "   DATEOF BIRTH"
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
         Top             =   1800
         Width           =   2655
      End
      Begin VB.Label Label4 
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
         Height          =   615
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   2655
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "     FULL NAME"
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
         Width           =   2655
      End
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   495
      Left            =   9600
      TabIndex        =   11
      Top             =   5280
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "            DRIVING LICENCE"
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
      Height          =   735
      Left            =   5280
      TabIndex        =   1
      Top             =   1080
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "                               REGIONAL TRANSPORT OFFICE"
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
Attribute VB_Name = "Form9"
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
Text8.Text = ""
Text9.Text = ""
Text10.Text = ""
Text11.Text = ""
Text12.Text = ""
End Sub

Private Sub Command2_Click()
Form9.Hide
Form11.Show
End Sub

Private Sub Command3_Click()
Form9.Hide
Form4.Show
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii > 65 And KeyAscii < 90) Or (KeyAscii > 96 And KeyAscii < 122) Or kayacsii = 32 Or KeyAscii = 8 Then
ElseIf Text1.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text1.Text = ""
MsgBox "Enter valid name"
End If
End Sub

Private Sub Text10_KeyPress(KeyAscii As Integer)
If (KeyAscii > 65 And KeyAscii < 90) Or (KeyAscii > 96 And KeyAscii < 122) Or kayacsii = 32 Or KeyAscii = 8 Or (KeyAscii > 48 And KeyAscii < 57) Then
ElseIf Text10.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text10.Text = ""
MsgBox "Enter valid number"
End If
End Sub

Private Sub Text11_KeyPress(KeyAscii As Integer)
If (KeyAscii > 65 And KeyAscii < 90) Or (KeyAscii > 96 And KeyAscii < 122) Or kayacsii = 32 Or KeyAscii = 8 Then
ElseIf Text2.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text2.Text = ""
MsgBox "Enter valid name"
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii > 65 And KeyAscii < 90) Or (KeyAscii > 96 And KeyAscii < 122) Or kayacsii = 32 Or KeyAscii = 8 Then
ElseIf Text2.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text2.Text = ""
MsgBox "Enter valid name"
End If
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
If (KeyAscii > 65 And KeyAscii < 90) Or (KeyAscii > 96 And KeyAscii < 122) Or kayacsii = 32 Or KeyAscii = 8 Then
ElseIf Text3.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text3.Text = ""
MsgBox "Enter valid name"
End If
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
If (KeyAscii > 48 And KeyAscii < 57) Or KeyAscii = 8 Then
ElseIf Text4.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text4.Text = ""
MsgBox "Enter valid number"
End If
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
If (KeyAscii > 65 And KeyAscii < 90) Or (KeyAscii > 97 And KeyAscii < 122) Or kayacsii = 32 Or KeyAscii = 8 Or KeyAscii = 64 Or KeyAscii = 95 Or (KeyAscii > 48 And KeyAscii < 57) Then
ElseIf Text5.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
MsgBox "Enter valid mail id"
End If
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
If (KeyAscii > 65 And KeyAscii < 90) Or (KeyAscii > 96 And KeyAscii < 122) Or kayacsii = 32 Or KeyAscii = 8 Or (KeyAscii > 48 And KeyAscii < 57) Then
ElseIf Text3.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text3.Text = ""
MsgBox "Enter valid address"
End If
End Sub

Private Sub Text7_KeyPress(KeyAscii As Integer)
If (KeyAscii > 65 And KeyAscii < 90) Or (KeyAscii > 96 And KeyAscii < 122) Or kayacsii = 32 Or KeyAscii = 8 Or (KeyAscii > 48 And KeyAscii < 57) Then
ElseIf Text7.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text3.Text = ""
MsgBox "Enter valid address"
End If
End Sub

Private Sub Text8_KeyPress(KeyAscii As Integer)
If Option1 = True Then
Text7.Text = Indian
ElseIf Option2 = True Then
Text7.Text = Non - Indian
Else
MsgBox "Select nationality"
End If
End Sub

Private Sub Text9_KeyPress(KeyAscii As Integer)
If (KeyAscii > 65 And KeyAscii < 90) Or (KeyAscii > 96 And KeyAscii < 122) Or kayacsii = 32 Or KeyAscii = 8 Or (KeyAscii > 48 And KeyAscii < 57) Then
ElseIf Text9.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text9.Text = ""
MsgBox "Enter valid number"
End If
End Sub
