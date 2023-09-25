VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H000080FF&
   Caption         =   "Form2"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form2"
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3960
      TabIndex        =   8
      ToolTipText     =   "clears data you typed."
      Top             =   7680
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6600
      TabIndex        =   7
      ToolTipText     =   "go back "
      Top             =   7680
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "login "
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
      Height          =   5055
      Left            =   3120
      TabIndex        =   2
      Top             =   1560
      Width           =   9975
      Begin VB.TextBox Text2 
         Height          =   615
         IMEMode         =   3  'DISABLE
         Left            =   4440
         PasswordChar    =   "*"
         TabIndex        =   6
         ToolTipText     =   "enter password."
         Top             =   3000
         Width           =   4455
      End
      Begin VB.TextBox Text1 
         Height          =   615
         Left            =   4440
         TabIndex        =   4
         ToolTipText     =   "enter your username."
         Top             =   1080
         Width           =   4455
      End
      Begin VB.Label Label3 
         BackColor       =   &H000080FF&
         Caption         =   "      Password"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         TabIndex        =   5
         Top             =   3000
         Width           =   2775
      End
      Begin VB.Label Label2 
         BackColor       =   &H000080FF&
         Caption         =   "        Username"
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
         Left            =   720
         TabIndex        =   3
         Top             =   1080
         Width           =   2775
      End
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Caption         =   "Login"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   9240
      TabIndex        =   1
      ToolTipText     =   "login for next process.you can login with same username and password next time.the information is saved. "
      Top             =   7680
      Width           =   1815
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
      Height          =   1095
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim con As New ADODB.Connection
Dim rs As New ADODB.Recordset

Private Sub Command1_Click()
If ((Text1.Text = "") And (Text2.Text = "")) Then
MsgBox "please login with username and password"
End If
Form3.Show
Form2.Hide
End Sub

Private Sub Command2_Click()
Text1.Text = ""
Text2.Text = ""
Form2.Hide
Form1.Show
End Sub

Private Sub Command3_Click()
Text1.Text = ""
Text2.Text = ""
End Sub

Private Sub Form_Load()
con.Open "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=False;Initial Catalog=rto;Data Source=RANJITHA-5F079E"
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii > 65 And KeyAscii < 90) Or (KeyAscii > 96 And KeyAscii < 122) Or (KeyAscii > 46 And KeyAscii < 55) Or KeyAscii = 8 Then
Else
KeyAscii = 0
MsgBox "Give validcharacters only"
End If

End Sub

