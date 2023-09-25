VERSION 5.00
Object = "Word.Document.12"; "WINWORD.EXE"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form11 
   BackColor       =   &H000080FF&
   Caption         =   "Form11"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form11"
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Finish"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   14400
      TabIndex        =   8
      Top             =   8640
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   975
      Left            =   10920
      TabIndex        =   7
      ToolTipText     =   "specify correct name with initial."
      Top             =   7440
      Width           =   2895
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2640
      TabIndex        =   4
      Top             =   2280
      Width           =   1455
   End
   Begin MSComCtl2.DTPicker DTPicker1 
      Height          =   375
      Left            =   2640
      TabIndex        =   2
      Top             =   1560
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      _Version        =   393216
      Format          =   16580609
      CurrentDate     =   42221
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "     Applicant name "
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
      Left            =   10920
      TabIndex        =   6
      Top             =   6600
      Width           =   2895
   End
   Begin WordCtl.Document Document1 
      Height          =   2535
      Left            =   2880
      OleObjectBlob   =   "agreement.frx":0000
      TabIndex        =   5
      Top             =   3480
      Width           =   12600
   End
   Begin VB.Label Label3 
      BackColor       =   &H000080FF&
      Caption         =   "   PLACE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   840
      TabIndex        =   3
      Top             =   2280
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      Caption         =   "   DATE:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   1560
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "                              REGIONAL TRANSPORT OFFICE"
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
Attribute VB_Name = "Form11"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form11.Hide
Form12.Show
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
If (KeyAscii > 65 And KeyAscii < 90) Or (KeyAscii > 96 And KeyAscii < 122) Or kayacsii = 32 Or KeyAscii = 8 Then
ElseIf Text1.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text1.Text = ""
MsgBox "Enter only characters"
End If
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
If (KeyAscii > 65 And KeyAscii < 90) Or (KeyAscii > 96 And KeyAscii < 122) Or kayacsii = 32 Or KeyAscii = 8 Then
ElseIf Text2.Text = "" Then
KeyAscii = 0
MsgBox "Enter all the fields"
Else
Text2.Text = ""
MsgBox "Enter only characters"
End If
End Sub
