VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H000080FF&
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   FillColor       =   &H8000000F&
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   24
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture2 
      Height          =   4095
      Left            =   5040
      Picture         =   "main frame.frx":0000
      ScaleHeight     =   4035
      ScaleWidth      =   5955
      TabIndex        =   4
      Top             =   2040
      Width           =   6015
   End
   Begin VB.PictureBox Picture1 
      ForeColor       =   &H000040C0&
      Height          =   1335
      Left            =   960
      Picture         =   "main frame.frx":5037
      ScaleHeight     =   1275
      ScaleWidth      =   1395
      TabIndex        =   3
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00000000&
      Caption         =   "Proceed"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   12960
      MaskColor       =   &H00000000&
      TabIndex        =   2
      ToolTipText     =   "click to follow next process"
      Top             =   8280
      Width           =   2175
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   " KARNATAKA"
      ForeColor       =   &H000080FF&
      Height          =   495
      Left            =   9480
      TabIndex        =   1
      Top             =   6840
      Width           =   3375
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"main frame.frx":5EFD
      ForeColor       =   &H000080FF&
      Height          =   1335
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   16215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form1.Hide
Form2.Show

End Sub

