VERSION 5.00
Begin VB.Form Form4 
   BackColor       =   &H000080FF&
   Caption         =   "Form4"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
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
      Left            =   5760
      TabIndex        =   7
      ToolTipText     =   "click to go back."
      Top             =   8880
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
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
      Left            =   8520
      TabIndex        =   6
      ToolTipText     =   "click to proceed."
      Top             =   8880
      Width           =   2175
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H000080FF&
      Height          =   6015
      Left            =   1440
      TabIndex        =   2
      ToolTipText     =   "if your dealer select either of option and click next."
      Top             =   2640
      Width           =   13575
      Begin VB.OptionButton Option4 
         BackColor       =   &H00000000&
         Caption         =   "Driving Licence registration"
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
         Left            =   7560
         TabIndex        =   10
         Top             =   4200
         Width           =   5775
      End
      Begin VB.OptionButton Option3 
         BackColor       =   &H00000000&
         Caption         =   "Leaner's Licence Registration"
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
         Left            =   7560
         TabIndex        =   9
         Top             =   2040
         Width           =   5775
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H00000000&
         Caption         =   "Insurance Registration"
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
         Left            =   480
         TabIndex        =   5
         Top             =   4200
         Width           =   5295
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H00000000&
         Caption         =   "Vehicle Registration"
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
         Left            =   480
         TabIndex        =   4
         Top             =   2040
         Width           =   5295
      End
      Begin VB.Label Label4 
         BackColor       =   &H00000000&
         Caption         =   "          RETAILER"
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
         Left            =   8400
         TabIndex        =   8
         Top             =   360
         Width           =   3975
      End
      Begin VB.Line Line1 
         X1              =   6840
         X2              =   6840
         Y1              =   120
         Y2              =   6000
      End
      Begin VB.Label Label3 
         BackColor       =   &H00000000&
         Caption         =   "           DEALER"
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
         Left            =   1080
         TabIndex        =   3
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Label Label2 
      BackColor       =   &H000080FF&
      Caption         =   " Click on the Recruitments you need to perform ?"
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
      Height          =   615
      Left            =   1920
      TabIndex        =   1
      Top             =   1680
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
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
If Option1 = True Then
Form4.Hide
Form5.Show
ElseIf Option2 = True Then
Form4.Hide
Form7.Show
ElseIf Option3 = True Then
Form4.Hide
form8.Show
ElseIf Option4 = True Then
Form4.Hide
Form9.Show
Else
MsgBox "Click on any option to perform"
End If
End Sub


Private Sub Command2_Click()
Form4.Hide
Form3.Show
End Sub
