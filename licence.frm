VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form6 
   BackColor       =   &H000080FF&
   Caption         =   "Form6"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form6"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "exit"
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
      Left            =   9840
      TabIndex        =   21
      Top             =   8520
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "next"
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
      Left            =   8160
      TabIndex        =   19
      Top             =   8520
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "save"
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
      Left            =   6240
      TabIndex        =   18
      Top             =   8520
      Width           =   1695
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
      Left            =   1800
      TabIndex        =   2
      Top             =   2040
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
         Format          =   16580609
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
         Width           =   3015
      End
      Begin VB.TextBox Text4 
         Height          =   375
         Left            =   3000
         TabIndex        =   11
         Top             =   4440
         Width           =   3375
      End
      Begin VB.TextBox Text3 
         Height          =   495
         Left            =   3000
         TabIndex        =   10
         Top             =   3480
         Width           =   2775
      End
      Begin VB.TextBox Text2 
         Height          =   495
         Left            =   3000
         TabIndex        =   7
         Top             =   2520
         Width           =   2775
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
      Caption         =   "   REGIONALTRANSPORT OFFICE"
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
      Height          =   735
      Left            =   3960
      TabIndex        =   0
      Top             =   240
      Width           =   8655
   End
End
Attribute VB_Name = "Form6"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
