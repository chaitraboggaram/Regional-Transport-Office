VERSION 5.00
Object = "Word.Document.12"; "WINWORD.EXE"
Begin VB.Form Form12 
   BackColor       =   &H000080FF&
   Caption         =   "c"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form12"
   ScaleHeight     =   11010
   ScaleWidth      =   20370
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Finish"
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
      Left            =   9720
      TabIndex        =   4
      Top             =   8400
      Width           =   2775
   End
   Begin WordCtl.Document Document1 
      Height          =   4065
      Left            =   5715
      OleObjectBlob   =   "contact s.frx":0000
      TabIndex        =   3
      Top             =   4320
      Width           =   4020
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      Caption         =   "           contact us "
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
      Left            =   6360
      TabIndex        =   2
      Top             =   2760
      Width           =   3135
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "  Any changes to be done in the application or any verification."
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
      Left            =   3240
      TabIndex        =   1
      Top             =   1560
      Width           =   10095
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
Attribute VB_Name = "Form12"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
