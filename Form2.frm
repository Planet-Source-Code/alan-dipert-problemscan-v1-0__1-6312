VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ServerScan v1.0 Info"
   ClientHeight    =   6570
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6570
   ScaleWidth      =   5550
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton Command2 
      BackColor       =   &H000080FF&
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4680
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   6120
      Width           =   735
   End
   Begin VB.Label Label17 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "http://www.neologic.net/ps"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2640
      TabIndex        =   17
      Top             =   480
      Width           =   2895
   End
   Begin VB.Label Label16 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "              version 1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label Label15 
      BackColor       =   &H00000000&
      Caption         =   "- User defined scan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   14
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label14 
      BackColor       =   &H00000000&
      Caption         =   "- PWS Scanner"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   6000
      Width           =   1335
   End
   Begin VB.Label Label13 
      BackColor       =   &H00000000&
      Caption         =   "- Removal of trojans if already installed"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   5760
      Width           =   3375
   End
   Begin VB.Label Label12 
      BackColor       =   &H00000000&
      Caption         =   "- Downloadable trojan database "
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   5520
      Width           =   2775
   End
   Begin VB.Label Label11 
      BackColor       =   &H00000000&
      Caption         =   "- Expanded trojan list"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   5280
      Width           =   1935
   End
   Begin VB.Line Line3 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   5400
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Label Label10 
      BackColor       =   &H00000000&
      Caption         =   "What to look for in v2.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   4920
      Width           =   2415
   End
   Begin VB.Label Label9 
      BackColor       =   &H00000000&
      Caption         =   "- Prompts for deletion if a trojan server is discovered"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   4560
      Width           =   4575
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   5400
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Label Label8 
      BackColor       =   &H00000000&
      Caption         =   $"Form2.frx":0442
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   120
      TabIndex        =   6
      Top             =   3600
      Width           =   5295
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   120
      X2              =   5400
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Label Label7 
      BackColor       =   &H00000000&
      Caption         =   "ProblemScan v1.0 Features"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3240
      Width           =   2415
   End
   Begin VB.Label Label6 
      BackColor       =   &H00000000&
      Caption         =   $"Form2.frx":0507
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   5295
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      Caption         =   $"Form2.frx":066D
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   120
      TabIndex        =   3
      Top             =   1080
      Width           =   5295
   End
   Begin VB.Label Label4 
      BackColor       =   &H00000000&
      Caption         =   "About ProblemScan v1.0"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   720
      Width           =   2055
   End
   Begin VB.Label Label3 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "by Alan Dipert  alan@neologic.net"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   255
      Left            =   2640
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      Caption         =   "Scan"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   1200
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label2 
      BackColor       =   &H00000000&
      Caption         =   "Problem"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Unload Me
End Sub

