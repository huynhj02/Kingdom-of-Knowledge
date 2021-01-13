VERSION 5.00
Begin VB.Form frmAbout 
   ClientHeight    =   7200
   ClientLeft      =   15465
   ClientTop       =   15030
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   Picture         =   "frmAbout.frx":0000
   ScaleHeight     =   7200
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNext 
      Height          =   1575
      Left            =   6360
      Picture         =   "frmAbout.frx":AE7B
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5160
      Width           =   2175
   End
   Begin VB.CommandButton cmdMainMenu 
      Height          =   1575
      Left            =   9120
      Picture         =   "frmAbout.frx":F528
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5160
      Width           =   2175
   End
   Begin VB.Image Image2 
      Height          =   135
      Left            =   5160
      Picture         =   "frmAbout.frx":14F27
      Stretch         =   -1  'True
      Top             =   6120
      Width           =   135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "All Rights Reserved"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3000
      TabIndex        =   8
      Top             =   6120
      Width           =   2775
   End
   Begin VB.Label lbldescriptioninfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmAbout.frx":15303
      BeginProperty Font 
         Name            =   "DK Crayon Crumble"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3735
      Left            =   480
      TabIndex        =   7
      Top             =   1560
      Width           =   6255
   End
   Begin VB.Image Image1 
      Height          =   1215
      Left            =   600
      Picture         =   "frmAbout.frx":1542D
      Stretch         =   -1  'True
      Top             =   5400
      Width           =   2055
   End
   Begin VB.Label lblDescription 
      BackStyle       =   0  'Transparent
      Caption         =   "Game version 1.10.2"
      BeginProperty Font 
         Name            =   "DK Crayon Crumble"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   1200
      TabIndex        =   6
      Top             =   960
      Width           =   4695
   End
   Begin VB.Label lblDevelopers 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Developers"
      BeginProperty Font 
         Name            =   "DK Crayon Crumble"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   6960
      TabIndex        =   4
      Top             =   720
      Width           =   4335
   End
   Begin VB.Label lblName2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Domenico Didiano"
      BeginProperty Font 
         Name            =   "DK Crayon Crumble"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   9360
      TabIndex        =   3
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Label lblName1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Jimmy Huynh"
      BeginProperty Font 
         Name            =   "DK Crayon Crumble"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   7320
      TabIndex        =   2
      Top             =   3720
      Width           =   1335
   End
   Begin VB.Image imgDomenico 
      Height          =   2175
      Left            =   9360
      Picture         =   "frmAbout.frx":170F8
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Image imgJimmy 
      Height          =   2175
      Left            =   6840
      Picture         =   "frmAbout.frx":1F120
      Stretch         =   -1  'True
      Top             =   1440
      Width           =   2175
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "About"
      BeginProperty Font 
         Name            =   "DK Crayon Crumble"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   4800
      TabIndex        =   1
      Top             =   120
      Width           =   2775
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Jimmy Huynh and Domenico Didiano
'January 22, 2018
'ICS ISU - Kingdom of Knowledge - About Form
'The about form, introducing the developers, the version of Kingdom of Knowledge, and its purpose.
Option Explicit

Private Sub cmdMainMenu_Click()

    frmAbout.Hide
    frmMainMenu.Show

End Sub

Private Sub cmdNext_Click()

    frmAbout.Hide
    frmDescription.Show
    
End Sub

