VERSION 5.00
Begin VB.Form frmTrigonometry 
   Caption         =   "Kingdom of Knowledge"
   ClientHeight    =   7200
   ClientLeft      =   4665
   ClientTop       =   -6510
   ClientWidth     =   11955
   Icon            =   "frmTrigonometry.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmTrigonometry.frx":424A
   ScaleHeight     =   7200
   ScaleWidth      =   11955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTangent 
      Height          =   1575
      Left            =   7680
      Picture         =   "frmTrigonometry.frx":F0C5
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2813
      Width           =   2895
   End
   Begin VB.CommandButton cmdCosine 
      Height          =   1575
      Left            =   4320
      Picture         =   "frmTrigonometry.frx":13BC9
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2813
      Width           =   2895
   End
   Begin VB.CommandButton cmdSine 
      DisabledPicture =   "frmTrigonometry.frx":18321
      Height          =   1575
      Left            =   960
      Picture         =   "frmTrigonometry.frx":1EFA3
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2822
      Width           =   2895
   End
   Begin VB.Label lblPoints 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "DK Crayon Crumble"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   9240
      TabIndex        =   4
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Trigonometry"
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
      Height          =   1455
      Left            =   1110
      TabIndex        =   0
      Top             =   480
      Width           =   9735
   End
End
Attribute VB_Name = "frmTrigonometry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Jimmy Huynh and Domenico Didiano
'January 22, 2018
'ICS ISU - Kingdom of Knowledge - Trigonometry Form
'The user can choose which chapter of trigonometry that they would like to complete.
Option Explicit


Private Sub cmdCosine_Click()

    frmTrigonometry.Hide
    frmCosLesson.Show
    

End Sub

Private Sub cmdSine_Click()

    frmTrigonometry.Hide
    frmSineLesson.Show

End Sub

Private Sub cmdTangent_Click()

    frmTrigonometry.Hide
    frmTanLesson.Show


End Sub

Private Sub Form_Load()

    lblPoints.Caption = "Points: " & frmUnits.intPoints

End Sub

