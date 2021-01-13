VERSION 5.00
Begin VB.Form frmAlgebra 
   Caption         =   "Kingdom of Knowledge"
   ClientHeight    =   7200
   ClientLeft      =   4665
   ClientTop       =   -6510
   ClientWidth     =   11955
   Icon            =   "frmAlgebra.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmAlgebra.frx":424A
   ScaleHeight     =   7200
   ScaleWidth      =   11955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdElim 
      Height          =   1575
      Left            =   6585
      Picture         =   "frmAlgebra.frx":F0C5
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2803
      Width           =   2895
   End
   Begin VB.CommandButton cmdSub 
      DisabledPicture =   "frmAlgebra.frx":15BF6
      Height          =   1575
      Left            =   2475
      Picture         =   "frmAlgebra.frx":1C878
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
      TabIndex        =   3
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Algebra/Linear Systems"
      BeginProperty Font 
         Name            =   "DK Crayon Crumble"
         Size            =   51.75
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
      Top             =   960
      Width           =   9735
   End
End
Attribute VB_Name = "frmAlgebra"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Jimmy Huynh and Domenico Didiano
'January 22, 2018
'ICS ISU - Kingdom of Knowledge - Algebra Chapters Form
'The user can choose which chapter of algebra/linear systems that they would like to complete.
Option Explicit

Private Sub cmdElim_Click()
    
    frmAlgebra.Hide
    frmElimLesson.Show


End Sub

Private Sub cmdSub_Click()
    
    frmAlgebra.Hide
    frmSubLesson.Show
    
End Sub

Private Sub Form_Load()

    lblPoints.Caption = "Points: " & frmUnits.intPoints

End Sub
