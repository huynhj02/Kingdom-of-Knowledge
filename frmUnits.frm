VERSION 5.00
Begin VB.Form frmUnits 
   Caption         =   "Kingdom of Knowledge"
   ClientHeight    =   7200
   ClientLeft      =   2055
   ClientTop       =   15225
   ClientWidth     =   11955
   FillColor       =   &H00FFFFFF&
   Icon            =   "frmUnits.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmUnits.frx":424A
   ScaleHeight     =   7200
   ScaleWidth      =   11955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdTrig 
      Height          =   1575
      Left            =   6705
      Picture         =   "frmUnits.frx":F0C5
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4500
      Width           =   2895
   End
   Begin VB.CommandButton cmdGeometry 
      Height          =   1575
      Left            =   2385
      Picture         =   "frmUnits.frx":14C6C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4500
      Width           =   2895
   End
   Begin VB.CommandButton cmdQuadratics 
      Height          =   1575
      Left            =   6675
      Picture         =   "frmUnits.frx":1B789
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   2340
      Width           =   2895
   End
   Begin VB.CommandButton cmdAlgebra 
      DisabledPicture =   "frmUnits.frx":22219
      Height          =   1575
      Left            =   2355
      Picture         =   "frmUnits.frx":2924D
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2340
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
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   9240
      TabIndex        =   5
      Top             =   480
      Width           =   2295
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Units"
      BeginProperty Font 
         Name            =   "DK Crayon Crumble"
         Size            =   72
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   3510
      TabIndex        =   0
      Top             =   360
      Width           =   4935
   End
End
Attribute VB_Name = "frmUnits"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Jimmy Huynh and Domenico Didiano
'January 22, 2018
'ICS ISU - Kingdom of Knowledge - Units
'A form which displays all the units that the user can play.
Option Explicit

Public intAnswerStreak As Integer
Public intPoints As Integer

Private Sub cmdAlgebra_Click()

    frmUnits.Hide
    frmAlgebra.Show
    
    cmdAlgebra.Enabled = False
        
End Sub

Private Sub cmdQuadratics_Click()

    frmUnits.Hide
    frmQuadratics.Show

    cmdQuadratics.Enabled = False
    
    
End Sub

Private Sub cmdGeometry_Click()

    frmUnits.Hide
    frmGeometry.Show

    cmdGeometry.Enabled = False
    
End Sub

Private Sub cmdTrig_Click()

    frmUnits.Hide
    frmTrigonometry.Show
    
    cmdTrig.Enabled = False
    

End Sub

Private Sub Form_Load()

    lblPoints.Caption = "Points: " & intPoints

End Sub
