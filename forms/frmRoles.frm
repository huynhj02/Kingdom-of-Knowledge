VERSION 5.00
Begin VB.Form frmRoles 
   Caption         =   "Kingdom of Knowledge"
   ClientHeight    =   7200
   ClientLeft      =   22065
   ClientTop       =   2970
   ClientWidth     =   11970
   Icon            =   "frmRoles.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmRoles.frx":424A
   ScaleHeight     =   7200
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdMainMenu 
      Height          =   1335
      Left            =   4598
      Picture         =   "frmRoles.frx":F0C5
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
      Width           =   2775
   End
   Begin VB.CommandButton cmdPeasant 
      Height          =   1335
      Left            =   7838
      Picture         =   "frmRoles.frx":14AC4
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   2775
   End
   Begin VB.CommandButton cmdWizard 
      Height          =   1335
      Left            =   4598
      Picture         =   "frmRoles.frx":19F5C
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3720
      Width           =   2775
   End
   Begin VB.CommandButton cmdKnight 
      Height          =   1335
      Left            =   1358
      Picture         =   "frmRoles.frx":1F0E0
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   3720
      Width           =   2775
   End
   Begin VB.Label lblPeasant 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Will you be the hard working peasant who gets a point bonus after completing all his work"
      BeginProperty Font 
         Name            =   "BrookeShappell8"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1455
      Left            =   7800
      TabIndex        =   6
      Top             =   2160
      Width           =   3015
   End
   Begin VB.Label lblWizard 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Will you be the Wizard who uses his magical powers to gain a two times point multiplier"
      BeginProperty Font 
         Name            =   "BrookeShappell8"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   4680
      TabIndex        =   5
      Top             =   2160
      Width           =   2775
   End
   Begin VB.Label lblKnight 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Will you be the Knight who uses his shield to defend himself from incorrect answers"
      BeginProperty Font 
         Name            =   "BrookeShappell8"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   1440
      TabIndex        =   4
      Top             =   2160
      Width           =   2655
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Choose Your Role!"
      BeginProperty Font 
         Name            =   "DK Crayon Crumble"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   2520
      TabIndex        =   3
      Top             =   600
      Width           =   7095
   End
End
Attribute VB_Name = "frmRoles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Jimmy Huynh and Domenico Didiano
'January 22, 2018
'ICS ISU - Kingdom of Knowledge - Roles
'The form where the user is able to choose their role. This will decide what powerup they will get and when.
Option Explicit

Public intRole As Integer
Public intBlock As Integer
Public intMultiply As Integer


Private Sub cmdKnight_Click()

    intRole = 1
    
    frmRoles.Hide
    frmUnits.Show
    
End Sub

Private Sub cmdWizard_Click()

    intRole = 2
    
    frmRoles.Hide
    frmUnits.Show

End Sub

Private Sub cmdPeasant_Click()

    intRole = 3
    
    frmRoles.Hide
    frmUnits.Show

End Sub

Private Sub cmdMainMenu_Click()

    frmRoles.Hide
    frmMainMenu.Show
    
End Sub

