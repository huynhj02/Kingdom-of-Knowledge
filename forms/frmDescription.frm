VERSION 5.00
Begin VB.Form frmDescription 
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   12255
   ClientTop       =   -12045
   ClientWidth     =   11970
   LinkTopic       =   "Form1"
   Picture         =   "frmDescription.frx":0000
   ScaleHeight     =   7200
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      Height          =   1095
      Left            =   9360
      Picture         =   "frmDescription.frx":AE7B
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   2175
   End
   Begin VB.Label lblDescriptiontext 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   $"frmDescription.frx":B2C6
      BeginProperty Font 
         Name            =   "DK Crayon Crumble"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   4695
      Left            =   480
      TabIndex        =   1
      Top             =   1320
      Width           =   11055
   End
   Begin VB.Label lblDescription 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Game Description"
      BeginProperty Font 
         Name            =   "DK Crayon Crumble"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   3720
      TabIndex        =   0
      Top             =   480
      Width           =   4695
   End
End
Attribute VB_Name = "frmDescription"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Jimmy Huynh and Domenico Didiano
'January 22, 2018
'ICS ISU - Kingdom of Knowledge - Description Form
'An extension of the about form, explaning the premise of the game, including a story behind the characters.
Option Explicit

Private Sub cmdBack_Click()

frmDescription.Hide
frmAbout.Show

End Sub

