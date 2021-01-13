VERSION 5.00
Begin VB.Form frmSubLesson 
   Caption         =   "Kingdom of Knowledge"
   ClientHeight    =   7200
   ClientLeft      =   3285
   ClientTop       =   2400
   ClientWidth     =   11955
   Icon            =   "frmSubLesson.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmSubLesson.frx":424A
   ScaleHeight     =   7200
   ScaleWidth      =   11955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNext 
      Height          =   1575
      Left            =   8640
      Picture         =   "frmSubLesson.frx":F0C5
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5280
      Width           =   2775
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "The Method of Substitution"
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
      Height          =   1215
      Left            =   330
      TabIndex        =   0
      Top             =   360
      Width           =   11295
   End
End
Attribute VB_Name = "frmSubLesson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Jimmy Huynh and Domenico Didiano
'January 22, 2018
'ICS ISU - Kingdom of Knowledge - About Form
'A lesson form, teaching the user the method of substitution.
Option Explicit

Private Sub cmdNext_Click()

    frmSubLesson.Hide
    frmSubExample.Show
    
End Sub
