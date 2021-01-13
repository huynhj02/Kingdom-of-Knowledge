VERSION 5.00
Begin VB.Form frmFactor2 
   Caption         =   "Kingdom of Knowledge"
   ClientHeight    =   7200
   ClientLeft      =   -1545
   ClientTop       =   6525
   ClientWidth     =   11955
   Icon            =   "frmFactor2.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmFactor2.frx":424A
   ScaleHeight     =   7200
   ScaleWidth      =   11955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdhelp 
      Height          =   975
      Left            =   9480
      Picture         =   "frmFactor2.frx":F0C5
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5760
      Width           =   1815
   End
   Begin VB.OptionButton optD 
      BackColor       =   &H00C0FFFF&
      Caption         =   "D k(k + 5) + 4"
      BeginProperty Font 
         Name            =   "BrookeShappell8"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      TabIndex        =   5
      Top             =   3720
      Width           =   2775
   End
   Begin VB.OptionButton optC 
      BackColor       =   &H00C0FFFF&
      Caption         =   "C k(k + 4) (k + 1)"
      BeginProperty Font 
         Name            =   "BrookeShappell8"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6480
      TabIndex        =   4
      Top             =   2640
      Width           =   2775
   End
   Begin VB.OptionButton optB 
      BackColor       =   &H00C0FFFF&
      Caption         =   "B (k + 4) (k + 1)"
      BeginProperty Font 
         Name            =   "BrookeShappell8"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   3
      Top             =   3720
      Width           =   2775
   End
   Begin VB.OptionButton optA 
      BackColor       =   &H00C0FFFF&
      Caption         =   "A (k + 5) (k + 4)"
      BeginProperty Font 
         Name            =   "BrookeShappell8"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   2
      Top             =   2640
      Width           =   2775
   End
   Begin VB.CommandButton cmdSubmit 
      Height          =   1575
      Left            =   4440
      Picture         =   "frmFactor2.frx":1382A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Label lblQuestion 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Factor k**2 + 5k + 4"
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
      Height          =   1695
      Left            =   1320
      TabIndex        =   1
      Top             =   1440
      Width           =   9015
   End
End
Attribute VB_Name = "frmFactor2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
 'Domenico Didiano and Jimmy Huynh
 '1/24/2018
 'ICS ISU Kingdom of Knowledge Second Factor Question
 ' The Second of three questions of the Factoring Lesson
 Option Explicit
 
 
Private Sub cmdSubmit_Click()
    
    'Check if blank
   If optA.Value = False And optB.Value = False And optC.Value = False And optD.Value = False Then
    
        MsgBox "Please select an answer.", vbOKOnly, "Select an answer"
   Else
         'Wizard Power Up
        If optB.Value = True And frmRoles.intMultiply > 1 Then
             
             MsgBox "Correct, the equation can be factored by finding two numbers that sum up to the second coefficient and multiply into the number without a variable.", _
             vbOKOnly, "Correct Answer (Wizard Multiplier)"
         
             frmRoles.intMultiply = frmRoles.intMultiply - 1
         
         'No Power Up
        ElseIf optB.Value = True Then
         
             MsgBox "Correct, the equation can be factored by finding two numbers that sum up to the second coefficient and multiply into the number without a variable.", _
             vbOKOnly, "Correct Answer"
         
             frmUnits.intAnswerStreak = frmUnits.intAnswerStreak + 1
             frmUnits.intPoints = frmUnits.intPoints + 1
             
         'Knight Power Up
        ElseIf optB.Value = False And frmRoles.intBlock > 1 Then
             
             MsgBox "Incorrect, the equation can be factored by finding two numbers that sum up to the second coefficient and multiply into the number without a variable, The first step would be k**2 + 4k + k + 4", _
             vbOKOnly, "Incorrect Answer"
             MsgBox "You have succesfully blocked the incorrect answer, it will not affect your streak.", vbOKOnly, "Incorrect Answer Blocked"
             
             frmRoles.intBlock = frmRoles.intBlock - 1
             
         Else
             MsgBox "Incorrect, the equation can be factored by finding two numbers that sum up to the second coefficient and multiply into the number without a variable, The first step would be k**2 + 4k + k + 4", _
             vbOKOnly, "Incorrect Answer"
             frmUnits.intAnswerStreak = 0
             
         End If
         
         frmFactor2.Hide
         frmFactor3.Show
         
    End If
 

 
End Sub
 

