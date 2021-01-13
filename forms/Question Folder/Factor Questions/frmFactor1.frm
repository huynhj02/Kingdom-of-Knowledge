VERSION 5.00
Begin VB.Form frmFactor1 
   Caption         =   "Kingdom of Knowledge"
   ClientHeight    =   7200
   ClientLeft      =   -1545
   ClientTop       =   6525
   ClientWidth     =   11955
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   30
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmFactor1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmFactor1.frx":424A
   ScaleHeight     =   7200
   ScaleWidth      =   11955
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAnswer 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2880
      TabIndex        =   3
      Top             =   3240
      Width           =   6015
   End
   Begin VB.CommandButton cmdhelp 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   9480
      Picture         =   "frmFactor1.frx":F0C5
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5760
      Width           =   1815
   End
   Begin VB.CommandButton cmdSubmit 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4440
      Picture         =   "frmFactor1.frx":1382A
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Label lblQuestion 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Factor j**2 + 12j +27 Make sure it looks like this Spaces and all or it maybe marked as incorrect (x + ?) (x + ?). "
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
      Height          =   2895
      Left            =   1440
      TabIndex        =   0
      Top             =   720
      Width           =   9015
   End
End
Attribute VB_Name = "frmFactor1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Jimmy Huynh and Domenico Didiano
'January 22, 2018
'ICS ISU - Kingdom of Knowledge - Factor Question 1
'The first of three consecutive questions on the lesson about Factored form.
Option Explicit

    

Private Sub cmdSubmit_Click()

    Dim strAnswer As String
    
    Const strCorrect As String = "(j + 9) (j + 3)"
    
    strAnswer = txtAnswer.Text
    
    'Check if blank
   If strAnswer = 0 Then
    
        MsgBox "Please select an answer.", vbOKOnly, "Select an answer"
        
    
    'Wizard Power Up
   ElseIf strAnswer = strCorrect And frmRoles.intMultiply > 1 Then
        
        MsgBox "Correct, the equation can be factored by finding two numbers that sum up to the second coefficient and multiply into the number without a variable.", _
        vbOKOnly, "Correct Answer (Wizard Multiplier)"
    
        frmRoles.intMultiply = frmRoles.intMultiply - 1
    
    'No Power Up
   ElseIf strAnswer = strCorrect Then
    
        MsgBox "Correct, the equation can be factored by finding two numbers that sum up to the second coefficient and multiply into the number without a variable.", _
        vbOKOnly, "Correct Answer"
    
        frmUnits.intAnswerStreak = frmUnits.intAnswerStreak + 1
        frmUnits.intPoints = frmUnits.intPoints + 1
        
    'Knight Power Up
   ElseIf strAnswer <> strCorrect And frmRoles.intBlock > 1 Then
        
        MsgBox "Incorrect, the equation can be factored by finding two numbers that sum up to the second coefficient and multiply into the number without a variable, The first step would be j**2 + 9j + 3j + 27. ", _
        vbOKOnly, "Incorrect Answer"
        MsgBox "You have succesfully blocked the incorrect answer, it will not affect your streak.", vbOKOnly, "Incorrect Answer Blocked"
        
        frmRoles.intBlock = frmRoles.intBlock - 1
        
    Else
        MsgBox "Incorrect, the equation can be factored by finding two numbers that sum up to the second coefficient and multiply into the number without a variable, The first step would be j**2 + 9j + 3j + 27.", _
        vbOKOnly, "Incorrect Answer"
        frmUnits.intAnswerStreak = 0
        
    End If
 
    frmFactor1.Hide
    frmFactor2.Show
 
End Sub







     





