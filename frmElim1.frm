VERSION 5.00
Begin VB.Form frmElim1 
   Caption         =   "Kingdom of Knowledge"
   ClientHeight    =   7200
   ClientLeft      =   -1545
   ClientTop       =   6525
   ClientWidth     =   11955
   Icon            =   "frmElim1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmElim1.frx":424A
   ScaleHeight     =   7200
   ScaleWidth      =   11955
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer tmrCountdown 
      Interval        =   1000
      Left            =   6840
      Top             =   960
   End
   Begin VB.CommandButton cmdhelp 
      Height          =   855
      Left            =   9720
      Picture         =   "frmElim1.frx":F0C5
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5880
      Width           =   1575
   End
   Begin VB.OptionButton optD 
      BackColor       =   &H00C0FFFF&
      Caption         =   "D (2,2)"
      BeginProperty Font 
         Name            =   "BrookeShappell8"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      TabIndex        =   5
      Top             =   4080
      Width           =   2775
   End
   Begin VB.OptionButton optC 
      BackColor       =   &H00C0FFFF&
      Caption         =   "C (17,2)"
      BeginProperty Font 
         Name            =   "BrookeShappell8"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   6360
      TabIndex        =   4
      Top             =   3120
      Width           =   2775
   End
   Begin VB.OptionButton optB 
      BackColor       =   &H00C0FFFF&
      Caption         =   "B (17,31)"
      BeginProperty Font 
         Name            =   "BrookeShappell8"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   3
      Top             =   4080
      Width           =   2775
   End
   Begin VB.OptionButton optA 
      BackColor       =   &H00C0FFFF&
      Caption         =   "A (0,0)"
      BeginProperty Font 
         Name            =   "BrookeShappell8"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   2760
      TabIndex        =   2
      Top             =   3120
      Width           =   2775
   End
   Begin VB.CommandButton cmdSubmit 
      Height          =   1335
      Left            =   4920
      Picture         =   "frmElim1.frx":1382A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Image imgClock 
      Height          =   1125
      Left            =   5400
      Picture         =   "frmElim1.frx":188C1
      Top             =   480
      Width           =   1125
   End
   Begin VB.Label lblQuestion 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Find the point of intersection of the lines: 50x + 35y = 170 and 140x + 35y = 350."
      BeginProperty Font 
         Name            =   "DK Crayon Crumble"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1575
      Left            =   2400
      TabIndex        =   1
      Top             =   1680
      Width           =   7335
   End
End
Attribute VB_Name = "frmElim1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Jimmy Huynh and Domenico Didiano
'January 24, 2018
'ICS ISU - Kingdom of Knowledge - Elimination Question #1
'The first of three consecutive questions on the lesson about the method of elimination.
Option Explicit

Dim intCount As Integer

Private Sub cmdSubmit_Click()
    
        'Check if blank
    If optA.Value = False And optB.Value = False And optC.Value = False And optD.Value = False Then
    
        MsgBox "Please select an answer.", vbOKOnly, "Select an answer"
        
    Else
    
        'Wizard Power Up
        If optD.Value = True And frmRoles.intMultiply > 1 Then
            
            MsgBox "Correct, in this question, both equations have the same coefficients in front of their y values." _
            & "This allows for you to easily eliminate y and solve for x. Once you have found x, you can re input it into" _
            & "the equation and solve for y." _
            , vbOKOnly, "Correct Answer (Wizard Multiplier)"
        
            frmRoles.intMultiply = frmRoles.intMultiply - 1
        
        'No Power Up
        ElseIf optD.Value = True Then
        
            MsgBox "Correct, in this question, both equations have the same coefficients in front of their y values." _
            & "This allows for you to easily eliminate y and solve for x. Once you have found x, you can re input it into" _
            & "the equation and solve for y." _
            , vbOKOnly, "Correct Answer"
        
            frmUnits.intAnswerStreak = frmUnits.intAnswerStreak + 1
            frmUnits.intPoints = frmUnits.intPoints + 1
            
        'Knight Power Up
        ElseIf optD.Value = False And frmRoles.intBlock > 1 Then
            
            MsgBox "Incorrect, in this question, both equations have the same coefficients in front of their y values." _
            & "This allows for you to easily eliminate y and solve for x. Once you have found x, you can re input it into" _
            & "the equation and solve for y. The correct answer was D:(2,2)" _
            , vbOKOnly, "Incorrect Answer"
            MsgBox "You have succesfully blocked the incorrect answer, it will not affect your streak.", vbOKOnly, "Incorrect Answer Blocked"
            
            frmRoles.intBlock = frmRoles.intBlock - 1
            
        'Incorrect Answer
        Else
            MsgBox "Incorrect, in this question, both equations have the same coefficients in front of their y values." _
            & "This allows for you to easily eliminate y and solve for x. Once you have found x, you can re input it into" _
            & "the equation and solve for y. The correct answer was D:(2,2)" _
            , vbOKOnly, "Incorrect Answer"
            frmUnits.intAnswerStreak = 0
            
        End If
    
        frmElim1.Hide
        frmElim2.Show
        
    End If

End Sub

Private Sub TmrCountdown_Timer()
    
    intCount = intCount + 1

    Do While intCount = 15
        
        MsgBox "Time is up! In this question, both equations have the same coefficients in front of their y values." _
        & "This allows for you to easily eliminate y and solve for x. Once you have found x, you can re input it into" _
        & "the equation and solve for y. The correct answer was D:(2,2)", vbOKOnly, "Timer"
        
        frmElim1.Hide
        
        frmUnits.intAnswerStreak = 0
        
        intCount = intCount + 1
        
        frmElim2.Show
        
    Loop
    
End Sub

Private Sub cmdhelp_Click()

    frmHelp.Show

End Sub
