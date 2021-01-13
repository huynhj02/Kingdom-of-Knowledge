VERSION 5.00
Begin VB.Form frmSubBonus 
   Caption         =   "Kingdom of Knowledge"
   ClientHeight    =   7200
   ClientLeft      =   -1740
   ClientTop       =   -9075
   ClientWidth     =   11970
   Icon            =   "frmSubBonus.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmSubBonus.frx":424A
   ScaleHeight     =   7200
   ScaleWidth      =   11970
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer TmrCountdown 
      Interval        =   1000
      Left            =   6840
      Top             =   840
   End
   Begin VB.CommandButton cmdExit 
      Height          =   855
      Left            =   9720
      Picture         =   "frmSubBonus.frx":F0C5
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton cmdSubmit 
      Height          =   1335
      Left            =   4920
      Picture         =   "frmSubBonus.frx":1382A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5400
      Width           =   2175
   End
   Begin VB.OptionButton optD 
      BackColor       =   &H00C0FFFF&
      Caption         =   "D 550"
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
      Top             =   4320
      Width           =   2775
   End
   Begin VB.OptionButton optC 
      BackColor       =   &H00C0FFFF&
      Caption         =   "C 275"
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
      Top             =   3360
      Width           =   2775
   End
   Begin VB.OptionButton optB 
      BackColor       =   &H00C0FFFF&
      Caption         =   "B 575"
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
      Top             =   4320
      Width           =   2775
   End
   Begin VB.OptionButton optA 
      BackColor       =   &H00C0FFFF&
      Caption         =   "A 250"
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
      Top             =   3360
      Width           =   2775
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "BONUS ROUND!!!"
      BeginProperty Font 
         Name            =   "Milk Mustache BB"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   855
      Left            =   3878
      TabIndex        =   7
      Top             =   1440
      Width           =   4215
   End
   Begin VB.Label lblBonusQuestion 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Find the d value of the point of intersection of the lines C=0.20d+525 and C=0.30d+500."
      BeginProperty Font 
         Name            =   "DK Crayon Crumble"
         Size            =   20.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   2400
      TabIndex        =   1
      Top             =   2160
      Width           =   6975
   End
   Begin VB.Image Image1 
      Height          =   1125
      Left            =   5400
      Picture         =   "frmSubBonus.frx":188C1
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1125
   End
End
Attribute VB_Name = "frmSubBonus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Jimmy Huynh and Domenico Didiano
'January 22, 2018
'ICS ISU - Kingdom of Knowledge - Bonus Question for Substitution
'The question that comes after the three consecutive one point questions. This question is more difficult and is worth 5 points.
'Correctly answering three consecutive questions along with this question can allow a user to achieve a power up for their class.
Option Explicit

Dim intCount As Integer

Private Sub cmdSubmit_Click()
    
        'Check if blank
    If optA.Value = False And optB.Value = False And optC.Value = False And optD.Value = False Then
    
        MsgBox "Please select an answer.", vbOKOnly, "Select an answer"
    
    Else
        
        'Knight Powerup
        ElseIf frmUnits.intAnswerStreak >= 3 And optA.Value = True And frmRoles.intRole = 1 Then
        
            MsgBox "Correct, make the two equations equal to find the value of d.", vbOKOnly, "Correct Answer"
            
            MsgBox "Congratulations! You have earned a power up. You can now block the next wrong answer.", vbOKOnly, _
            "Knight Power Up!"
            
            frmUnits.intAnswerStreak = frmUnits.intAnswerStreak + 1
            frmUnits.intPoints = frmUnits.intPoints + 5
            
        'Wizard Poweup
        ElseIf frmUnits.intAnswerStreak >= 3 And optA.Value = True And frmRoles.intRole = 2 Then
            
            MsgBox "Correct, make the two equations equal to find the value of d.", vbOKOnly, "Correct Answer"
            
            MsgBox "Congratulations! You have earned a power up. A 2x multiplier will now be applied to the next 3 questions.", _
            vbOKOnly, "Wizard Power Up!"
            
            frmUnits.intAnswerStreak = frmUnits.intAnswerStreak + 1
            frmUnits.intPoints = frmUnits.intPoints + 5
            
        'Peasant Powerup (Point Randomize)
        ElseIf frmUnits.inAnswerStreak >= 3 And optA.Value = True And frmRoles.intRole = 3 Then
            
            MsgBox "Correct, make the two equations equal to find the value of d.", vbOKOnly, "Correct Answer"
        
            MsgBox "Congratulations! You have earned a power up. A random number of bonus points will be added to your score.", _
            vbOKOnly, "Peasant Power Up!"
            
            frmUnits.intAnswerStreak = frmUnits.intAnswerStreak + 1
            frmUnits.intPoints = frmUnits.intPoints + 5 + (8) * Rnd + 5
            
        'Correct Answer
        ElseIf frmUnits.intAnswerStreak < 3 And optA.Value = True Then
            
            MsgBox "Correct, make the two equations equal to find the value of d."
            
            frmUnits.intAnswerStreak = frmUnits.intAnswerStreak + 1
            frmUnits.intPoints = frmUnits.intPoints + 5
        
        'Knight Power Up Active
        ElseIf optA.Value = False And frmRoles.intBlock > 0 Then
        
            MsgBox "Incorrect, make the two equations equal to find the value of d. The answer was A:250" _
            , vbOKOnly, "Incorrect Answer"
            
            MsgBox "You have succesfully blocked the incorrect answer, it will not affect your streak.", vbOKOnly, "Incorrect Answer Blocked"
            
            frmRoles.intBlock = frmRoles.intBlock - 1
        
        'Wrong Answer
        ElseIf optA.Value = False Then
            MsgBox "Incorrect, make the two equations equal to find the value of d. The answer was A:250" _
            , vbOKOnly, "Incorrect Answer"
            
            frmUnits.intAnswerStreak = 0
            
            
        End If
        
        frmUnits.lblPoints = "Points: " & intPoints
        frmAlgebra.lblPoints = "Points: " & intPoints
        frmGeometry.lblPoints = "Points: " & intPoints
        frmQuadratics.lblPoints = "Points: " & intPoints
        frmTrigonometry.lblPoints = "Points: " & intPoints
        
        frmSubBonus.Hide
        
        If frmAlgebra.cmdSub.Enabled = False And frmAlgebra.cmdElim.Enabled = False And frmUnits.cmdAlgebra.Enabled = False And _
            frmUnits.cmdGeometry.Enabled = False And frmUnits.cmdQuadratics.Enabled = False And frmUnits.cmdTrig.Enabled = False Then
                
            frmLeaderboard.Show
                
        ElseIf frmAlgebra.cmdSub.Enabled = False And frmAlgebra.cmdElim.Enabled = False Then
            
            frmUnits.Show
            
        Else
            frmAlgebra.Show
            
        End If
        
    End If

End Sub

Private Sub TmrCountdown_Timer()

    intCount = intCount + 1

    Do While intCount = 15
        
        MsgBox "Time is up! The answer is A:250", vbOKOnly, "Timer"
        
        frmSubBonus.Hide
            
        frmUnits.lblPoints = "Points: " & intPoints
        frmAlgebra.lblPoints = "Points: " & intPoints
        frmGeometry.lblPoints = "Points: " & intPoints
        frmQuadratics.lblPoints = "Points: " & intPoints
        frmTrigonometry.lblPoints = "Points: " & intPoints
            
        If frmAlgebra.cmdSub.Enabled = False And frmAlgebra.cmdElim.Enabled = False And frmUnits.cmdAlgebra.Enabled = False And _
        frmUnits.cmdGeometry.Enabled = False And frmUnits.cmdQuadratics.Enabled = False And frmUnits.cmdTrig.Enabled = False Then
                
            frmLeaderboard.Show
                
        ElseIf frmAlgebra.cmdSub.Enabled = False And frmAlgebra.cmdElim.Enabled = False Then
            
            frmUnits.Show
            
        Else
            frmAlgebra.Show
            
        End If
             
        frmUnits.intAnswerStreak = 0
        
        intCount = intCount + 1
        
    Loop

End Sub

Private Sub cmdhelp_Click()

    frmHelp.Show

End Sub
