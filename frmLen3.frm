VERSION 5.00
Begin VB.Form frmLen3 
   Caption         =   "Kingdom of Knowledge"
   ClientHeight    =   7200
   ClientLeft      =   -1545
   ClientTop       =   6525
   ClientWidth     =   11955
   Icon            =   "frmLen3.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmLen3.frx":424A
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
      Picture         =   "frmLen3.frx":F0C5
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5880
      Width           =   1575
   End
   Begin VB.OptionButton optD 
      BackColor       =   &H00C0FFFF&
      Caption         =   "D sqrt(172)"
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
      Caption         =   "C 14 units"
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
      Caption         =   "B 13 units"
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
      Caption         =   "A sqrt(152)"
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
      Picture         =   "frmLen3.frx":1382A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Image imgClock 
      Height          =   1125
      Left            =   5400
      Picture         =   "frmLen3.frx":188C1
      Top             =   480
      Width           =   1125
   End
   Begin VB.Label lblQuestion 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Find the length of a line segment if its endpoints are (8,5) and (3,17)."
      BeginProperty Font 
         Name            =   "DK Crayon Crumble"
         Size            =   21.75
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
Attribute VB_Name = "frmLen3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Jimmy Huynh and Domenico Didiano
'January 24, 2018
'ICS ISU - Kingdom of Knowledge - Length Question #3
'The third of three consecutive questions on the lesson about the length formula.


Private Sub cmdSubmit_Click()
    
    'Check if blank
    If optA.Value = False And optB.Value = False And optC.Value = False And optD.Value = False Then
    
        MsgBox "Please select an answer.", vbOKOnly, "Select an answer"
        
    Else
        
        'Wizard Power Up
        If optB.Value = True And frmRoles.intMultiply > 1 Then
                
            MsgBox "Correct, this distance can be found using the formula sqrt((x2-x1)^2)+(y2-y1)^2)" _
            , vbOKOnly, "Correct Answer (Wizard Multiplier)"
            
            frmRoles.intMultiply = frmRoles.intMultiply - 1
            
        'No Power Up
        ElseIf optB.Value = True Then
            
            MsgBox "Correct, this distance can be found using the formula sqrt((x2-x1)^2)+(y2-y1)^2)" _
            , vbOKOnly, "Correct Answer"
            
            frmUnits.intAnswerStreak = frmUnits.intAnswerStreak + 1
            frmUnits.intPoints = frmUnits.intPoints + 1
                
        'Knight Power Up
        ElseIf optB.Value = False And frmRoles.intBlock > 1 Then
                
            MsgBox "Incorrect, this distance can be found using the formula sqrt((x2-x1)^2)+(y2-y1)^2). The correct answer was B:13 units" _
            , vbOKOnly, "Incorrect Answer"
            MsgBox "You have succesfully blocked the incorrect answer, it will not affect your streak.", vbOKOnly, "Incorrect Answer Blocked"
                
            frmRoles.intBlock = frmRoles.intBlock - 1
                
        'Incorrect Answer
        Else
            MsgBox "Incorrect, this distance can be found using the formula sqrt((x2-x1)^2)+(y2-y1)^2). The correct answer was B:13 units" _
            , vbOKOnly, "Incorrect Answer"
            frmUnits.intAnswerStreak = 0
                
        End If
            
    
        frmLen3.Hide
        frmLenBonus.Show
    
    End If

End Sub


Private Sub cmdhelp_Click()

    frmHelp.Show

End Sub
