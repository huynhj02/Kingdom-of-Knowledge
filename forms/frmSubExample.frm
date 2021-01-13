VERSION 5.00
Begin VB.Form frmSubExample 
   Caption         =   "Kingdom of Knowledge"
   ClientHeight    =   7200
   ClientLeft      =   -1545
   ClientTop       =   6525
   ClientWidth     =   11955
   Icon            =   "frmSubExample.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmSubExample.frx":424A
   ScaleHeight     =   7200
   ScaleWidth      =   11955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdhelp 
      Height          =   855
      Left            =   9720
      Picture         =   "frmSubExample.frx":F0C5
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5880
      Width           =   1575
   End
   Begin VB.OptionButton optD 
      BackColor       =   &H00C0FFFF&
      Caption         =   "D (3,17)"
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
      Caption         =   "C (15,18)"
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
      Caption         =   "B (2,24)"
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
      Caption         =   "A (3,19)"
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
      Picture         =   "frmSubExample.frx":1382A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5040
      Width           =   2175
   End
   Begin VB.Label lblQuestion 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Find the point of intersection of the lines of the lines y = 3x + 10 and y = 4x + 7."
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
Attribute VB_Name = "frmSubExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Jimmy Huynh and Domenico Didiano
'January 24, 2018
'ICS ISU - Kingdom of Knowledge - Substitution Example Question
'The example question on the substitution lesson.
Option Explicit

Private Sub cmdSubmit_Click()
    
    'Check if blank
    If optA.Value = False And optB.Value = False And optC.Value = False And optD.Value = False Then
    
        MsgBox "Please select an answer.", vbOKOnly, "Select an answer"
    
    Else
    
        'Correct Answer
        If optA.Value = True Then
            
            MsgBox "Correct, this question can be solved as one of the expressions can be substituted in place of the Y value of another expression." _
            , vbOKOnly, "Correct Answer"
           
        'Incorrect Answer
        Else
            MsgBox "Incorrect, since both equations are expressed to solve for y, substitute one equation into the y value of the other equation." _
            & "The correct answer is A:(3,19)", vbOKOnly, "Incorrect Answer"
            
        End If
    
        frmSubExample.Hide
        frmSubstitution1.Show
        
    End If

End Sub

Private Sub cmdhelp_Click()

    frmHelp.Show

End Sub
