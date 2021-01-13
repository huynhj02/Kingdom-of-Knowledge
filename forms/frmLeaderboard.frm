VERSION 5.00
Begin VB.Form frmLeaderboard 
   Caption         =   "Kingdom of Knowledge"
   ClientHeight    =   7200
   ClientLeft      =   22065
   ClientTop       =   -6120
   ClientWidth     =   11955
   Icon            =   "frmLeaderboard.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmLeaderboard.frx":424A
   ScaleHeight     =   7200
   ScaleWidth      =   11955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Height          =   1575
      Left            =   8520
      Picture         =   "frmLeaderboard.frx":F0C5
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   5280
      Width           =   2895
   End
   Begin VB.CommandButton cmdPlayAgain 
      Height          =   1575
      Left            =   8520
      Picture         =   "frmLeaderboard.frx":1376F
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   3600
      Width           =   2895
   End
   Begin VB.PictureBox picLeaderboard 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Milk Mustache BB"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5895
      Left            =   1230
      Picture         =   "frmLeaderboard.frx":18958
      ScaleHeight     =   5895
      ScaleWidth      =   9495
      TabIndex        =   0
      Top             =   653
      Width           =   9495
   End
End
Attribute VB_Name = "frmLeaderboard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Jimmy Huynh and Domenico Didiano
'January 22, 2018
'ICS ISU - Kingdom of Knowledge - Leaderboard
'After the user answers all the questions, his/her points are displayed, as well as the intials he/she chose.

Option Explicit
Dim intFirstTime As Integer
Dim strName As String

Private Sub Form_Load()

strName = InputBox("Hello young " & strRole & " what is your name?", "What is your name?")

If intFirstTime = 0 Then
picLeaderboard.Print
picLeaderboard.Print
picLeaderboard.Print
picLeaderboard.Print Space(20) & strName & Space(50) & frmUnits.intPoints
intFirstTime = intFirstTime + 1

ElseIf intFirstTime >= 1 Then

picLeaderboard.Print "hello"

End If

End Sub

Private Sub cmdExit_Click()

    Unload frmAbout
    Unload frmAlgebra
    Unload frmCosLesson
    Unload frmCosExample
    Unload frmCosine1
    Unload frmCosine2
    Unload frmCosine3
    Unload frmCosineBonus
    Unload frmDescription
    Unload frmElim1
    Unload frmElim2
    Unload frmElim3
    Unload frmElimBonus
    Unload frmElimExample
    Unload frmElimLesson
    Unload frmFacLesson
    Unload frmFactor1
    Unload frmFactor2
    Unload frmFactor3
    Unload frmFacBonus
    Unload frmFacExample
    Unload frmGeometry
    Unload frmHelp
    Unload frmLeaderboard
    Unload frmLen1
    Unload frmLen2
    Unload frmLen3
    Unload frmLenBonus
    Unload frmLenExample
    Unload frmLenLesson
    Unload frmMainMenu
    Unload frmMid1
    Unload frmMid2
    Unload frmMid3
    Unload frmMidBonus
    Unload frmMidExample
    Unload frmMidLesson
    Unload frmQuadratics
    Unload frmRoles
    Unload frmSineExample
    Unload frmSineLesson
    Unload frmSine1
    Unload frmSine2
    Unload frmSine3
    Unload frmSineBonus
    Unload frmSinExample
    Unload frmSubBonus
    Unload frmSubExample
    Unload frmSubLesson
    Unload frmSubstitution1
    Unload frmSubstitution2
    Unload frmSubstitution3
    Unload frmTanLesson
    Unload frmTanExample
    Unload frmTangent1
    Unload frmTangent2
    Unload frmTangent3
    Unload frmTangentBonus
    Unload frmTrigonometry
    Unload frmUI1
    Unload frmUI2
    Unload frmUI3
    Unload frmUI4
    Unload frmUI5
    Unload frmUI6
    Unload frmUI7
    Unload frmUI8
    Unload frmUnits
    Unload frmVerLesson
    Unload frmVerExample
    Unload frmVertex1
    Unload frmVertex2
    Unload frmVertex3
    Unload frmVerBonus

End Sub

