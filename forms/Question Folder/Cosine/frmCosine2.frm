VERSION 5.00
Begin VB.Form frmCosine2 
   Caption         =   "Kingdom of Knowledge"
   ClientHeight    =   7200
   ClientLeft      =   -1545
   ClientTop       =   6525
   ClientWidth     =   11955
   Icon            =   "frmCosine2.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmCosine2.frx":424A
   ScaleHeight     =   7200
   ScaleWidth      =   11955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdhelp 
      Height          =   975
      Left            =   9480
      Picture         =   "frmCosine2.frx":F0C5
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5760
      Width           =   1815
   End
   Begin VB.OptionButton optD 
      BackColor       =   &H00C0FFFF&
      Caption         =   "D 62"
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
      Left            =   6480
      TabIndex        =   5
      Top             =   3720
      Width           =   2775
   End
   Begin VB.OptionButton optC 
      BackColor       =   &H00C0FFFF&
      Caption         =   "C 83"
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
      Left            =   6480
      TabIndex        =   4
      Top             =   2640
      Width           =   2775
   End
   Begin VB.OptionButton optB 
      BackColor       =   &H00C0FFFF&
      Caption         =   "B 65"
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
      Top             =   3720
      Width           =   2775
   End
   Begin VB.OptionButton optA 
      BackColor       =   &H00C0FFFF&
      Caption         =   "A 63"
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
      Top             =   2640
      Width           =   2775
   End
   Begin VB.CommandButton cmdSubmit 
      Height          =   1575
      Left            =   4440
      Picture         =   "frmCosine2.frx":1382A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Label lblQuestion 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Find the measure of the angle to the nearest degree CosA = 5/11"
      BeginProperty Font 
         Name            =   "DK Crayon Crumble"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1695
      Left            =   1560
      TabIndex        =   1
      Top             =   840
      Width           =   9015
   End
End
Attribute VB_Name = "frmCosine2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Jimmy Huynh and Domenico Didiano
'January 22, 2018
'ICS ISU - Kingdom of Knowledge - Cosine question 2
'The Second Cosine Question
Option Explicit

Dim intCount As Integer

Private Sub cmdSubmit_Click()
    
    'Check if blank
   If optA.Value = False And optB.Value = False And optC.Value = False And optD.Value = False Then
    
        MsgBox "Please select an answer.", vbOKOnly, "Select an answer"
    
    'Wizard Power Up
   ElseIf optA.Value = True And frmRoles.intMultiply > 1 Then
        
        MsgBox "Correct, angles can be solved for with Cos if you Isolate the Angle.", vbOKOnly, "Correct Answer (Wizard Multiplier)"
    
        frmRoles.intMultiply = frmRoles.intMultiply - 1
    
    'No Power Up
   ElseIf optA.Value = True Then
    
        MsgBox "Correct, angles can be solved for with Cos if you Isolate the Angle.", vbOKOnly, "Correct Answer"
    
        frmUnits.intAnswerStreak = frmUnits.intAnswerStreak + 1
        frmUnits.intPoints = frmUnits.intPoints + 1
        
    'Knight Power Up
   ElseIf optA.Value = False And frmRoles.intBlock > 1 Then
        
        MsgBox "Incorrect, remember to isolate the angle and move Cos over to become Cos-1 ", vbOKOnly, "Incorrect Answer"
        MsgBox "You have succesfully blocked the incorrect answer, it will not affect your streak.", vbOKOnly, "Incorrect Answer Blocked"
        
        frmRoles.intBlock = frmRoles.intBlock - 1
        
    Else
        MsgBox "Incorrect, remember to isolate the angle and move Cos over to become Cos-1", vbOKOnly, "Incorrect Answer"
        frmUnits.intAnswerStreak = 0
        
    End If
 
    frmCosine2.Hide
    frmCosine3.Show
 
End Sub






