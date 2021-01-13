VERSION 5.00
Begin VB.Form frmFacExample 
   Caption         =   "Kingdom of Knowledge"
   ClientHeight    =   7200
   ClientLeft      =   -1545
   ClientTop       =   6525
   ClientWidth     =   11955
   Icon            =   "frmFacExample.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmFacExample.frx":424A
   ScaleHeight     =   7200
   ScaleWidth      =   11955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdhelp 
      Height          =   975
      Left            =   9480
      Picture         =   "frmFacExample.frx":F0C5
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5760
      Width           =   1815
   End
   Begin VB.OptionButton optD 
      BackColor       =   &H00C0FFFF&
      Caption         =   "D x(x + 7) + 10"
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
      Caption         =   "C (x + 5) (x + 2)"
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
      Caption         =   "B (x + 2)**2 (x +5)**2"
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
      Caption         =   "A (x + 7) (x +10)"
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
      Picture         =   "frmFacExample.frx":1382A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Label lblExample 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Example"
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
      Left            =   3360
      TabIndex        =   7
      Top             =   360
      Width           =   5175
   End
   Begin VB.Label lblQuestion 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Factor x**2 + 7x +10"
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
Attribute VB_Name = "frmFacExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Jimmy Huynh and Domenico Didiano
'January 22, 2018
'ICS ISU - Kingdom of Knowledge - Factor Example
'The Factor form example.
Option Explicit
    

Private Sub cmdSubmit_Click()

 'Check if blank
    If optA.Value = False And optB.Value = False And optC.Value = False And optD.Value = False Then
        
        MsgBox "Please select an answer.", vbOKOnly, "Select an answer"
        
    'Check if Answer is correct
    ElseIf optC.Value = True Then
        
        MsgBox "Correct, the equation can be factored by finding two numbers that sum up to the second coefficient and multiply into the number without a variable.", vbOKOnly, "Correct Answer"
    
    'Check if the answer is incorrect
    Else
    
        MsgBox "Incorrect, the equation can be factored by finding two numbers that sum up to the second coefficient and multiply into the number without a variable, The first step would be x**2 + 5x + 2x + 10", _
        vbOKOnly, "Incorrect Answer"
    
    End If
    'Goes to the real questions
    frmFacExample.Hide
    frmFactor1.Show

End Sub
