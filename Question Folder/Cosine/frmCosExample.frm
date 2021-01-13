VERSION 5.00
Begin VB.Form frmCosExample 
   Caption         =   "Kingdom of Knowledge"
   ClientHeight    =   7200
   ClientLeft      =   -1545
   ClientTop       =   6525
   ClientWidth     =   11955
   Icon            =   "frmCosExample.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmCosExample.frx":424A
   ScaleHeight     =   7200
   ScaleWidth      =   11955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdhelp 
      Height          =   975
      Left            =   9480
      Picture         =   "frmCosExample.frx":F0C5
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5760
      Width           =   1815
   End
   Begin VB.OptionButton optD 
      BackColor       =   &H00C0FFFF&
      Caption         =   "D 31"
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
      Caption         =   "C 62"
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
      Caption         =   "B 42"
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
      Picture         =   "frmCosExample.frx":1382A
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5040
      Width           =   2775
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Example"
      BeginProperty Font 
         Name            =   "DK Crayon Crumble"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1095
      Left            =   3000
      TabIndex        =   7
      Top             =   240
      Width           =   6135
   End
   Begin VB.Label lblQuestion 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Find the measure of the angle to the nearest degree CosA = 0.4481"
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
      Top             =   1320
      Width           =   9015
   End
End
Attribute VB_Name = "frmCosExample"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Jimmy Huynh and Domenico Didiano
'January 22, 2018
'ICS ISU - Kingdom of Knowledge - Cosine Example Question
'The Example Question for Cosine.
Option Explicit




Private Sub cmdSubmit_Click()

    'Check if blank
    If optA.Value = False And optB.Value = False And optC.Value = False And optD.Value = False Then
        
        MsgBox "Please select an answer.", vbOKOnly, "Select an answer"
        
    'Check if Answer is correct
    ElseIf optA.Value = True Then
        
        MsgBox "Correct, angles can be solved for with Cos if you Isolate the Angle.", vbOKOnly, "Correct Answer"
    
    'Check if the answer is incorrect
    Else
    
        MsgBox "Incorrect, remember to isolate the angle and move Cos over to become Cos-1.", vbOKOnly, "Incorrect Answer"
    
    End If
    
    'Goes to the real questions
    frmCosExample.Hide
    frmCosine1.Show
End Sub
