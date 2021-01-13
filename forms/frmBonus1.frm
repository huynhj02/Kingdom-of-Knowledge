VERSION 5.00
Begin VB.Form frmBonus1 
   Caption         =   "Kingdom of Knowledge"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11970
   Icon            =   "frmBonus1.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "frmBonus1.frx":424A
   ScaleHeight     =   7200
   ScaleWidth      =   11970
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdExit 
      Height          =   855
      Left            =   9720
      Picture         =   "frmBonus1.frx":F0C5
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton cmdSubmit 
      Height          =   1335
      Left            =   4920
      Picture         =   "frmBonus1.frx":1382A
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5400
      Width           =   2175
   End
   Begin VB.OptionButton optAnswer4 
      BackColor       =   &H00C0FFFF&
      Caption         =   "172"
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
      Top             =   4320
      Width           =   2775
   End
   Begin VB.OptionButton optAnswer3 
      BackColor       =   &H00C0FFFF&
      Caption         =   "168"
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
      TabIndex        =   3
      Top             =   3360
      Width           =   2775
   End
   Begin VB.OptionButton optAnswer2 
      BackColor       =   &H00C0FFFF&
      Caption         =   "165"
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
      Top             =   4320
      Width           =   2775
   End
   Begin VB.OptionButton optAnswer1 
      BackColor       =   &H00C0FFFF&
      Caption         =   "162"
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
      TabIndex        =   1
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
      Caption         =   "The average of 13 consecutive integers is 162. What is the greatest of these integers?"
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
      TabIndex        =   0
      Top             =   2160
      Width           =   6975
   End
   Begin VB.Image Image1 
      Height          =   1125
      Left            =   5400
      Picture         =   "frmBonus1.frx":188C1
      Stretch         =   -1  'True
      Top             =   360
      Width           =   1125
   End
End
Attribute VB_Name = "frmBonus1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
