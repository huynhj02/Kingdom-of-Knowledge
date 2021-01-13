VERSION 5.00
Begin VB.Form frmFacLesson 
   Caption         =   "Form1"
   ClientHeight    =   7200
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11955
   BeginProperty Font 
      Name            =   "DK Crayon Crumble"
      Size            =   14.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   Picture         =   "frmFacLesson.frx":0000
   ScaleHeight     =   7200
   ScaleWidth      =   11955
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Height          =   1575
      Left            =   8640
      Picture         =   "frmFacLesson.frx":AE7B
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5280
      Width           =   2775
   End
   Begin VB.Label lblAnswer 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "0 = (x + 2) (x + 3)"
      BeginProperty Font 
         Name            =   "DK Crayon Crumble"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3840
      TabIndex        =   11
      Top             =   6120
      Width           =   4335
   End
   Begin VB.Label lblStep4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Now your group your similar brackets and group the two outer bracket numbers into a bracket Your Final Answer is"
      ForeColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   3480
      TabIndex        =   10
      Top             =   4920
      Width           =   5055
   End
   Begin VB.Label lblStep3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x(x + 2) + 3(x + 2) = 0"
      BeginProperty Font 
         Name            =   "DK Crayon Crumble"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2400
      TabIndex        =   9
      Top             =   4440
      Width           =   6735
   End
   Begin VB.Label lblStep2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Now factor the first two numbers and the second two numbers together"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   2160
      TabIndex        =   8
      Top             =   3720
      Width           =   7335
   End
   Begin VB.Label lblEquation2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x**2 + 2x +3x + 6 = 0"
      BeginProperty Font 
         Name            =   "DK Crayon Crumble"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   3000
      TabIndex        =   7
      Top             =   3360
      Width           =   5535
   End
   Begin VB.Label lblExplanation2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "And the same two numbers that multiply into this number"
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   6600
      TabIndex        =   6
      Top             =   2760
      Width           =   3615
   End
   Begin VB.Label lblExplanation1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "To Factor you need two numbers that add up to this coefficient "
      ForeColor       =   &H00FFFFFF&
      Height          =   735
      Left            =   1440
      TabIndex        =   5
      Top             =   2760
      Width           =   4215
   End
   Begin VB.Line lin2 
      BorderWidth     =   3
      X1              =   6480
      X2              =   6840
      Y1              =   2280
      Y2              =   2640
   End
   Begin VB.Line linOne 
      BorderWidth     =   3
      X1              =   5640
      X2              =   5280
      Y1              =   2280
      Y2              =   2640
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Factoring changes Standard form Quadratic equations into Factored form Quadratic equations by using simple factoring."
      ForeColor       =   &H00FFFFFF&
      Height          =   855
      Left            =   1800
      TabIndex        =   4
      Top             =   1200
      Width           =   8535
   End
   Begin VB.Label lblNote 
      BackStyle       =   0  'Transparent
      Caption         =   "Two asteriks means ""to the power of"" 2**2 means 2 squared"
      BeginProperty Font 
         Name            =   "DK Crayon Crumble"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1335
      Left            =   600
      TabIndex        =   3
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblEquation 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "x**2 + 5x +6 = 0"
      BeginProperty Font 
         Name            =   "DK Crayon Crumble"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   495
      Left            =   3000
      TabIndex        =   2
      Top             =   1920
      Width           =   5775
   End
   Begin VB.Label lblTitle 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Factoring Quadratic Equations"
      BeginProperty Font 
         Name            =   "DK Crayon Crumble"
         Size            =   36
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   975
      Left            =   2040
      TabIndex        =   1
      Top             =   360
      Width           =   7815
   End
End
Attribute VB_Name = "frmFacLesson"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
