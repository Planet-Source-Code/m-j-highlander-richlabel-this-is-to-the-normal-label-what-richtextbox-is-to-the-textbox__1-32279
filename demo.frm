VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3525
   ClientLeft      =   2730
   ClientTop       =   1935
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   ScaleHeight     =   3525
   ScaleWidth      =   6585
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   555
      Left            =   585
      TabIndex        =   1
      Top             =   1710
      Width           =   2850
   End
   Begin Project1.RichLabel rblX 
      Height          =   600
      Left            =   630
      TabIndex        =   0
      Top             =   450
      Width           =   5235
      _ExtentX        =   9234
      _ExtentY        =   1058
      BackColor       =   16777215
      Font.Name       =   "Verdana"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

rblX.Caption = "This is my <B>cool</B> <C1>Color</C> <I><C3>Label</C></I><BR>It is even <B>multiline</b>"



End Sub


