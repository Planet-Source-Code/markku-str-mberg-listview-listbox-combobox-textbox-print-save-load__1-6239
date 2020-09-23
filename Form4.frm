VERSION 5.00
Begin VB.Form Form4 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "TextBox demo load and Save"
   ClientHeight    =   4260
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form4"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4260
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Print values"
      Height          =   495
      Left            =   3120
      TabIndex        =   7
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load values"
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   2040
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save values"
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   2040
      Width           =   1455
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Tag             =   "T"
      Text            =   "and other goodies...."
      Top             =   1560
      Width           =   4335
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Tag             =   "T"
      Text            =   "Banana"
      Top             =   1080
      Width           =   4335
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Tag             =   "T"
      Text            =   "Orange"
      Top             =   600
      Width           =   4335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Tag             =   "T"
      Text            =   "Apple"
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Form4.frx":0000
      Height          =   1215
      Left            =   120
      TabIndex        =   6
      Top             =   2640
      Width           =   4215
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call SaveTxtBoxes(Form4, "\textBox.txt")
    Call EmbtyText(Form4)
End Sub

Private Sub Command2_Click()
    Call LoadTxtBoxes(Form4, "\textBox.txt")
End Sub

Private Sub Command3_Click()
    Call PrintTxtBoxes(Form4)
End Sub
