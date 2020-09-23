VERSION 5.00
Begin VB.Form Form2 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "ComboBox demo Save abd Load"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4290
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4290
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Print Combo"
      Height          =   495
      Left            =   2880
      TabIndex        =   4
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load Combo"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save Combo"
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   2520
      Width           =   1215
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   120
      Width           =   3855
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "We just add some items to Combo when the Form is loaded so you could test these functions."
      Height          =   1695
      Left            =   240
      TabIndex        =   3
      Top             =   600
      Width           =   3855
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call SaveCombo(Combo1)
End Sub

Private Sub Command2_Click()
    Call LoadCombo(Combo1)
End Sub

Private Sub Command3_Click()
    Call PrintCombo(Combo1)
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 0 To 50
        Combo1.AddItem "Line " & i
    Next i
End Sub
