VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "LisatBox demo Save and Load"
   ClientHeight    =   3930
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4680
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3930
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command3 
      Caption         =   "Print ListBox"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Load ListBox"
      Height          =   495
      Left            =   1560
      TabIndex        =   2
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Save ListBox"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "We just add some items to Combo when the Form is loaded so you could test these functions."
      Height          =   735
      Left            =   120
      TabIndex        =   3
      Top             =   3000
      Width           =   4215
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    Call SaveL(List1, App.Path & "\ListBox.txt")
    List1.Clear
End Sub

Private Sub Command2_Click()
    Call LoadL(List1, App.Path & "\ListBox.txt")
End Sub

Private Sub Command3_Click()
    Call PrintL(List1)
End Sub

Private Sub Form_Load()
    Dim i As Integer
    For i = 0 To 50
        List1.AddItem "Testiline " & i
    Next i
End Sub
