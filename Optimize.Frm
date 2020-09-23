VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6720
   ClientLeft      =   960
   ClientTop       =   1530
   ClientWidth     =   4845
   LinkTopic       =   "Form3"
   PaletteMode     =   1  'UseZOrder
   ScaleHeight     =   6720
   ScaleWidth      =   4845
   Begin VB.CommandButton Command7 
      Caption         =   "Print ListView"
      Height          =   495
      Left            =   3240
      TabIndex        =   10
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton Command6 
      Caption         =   "TextBox demo"
      Height          =   375
      Left            =   120
      TabIndex        =   9
      Top             =   6120
      Width           =   4575
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ListBox demo"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   5640
      Width           =   4575
   End
   Begin VB.CommandButton Command4 
      Caption         =   "Combo demo"
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   5160
      Width           =   4575
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load ListView"
      Height          =   495
      Left            =   1560
      TabIndex        =   6
      Top             =   4560
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   3720
      TabIndex        =   4
      Text            =   "20"
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton cmdTallenna 
      Caption         =   "Save ListView"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Add item to ListView (optimize)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   60
      TabIndex        =   2
      Top             =   960
      Width           =   4695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Add items to ListView witout optimization"
      Height          =   375
      Left            =   60
      TabIndex        =   1
      Top             =   540
      Width           =   4695
   End
   Begin ComctlLib.ListView ListView1 
      Height          =   3015
      Left            =   60
      TabIndex        =   0
      Top             =   1440
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   5318
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Teksti"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Subitem1"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Subitem2"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Number of items to generate to ListView:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   240
      TabIndex        =   5
      Top             =   120
      Width           =   3465
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Lista As ListItem
Private Declare Sub InvalidateRect Lib "user32" (ByVal hWnd As Long, ByVal t As Long, ByVal bErase As Long)
Private Declare Sub ValidateRect Lib "user32" (ByVal hWnd As Long, ByVal t As Long)

Private Sub cmdTallenna_Click()
    SaveLW ListView1, App.Path + "\testi.txt"
End Sub

Private Sub Command1_Click()
    Dim i As Long
    ListView1.View = 3
    Set Lista = ListView1.ListItems.Add
    With ListView1.ListItems
        .Clear
        For i = 1 To Text1.Text
            '****Muutos alkaa*********
            Set Lista = ListView1.ListItems.Add
            Lista.Text = "Testi "
            Lista.SubItems(1) = CStr(i)
            Lista.SubItems(2) = "Sub 2"
            '****Muutos päättyy*********
            DoEvents
        Next i
    End With
End Sub

Private Sub Command2_Click()
    Dim i As Long
    ListView1.View = 3
    With ListView1.ListItems
        .Clear
        For i = 1 To Text1.Text
            Set Lista = ListView1.ListItems.Add
            Lista.Text = "Testi " & CStr(i)
            Lista.SubItems(1) = CStr(i)
            Lista.SubItems(2) = "Sub 2"
            ValidateRect ListView1.hWnd, 0&
            If (i Mod 10) = 0 Then InvalidateRect ListView1.hWnd, 0&, 0&
            DoEvents
        Next i
    End With
    InvalidateRect ListView1.hWnd, 0&, 0&
End Sub

Private Sub Command3_Click()
    ListView1.View = 3
    LoadLW ListView1, App.Path & "\testi.txt"
End Sub

Private Sub Command4_Click()
    Form2.Show
End Sub

Private Sub Command5_Click()
    Form3.Show
End Sub

Private Sub Command6_Click()
    Form4.Show
End Sub

Private Sub Command7_Click()
    PrintLW ListView1
End Sub
