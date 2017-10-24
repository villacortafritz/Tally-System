VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6120
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   ScaleHeight     =   6120
   ScaleWidth      =   6450
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List3 
      Height          =   2010
      Left            =   720
      TabIndex        =   4
      Top             =   3240
      Width           =   2295
   End
   Begin VB.ListBox List2 
      Height          =   2010
      Left            =   3240
      TabIndex        =   3
      Top             =   1080
      Width           =   2295
   End
   Begin VB.ListBox List1 
      Height          =   2010
      Left            =   720
      TabIndex        =   2
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label ctr3 
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      Top             =   4320
      Width           =   735
   End
   Begin VB.Label ctr2 
      Height          =   255
      Left            =   4800
      TabIndex        =   9
      Top             =   3840
      Width           =   735
   End
   Begin VB.Label ctr1 
      Height          =   255
      Left            =   4800
      TabIndex        =   8
      Top             =   3360
      Width           =   735
   End
   Begin VB.Label Label4 
      Caption         =   "List 3 Counter:"
      Height          =   255
      Left            =   3360
      TabIndex        =   7
      Top             =   4320
      Width           =   1215
   End
   Begin VB.Label Label3 
      Caption         =   "List 2 Counter:"
      Height          =   255
      Left            =   3360
      TabIndex        =   6
      Top             =   3840
      Width           =   1215
   End
   Begin VB.Label Label2 
      Caption         =   "List 1 Counter:"
      Height          =   255
      Left            =   3360
      TabIndex        =   5
      Top             =   3360
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Input Text:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   0
      Top             =   480
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub List1_Click()
    List3.AddItem (List1.Text)
    ctr3.Caption = List3.ListCount
End Sub

Private Sub List1_DblClick()
    Dim lol As String
    lol = List1.Text
    For i = 0 To List2.ListCount - 1
        For j = 0 To List2.ListCount - 1
            If lol = List2.List(i) Then List2.RemoveItem (i)
            
        Next
    Next
    List1.RemoveItem (List1.ListIndex)
    ctr1.Caption = List1.ListCount
    ctr2.Caption = List2.ListCount
    ctr3.Caption = List3.ListCount
End Sub

Private Sub List2_DblClick()
    Dim lol As String
    lol = List2.Text
    For i = 0 To List2.ListCount - 1
        For j = 0 To List2.ListCount - 1
            If lol = List2.List(j) Then List2.RemoveItem
        Next
    Next
    ctr2.Caption = List2.ListCount
End Sub

Private Sub List3_Click()
    Dim lol As String
    lol = List3.Text
    For i = 0 To List3.ListCount - 1
        For j = 0 To List3.ListCount - 1
            If lol = List3.List(j) Then List3.RemoveItem
        Next
    Next
    ctr3.Caption = List3.ListCount
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    Dim a As Integer
    a = 0
    If KeyAscii = 13 Then
        For i = 0 To List1.ListCount - 1
            If Text1.Text = List1.List(i) Then
            List2.AddItem (Text1.Text)
            a = 1
        End If
        Next
        If a = 0 Then List1.AddItem (Text1.Text)
        Text1.Text = ""
        ctr1.Caption = List1.ListCount
        ctr2.Caption = List2.ListCount
    End If
End Sub
