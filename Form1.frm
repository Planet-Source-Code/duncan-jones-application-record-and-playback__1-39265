VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3675
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3675
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdButton 
      Caption         =   "5"
      Height          =   375
      Index           =   5
      Left            =   2520
      TabIndex        =   11
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "4"
      Height          =   375
      Index           =   4
      Left            =   2040
      TabIndex        =   10
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "3"
      Height          =   375
      Index           =   3
      Left            =   1560
      TabIndex        =   9
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "2"
      Height          =   375
      Index           =   2
      Left            =   1080
      TabIndex        =   8
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "1"
      Height          =   375
      Index           =   1
      Left            =   600
      TabIndex        =   7
      Top             =   2520
      Width           =   495
   End
   Begin VB.CommandButton cmdButton 
      Caption         =   "0"
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   6
      Top             =   2520
      Width           =   495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   975
      Left            =   2760
      TabIndex        =   4
      Top             =   960
      Width           =   1455
      Begin VB.TextBox Text2 
         Height          =   375
         Index           =   2
         Left            =   120
         TabIndex        =   5
         Text            =   "Text2"
         Top             =   360
         Width           =   1095
      End
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Index           =   1
      Left            =   360
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1800
      Width           =   2175
   End
   Begin VB.TextBox Text2 
      Height          =   615
      Index           =   0
      Left            =   360
      TabIndex        =   1
      Text            =   "Text2"
      Top             =   1080
      Width           =   2175
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   240
      Width           =   2175
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Left            =   2760
      TabIndex        =   3
      Top             =   360
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()

End Sub

Private Sub cmdButton_Click(Index As Integer)

Text1.Text = cmdButton(Index).Caption

End Sub

Private Sub cmdButton_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Debug.Print "Mouse down at " & X / Screen.TwipsPerPixelX & "," & Y / Screen.TwipsPerPixelY

End Sub

Private Sub cmdButton_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)

Debug.Print "Mouse up at " & X / Screen.TwipsPerPixelX & "," & Y / Screen.TwipsPerPixelY

End Sub


Private Sub Form_Load()

'\\ Move a bit one way or another to stimulat ethat the recorder might not load the
'\\ form in the same position as the playback app
Me.Left = Me.Left + 20 - (Rnd() * 40)
Me.Top = Me.Top + 20 - (Rnd() * 40)

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Debug.Print UnloadMode

End Sub

