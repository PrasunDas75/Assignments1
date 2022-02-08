VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000018&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tic Tac Toe"
   ClientHeight    =   5505
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3375
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5505
   ScaleWidth      =   3375
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      BackColor       =   &H80000018&
      Caption         =   "Reset"
      Height          =   375
      Left            =   1080
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   4920
      Width           =   1215
   End
   Begin VB.OptionButton optO 
      BackColor       =   &H80000018&
      Caption         =   "O"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1800
      TabIndex        =   10
      Top             =   1320
      Width           =   615
   End
   Begin VB.OptionButton optX 
      BackColor       =   &H80000018&
      Caption         =   "X"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   960
      TabIndex        =   9
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Choose First Move"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   720
      TabIndex        =   11
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label Label9 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   8
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   7
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   6
      Top             =   3960
      Width           =   495
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   5
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   4
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      TabIndex        =   2
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   2520
      Width           =   495
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "MS UI Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   0
      Top             =   2520
      Width           =   495
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As String
Dim o As String
Dim c As Integer


Private Sub Command1_Click()
Label1.Caption = ""
Label2.Caption = ""
Label3.Caption = ""
Label4.Caption = ""
Label5.Caption = ""
Label6.Caption = ""
Label7.Caption = ""
Label8.Caption = ""
Label9.Caption = ""
c = 1
End Sub

Private Sub Form_Load()
x = "X"
o = "O"
c = 1
End Sub

Private Sub Label1_Click()
If c Mod 2 = 0 Then
    Label1.Caption = o
Else
    Label1.Caption = x
End If
    c = c + 1
win
End Sub



Private Sub Label2_Click()
If c Mod 2 = 0 Then
    Label2.Caption = o
Else
    Label2.Caption = x
End If
c = c + 1
win
End Sub

Private Sub Label3_Click()
If c Mod 2 = 0 Then
    Label3.Caption = o
Else
    Label3.Caption = x
End If

c = c + 1
win
End Sub

Private Sub Label4_Click()
If c Mod 2 = 0 Then
    Label4.Caption = o
Else
    Label4.Caption = x
End If
c = c + 1
win
End Sub

Private Sub Label5_Click()
If c Mod 2 = 0 Then
    Label5.Caption = o
Else
    Label5.Caption = x
End If
c = c + 1
win
End Sub

Private Sub Label6_Click()
If c Mod 2 = 0 Then
    Label6.Caption = o
Else
    Label6.Caption = x
End If
c = c + 1
win
End Sub

Private Sub Label7_Click()
If c Mod 2 = 0 Then
    Label7.Caption = o
Else
    Label7.Caption = x
End If
c = c + 1
win
End Sub

Private Sub Label8_Click()
If c Mod 2 = 0 Then
    Label8.Caption = o
Else
    Label8.Caption = x
End If
c = c + 1
win
End Sub

Private Sub Label9_Click()
If c Mod 2 = 0 Then
    Label9.Caption = o
Else
    Label9.Caption = x
End If
c = c + 1
win
End Sub

Private Sub win()
If (Label1.Caption = "X" And Label2.Caption = "X" And Label3.Caption = "X") Then
 MsgBox Label3.Caption & " wins!"
ElseIf (Label1.Caption = "O" And Label2.Caption = "O" And Label3.Caption = "O") Then
 MsgBox Label3.Caption & " wins!"

ElseIf (Label4.Caption = "X" And Label5.Caption = "X" And Label6.Caption = "X") Then
 MsgBox Label6.Caption & " wins!"
ElseIf (Label4.Caption = "O" And Label5.Caption = "O" And Label6.Caption = "O") Then
 MsgBox Label6.Caption & " wins!"

ElseIf (Label7.Caption = "X" And Label8.Caption = "X" And Label9.Caption = "X") Then
 MsgBox Label9.Caption & " wins!"
ElseIf (Label7.Caption = "O" And Label8.Caption = "O" And Label9.Caption = "O") Then
 MsgBox Label9.Caption & " wins!"

ElseIf (Label1.Caption = "X" And Label5.Caption = "X" And Label9.Caption = "X") Then
 MsgBox Label9.Caption & " wins!"
ElseIf (Label1.Caption = "O" And Label5.Caption = "O" And Label9.Caption = "O") Then
 MsgBox Label9.Caption & " wins!"
 
ElseIf (Label3.Caption = "X" And Label5.Caption = "X" And Label7.Caption = "X") Then
 MsgBox Label7.Caption & " wins!"
ElseIf (Label3.Caption = "O" And Label5.Caption = "O" And Label7.Caption = "O") Then
 MsgBox Label7.Caption & " wins!"

ElseIf (Label1.Caption = "X" And Label4.Caption = "X" And Label7.Caption = "X") Then
 MsgBox Label7.Caption & " wins!"
ElseIf (Label1.Caption = "O" And Label4.Caption = "O" And Label7.Caption = "O") Then
 MsgBox Label7.Caption & " wins!"

ElseIf (Label2.Caption = "X" And Label5.Caption = "X" And Label8.Caption = "X") Then
 MsgBox Label8.Caption & " wins!"
ElseIf (Label2.Caption = "O" And Label5.Caption = "O" And Label8.Caption = "O") Then
 MsgBox Label8.Caption & " wins!"

ElseIf (Label3.Caption = "X" And Label6.Caption = "X" And Label9.Caption = "X") Then
 MsgBox Label9.Caption & " wins!"
ElseIf (Label3.Caption = "O" And Label6.Caption = "O" And Label9.Caption = "O") Then
 MsgBox Label9.Caption & " wins!"
 
End If

End Sub


Private Sub optO_Click()
x = "O"
o = "X"
End Sub

Private Sub optX_Click()
x = "X"
o = "O"
End Sub
