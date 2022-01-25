VERSION 5.00
Begin VB.Form frmGame 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Number Puzzle"
   ClientHeight    =   4425
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4050
   LinkTopic       =   "Form9"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4425
   ScaleWidth      =   4050
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdShuffle 
      Caption         =   "Shuffle"
      Height          =   495
      Left            =   1440
      TabIndex        =   16
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton Command16 
      Caption         =   "15"
      Height          =   615
      Left            =   2640
      TabIndex        =   15
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Command15 
      Caption         =   "14"
      Height          =   615
      Left            =   2040
      TabIndex        =   14
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Command14 
      Caption         =   "13"
      Height          =   615
      Left            =   1440
      TabIndex        =   13
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Command13 
      Caption         =   "12"
      Height          =   615
      Left            =   840
      TabIndex        =   12
      Top             =   2520
      Width           =   615
   End
   Begin VB.CommandButton Command12 
      Caption         =   "11"
      Height          =   615
      Left            =   2640
      TabIndex        =   11
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Command11 
      Caption         =   "10"
      Height          =   615
      Left            =   2040
      TabIndex        =   10
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Command10 
      Caption         =   "9"
      Height          =   615
      Left            =   1440
      TabIndex        =   9
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Command9 
      Caption         =   "8"
      Height          =   615
      Left            =   840
      TabIndex        =   8
      Top             =   1920
      Width           =   615
   End
   Begin VB.CommandButton Command8 
      Caption         =   "7"
      Height          =   615
      Left            =   2640
      TabIndex        =   7
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command7 
      Caption         =   "6"
      Height          =   615
      Left            =   2040
      TabIndex        =   6
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command6 
      Caption         =   "5"
      Height          =   615
      Left            =   1440
      TabIndex        =   5
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command5 
      Caption         =   "4"
      Height          =   615
      Left            =   840
      TabIndex        =   4
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton Command4 
      Caption         =   "3"
      Height          =   615
      Left            =   2640
      TabIndex        =   3
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command3 
      Caption         =   "2"
      Height          =   615
      Left            =   2040
      TabIndex        =   2
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "1"
      Height          =   615
      Left            =   1440
      TabIndex        =   1
      Top             =   720
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Height          =   615
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   615
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdShuffle_Click()
    Dim a(15), i, j, r As Integer
    Dim f As Boolean
    
    f = False
    
    i = 1
    a(j) = 1
    
    While i <= 15
        r = (Int(15 * Rnd()) + 1)
        
        For j = 1 To i
            If (a(j) = r) Then
                f = True
                Exit For
            End If
        Next
        If f = True Then
            f = False
        Else
            a(i) = r
            i = i + 1
        End If
    Wend
    
    
    Command2.Caption = a(2)
    Command3.Caption = a(3)
    Command4.Caption = a(4)
    Command5.Caption = a(5)
    Command6.Caption = a(6)
    Command7.Caption = a(7)
    Command8.Caption = a(8)
    Command9.Caption = a(9)
    Command10.Caption = a(10)
    Command11.Caption = a(11)
    Command12.Caption = a(12)
    Command13.Caption = a(13)
    Command14.Caption = a(14)
    Command15.Caption = a(1)
    Command16.Caption = a(15)
    Command1.Caption = ""
End Sub

Private Sub Command1_Click()
If Command2.Caption = "" Then
    Command2.Caption = Command1.Caption
    Command1.Caption = ""
ElseIf Command5.Caption = "" Then
    Command5.Caption = Command1.Caption
    Command1.Caption = ""
End If
Winner
End Sub

Private Sub Command10_Click()
If Command6.Caption = "" Then
    Command6.Caption = Command10.Caption
    Command10.Caption = ""
ElseIf Command9.Caption = "" Then
    Command9.Caption = Command10.Caption
    Command10.Caption = ""
ElseIf Command11.Caption = "" Then
    Command11.Caption = Command10.Caption
    Command10.Caption = ""
ElseIf Command14.Caption = "" Then
    Command14.Caption = Command10.Caption
    Command10.Caption = ""
End If
Winner
End Sub

Private Sub Command11_Click()
If Command7.Caption = "" Then
    Command7.Caption = Command11.Caption
    Command11.Caption = ""
ElseIf Command10.Caption = "" Then
    Command10.Caption = Command11.Caption
    Command11.Caption = ""
ElseIf Command12.Caption = "" Then
    Command12.Caption = Command11.Caption
    Command11.Caption = ""
ElseIf Command15.Caption = "" Then
    Command15.Caption = Command11.Caption
    Command11.Caption = ""
End If
Winner
End Sub

Private Sub Command12_Click()
If Command8.Caption = "" Then
    Command8.Caption = Command12.Caption
    Command12.Caption = ""
ElseIf Command11.Caption = "" Then
    Command11.Caption = Command12.Caption
    Command12.Caption = ""
ElseIf Command16.Caption = "" Then
    Command16.Caption = Command12.Caption
    Command12.Caption = ""
End If
Winner
End Sub

Private Sub Command13_Click()
If Command9.Caption = "" Then
    Command9.Caption = Command13.Caption
    Command13.Caption = ""
ElseIf Command14.Caption = "" Then
    Command14.Caption = Command13.Caption
    Command13.Caption = ""
End If
Winner
End Sub

Private Sub Command14_Click()
If Command10.Caption = "" Then
    Command10.Caption = Command14.Caption
    Command14.Caption = ""
ElseIf Command13.Caption = "" Then
    Command13.Caption = Command14.Caption
    Command14.Caption = ""
ElseIf Command15.Caption = "" Then
    Command15.Caption = Command14.Caption
    Command14.Caption = ""
End If
Winner
End Sub

Private Sub Command15_Click()
If Command11.Caption = "" Then
    Command11.Caption = Command15.Caption
    Command15.Caption = ""
ElseIf Command14.Caption = "" Then
    Command14.Caption = Command15.Caption
    Command15.Caption = ""
ElseIf Command16.Caption = "" Then
    Command16.Caption = Command15.Caption
    Command15.Caption = ""
End If
Winner
End Sub

Private Sub Command16_Click()
If Command12.Caption = "" Then
    Command12.Caption = Command16.Caption
    Command16.Caption = ""
ElseIf Command15.Caption = "" Then
    Command15.Caption = Command16.Caption
    Command16.Caption = ""
End If
Winner
End Sub

Private Sub Command2_Click()
If Command1.Caption = "" Then
    Command1.Caption = Command2.Caption
    Command2.Caption = ""
ElseIf Command3.Caption = "" Then
    Command3.Caption = Command2.Caption
    Command2.Caption = ""
ElseIf Command6.Caption = "" Then
    Command6.Caption = Command2.Caption
    Command2.Caption = ""
End If
Winner
End Sub

Private Sub Command3_Click()
If Command2.Caption = "" Then
    Command2.Caption = Command3.Caption
    Command3.Caption = ""
ElseIf Command4.Caption = "" Then
    Command4.Caption = Command3.Caption
    Command3.Caption = ""
ElseIf Command7.Caption = "" Then
    Command7.Caption = Command3.Caption
    Command3.Caption = ""
End If
Winner
End Sub

Private Sub Command4_Click()
If Command3.Caption = "" Then
    Command3.Caption = Command4.Caption
    Command4.Caption = ""
ElseIf Command8.Caption = "" Then
    Command8.Caption = Command4.Caption
    Command4.Caption = ""
End If
Winner
End Sub

Private Sub Command5_Click()
If Command1.Caption = "" Then
    Command1.Caption = Command5.Caption
    Command5.Caption = ""
ElseIf Command6.Caption = "" Then
    Command6.Caption = Command5.Caption
    Command5.Caption = ""
ElseIf Command9.Caption = "" Then
    Command9.Caption = Command5.Caption
    Command5.Caption = ""
End If
Winner
End Sub

Private Sub Command6_Click()
If Command2.Caption = "" Then
    Command2.Caption = Command6.Caption
    Command6.Caption = ""
ElseIf Command5.Caption = "" Then
    Command5.Caption = Command6.Caption
    Command6.Caption = ""
ElseIf Command7.Caption = "" Then
    Command7.Caption = Command6.Caption
    Command6.Caption = ""
ElseIf Command10.Caption = "" Then
    Command10.Caption = Command6.Caption
    Command6.Caption = ""
End If
Winner
End Sub

Private Sub Command7_Click()
If Command3.Caption = "" Then
    Command3.Caption = Command7.Caption
    Command7.Caption = ""
ElseIf Command6.Caption = "" Then
    Command6.Caption = Command7.Caption
    Command7.Caption = ""
ElseIf Command8.Caption = "" Then
    Command8.Caption = Command7.Caption
    Command7.Caption = ""
ElseIf Command11.Caption = "" Then
    Command11.Caption = Command7.Caption
    Command7.Caption = ""
End If
Winner
End Sub

Private Sub Command8_Click()
If Command4.Caption = "" Then
    Command4.Caption = Command8.Caption
    Command8.Caption = ""
ElseIf Command7.Caption = "" Then
    Command7.Caption = Command8.Caption
    Command8.Caption = ""
ElseIf Command12.Caption = "" Then
    Command12.Caption = Command8.Caption
    Command8.Caption = ""
End If
Winner
End Sub

Private Sub Command9_Click()
If Command5.Caption = "" Then
    Command5.Caption = Command9.Caption
    Command9.Caption = ""
ElseIf Command10.Caption = "" Then
    Command10.Caption = Command9.Caption
    Command9.Caption = ""
ElseIf Command13.Caption = "" Then
    Command13.Caption = Command9.Caption
    Command9.Caption = ""
End If
Winner
End Sub

Private Sub Winner()
If Command1.Caption = Str(1) And Command2.Caption = Str(2) And Command3.Caption = Str(3) And Command4.Caption = Str(4) And Command5.Caption = Str(5) And Command6.Caption = Str(6) And Command7.Caption = Str(7) And Command8.Caption = Str(8) And Command9.Caption = Str(9) And Command10.Caption = Str(10) And Command11.Caption = Str(11) And Command12.Caption = Str(12) And Command13.Caption = Str(13) And Command14.Caption = Str(14) And Command15.Caption = Str(15) And Command16.Caption = "" Then
    MsgBox "Cheers! You Win!"
End If
End Sub


