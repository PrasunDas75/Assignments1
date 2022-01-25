VERSION 5.00
Begin VB.Form frmCalculator 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Calculator"
   ClientHeight    =   5715
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4350
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5715
   ScaleWidth      =   4350
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdBack 
      BackColor       =   &H80000000&
      Caption         =   "<"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   1200
      Width           =   855
   End
   Begin VB.CommandButton cmdEquals 
      BackColor       =   &H8000000D&
      Caption         =   "="
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      MaskColor       =   &H00FFFFFF&
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   4800
      Width           =   855
   End
   Begin VB.CommandButton cmdMult 
      BackColor       =   &H80000002&
      Caption         =   "*"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   3360
      Width           =   855
   End
   Begin VB.CommandButton cmdDiv 
      BackColor       =   &H80000002&
      Caption         =   "/"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton cmdMinus 
      BackColor       =   &H80000002&
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   2640
      Width           =   855
   End
   Begin VB.CommandButton cmdPlus 
      BackColor       =   &H80000002&
      Caption         =   "+"
      BeginProperty Font 
         Name            =   "Myanmar Text"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   1920
      Width           =   855
   End
   Begin VB.Frame frmCalculator 
      BackColor       =   &H80000018&
      BorderStyle     =   0  'None
      Height          =   4695
      Left            =   120
      TabIndex        =   1
      Top             =   960
      Width           =   3135
      Begin VB.CommandButton cmdDot 
         Caption         =   "."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2040
         TabIndex        =   13
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton Command0 
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1080
         TabIndex        =   12
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton cmdC 
         BackColor       =   &H80000001&
         Caption         =   "C"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         MaskColor       =   &H000000FF&
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3480
         Width           =   855
      End
      Begin VB.CommandButton Command9 
         Caption         =   "9"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2040
         TabIndex        =   10
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton Command8 
         Caption         =   "8"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1080
         TabIndex        =   9
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton Command7 
         Caption         =   "7"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   8
         Top             =   2400
         Width           =   855
      End
      Begin VB.CommandButton Command6 
         Caption         =   "6"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2040
         TabIndex        =   7
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton Command5 
         Caption         =   "5"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1080
         TabIndex        =   6
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton Command4 
         Caption         =   "4"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   5
         Top             =   1320
         Width           =   855
      End
      Begin VB.CommandButton Command3 
         Caption         =   "3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   2040
         TabIndex        =   4
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command2 
         Caption         =   "2"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   1080
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton Command1 
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   975
         Left            =   120
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
   End
   Begin VB.TextBox txtResult 
      Alignment       =   1  'Right Justify
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   17.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   240
      Width           =   4095
   End
End
Attribute VB_Name = "frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim op As String
Dim r As Double
Dim c As Integer
Dim cm As Integer
Dim cm1 As Integer
Dim cd As Integer

Private Sub cmdBack_Click()
If Val(txtResult.Text) >= 10 Then
    txtResult.Text = Fix(txtResult.Text \ 10)
Else
    txtResult.Text = ""
End If
End Sub

Private Sub cmdC_Click()
txtResult.Text = ""
r = 0
op = Empty
End Sub


Private Sub cmdEquals_Click()
If op = "+" Then
txtResult.Text = r + txtResult.Text

End If
If op = "-" Then
txtResult.Text = r - txtResult.Text

End If
If op = "*" Then
txtResult.Text = r * txtResult.Text

End If
If op = "/" Then
txtResult.Text = r / txtResult.Text

End If
End Sub

Private Sub cmdDiv_Click()
    op = "/"
If cd = 0 Then
    r = txtResult.Text
    txtResult.Text = ""
Else
    txtResult.Text = r / Val(txtResult.Text)
    r = txtResult.Text
End If
cd = cd + 1
End Sub


Private Sub cmdMinus_Click()
    op = "-"
If cm1 = 0 Then
    r = txtResult.Text
    txtResult.Text = ""
Else
    txtResult.Text = r - Val(txtResult.Text)
    r = txtResult.Text
End If
cm1 = cm1 + 1
End Sub

Private Sub cmdMult_Click()
    op = "*"
If cm = 0 Then
    r = txtResult.Text
    txtResult.Text = ""
Else
    txtResult.Text = r * Val(txtResult.Text)
    r = txtResult.Text
End If
cm = cm + 1
End Sub

Private Sub cmdPlus_Click()

    op = "+"
If c = 0 Then
    r = txtResult.Text
    txtResult.Text = ""
Else
    txtResult.Text = r + Val(txtResult.Text)
    r = txtResult.Text
End If
c = c + 1
End Sub

Private Sub Command1_Click()
If op = Empty Then
    txtResult.Text = txtResult.Text & Command1.Caption
Else
    txtResult.Text = ""
    txtResult.Text = txtResult.Text & Command1.Caption
End If
End Sub

Private Sub Command2_Click()
If op = Empty Then
    txtResult.Text = txtResult.Text & Command2.Caption
Else
    txtResult.Text = ""
    txtResult.Text = txtResult.Text & Command2.Caption
End If
End Sub

Private Sub Command3_Click()
If op = Empty Then
    txtResult.Text = txtResult.Text & Command3.Caption
Else
    txtResult.Text = ""
    txtResult.Text = txtResult.Text & Command3.Caption
End If
End Sub

Private Sub Command4_Click()
If op = Empty Then
    txtResult.Text = txtResult.Text & Command4.Caption
Else
    txtResult.Text = ""
    txtResult.Text = txtResult.Text & Command4.Caption
End If
End Sub

Private Sub Command5_Click()
If op = Empty Then
    txtResult.Text = txtResult.Text & Command5.Caption
Else
    txtResult.Text = ""
    txtResult.Text = txtResult.Text & Command5.Caption
End If
End Sub

Private Sub Command6_Click()
If op = Empty Then
    txtResult.Text = txtResult.Text & Command6.Caption
Else
    txtResult.Text = ""
    txtResult.Text = txtResult.Text & Command6.Caption
End If
End Sub

Private Sub Command7_Click()
If op = Empty Then
    txtResult.Text = txtResult.Text & Command7.Caption
Else
    txtResult.Text = ""
    txtResult.Text = txtResult.Text & Command7.Caption
End If
End Sub

Private Sub Command8_Click()
If op = Empty Then
    txtResult.Text = txtResult.Text & Command8.Caption
Else
    txtResult.Text = ""
    txtResult.Text = txtResult.Text & Command8.Caption
End If
End Sub

Private Sub Command9_Click()
If op = Empty Then
    txtResult.Text = txtResult.Text & Command9.Caption
Else
    txtResult.Text = ""
    txtResult.Text = txtResult.Text & Command9.Caption
End If
End Sub

Private Sub Command0_Click()
If op = Empty Then
    txtResult.Text = txtResult.Text & Command0.Caption
Else
    txtResult.Text = ""
    txtResult.Text = txtResult.Text & Command0.Caption
End If
End Sub

Private Sub cmdDot_Click()
If op = Empty Then
    txtResult.Text = txtResult.Text & cmdDot.Caption
Else
    txtResult.Text = ""
    txtResult.Text = txtResult.Text & cmdDot.Caption
End If
End Sub
