VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Student Info"
   ClientHeight    =   7500
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6600
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7500
   ScaleWidth      =   6600
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdDisplay 
      BackColor       =   &H80000005&
      Caption         =   "Display"
      Height          =   615
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2280
      Width           =   1335
   End
   Begin VB.CommandButton cmdSubmit 
      BackColor       =   &H80000005&
      Caption         =   "Submit"
      Height          =   615
      Left            =   5040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   840
      Width           =   1335
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Open Sans Semibold"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3960
      Left            =   480
      TabIndex        =   4
      Top             =   3120
      Width           =   5655
   End
   Begin VB.TextBox txtAge 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   3
      Top             =   1200
      Width           =   3255
   End
   Begin VB.TextBox txtName 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      TabIndex        =   2
      Top             =   480
      Width           =   3255
   End
   Begin VB.Label Label2 
      Caption         =   "Age"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   1200
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type std
name As String
age As String
End Type

Dim i As Integer
Dim student(10) As std

Private Sub cmdSubmit_Click()

If txtName.Text <> "" And txtAge.Text <> "" And IsNumeric(txtAge.Text) Then
    With student(i)
        .name = txtName.Text
        .age = txtAge.Text
    End With
    
    txtName.Text = ""
    txtAge.Text = ""
    
    i = i + 1
Else
MsgBox "Correct input value needed!"
End If

End Sub

Private Sub cmdDisplay_Click()
List1.Clear
While c < i
List1.AddItem "Name: " & student(c).name & ", " & "Age: " & student(c).age & ";"
c = c + 1
Wend
End Sub

Private Sub Form_Load()
txtName.Text = ""
txtAge.Text = ""
End Sub
