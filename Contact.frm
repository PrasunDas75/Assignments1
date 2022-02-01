VERSION 5.00
Begin VB.Form Form5 
   Caption         =   "Form5"
   ClientHeight    =   2790
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9585
   LinkTopic       =   "Form5"
   ScaleHeight     =   2790
   ScaleWidth      =   9585
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   1815
      Left            =   4800
      TabIndex        =   8
      Top             =   840
      Width           =   3495
   End
   Begin VB.TextBox txts 
      Height          =   375
      Left            =   4800
      TabIndex        =   7
      Top             =   360
      Width           =   3495
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search Name"
      Height          =   375
      Left            =   8400
      TabIndex        =   6
      Top             =   360
      Width           =   1095
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Add"
      Height          =   375
      Left            =   1920
      TabIndex        =   5
      Top             =   2040
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Height          =   2535
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4575
      Begin VB.TextBox txtNo 
         Height          =   375
         Left            =   840
         TabIndex        =   10
         Top             =   1200
         Width           =   3495
      End
      Begin VB.TextBox txtAddr 
         Height          =   375
         Left            =   840
         TabIndex        =   4
         Top             =   720
         Width           =   3495
      End
      Begin VB.TextBox txtName 
         Height          =   375
         Left            =   840
         TabIndex        =   3
         Top             =   240
         Width           =   3495
      End
      Begin VB.Label Label3 
         Caption         =   "Ph No."
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Address"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   375
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   735
      End
   End
End
Attribute VB_Name = "Form5"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
If txtName.Text <> "" And txtAddr.Text <> "" And txtNo.Text <> "" Then
    If Not IsNumeric(txtNo.Text) Then
        MsgBox "Ph No. should be Number"
    Else
        Dim fsObj As New FileSystemObject
        Dim myTextFile As TextStream
        
        Set myTextFile = fsObj.OpenTextFile(App.Path & "\name.ini", ForAppending)
        myTextFile.WriteLine txtName.Text & " : " & txtAddr.Text & " = " & txtNo.Text
        myTextFile.Close
        Set fsObj = Nothing
        txtName.Text = ""
        txtAddr.Text = ""
        txtNo.Text = ""
        MsgBox "Added Successfully!"
    End If
    
Else
 MsgBox "fileds can not be empty"
End If
End Sub


Private Sub cmdSearch_Click()
If txts.Text <> "" Then
    Dim fsObj As New FileSystemObject
    
    Dim fileContent As String
    
    fileContent = fsObj.OpenTextFile(App.Path & "\name.ini", ForReading).ReadAll
    
    Set fsObj = Nothing
    If getName(fileContent, txts.Text) <> "" Then
        List1.AddItem "Name: " & getName(fileContent, txts.Text)
        List1.AddItem "Address:" & getDataFrom(fileContent, txts.Text)
        List1.AddItem "Ph No.:" & getNoFrom(fileContent, txts.Text)
        List1.AddItem vbCrLf
    Else
    MsgBox "no record"
    End If
Else
    MsgBox "Nothing to Search"
End If
End Sub

Private Function getName(ByVal fileContent As String, Name As String)

Dim charIndex As Integer
Dim Val As String
On Error Resume Next
charIndex = InStr(1, fileContent, Name, vbTextCompare)
Val = Mid$(fileContent, charIndex, Len(fileContent))

charIndex = InStr(1, Val, ":", vbTextCompare)
Val = Mid$(Val, 1, charIndex - 1)
getName = Val
End Function



Private Function getDataFrom(ByVal fileContent As String, Name As String)
Dim charIndex As Integer
Dim returnVal As String
On Error Resume Next
charIndex = InStr(1, fileContent, Name, vbTextCompare)
returnVal = Mid$(fileContent, charIndex, Len(fileContent))


charIndex = InStr(1, returnVal, "=", vbTextCompare)
returnVal = Mid$(returnVal, 1, charIndex - 1)

charIndex = InStr(1, returnVal, ":", vbTextCompare)
returnVal = Mid$(returnVal, charIndex + 1, Len(returnVal))

getDataFrom = returnVal
End Function

Private Function getNoFrom(ByVal fileContent As String, Name As String)
Dim charIndex As Integer
Dim returnVal As String
On Error Resume Next
charIndex = InStr(1, fileContent, Name, vbTextCompare)
returnVal = Mid$(fileContent, charIndex, Len(fileContent))


charIndex = InStr(1, returnVal, vbCrLf, vbTextCompare)
returnVal = Mid$(returnVal, 1, charIndex - 1)

charIndex = InStr(1, returnVal, "=", vbTextCompare)
returnVal = Mid$(returnVal, charIndex + 1, Len(returnVal))

getNoFrom = returnVal
End Function



