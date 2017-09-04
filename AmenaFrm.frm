VERSION 5.00
Begin VB.Form AmenaFrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Amena Recarga"
   ClientHeight    =   3135
   ClientLeft      =   3810
   ClientTop       =   3390
   ClientWidth     =   4740
   Icon            =   "AmenaFrm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "AmenaFrm.frx":030A
   ScaleHeight     =   3135
   ScaleWidth      =   4740
   Begin VB.CommandButton Command1 
      Caption         =   "Generar código"
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   2400
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackColor       =   &H00F3F3F3&
      Caption         =   "0000000000000000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1680
      TabIndex        =   1
      Top             =   2400
      Width           =   2895
   End
End
Attribute VB_Name = "AmenaFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim f As Integer
Dim txt As String
Dim num As String
Private Sub Command1_Click()
inicio:
f = FreeFile
num = ""
While Len(num) < 16
num = num & Left(Int(Rnd * 100), 1)
Wend
On Error GoTo error
Open "c:\no.txt" For Input As f
Do Until EOF(f)
Line Input #f, txt
If txt = num Then GoTo inicio
Loop
Close
Label1.Caption = Left(num, 16)
Open "C:\no.txt" For Append As f
Print #f, num
Close
error:
Select Case Err.Number
 Case 53
 f = FreeFile
 Open "C:\no.txt" For Append As f
 Print #f, "0000000000000000"
 Close
 GoTo inicio
End Select
End Sub
Private Sub Form_DblClick()
If Command1.Visible = False And Label1.Visible = False And Me.Picture = AmenaAtras.Picture Then
Command1.Visible = True
Label1.Visible = True
Me.Picture = normal.Picture
ElseIf Command1.Visible = True And Label1.Visible = True And Me.Picture = AmenaFrm.Picture Then
Me.Picture = AmenaAtras.Picture
Command1.Visible = False
Label1.Visible = False
End If
End Sub
Private Sub Form_Load()
Label1.Caption = ""
Me.Caption = "Amena Recarga " & App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub Form_Terminate()
Close f
End Sub
