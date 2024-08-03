VERSION 5.00
Begin VB.Form InputForm 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Input"
   ClientHeight    =   2085
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2085
   ScaleWidth      =   4575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.TextBox Text2 
      Height          =   315
      Left            =   120
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   4335
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   315
      Left            =   3000
      TabIndex        =   4
      Top             =   1680
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&OK"
      Height          =   315
      Left            =   120
      TabIndex        =   3
      Top             =   1680
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   900
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   4335
   End
   Begin VB.Label Label1 
      Height          =   420
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "InputForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private PropSearchMode As Boolean
Private PropPrompt As String, PropDefaultText As String
Private PropResult As String

Public Property Get SearchMode() As Boolean
SearchMode = PropSearchMode
End Property

Public Property Let SearchMode(ByVal Value As Boolean)
PropSearchMode = Value
End Property

Public Property Get Prompt() As String
Prompt = PropPrompt
End Property

Public Property Let Prompt(ByVal Value As String)
PropPrompt = Value
End Property

Public Property Get DefaultText() As String
DefaultText = PropDefaultText
End Property

Public Property Let DefaultText(ByVal Value As String)
PropDefaultText = Value
End Property

Public Property Get Result() As String
Result = PropResult
End Property

Private Sub Form_Load()
Call SetupVisualStylesFixes(Me)
Label1.Caption = PropPrompt
If PropSearchMode = True Then
    Text2.Visible = True
    Text1.Visible = False
    Command1.Default = True
Else
    Text1.Text = PropDefaultText
End If
End Sub

Private Sub Command1_Click()
If PropSearchMode = True Then
    PropResult = Text2.Text
Else
    PropResult = Text1.Text
End If
If PropResult = vbNullString Then PropResult = ""
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub
