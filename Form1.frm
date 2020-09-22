VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   1995
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   3420
   LinkTopic       =   "Form1"
   ScaleHeight     =   1995
   ScaleWidth      =   3420
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command3 
      Caption         =   "Sum the numbers"
      Height          =   390
      Left            =   255
      TabIndex        =   3
      Top             =   1080
      Width           =   2805
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Message"
      Height          =   390
      Left            =   1530
      TabIndex        =   2
      Top             =   675
      Width           =   1530
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Change Text"
      Height          =   390
      Left            =   270
      TabIndex        =   1
      Top             =   675
      Width           =   1245
   End
   Begin VB.TextBox Text1 
      Height          =   390
      Left            =   285
      TabIndex        =   0
      Text            =   "Return values of the functions"
      Top             =   195
      Width           =   2790
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "All about the function or Sub"
      Height          =   195
      Left            =   705
      TabIndex        =   4
      Top             =   1620
      Width           =   2010
   End
   Begin VB.Menu mnuauthor 
      Caption         =   "Author"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'For explanation look at Function.bas

Private Sub Command1_Click()
Dim t As String
TextChange t ' Check this function
Text1.Text = t
End Sub

Private Sub Command2_Click()
Call Hi ' parameter not included
Call Hi("No way . Dont use optional string") ' parameter included
End Sub

Private Sub Command3_Click()
Text1.Text = Sum(1, 2, 3, 4, 5, 6, 7, 8, 9, 10)
End Sub

Private Sub Form_Load()
Me.Caption = ChangeCaption("Hi, I'm Changed") 'Check this function
End Sub

Private Sub mnuauthor_Click()
frmAuthor.Show 1
End Sub
