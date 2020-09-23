VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Type Writer Effect Demo- frm Satheesh"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      BackColor       =   &H000080FF&
      Caption         =   "Show effect"
      Height          =   330
      Left            =   1680
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1050
      Width           =   1170
   End
   Begin VB.TextBox Text1 
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      Height          =   540
      Left            =   420
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "typeEffect.frx":0000
      Top             =   315
      Width           =   3585
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2100
      Top             =   2415
   End
   Begin VB.Label Label1 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   975
      Left            =   420
      TabIndex        =   0
      Top             =   1575
      Width           =   3510
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Please e-mail your comments to sendsatheesh@yahoo.com
Private Sub Command1_Click()
Form1.Caption = "Effect has started"
Timer1.Enabled = True 'to start the effect when ueser clicks the button
End Sub

Private Sub Command1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = vbRed
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Command1.BackColor = &H80FF&
End Sub

Private Sub Text1_Click()
Text1.Text = ""
End Sub

Private Sub Timer1_Timer()
'typeEfect takes in two parameters, one the text you want to creat the effect and
'another is the object in which the result have to be displayed on
'In this example, i am sending text from the text box and giving the label1 as my displayObj
If typeEffect(Text1.Text, Label1) = False Then 'to wait still the function has completed the effect
    Timer1.Enabled = False 'to  stop the timer if the function has completed the effect
    Form1.Caption = "Effect has stopped"
End If
End Sub

Private Function typeEffect(ByRef Text As String, ByRef DisplayObj As Object) As Boolean
Static i, lenOfString As Integer 'Variable for controling the effect

typeEffect = True 'typeEffect wiil be true until the effect has completed and stoped
lenOfString = Len(Text) 'Storing the len of the passed Str

i = i + 1 'variable used control to create the effect

    If i > lenOfString Then 'This whole part deals
    i = 0                   'is used to reset and exit
    lenOfString = 0         'this function and its static variables
    typeEffect = False
    DisplayObj.Caption = Text
    Exit Function
    End If

     'to display the result to the DisplayObj
    DisplayObj.Caption = Left(Text, Len(Text) - Len(Text) + i) + "|"  ' display the result with cursor
End Function

