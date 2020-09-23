VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00C0C0C0&
   Caption         =   "DM inputbox Example"
   ClientHeight    =   3240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   ScaleHeight     =   3240
   ScaleWidth      =   4695
   Begin VB.TextBox txt2 
      Height          =   285
      Left            =   555
      TabIndex        =   6
      Text            =   "What is your name ?"
      Top             =   1365
      Width           =   3150
   End
   Begin VB.TextBox txt1 
      Height          =   630
      Left            =   150
      MultiLine       =   -1  'True
      TabIndex        =   4
      Text            =   "Example.frx":0000
      Top             =   510
      Width           =   4410
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Exit"
      Height          =   375
      Left            =   105
      TabIndex        =   2
      Top             =   2775
      Width           =   2640
   End
   Begin VB.CommandButton Command2 
      Caption         =   "About"
      Height          =   375
      Left            =   105
      TabIndex        =   1
      Top             =   2310
      Width           =   2640
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show Input-box"
      Height          =   375
      Left            =   90
      TabIndex        =   0
      Top             =   1815
      Width           =   2640
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "Title"
      Height          =   195
      Left            =   150
      TabIndex        =   5
      Top             =   1410
      Width           =   300
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "Prompt"
      Height          =   195
      Left            =   150
      TabIndex        =   3
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim iBox As New clsMain
Dim Ans ' Used to hold the inputbox return result

    iBox.BackColour = &HE0E0E0  ' This sets the backcolour for the input-box
    iBox.PromptFontStyle = dmitalic ' Set the font-sytle
    iBox.FontName = "Verdana"   ' Set the font-name
    iBox.FontSize = 8          ' Set the font-size
    iBox.PromptForeColour = &HFF0000  ' Set the prompt fore-colour
    
    Ans = iBox.TInputbox(txt1, txt2, , , , dmUserLogo)
    
    If Ans = "" Then
        MsgBox "You did not enter your name please try latter"
    Else
        MsgBox "Hello " & Ans & " I am glad to meet you"
    End If
    
End Sub

Private Sub Command2_Click()
    ' It just shows a message box about this project
    MsgBox "Input box Replacement for Visual Basic Programmers" _
    & vbNewLine & "Make by Ben Jones" _
    & vbNewLine & vbNewLine & "Please vote if you like my code", vbInformation
    
End Sub

Private Sub Command3_Click()
    Unload Me ' unload thr form
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set Form1 = Nothing
    
End Sub
