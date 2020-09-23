VERSION 5.00
Begin VB.Form frminbox 
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   1770
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5325
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1770
   ScaleWidth      =   5325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtvalue 
      BackColor       =   &H00FFFFFF&
      Height          =   300
      Left            =   90
      TabIndex        =   0
      Top             =   1320
      Width           =   5145
   End
   Begin VB.CommandButton cmdcancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   4275
      TabIndex        =   3
      Top             =   660
      Width           =   930
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "&OK"
      Height          =   345
      Left            =   4275
      TabIndex        =   1
      Top             =   225
      Width           =   930
   End
   Begin VB.Label lblmess 
      BackStyle       =   0  'Transparent
      Height          =   1005
      Left            =   705
      TabIndex        =   2
      Top             =   165
      Width           =   3435
   End
   Begin VB.Image imgIcon 
      Height          =   480
      Left            =   105
      Top             =   135
      Width           =   555
   End
End
Attribute VB_Name = "frminbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' DM input-box form
' By Ben Jones
' Email: dreamvb@yahoo.com or vbdream2k@yahoo.com

Private Sub cmdcancel_Click()
    ReturnStr = ""  ' This means the user has pressed the cancel button
    Unload frminbox ' Unload the form
    
End Sub

Private Sub cmdok_Click()
    ReturnStr = txtvalue.Text   ' Pass the string to ReturnStr
    Unload frminbox ' Unload the form
    
End Sub

Private Sub Form_Load()
    frminbox.Icon = Nothing ' Remove the forms icon

End Sub

Private Sub Form_Paint()
    frminbox.AutoRedraw = True  ' Se the forms auto draw to true
    frminbox.Line (frminbox.ScaleWidth, 0)-(0, 0), &HFFFFFF ' This just draw's the small line at top of the form
    txtvalue.SelStart = 0   ' Set sel start to zero
    txtvalue.SelLength = Len(txtvalue.Text) ' Assign the text length as the sel length
    txtvalue.SetFocus   ' Set the focus on the textbox
    SetWindowTop frminbox.Hwnd, True ' Put the form ontop of all the others
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frminbox = Nothing  ' Clean up

End Sub

Private Sub txtvalue_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        cmdok_Click
    End If
    
End Sub
