VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' DM input-box class
' By Ben Jones
' Email: dreamvb@yahoo.com or vbdream2k@yahoo.com

Public Enum FontS
    dmBold = 63
    dmitalic = 64
    dmUnderline = 128
End Enum

Public Enum tIcon
    dmInformation = 1
    dmCritical
    dmExclamation
    dmQuestion
    dmPassword
    dmWinLogo
    dmGlobeLogo
    dmUserLogo
    dmHelpLogo
End Enum

Public Function TInputbox(Prompt, Optional Title, Optional Default, Optional Xpos, Optional Ypos, Optional IconType As tIcon) As String
    With frminbox
        .lblmess.Caption = Prompt   ' Set the lblmess with the prompt message
        
        If IsMissing(Title) = True Then
            .Caption = App.EXEName  ' If no title is found then set title to the programs exe name
        Else
            .Caption = Title        ' Set the title with the title found
        End If
        
        If IsMissing(Default) = True Then
            .txtvalue = ""  ' Do nothing
        Else
            .txtvalue = Default ' Set the txtvalue with the Default value found
        End If
        
        If IsMissing(Xpos) Or IsMissing(Ypos) Then
            ' This is used to center the form
            ' if no xpos and ypos positions have been found
            .Top = (Screen.Height - .Height) / 2
            .Left = (Screen.Width - .Width) / 2
        Else
            .Top = Xpos     ' Value for the form's top position
            .Left = Xpos    ' Value for the form's left position
        End If
        
        ' Adding the Message icon image
        If IconType <= 0 Or IconType > 9 Then
            ' The above is just a check to see
            ' if a image exists for the value we are passing
            ' you may want to change if you add more of your own images
            .imgIcon.Visible = False ' Hide the image
            .lblmess.Left = 105 ' Arrange the lable
        Else
            ' Load in the icon from the resource file
            .imgIcon.Picture = LoadResPicture(IconType, vbResIcon)
        End If
        
        .Show vbModal       ' Show the form
        TInputbox = ReturnStr  ' Return the result back
        ReturnStr = ""  ' Clear the return string
    End With
    
End Function

Public Property Let PromptForeColour(ByVal vNewValue As Long)
    ' Set the message's Forecolour
    frminbox.lblmess.ForeColor = vNewValue
End Property

Public Property Let PromptFontStyle(ByVal vNewValue As FontS)
' The code below is used to set the font style of the message text
    Select Case vNewValue
        Case dmBold
            frminbox.lblmess.FontBold = True
        Case dmitalic
            frminbox.lblmess.FontItalic = True
        Case dmUnderline
            frminbox.lblmess.FontUnderline = True
        Case dmBold + dmitalic + dmUnderline
            frminbox.lblmess.FontBold = True
            frminbox.lblmess.FontItalic = True
            frminbox.lblmess.FontUnderline = True
        Case dmBold + dmitalic
            frminbox.lblmess.FontBold = True
            frminbox.lblmess.FontItalic = True
        Case dmitalic + dmUnderline
            frminbox.lblmess.FontUnderline = True
            frminbox.lblmess.FontItalic = True
        Case dmBold + dmUnderline
            frminbox.lblmess.FontBold = True
            frminbox.lblmess.FontUnderline = True
    End Select
End Property

Public Property Let BackColour(ByVal vNewValue As Long)
    ' Set the dialogs Foreground colour
    frminbox.BackColor = vNewValue
End Property
Public Property Let FontName(ByVal vNewValue As String)
    ' Set the lblmess font-name
    frminbox.lblmess.FontName = vNewValue
End Property
Public Property Let FontSize(ByVal vNewValue As Long)
    ' Set the lblmess font-size
    frminbox.lblmess.FontSize = vNewValue
End Property
