Attribute VB_Name = "modHiLiteText"
Option Explicit
'~modHiLiteText.bas;
'Highlights all text in a textbox
'*******************************************************************
' modHiLiteText: The HiLiteText() function highlights all text in a textbox.
' In: txtBox as textbox control
'
'EXAMPLE: highlight the contents of the txtName field when it is picked
'Private Sub txtName_GotFocus()
'  HiliteText txtName
'End Sub
'*******************************************************************

Public Sub HiLiteText(txtBox As TextBox)
  With txtBox
    .SelStart = 0                           'move to start of text
    .SelLength = Len(.Text)                 'highlight to end of text
  End With
End Sub
