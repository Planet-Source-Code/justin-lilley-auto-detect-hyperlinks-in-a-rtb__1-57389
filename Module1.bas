Attribute VB_Name = "Module1"
'Text Adder
Public Sub AddTxt(ParamArray saElements() As Variant)
    Dim strTimeStamp As String, Data As String
    strTimeStamp = "[-]" & Time & "[-] "
On Error Resume Next

AddTextLen = AddTextLen + 1
If AutoClear = "1" Then
If AddTextLen >= AutoClearLines Then
frmMain.RTB1.Text = ""
AddTextLen = "0"
End If
End If

With frmMain.RTB1
.SelStart = Len(.Text)
.SelLength = 0
.SelColor = vbRed
.SelText = strTimeStamp
.SelStart = Len(.Text)
Data = strTimeStamp
End With
    
Dim k As Integer
For k = LBound(saElements) To UBound(saElements) Step 2
With frmMain.RTB1
.SelStart = Len(.Text)
.SelLength = 0
.SelColor = saElements(i)
.SelText = saElements(k + 1) & Left$(vbCrLf, -2 * CLng((k + 1) = UBound(saElements)))
.SelStart = Len(.Text)
Data = Data & saElements(k + 1)
End With
Next k
End Sub
