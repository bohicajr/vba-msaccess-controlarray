# VBA MSAccess Control Array Calculator Example

Demonstrate how the control array can drastically cut down lines of code by consolidating common events.  Simply download the Access file and run one of the sample forms.  Alternatively you can import the source code .txt file by using the following command on the Visual Basic Immediate Window.
If you choose to manually import the file, ensure that you have all the ControlArray classes in [src](https://github.com/bohicajr/vba-msaccess-controlarray/tree/master/src).

```VBA
Application.LoadFromText acForm, "frmCalculator", "PATH\frmCalculator.txt"
```

Note PATH is where you saved the frmCalculator.txt file on your machine.

### How frmCalculator works

When the form opens up we initialize two control array objects, NumericButtons and OperatorButtons.
The buttons have a tag property set to either "num" or "op" depending on if they are used to enter a number value or used as a calculator operation.

```VBA
Private WithEvents NumericButtons As ControlArray
Private WithEvents OperatorButtons As ControlArray
```
```VBA
Private Sub Form_Load()
    
    Set NumericButtons = ControlArray.Create(Me.Controls, ctCommandButton, "num")
    Set OperatorButtons = ControlArray.Create(Me.Controls, ctCommandButton, "op")
    
End Sub
```

The next part is to handle the events when the user clicks a numeric button. Notice how a reference to the button is passed into the event handling procedure, we will use this to evaluate which button was pushed and how it should be concatenated to the display value.

```VBA
Private Sub NumericButtons_CommandButtonClick(btn As CommandButton)

    If lblDisplay.Caption = "0" Then
        lblDisplay.Caption = btn.Caption
    Else
    
        If mblnResetDisplay = True Then
            lblDisplay.Caption = btn.Caption
            mblnResetDisplay = False
        ElseIf IsNumeric(lblDisplay.Caption & btn.Caption) Then
            lblDisplay.Caption = lblDisplay.Caption & btn.Caption
        End If
            
    End If
    
End Sub
```

The last part is to handle the events when the user clicks an operator key, again note how the button reference is passed in.

```VBA
Private Sub OperatorButtons_CommandButtonClick(btn As CommandButton)

    Select Case btn.Caption
        
        Case "C"
            
            mdblValueOne = 0
            mdblValueTwo = 0
            lblDisplay.Caption = "0"
            mMode = None
            
        Case "/"
            
            mdblValueOne = lblDisplay.Caption
            lblDisplay.Caption = 0
            mMode = Divide
        
        Case "*"
        
            mdblValueOne = lblDisplay.Caption
            lblDisplay.Caption = 0
            mMode = Multiply
        
        Case "+"
            
            mdblValueOne = lblDisplay.Caption
            lblDisplay.Caption = 0
            mMode = Add
        
        Case "-"
            
            mdblValueOne = lblDisplay.Caption
            lblDisplay.Caption = 0
            mMode = Subtract
        
        Case "="
            
            mdblValueTwo = lblDisplay.Caption
            mblnResetDisplay = True

            Select Case mMode
            
                Case Add
                    
                    lblDisplay.Caption = mdblValueOne + mdblValueTwo
                    
                Case Subtract
                    
                    lblDisplay.Caption = mdblValueOne - mdblValueTwo
                    
                Case Multiply
                    
                    lblDisplay.Caption = mdblValueOne * mdblValueTwo
                    
                Case Divide
                    
                    If mdblValueTwo = 0 Then
                        lblDisplay.Caption = "Div by 0 error"
                        mblnResetDisplay = True
                    Else
                        lblDisplay.Caption = mdblValueOne / mdblValueTwo
                    End If
                
            End Select
            
            'reset calculator
            mdblValueOne = 0
            mdblValueTwo = 0
            mMode = None
        
    End Select
End Sub
```

### How frmValidation works

When the form opens up we initialize two control array objects, NumberTextBoxes and AlphaTextBoxes.
The textboxes have a tag property set to either "num" or "alpha" depending on if they are used to enter a numeric value or alpha character value.

```VBA
Private WithEvents NumberTextBoxes As ControlArray
Private WithEvents AlphaTextBoxes As ControlArray
```
```VBA
Private Sub Form_Load()

    Set NumberTextBoxes = ControlArray.Create(Me.Controls, ctTextBox, "num")
    Set AlphaTextBoxes = ControlArray.Create(Me.Controls, ctTextBox, "alpha")

End Sub
```

The following three events are handled to check for different conditions of numeric textboxes.

```VBA
Private Sub NumberTextBoxes_TextBoxKeyPress(txt As TextBox, KeyAscii As Integer)
    'Verifys each keypress is a valid entry
End Sub

Private Sub NumberTextBoxes_TextBoxChange(txt As TextBox)
    'Verifys text if it is pasted in
End Sub

Private Sub NumberTextBoxes_TextBoxExit(txt As TextBox, Cancel As Integer)
    'Verify that the number meets the rules when the user tries to leave the textbox
    'cancels if the rule is broken
End Sub

the following two events are handled to check for different conditions of alpha textboxes.

```VBA
Private Sub AlphaTextBoxes_TextBoxKeyPress(txt As TextBox, KeyAscii As Integer)
    'Verifys each keypress is a valid entry
End Sub

Private Sub AlphaTextBoxes_TextBoxChange(txt As TextBox)
	'Verifys text if it is pasted in
End Sub
```

The point of this demonstration is not the logic for key validation, but to show that we can drastically cut down on the number of events we need to code.  There are 10 numeric textboxes, with three events, in traditional fashion we would have 30 events to code, now only three.
There are also 10 alpha textboxes, with two events, that would be 20 more events to code.  So we've cut 50 events down to 5!








The next part is to handle the events when the user clicks a numeric button. Notice how a reference to the button is passed into the event handling procedure, we will use this to evaluate which button was pushed and how it should be concatenated to the display value.

