# VBA MSAccess Control Array Calculator Example

Demonstrate how the control array can drastically cut down lines of code by consolidating common events.  Simply download the Access file and run the form frmCalculator.

### How it works

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
