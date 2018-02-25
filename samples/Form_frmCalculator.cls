VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Form_frmCalculator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
''
' VBA-MSAccess-ControlArray Calculator Sample
' (c) J Martinez - https://github.com/bohicajr/vba-msaccess-controlarray
'
'
''----------------------------------------------------------------------------''
Option Compare Binary
Option Explicit

Private WithEvents NumericButtons As ControlArray
Attribute NumericButtons.VB_VarHelpID = -1
Private WithEvents OperatorButtons As ControlArray
Attribute OperatorButtons.VB_VarHelpID = -1

Private Enum enMode
    Add
    Subtract
    Divide
    Multiply
    None
End Enum

Private mdblValueOne As Double
Private mdblValueTwo As Double
Private mMode As enMode
Private mblnResetDisplay As Boolean
Private mblnIsHighlighted As Boolean

Private Sub Form_Load()
    
    'load the two control arrays with command buttons based on there tag property
    'NumericButtons will add the number and decimal buttons
    'OperatorButtons will add the remaining buttons used to operate the calculator
    Set NumericButtons = ControlArray.Create(Me.Controls, ctCommandButton, "num")
    Set OperatorButtons = ControlArray.Create(Me.Controls, ctCommandButton, "op")
    
End Sub

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

'fix VBA_IDE enum bug that doesn't correct case when typing
#If False Then
    #Const Add = 0
    #Const Subtract = 0
    #Const Divide = 0
    #Const Multiply = 0
    #Const None = 0
#End If


