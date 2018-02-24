VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAListBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
''
' VBA-ControlArray v1.0.0
' (c) J Martinez - https://github.com/bohicajr/vba-msaccess-controlarray
'
'
''----------------------------------------------------------------------------''
Option Compare Binary
Option Explicit

Implements IControl

Private WithEvents AccessCtl As Access.ListBox
Attribute AccessCtl.VB_VarHelpID = -1

Private Type udtMembers
    Name As String
    ControlObject As Access.Control
    Parent As ControlArray
End Type

Private This As udtMembers

Public Function Create(ByRef ctl As Access.Control, ByRef Parent As ControlArray) As IControl

    With New CAListBox
    
        This.Name = ctl.Name
        Set This.ControlObject = ctl
        Set This.Parent = Parent
        Call .setMembers(This)
    
        Set Create = .Self
        
    End With
    
End Function

Friend Function Self() As IControl
    Set Self = Me
End Function

Friend Sub setMembers(ByRef Data As udtMembers)

    Const EVENT_PROCEDURE As String = "[Event Procedure]"
    Const EVENT_PREFIX As String = "On"
    Const MACRO_SUFFIX As String = "EmMacro"
    
    Dim prop As Property
    
    This = Data
    
    Set AccessCtl = Data.ControlObject
    
    AccessCtl.AfterUpdate = EVENT_PROCEDURE
    AccessCtl.BeforeUpdate = EVENT_PROCEDURE
    
    'find the OnEvent properties and set their value to Event Procedure
    For Each prop In AccessCtl.Properties
        
        If StrComp(Left(prop.Name, Len(EVENT_PREFIX)), EVENT_PREFIX, vbTextCompare) = 0 Then
 
            If InStr(1, prop.Name, MACRO_SUFFIX, vbTextCompare) = 0 Then 'exclude macro props
                
                prop.Value = EVENT_PROCEDURE
            
            End If
        
        End If
        
    Next
    
End Sub

Public Property Get ControlType() As AcControlType
    ControlType = AccessCtl.ControlType
End Property

Public Property Get Name() As String
    Name = This.Name
End Property

Public Property Get Object() As Access.Control
    Set Object = AccessCtl
End Property

Private Property Get iControl_ControlType() As AcControlType
    iControl_ControlType = Me.ControlType
End Property

Private Property Get iControl_Name() As String
    iControl_Name = Me.Name
End Property

Private Property Get iControl_Object() As Control
    Set iControl_Object = Me.Object
End Property

Private Sub AccessCtl_AfterUpdate()
    Call This.Parent.RaiseAfterUpdate(Me)
End Sub

Private Sub AccessCtl_BeforeUpdate(Cancel As Integer)
    Call This.Parent.RaiseBeforeUpdate(Me, Cancel)
End Sub

Private Sub AccessCtl_Click()
    Call This.Parent.RaiseClick(Me)
End Sub

Private Sub AccessCtl_DblClick(Cancel As Integer)
    Call This.Parent.RaiseDblClick(Me, Cancel)
End Sub

Private Sub AccessCtl_Enter()
    Call This.Parent.RaiseEnter(Me)
End Sub

Private Sub AccessCtl_Exit(Cancel As Integer)
    Call This.Parent.RaiseExit(Me, Cancel)
End Sub

Private Sub AccessCtl_GotFocus()
    Call This.Parent.RaiseGotFocus(Me)
End Sub

Private Sub AccessCtl_KeyDown(KeyCode As Integer, Shift As Integer)
    Call This.Parent.RaiseKeyDown(Me, KeyCode, Shift)
End Sub

Private Sub AccessCtl_KeyPress(KeyAscii As Integer)
    Call This.Parent.RaiseKeyPress(Me, KeyAscii)
End Sub

Private Sub AccessCtl_KeyUp(KeyCode As Integer, Shift As Integer)
    Call This.Parent.RaiseKeyUp(Me, KeyCode, Shift)
End Sub

Private Sub AccessCtl_LostFocus()
    Call This.Parent.RaiseLostFocus(Me)
End Sub

Private Sub AccessCtl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call This.Parent.RaiseMouseDown(Me, Button, Shift, X, Y)
End Sub

Private Sub AccessCtl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call This.Parent.RaiseMouseMove(Me, Button, Shift, X, Y)
End Sub

Private Sub AccessCtl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call This.Parent.RaiseMouseUp(Me, Button, Shift, X, Y)
End Sub
