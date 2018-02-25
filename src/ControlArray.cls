VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlArray"
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

#Const IncludeSpecificEvents = True

Public Enum eControlTypes
    ctAll = -1
    ctAttachment = acAttachment
    ctBoundObjectFrame = acBoundObjectFrame
    ctChart = 113
    ctCheckBox = acCheckBox
    ctComboBox = acComboBox
    ctCommandButton = acCommandButton
    ctImage = acImage
    ctLabel = acLabel
    ctListBox = acListBox
    ctNavigationButton = acNavigationButton
    ctNavigationControl = acNavigationControl
    ctObjectFrame = acObjectFrame
    ctOptionButton = acOptionButton
    ctOptionGroup = acOptionGroup
    ctPage = acPage
    ctRectangle = acRectangle
    ctSubForm = acSubform
    ctTabCtl = acTabCtl
    ctTextBox = acTextBox
    ctToggleButton = acToggleButton
    ctWebBrowserControl = acWebBrowser
End Enum

Private Type udtMembers
    Col As Collection
End Type

Private This As udtMembers

Public Event onAfterUpdate(ByRef ctl As IControl)
Public Event onAttachmentCurrent(ByRef ctl As IControl)
Public Event onBeforeUpdate(ByRef ctl As IControl, Cancel As Integer)
Public Event onBeforeNavigate2(ByRef ctl As IControl, ByVal pDisp As Object, ByRef URL As Variant, ByRef flags As Variant, ByRef TargetFrameName As Variant, ByRef PostData As Variant, ByRef Headers As Variant, ByRef Cancel As Boolean)
Public Event onChange(ByRef ctl As IControl)
Public Event onClick(ByRef ctl As IControl)
Public Event onDblClick(ByRef ctl As IControl, Cancel As Integer)
Public Event onDirty(ByRef ctl As IControl, Cancel As Integer)
Public Event onDocumentComplete(ByRef ctl As IControl, ByVal pDisp As Object, URL As Variant)
Public Event onEnter(ByRef ctl As IControl)
Public Event onExit(ByRef ctl As IControl, Cancel As Integer)
Public Event onGotFocus(ByRef ctl As IControl)
Public Event onKeyDown(ByRef ctl As IControl, KeyCode As Integer, Shift As Integer)
Public Event onKeyPress(ByRef ctl As IControl, KeyAscii As Integer)
Public Event onKeyUp(ByRef ctl As IControl, KeyCode As Integer, Shift As Integer)
Public Event onLostFocus(ByRef ctl As IControl)
Public Event onMouseDown(ByRef ctl As IControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event onMouseMove(ByRef ctl As IControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event onMouseUp(ByRef ctl As IControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event onNavigateError(ByRef ctl As IControl, ByVal pDisp As Object, URL As Variant, TargetFrameName As Variant, StatusCode As Variant, Cancel As Boolean)
Public Event onNotInList(ByRef ctl As IControl, NewData As String, Response As Integer)
Public Event onProgressChange(ByRef ctl As IControl, ByVal Progress As Long, ByVal ProgressMax As Long)
Public Event onUndo(ByRef ctl As IControl, Cancel As Integer)
Public Event onUpdated(ByRef ctl As IControl, Code As Integer)

#If IncludeSpecificEvents = True Then
    
    Public Event AttachmentAfterUpdate(ByRef attach As Access.Attachment)
    Public Event AttachmentCurrent(ByRef attach As Access.Attachment)
    Public Event AttachmentBeforeUpdate(ByRef attach As Access.Attachment, Cancel As Integer)
    Public Event AttachmentChange(ByRef attach As Access.Attachment)
    Public Event AttachmentClick(ByRef attach As Access.Attachment)
    Public Event AttachmentDblClick(ByRef attach As Access.Attachment, Cancel As Integer)
    Public Event AttachmentDirty(ByRef attach As Access.Attachment, Cancel As Integer)
    Public Event AttachmentEnter(ByRef attach As Access.Attachment)
    Public Event AttachmentExit(ByRef attach As Access.Attachment, Cancel As Integer)
    Public Event AttachmentGotFocus(ByRef attach As Access.Attachment)
    Public Event AttachmentKeyDown(ByRef attach As Access.Attachment, KeyCode As Integer, Shift As Integer)
    Public Event AttachmentKeyPress(ByRef attach As Access.Attachment, KeyAscii As Integer)
    Public Event AttachmentKeyUp(ByRef attach As Access.Attachment, KeyCode As Integer, Shift As Integer)
    Public Event AttachmentLostFocus(ByRef attach As Access.Attachment)
    Public Event AttachmentMouseDown(ByRef attach As Access.Attachment, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event AttachmentMouseMove(ByRef attach As Access.Attachment, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event AttachmentMouseUp(ByRef attach As Access.Attachment, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Public Event BoundObjectFrameAfterUpdate(ByRef bframe As Access.BoundObjectFrame)
    Public Event BoundObjectFrameBeforeUpdate(ByRef bframe As Access.BoundObjectFrame, Cancel As Integer)
    Public Event BoundObjectFrameClick(ByRef bframe As Access.BoundObjectFrame)
    Public Event BoundObjectFrameDblClick(ByRef bframe As Access.BoundObjectFrame, Cancel As Integer)
    Public Event BoundObjectFrameEnter(ByRef bframe As Access.BoundObjectFrame)
    Public Event BoundObjectFrameExit(ByRef bframe As Access.BoundObjectFrame, Cancel As Integer)
    Public Event BoundObjectFrameGotFocus(ByRef bframe As Access.BoundObjectFrame)
    Public Event BoundObjectFrameKeyDown(ByRef bframe As Access.BoundObjectFrame, KeyCode As Integer, Shift As Integer)
    Public Event BoundObjectFrameKeyPress(ByRef bframe As Access.BoundObjectFrame, KeyAscii As Integer)
    Public Event BoundObjectFrameKeyUp(ByRef bframe As Access.BoundObjectFrame, KeyCode As Integer, Shift As Integer)
    Public Event BoundObjectFrameLostFocus(ByRef bframe As Access.BoundObjectFrame)
    Public Event BoundObjectFrameMouseDown(ByRef bframe As Access.BoundObjectFrame, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event BoundObjectFrameMouseMove(ByRef bframe As Access.BoundObjectFrame, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event BoundObjectFrameMouseUp(ByRef bframe As Access.BoundObjectFrame, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event BoundObjectFrameUpdated(ByRef bframe As Access.BoundObjectFrame, Code As Integer)
    
    Public Event ChartClick(ByRef fra As Access.ObjectFrame)
    Public Event ChartDblClick(ByRef fra As Access.ObjectFrame, Cancel As Integer)
    Public Event ChartEnter(ByRef fra As Access.ObjectFrame)
    Public Event ChartExit(ByRef fra As Access.ObjectFrame, Cancel As Integer)
    Public Event ChartGotFocus(ByRef fra As Access.ObjectFrame)
    Public Event ChartLostFocus(ByRef fra As Access.ObjectFrame)
    Public Event ChartMouseDown(ByRef fra As Access.ObjectFrame, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ChartMouseMove(ByRef fra As Access.ObjectFrame, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ChartMouseUp(ByRef fra As Access.ObjectFrame, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ChartUpdated(ByRef fra As Access.ObjectFrame, Code As Integer)

    Public Event CheckboxAfterUpdate(ByRef chk As Access.CheckBox)
    Public Event CheckboxBeforeUpdate(ByRef chk As Access.CheckBox, Cancel As Integer)
    Public Event CheckboxClick(ByRef chk As Access.CheckBox)
    Public Event CheckboxDblClick(ByRef chk As Access.CheckBox, Cancel As Integer)
    Public Event CheckboxEnter(ByRef chk As Access.CheckBox)
    Public Event CheckboxExit(ByRef chk As Access.CheckBox, Cancel As Integer)
    Public Event CheckboxGotFocus(ByRef chk As Access.CheckBox)
    Public Event CheckboxKeyDown(ByRef chk As Access.CheckBox, KeyCode As Integer, Shift As Integer)
    Public Event CheckboxKeyPress(ByRef chk As Access.CheckBox, KeyAscii As Integer)
    Public Event CheckboxKeyUp(ByRef chk As Access.CheckBox, KeyCode As Integer, Shift As Integer)
    Public Event CheckboxLostFocus(ByRef chk As Access.CheckBox)
    Public Event CheckboxMouseDown(ByRef chk As Access.CheckBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event CheckboxMouseMove(ByRef chk As Access.CheckBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event CheckboxMouseUp(ByRef chk As Access.CheckBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Public Event ComboBoxAfterUpdate(ByRef cbo As Access.ComboBox)
    Public Event ComboBoxBeforeUpdate(ByRef cbo As Access.ComboBox, Cancel As Integer)
    Public Event ComboBoxChange(ByRef cbo As Access.ComboBox)
    Public Event ComboBoxClick(ByRef cbo As Access.ComboBox)
    Public Event ComboBoxDblClick(ByRef cbo As Access.ComboBox, Cancel As Integer)
    Public Event ComboBoxDirty(ByRef cbo As Access.ComboBox, Cancel As Integer)
    Public Event ComboBoxEnter(ByRef cbo As Access.ComboBox)
    Public Event ComboBoxExit(ByRef cbo As Access.ComboBox, Cancel As Integer)
    Public Event ComboBoxGotFocus(ByRef cbo As Access.ComboBox)
    Public Event ComboBoxKeyDown(ByRef cbo As Access.ComboBox, KeyCode As Integer, Shift As Integer)
    Public Event ComboBoxKeyPress(ByRef cbo As Access.ComboBox, KeyAscii As Integer)
    Public Event ComboBoxKeyUp(ByRef cbo As Access.ComboBox, KeyCode As Integer, Shift As Integer)
    Public Event ComboBoxLostFocus(ByRef cbo As Access.ComboBox)
    Public Event ComboBoxMouseDown(ByRef cbo As Access.ComboBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ComboBoxMouseMove(ByRef cbo As Access.ComboBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ComboBoxMouseUp(ByRef cbo As Access.ComboBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ComboBoxNotInList(ByRef cbo As Access.ComboBox, NewData As String, Response As Integer)
    Public Event ComboBoxUndo(ByRef cbo As Access.ComboBox, Cancel As Integer)
    
    Public Event CommandButtonClick(ByRef btn As Access.CommandButton)
    Public Event CommandButtonDblClick(ByRef btn As Access.CommandButton, Cancel As Integer)
    Public Event CommandButtonEnter(ByRef btn As Access.CommandButton)
    Public Event CommandButtonExit(ByRef btn As Access.CommandButton, Cancel As Integer)
    Public Event CommandButtonGotFocus(ByRef btn As Access.CommandButton)
    Public Event CommandButtonKeyDown(ByRef btn As Access.CommandButton, KeyCode As Integer, Shift As Integer)
    Public Event CommandButtonKeyPress(ByRef btn As Access.CommandButton, KeyAscii As Integer)
    Public Event CommandButtonKeyUp(ByRef btn As Access.CommandButton, KeyCode As Integer, Shift As Integer)
    Public Event CommandButtonLostFocus(ByRef btn As Access.CommandButton)
    Public Event CommandButtonMouseDown(ByRef btn As Access.CommandButton, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event CommandButtonMouseMove(ByRef btn As Access.CommandButton, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event CommandButtonMouseUp(ByRef btn As Access.CommandButton, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Public Event ImageClick(ByRef img As Access.Image)
    Public Event ImageDblClick(ByRef img As Access.Image, Cancel As Integer)
    Public Event ImageMouseDown(ByRef img As Access.Image, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ImageMouseMove(ByRef img As Access.Image, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ImageMouseUp(ByRef img As Access.Image, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Public Event LabelClick(ByRef lbl As Access.Label)
    Public Event LabelDblClick(ByRef lbl As Access.Label, Cancel As Integer)
    Public Event LabelMouseDown(ByRef lbl As Access.Label, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event LabelMouseMove(ByRef lbl As Access.Label, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event LabelMouseUp(ByRef lbl As Access.Label, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Public Event ListBoxAfterUpdate(ByRef lst As Access.ListBox)
    Public Event ListBoxBeforeUpdate(ByRef lst As Access.ListBox, Cancel As Integer)
    Public Event ListBoxClick(ByRef lst As Access.ListBox)
    Public Event ListBoxDblClick(ByRef lst As Access.ListBox, Cancel As Integer)
    Public Event ListBoxEnter(ByRef lst As Access.ListBox)
    Public Event ListBoxExit(ByRef lst As Access.ListBox, Cancel As Integer)
    Public Event ListBoxGotFocus(ByRef lst As Access.ListBox)
    Public Event ListBoxKeyDown(ByRef lst As Access.ListBox, KeyCode As Integer, Shift As Integer)
    Public Event ListBoxKeyPress(ByRef lst As Access.ListBox, KeyAscii As Integer)
    Public Event ListBoxKeyUp(ByRef lst As Access.ListBox, KeyCode As Integer, Shift As Integer)
    Public Event ListBoxLostFocus(ByRef lst As Access.ListBox)
    Public Event ListBoxMouseDown(ByRef lst As Access.ListBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ListBoxMouseMove(ByRef lst As Access.ListBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ListBoxMouseUp(ByRef lst As Access.ListBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
 
    Public Event NavigationButtonClick(ByRef navbtn As Access.NavigationButton)
    Public Event NavigationButtonDblClick(ByRef navbtn As Access.NavigationButton, Cancel As Integer)
    Public Event NavigationButtonEnter(ByRef navbtn As Access.NavigationButton)
    Public Event NavigationButtonExit(ByRef navbtn As Access.NavigationButton, Cancel As Integer)
    Public Event NavigationButtonGotFocus(ByRef navbtn As Access.NavigationButton)
    Public Event NavigationButtonKeyDown(ByRef navbtn As Access.NavigationButton, KeyCode As Integer, Shift As Integer)
    Public Event NavigationButtonKeyPress(ByRef navbtn As Access.NavigationButton, KeyAscii As Integer)
    Public Event NavigationButtonKeyUp(ByRef navbtn As Access.NavigationButton, KeyCode As Integer, Shift As Integer)
    Public Event NavigationButtonLostFocus(ByRef navbtn As Access.NavigationButton)
    Public Event NavigationButtonMouseDown(ByRef navbtn As Access.NavigationButton, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event NavigationButtonMouseMove(ByRef navbtn As Access.NavigationButton, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event NavigationButtonMouseUp(ByRef navbtn As Access.NavigationButton, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Public Event NavigationControlAfterUpdate(ByRef navctl As Access.NavigationControl)
    Public Event NavigationControlBeforeUpdate(ByRef navctl As Access.NavigationControl, Cancel As Integer)
    Public Event NavigationControlChange(ByRef navctl As Access.NavigationControl)
    Public Event NavigationControlClick(ByRef navctl As Access.NavigationControl)
    Public Event NavigationControlDblClick(ByRef navctl As Access.NavigationControl, Cancel As Integer)
    Public Event NavigationControlDirty(ByRef navctl As Access.NavigationControl, Cancel As Integer)
    Public Event NavigationControlEnter(ByRef navctl As Access.NavigationControl)
    Public Event NavigationControlExit(ByRef navctl As Access.NavigationControl, Cancel As Integer)
    Public Event NavigationControlGotFocus(ByRef navctl As Access.NavigationControl)
    Public Event NavigationControlKeyDown(ByRef navctl As Access.NavigationControl, KeyCode As Integer, Shift As Integer)
    Public Event NavigationControlKeyPress(ByRef navctl As Access.NavigationControl, KeyAscii As Integer)
    Public Event NavigationControlKeyUp(ByRef navctl As Access.NavigationControl, KeyCode As Integer, Shift As Integer)
    Public Event NavigationControlLostFocus(ByRef navctl As Access.NavigationControl)
    Public Event NavigationControlMouseDown(ByRef navctl As Access.NavigationControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event NavigationControlMouseMove(ByRef navctl As Access.NavigationControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event NavigationControlMouseUp(ByRef navctl As Access.NavigationControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event NavigationControlUndo(ByRef navctl As Access.NavigationControl, Cancel As Integer)
    
    Public Event ObjectFrameClick(ByRef fra As Access.ObjectFrame)
    Public Event ObjectFrameDblClick(ByRef fra As Access.ObjectFrame, Cancel As Integer)
    Public Event ObjectFrameEnter(ByRef fra As Access.ObjectFrame)
    Public Event ObjectFrameExit(ByRef fra As Access.ObjectFrame, Cancel As Integer)
    Public Event ObjectFrameGotFocus(ByRef fra As Access.ObjectFrame)
    Public Event ObjectFrameLostFocus(ByRef fra As Access.ObjectFrame)
    Public Event ObjectFrameMouseDown(ByRef fra As Access.ObjectFrame, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ObjectFrameMouseMove(ByRef fra As Access.ObjectFrame, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ObjectFrameMouseUp(ByRef fra As Access.ObjectFrame, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ObjectFrameUpdated(ByRef fra As Access.ObjectFrame, Code As Integer)
    
    Public Event OptionButtonAfterUpdate(ByRef opt As Access.OptionButton)
    Public Event OptionButtonBeforeUpdate(ByRef opt As Access.OptionButton, Cancel As Integer)
    Public Event OptionButtonClick(ByRef opt As Access.OptionButton)
    Public Event OptionButtonDblClick(ByRef opt As Access.OptionButton, Cancel As Integer)
    Public Event OptionButtonEnter(ByRef opt As Access.OptionButton)
    Public Event OptionButtonExit(ByRef opt As Access.OptionButton, Cancel As Integer)
    Public Event OptionButtonGotFocus(ByRef opt As Access.OptionButton)
    Public Event OptionButtonKeyDown(ByRef opt As Access.OptionButton, KeyCode As Integer, Shift As Integer)
    Public Event OptionButtonKeyPress(ByRef opt As Access.OptionButton, KeyAscii As Integer)
    Public Event OptionButtonKeyUp(ByRef opt As Access.OptionButton, KeyCode As Integer, Shift As Integer)
    Public Event OptionButtonLostFocus(ByRef opt As Access.OptionButton)
    Public Event OptionButtonMouseDown(ByRef opt As Access.OptionButton, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event OptionButtonMouseMove(ByRef opt As Access.OptionButton, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event OptionButtonMouseUp(ByRef opt As Access.OptionButton, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Public Event OptionGroupAfterUpdate(ByRef grp As Access.OptionGroup)
    Public Event OptionGroupBeforeUpdate(ByRef grp As Access.OptionGroup, Cancel As Integer)
    Public Event OptionGroupClick(ByRef grp As Access.OptionGroup)
    Public Event OptionGroupDblClick(ByRef grp As Access.OptionGroup, Cancel As Integer)
    Public Event OptionGroupEnter(ByRef grp As Access.OptionGroup)
    Public Event OptionGroupExit(ByRef grp As Access.OptionGroup, Cancel As Integer)
    Public Event OptionGroupMouseDown(ByRef grp As Access.OptionGroup, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event OptionGroupMouseMove(ByRef grp As Access.OptionGroup, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event OptionGroupMouseUp(ByRef grp As Access.OptionGroup, Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    Public Event PageClick(ByRef pg As Access.Page)
    Public Event PageDblClick(ByRef pg As Access.Page, Cancel As Integer)
    Public Event PageMouseDown(ByRef pg As Access.Page, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event PageMouseMove(ByRef pg As Access.Page, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event PageMouseUp(ByRef pg As Access.Page, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Public Event RectangleClick(ByRef rect As Access.Rectangle)
    Public Event RectangleDblClick(ByRef rect As Access.Rectangle, Cancel As Integer)
    Public Event RectangleMouseDown(ByRef rect As Access.Rectangle, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event RectangleMouseMove(ByRef rect As Access.Rectangle, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event RectangleMouseUp(ByRef rect As Access.Rectangle, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Public Event SubFormEnter(ByRef sfrm As Access.SubForm)
    Public Event SubFormExit(ByRef sfrm As Access.SubForm, Cancel As Integer)
    
    Public Event TabControlChange(ByRef tbctl As Access.TabControl)
    Public Event TabControlClick(ByRef tbctl As Access.TabControl)
    Public Event TabControlDblClick(ByRef tbctl As Access.TabControl, Cancel As Integer)
    Public Event TabControlKeyDown(ByRef tbctl As Access.TabControl, KeyCode As Integer, Shift As Integer)
    Public Event TabControlKeyPress(ByRef tbctl As Access.TabControl, KeyAscii As Integer)
    Public Event TabControlKeyUp(ByRef tbctl As Access.TabControl, KeyCode As Integer, Shift As Integer)
    Public Event TabControlMouseDown(ByRef tbctl As Access.TabControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event TabControlMouseMove(ByRef tbctl As Access.TabControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event TabControlMouseUp(ByRef tbctl As Access.TabControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Public Event TextBoxAfterUpdate(ByRef txt As Access.TextBox)
    Public Event TextBoxBeforeUpdate(ByRef txt As Access.TextBox, Cancel As Integer)
    Public Event TextBoxChange(ByRef txt As Access.TextBox)
    Public Event TextBoxClick(ByRef txt As Access.TextBox)
    Public Event TextBoxDblClick(ByRef txt As Access.TextBox, Cancel As Integer)
    Public Event TextBoxDirty(ByRef txt As Access.TextBox, Cancel As Integer)
    Public Event TextBoxEnter(ByRef txt As Access.TextBox)
    Public Event TextBoxExit(ByRef txt As Access.TextBox, Cancel As Integer)
    Public Event TextBoxGotFocus(ByRef txt As Access.TextBox)
    Public Event TextBoxKeyDown(ByRef txt As Access.TextBox, KeyCode As Integer, Shift As Integer)
    Public Event TextBoxKeyPress(ByRef txt As Access.TextBox, KeyAscii As Integer)
    Public Event TextBoxKeyUp(ByRef txt As Access.TextBox, KeyCode As Integer, Shift As Integer)
    Public Event TextBoxLostFocus(ByRef txt As Access.TextBox)
    Public Event TextBoxMouseDown(ByRef txt As Access.TextBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event TextBoxMouseMove(ByRef txt As Access.TextBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event TextBoxMouseUp(ByRef txt As Access.TextBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event TextBoxUndo(ByRef txt As Access.TextBox, Cancel As Integer)
    
    Public Event ToggleButtonAfterUpdate(ByRef tgl As Access.ToggleButton)
    Public Event ToggleButtonBeforeUpdate(ByRef tgl As Access.ToggleButton, Cancel As Integer)
    Public Event ToggleButtonClick(ByRef tgl As Access.ToggleButton)
    Public Event ToggleButtonDblClick(ByRef tgl As Access.ToggleButton, Cancel As Integer)
    Public Event ToggleButtonEnter(ByRef tgl As Access.ToggleButton)
    Public Event ToggleButtonExit(ByRef tgl As Access.ToggleButton, Cancel As Integer)
    Public Event ToggleButtonGotFocus(ByRef tgl As Access.ToggleButton)
    Public Event ToggleButtonKeyDown(ByRef tgl As Access.ToggleButton, KeyCode As Integer, Shift As Integer)
    Public Event ToggleButtonKeyPress(ByRef tgl As Access.ToggleButton, KeyAscii As Integer)
    Public Event ToggleButtonKeyUp(ByRef tgl As Access.ToggleButton, KeyCode As Integer, Shift As Integer)
    Public Event ToggleButtonLostFocus(ByRef tgl As Access.ToggleButton)
    Public Event ToggleButtonMouseDown(ByRef tgl As Access.ToggleButton, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ToggleButtonMouseMove(ByRef tgl As Access.ToggleButton, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ToggleButtonMouseUp(ByRef tgl As Access.ToggleButton, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Public Event WebBrowserAfterUpdate(ByRef webctl As Access.WebBrowserControl)
    Public Event WebBrowserBeforeUpdate(ByRef webctl As Access.WebBrowserControl, Cancel As Integer)
    Public Event WebBrowserBeforeNavigate2(ByRef webctl As Access.WebBrowserControl, ByVal pDisp As Object, ByRef URL As Variant, ByRef flags As Variant, ByRef TargetFrameName As Variant, ByRef PostData As Variant, ByRef Headers As Variant, ByRef Cancel As Boolean)
    Public Event WebBrowserChange(ByRef webctl As Access.WebBrowserControl)
    Public Event WebBrowserClick(ByRef webctl As Access.WebBrowserControl)
    Public Event WebBrowserDblClick(ByRef webctl As Access.WebBrowserControl, Cancel As Integer)
    Public Event WebBrowserDirty(ByRef webctl As Access.WebBrowserControl, Cancel As Integer)
    Public Event WebBrowserDocumentComplete(ByRef webctl As Access.WebBrowserControl, ByVal pDisp As Object, URL As Variant)
    Public Event WebBrowserEnter(ByRef webctl As Access.WebBrowserControl)
    Public Event WebBrowserExit(ByRef webctl As Access.WebBrowserControl, Cancel As Integer)
    Public Event WebBrowserGotFocus(ByRef webctl As Access.WebBrowserControl)
    Public Event WebBrowserKeyDown(ByRef webctl As Access.WebBrowserControl, KeyCode As Integer, Shift As Integer)
    Public Event WebBrowserKeyPress(ByRef webctl As Access.WebBrowserControl, KeyAscii As Integer)
    Public Event WebBrowserKeyUp(ByRef webctl As Access.WebBrowserControl, KeyCode As Integer, Shift As Integer)
    Public Event WebBrowserLostFocus(ByRef webctl As Access.WebBrowserControl)
    Public Event WebBrowserMouseDown(ByRef webctl As Access.WebBrowserControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event WebBrowserMouseMove(ByRef webctl As Access.WebBrowserControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event WebBrowserMouseUp(ByRef webctl As Access.WebBrowserControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event WebBrowserNavigateError(ByRef webctl As Access.WebBrowserControl, ByVal pDisp As Object, URL As Variant, TargetFrameName As Variant, StatusCode As Variant, Cancel As Boolean)
    Public Event WebBrowserProgressChange(ByRef webctl As Access.WebBrowserControl, ByVal Progress As Long, ByVal ProgressMax As Long)
    Public Event WebBrowserUpdated(ByRef webctl As Access.WebBrowserControl, Code As Integer)

#End If

Public Function Create(Optional ByRef ctls As Access.Controls, Optional ByVal eType As eControlTypes = ctAll, Optional ByVal strTagFilter As String) As ControlArray

'    On Error GoTo errHandler

    Dim ctl As Access.Control

    With New ControlArray
        
        Set This.Col = New Collection
    
        If ctls Is Nothing = False Then

            For Each ctl In ctls

                If Len(strTagFilter) = 0 Or StrComp(ctl.Tag & vbNullString, strTagFilter, vbTextCompare) = 0 Then
                    
                    Select Case ctl.ControlType
                        
                        Case ctAttachment And (eType = ctAll Or eType = ctAttachment)
                            Call This.Col.Add(CAAttachment.Create(ctl, .Self), ctl.Name)
                        Case ctBoundObjectFrame And (eType = ctAll Or eType = ctBoundObjectFrame)
                            Call This.Col.Add(CABoundObjectFrame.Create(ctl, .Self), ctl.Name)
                        Case ctChart And (eType = ctAll Or eType = ctChart)
                            Call This.Col.Add(CAChart.Create(ctl, .Self), ctl.Name)
                        Case ctCheckBox And (eType = ctAll Or eType = ctCheckBox)
                            Call This.Col.Add(CACheckBox.Create(ctl, .Self), ctl.Name)
                        Case ctComboBox And (eType = ctAll Or eType = ctComboBox)
                            Call This.Col.Add(CAComboBox.Create(ctl, .Self), ctl.Name)
                        Case ctCommandButton And (eType = ctAll Or eType = ctCommandButton)
                            Call This.Col.Add(CACommandButton.Create(ctl, .Self), ctl.Name)
                        Case ctImage And (eType = ctAll Or eType = ctImage)
                            Call This.Col.Add(CAImage.Create(ctl, .Self), ctl.Name)
                        Case ctLabel And (eType = ctAll Or eType = ctLabel)
                            Call This.Col.Add(CALabel.Create(ctl, .Self), ctl.Name)
                        Case ctListBox And (eType = ctAll Or eType = ctListBox)
                            Call This.Col.Add(CAListBox.Create(ctl, .Self), ctl.Name)
                        Case ctNavigationButton And (eType = ctAll Or eType = ctNavigationButton)
                            Call This.Col.Add(CANavigationButton.Create(ctl, .Self), ctl.Name)
                        Case ctNavigationControl And (eType = ctAll Or eType = ctNavigationControl)
                            Call This.Col.Add(CANavigationControl.Create(ctl, .Self), ctl.Name)
                        Case ctObjectFrame And (eType = ctAll Or eType = ctObjectFrame)
                            Call This.Col.Add(CAObjectFrame.Create(ctl, .Self), ctl.Name)
                        Case ctOptionButton And (eType = ctAll Or eType = ctOptionButton)
                            Call This.Col.Add(CAOptionButton.Create(ctl, .Self), ctl.Name)
                        Case ctOptionGroup And (eType = ctAll Or eType = ctOptionGroup)
                            Call This.Col.Add(CAOptionGroup.Create(ctl, .Self), ctl.Name)
                        Case ctPage And (eType = ctAll Or eType = ctPage)
                            Call This.Col.Add(CAPage.Create(ctl, .Self), ctl.Name)
                        Case ctRectangle And (eType = ctAll Or eType = ctRectangle)
                            Call This.Col.Add(CARectangle.Create(ctl, .Self), ctl.Name)
                        Case ctSubForm And (eType = ctAll Or eType = ctSubForm)
                            Call This.Col.Add(CASubForm.Create(ctl, .Self), ctl.Name)
                        Case ctTabCtl And (eType = ctAll Or eType = ctTabCtl)
                            Call This.Col.Add(CATabControl.Create(ctl, .Self), ctl.Name)
                        Case ctTextBox And (eType = ctAll Or eType = ctTextBox)
                            Call This.Col.Add(CATextBox.Create(ctl, .Self), ctl.Name)
                        Case ctToggleButton And (eType = ctAll Or eType = ctToggleButton)
                            Call This.Col.Add(CAToggleButton.Create(ctl, .Self), ctl.Name)
                        Case ctWebBrowserControl And (eType = ctAll Or eType = ctWebBrowserControl)
                            Call This.Col.Add(CAWebBrowserControl.Create(ctl, .Self), ctl.Name)
                        Case Else
                            'unsupported control for control array,ignore
                    End Select
                
                End If
                
            Next
        
        End If
        
        Call .setMembers(This)
        
        Set Create = .Self
        
    End With
'
'    Exit Function
'
'errHandler:
'    Debug.Assert False
'    Resume
End Function

Friend Function Self() As ControlArray
    Set Self = Me
End Function

Friend Sub setMembers(ByRef Data As udtMembers)
    This = Data
End Sub

Public Sub Add(ByRef ctl As IControl)
    Call This.Col.Add(ctl, ctl.Name)
End Sub

Public Sub Remove(ByVal Index As Variant)
    Call This.Col.Remove(Index)
End Sub

Public Function Count() As Long
    Count = This.Col.Count
End Function

Public Sub Clear()
    Set This.Col = Nothing
    Set This.Col = New Collection
End Sub

Public Function NewEnum() As IUnknown
Attribute NewEnum.VB_UserMemId = -4
    Set NewEnum = This.Col.[_NewEnum]
End Function

Public Function Item(ByVal Index As Variant) As IControl
Attribute Item.VB_UserMemId = 0
    Set Item = This.Col.Item(Index)
End Function

Public Property Get Exists(ByVal Index As Variant) As Boolean
    
    On Error Resume Next

    Call Me.Item(Index)

    If Err.Number = 0 Then
        Exists = True
    Else
        Exists = False
    End If

End Property

Friend Sub RaiseAfterUpdate(ByRef ctl As IControl)
    
    RaiseEvent onAfterUpdate(ctl)

    #If IncludeSpecificEvents = True Then
    
        Select Case ctl.ControlType
            Case ctAttachment: RaiseEvent AttachmentAfterUpdate(ctl.Object)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameAfterUpdate(ctl.Object)
            Case ctCheckBox: RaiseEvent CheckboxAfterUpdate(ctl.Object)
            Case ctComboBox: RaiseEvent ComboBoxAfterUpdate(ctl.Object)
            Case ctListBox: RaiseEvent ListBoxAfterUpdate(ctl.Object)
            Case ctNavigationControl: RaiseEvent NavigationControlAfterUpdate(ctl.Object)
            Case ctOptionButton: RaiseEvent OptionButtonAfterUpdate(ctl.Object)
            Case ctOptionGroup: RaiseEvent OptionGroupAfterUpdate(ctl.Object)
            Case ctTextBox: RaiseEvent TextBoxAfterUpdate(ctl.Object)
            Case ctToggleButton: RaiseEvent ToggleButtonAfterUpdate(ctl.Object)
            Case ctWebBrowserControl: RaiseEvent WebBrowserAfterUpdate(ctl.Object)
        End Select
        
    #End If
    
End Sub

Friend Sub RaiseAttachmentCurrent(ByRef ctl As IControl)

    RaiseEvent onAttachmentCurrent(ctl)
    
    #If IncludeSpecificEvents = True Then
        
        Select Case ctl.ControlType
            Case ctAttachment: RaiseEvent AttachmentCurrent(ctl.Object)
        End Select
    
    #End If
    
End Sub

Friend Sub RaiseBeforeNavigate2(ByRef ctl As IControl, ByVal pDisp As Object, ByRef URL As Variant, ByRef flags As Variant, ByRef TargetFrameName As Variant, ByRef PostData As Variant, ByRef Headers As Variant, ByRef Cancel As Boolean)
    
    RaiseEvent onBeforeNavigate2(ctl, pDisp, URL, flags, TargetFrameName, PostData, Headers, Cancel)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case ctl.ControlType
            Case ctWebBrowserControl: RaiseEvent WebBrowserBeforeNavigate2(ctl.Object, pDisp, URL, flags, TargetFrameName, PostData, Headers, Cancel)
        End Select
        
    #End If
    
End Sub

Friend Sub RaiseBeforeUpdate(ByRef ctl As IControl, Cancel As Integer)
    
    RaiseEvent onBeforeUpdate(ctl, Cancel)
        
    #If IncludeSpecificEvents = True Then
        
        Select Case ctl.ControlType
            Case ctAttachment: RaiseEvent AttachmentBeforeUpdate(ctl.Object, Cancel)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameBeforeUpdate(ctl.Object, Cancel)
            Case ctCheckBox: RaiseEvent CheckboxBeforeUpdate(ctl.Object, Cancel)
            Case ctComboBox: RaiseEvent ComboBoxBeforeUpdate(ctl.Object, Cancel)
            Case ctListBox: RaiseEvent ListBoxBeforeUpdate(ctl.Object, Cancel)
            Case ctNavigationControl: RaiseEvent NavigationControlBeforeUpdate(ctl.Object, Cancel)
            Case ctOptionButton: RaiseEvent OptionButtonBeforeUpdate(ctl.Object, Cancel)
            Case ctOptionGroup: RaiseEvent OptionGroupBeforeUpdate(ctl.Object, Cancel)
            Case ctTextBox: RaiseEvent TextBoxBeforeUpdate(ctl.Object, Cancel)
            Case ctToggleButton: RaiseEvent ToggleButtonBeforeUpdate(ctl.Object, Cancel)
            Case ctWebBrowserControl: RaiseEvent WebBrowserBeforeUpdate(ctl.Object, Cancel)
        End Select
        
    #End If
    
End Sub

Friend Sub RaiseChange(ByRef ctl As IControl)
    
    RaiseEvent onChange(ctl)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case ctl.ControlType
            Case ctAttachment: RaiseEvent AttachmentChange(ctl.Object)
            Case ctComboBox: RaiseEvent ComboBoxChange(ctl.Object)
            Case ctNavigationControl: RaiseEvent NavigationControlChange(ctl.Object)
            Case ctTabCtl: RaiseEvent TabControlChange(ctl.Object)
            Case ctTextBox: RaiseEvent TextBoxChange(ctl.Object)
            Case ctWebBrowserControl: RaiseEvent WebBrowserChange(ctl.Object)
        End Select
    
    #End If
    
End Sub

Friend Sub RaiseClick(ByRef ctl As IControl)
    
    RaiseEvent onClick(ctl)

    #If IncludeSpecificEvents = True Then
    
        Select Case ctl.ControlType
            Case ctAttachment: RaiseEvent AttachmentClick(ctl.Object)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameClick(ctl.Object)
            Case ctChart: RaiseEvent ChartClick(ctl.Object)
            Case ctCheckBox: RaiseEvent CheckboxClick(ctl.Object)
            Case ctComboBox: RaiseEvent ComboBoxClick(ctl.Object)
            Case ctCommandButton: RaiseEvent CommandButtonClick(ctl.Object)
            Case ctImage: RaiseEvent ImageClick(ctl.Object)
            Case ctLabel: RaiseEvent LabelClick(ctl.Object)
            Case ctListBox: RaiseEvent ListBoxClick(ctl.Object)
            Case ctNavigationButton: RaiseEvent NavigationButtonClick(ctl.Object)
            Case ctNavigationControl: RaiseEvent NavigationControlClick(ctl.Object)
            Case ctObjectFrame: RaiseEvent ObjectFrameClick(ctl.Object)
            Case ctOptionButton: RaiseEvent OptionButtonClick(ctl.Object)
            Case ctOptionGroup: RaiseEvent OptionGroupClick(ctl.Object)
            Case ctPage: RaiseEvent PageClick(ctl.Object)
            Case ctRectangle: RaiseEvent RectangleClick(ctl.Object)
            Case ctTabCtl: RaiseEvent TabControlClick(ctl.Object)
            Case ctTextBox: RaiseEvent TextBoxClick(ctl.Object)
            Case ctToggleButton: RaiseEvent ToggleButtonClick(ctl.Object)
            Case ctWebBrowserControl: RaiseEvent WebBrowserClick(ctl.Object)
        End Select
        
    #End If
        
End Sub

Friend Sub RaiseDblClick(ByRef ctl As IControl, Cancel As Integer)
    
    RaiseEvent onDblClick(ctl, Cancel)
    
    #If IncludeSpecificEvents = True Then
        
        Select Case ctl.ControlType
            Case ctAttachment: RaiseEvent AttachmentDblClick(ctl.Object, Cancel)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameDblClick(ctl.Object, Cancel)
            Case ctChart: RaiseEvent ChartDblClick(ctl.Object, Cancel)
            Case ctCheckBox: RaiseEvent CheckboxDblClick(ctl.Object, Cancel)
            Case ctComboBox: RaiseEvent ComboBoxDblClick(ctl.Object, Cancel)
            Case ctCommandButton: RaiseEvent CommandButtonDblClick(ctl.Object, Cancel)
            Case ctImage: RaiseEvent ImageDblClick(ctl.Object, Cancel)
            Case ctLabel: RaiseEvent LabelDblClick(ctl.Object, Cancel)
            Case ctListBox: RaiseEvent ListBoxDblClick(ctl.Object, Cancel)
            Case ctNavigationButton: RaiseEvent NavigationButtonDblClick(ctl.Object, Cancel)
            Case ctNavigationControl: RaiseEvent NavigationControlDblClick(ctl.Object, Cancel)
            Case ctObjectFrame: RaiseEvent ObjectFrameDblClick(ctl.Object, Cancel)
            Case ctOptionButton: RaiseEvent OptionButtonDblClick(ctl.Object, Cancel)
            Case ctOptionGroup: RaiseEvent OptionGroupDblClick(ctl.Object, Cancel)
            Case ctPage: RaiseEvent PageDblClick(ctl.Object, Cancel)
            Case ctRectangle: RaiseEvent RectangleDblClick(ctl.Object, Cancel)
            Case ctTabCtl: RaiseEvent TabControlDblClick(ctl.Object, Cancel)
            Case ctTextBox: RaiseEvent TextBoxDblClick(ctl.Object, Cancel)
            Case ctToggleButton: RaiseEvent ToggleButtonDblClick(ctl.Object, Cancel)
            Case ctWebBrowserControl: RaiseEvent WebBrowserDblClick(ctl.Object, Cancel)
        End Select
            
    #End If
    
End Sub

Friend Sub RaiseDirty(ByRef ctl As IControl, Cancel As Integer)
    
    RaiseEvent onDirty(ctl, Cancel)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case ctl.ControlType
            Case ctAttachment: RaiseEvent AttachmentDirty(ctl.Object, Cancel)
            Case ctComboBox: RaiseEvent ComboBoxDirty(ctl.Object, Cancel)
            Case ctNavigationControl: RaiseEvent NavigationControlDirty(ctl.Object, Cancel)
            Case ctTextBox: RaiseEvent TextBoxDirty(ctl.Object, Cancel)
            Case ctWebBrowserControl: RaiseEvent WebBrowserDirty(ctl.Object, Cancel)
        End Select
            
    #End If
    
End Sub

Friend Sub RaiseDocumentComplete(ByRef ctl As IControl, ByVal pDisp As Object, URL As Variant)
    
    RaiseEvent onDocumentComplete(ctl, pDisp, URL)

    #If IncludeSpecificEvents = True Then

        Select Case ctl.ControlType
            Case ctWebBrowserControl: RaiseEvent WebBrowserDocumentComplete(ctl.Object, pDisp, URL)
        End Select
        
    #End If
    
End Sub

Friend Sub RaiseEnter(ByRef ctl As IControl)
    
    RaiseEvent onEnter(ctl)
    
    #If IncludeSpecificEvents = True Then
        
        Select Case ctl.ControlType
            Case ctAttachment: RaiseEvent AttachmentEnter(ctl.Object)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameEnter(ctl.Object)
            Case ctChart: RaiseEvent ChartEnter(ctl.Object)
            Case ctCheckBox: RaiseEvent CheckboxEnter(ctl.Object)
            Case ctComboBox: RaiseEvent ComboBoxEnter(ctl.Object)
            Case ctCommandButton: RaiseEvent CommandButtonEnter(ctl.Object)
            Case ctListBox: RaiseEvent ListBoxEnter(ctl.Object)
            Case ctNavigationButton: RaiseEvent NavigationButtonEnter(ctl.Object)
            Case ctNavigationControl: RaiseEvent NavigationControlEnter(ctl.Object)
            Case ctObjectFrame: RaiseEvent ObjectFrameEnter(ctl.Object)
            Case ctOptionButton: RaiseEvent OptionButtonEnter(ctl.Object)
            Case ctOptionGroup: RaiseEvent OptionGroupEnter(ctl.Object)
            Case ctSubForm: RaiseEvent SubFormEnter(ctl.Object)
            Case ctTextBox: RaiseEvent TextBoxEnter(ctl.Object)
            Case ctToggleButton: RaiseEvent ToggleButtonEnter(ctl.Object)
            Case ctWebBrowserControl: RaiseEvent WebBrowserEnter(ctl.Object)
        End Select
    
    #End If
    
End Sub

Friend Sub RaiseExit(ByRef ctl As IControl, Cancel As Integer)
    
    RaiseEvent onExit(ctl, Cancel)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case ctl.ControlType
            Case ctAttachment: RaiseEvent AttachmentExit(ctl.Object, Cancel)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameExit(ctl.Object, Cancel)
            Case ctChart: RaiseEvent ChartExit(ctl.Object, Cancel)
            Case ctCheckBox: RaiseEvent CheckboxExit(ctl.Object, Cancel)
            Case ctComboBox: RaiseEvent ComboBoxExit(ctl.Object, Cancel)
            Case ctCommandButton: RaiseEvent CommandButtonExit(ctl.Object, Cancel)
            Case ctListBox: RaiseEvent ListBoxExit(ctl.Object, Cancel)
            Case ctNavigationButton: RaiseEvent NavigationButtonExit(ctl.Object, Cancel)
            Case ctNavigationControl: RaiseEvent NavigationControlExit(ctl.Object, Cancel)
            Case ctObjectFrame: RaiseEvent ObjectFrameExit(ctl.Object, Cancel)
            Case ctOptionButton: RaiseEvent OptionButtonExit(ctl.Object, Cancel)
            Case ctOptionGroup: RaiseEvent OptionGroupExit(ctl.Object, Cancel)
            Case ctSubForm: RaiseEvent SubFormExit(ctl.Object, Cancel)
            Case ctTextBox: RaiseEvent TextBoxExit(ctl.Object, Cancel)
            Case ctToggleButton: RaiseEvent ToggleButtonExit(ctl.Object, Cancel)
            Case ctWebBrowserControl: RaiseEvent WebBrowserExit(ctl.Object, Cancel)
        End Select
        
    #End If
    
End Sub

Friend Sub RaiseGotFocus(ByRef ctl As IControl)
    
    RaiseEvent onGotFocus(ctl)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case ctl.ControlType
            Case ctAttachment: RaiseEvent AttachmentGotFocus(ctl.Object)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameGotFocus(ctl.Object)
            Case ctChart: RaiseEvent ChartGotFocus(ctl.Object)
            Case ctCheckBox: RaiseEvent CheckboxGotFocus(ctl.Object)
            Case ctComboBox: RaiseEvent ComboBoxGotFocus(ctl.Object)
            Case ctCommandButton: RaiseEvent CommandButtonGotFocus(ctl.Object)
            Case ctListBox: RaiseEvent ListBoxGotFocus(ctl.Object)
            Case ctNavigationButton: RaiseEvent NavigationButtonGotFocus(ctl.Object)
            Case ctNavigationControl: RaiseEvent NavigationControlGotFocus(ctl.Object)
            Case ctObjectFrame: RaiseEvent ObjectFrameGotFocus(ctl.Object)
            Case ctOptionButton: RaiseEvent OptionButtonGotFocus(ctl.Object)
            Case ctTextBox: RaiseEvent TextBoxGotFocus(ctl.Object)
            Case ctToggleButton: RaiseEvent ToggleButtonGotFocus(ctl.Object)
            Case ctWebBrowserControl: RaiseEvent WebBrowserGotFocus(ctl.Object)
        End Select
        
    #End If
    
End Sub

Friend Sub RaiseKeyDown(ByRef ctl As IControl, KeyCode As Integer, Shift As Integer)
    
    RaiseEvent onKeyDown(ctl, KeyCode, Shift)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case ctl.ControlType
            Case ctAttachment: RaiseEvent AttachmentKeyDown(ctl.Object, KeyCode, Shift)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameKeyDown(ctl.Object, KeyCode, Shift)
            Case ctCheckBox: RaiseEvent CheckboxKeyDown(ctl.Object, KeyCode, Shift)
            Case ctComboBox: RaiseEvent ComboBoxKeyDown(ctl.Object, KeyCode, Shift)
            Case ctCommandButton: RaiseEvent CommandButtonKeyDown(ctl.Object, KeyCode, Shift)
            Case ctListBox: RaiseEvent ListBoxKeyDown(ctl.Object, KeyCode, Shift)
            Case ctNavigationButton: RaiseEvent NavigationButtonKeyDown(ctl.Object, KeyCode, Shift)
            Case ctNavigationControl: RaiseEvent NavigationControlKeyDown(ctl.Object, KeyCode, Shift)
            Case ctOptionButton: RaiseEvent OptionButtonKeyDown(ctl.Object, KeyCode, Shift)
            Case ctTabCtl: RaiseEvent TabControlKeyDown(ctl.Object, KeyCode, Shift)
            Case ctTextBox: RaiseEvent TextBoxKeyDown(ctl.Object, KeyCode, Shift)
            Case ctToggleButton: RaiseEvent ToggleButtonKeyDown(ctl.Object, KeyCode, Shift)
            Case ctWebBrowserControl: RaiseEvent WebBrowserKeyDown(ctl.Object, KeyCode, Shift)
        End Select
        
    #End If
    
End Sub

Friend Sub RaiseKeyPress(ByRef ctl As IControl, KeyAscii As Integer)
    
    RaiseEvent onKeyPress(ctl, KeyAscii)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case ctl.ControlType
            Case ctAttachment: RaiseEvent AttachmentKeyPress(ctl.Object, KeyAscii)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameKeyPress(ctl.Object, KeyAscii)
            Case ctCheckBox: RaiseEvent CheckboxKeyPress(ctl.Object, KeyAscii)
            Case ctComboBox: RaiseEvent ComboBoxKeyPress(ctl.Object, KeyAscii)
            Case ctCommandButton: RaiseEvent CommandButtonKeyPress(ctl.Object, KeyAscii)
            Case ctListBox: RaiseEvent ListBoxKeyPress(ctl.Object, KeyAscii)
            Case ctNavigationButton: RaiseEvent NavigationButtonKeyPress(ctl.Object, KeyAscii)
            Case ctNavigationControl: RaiseEvent NavigationControlKeyPress(ctl.Object, KeyAscii)
            Case ctOptionButton: RaiseEvent OptionButtonKeyPress(ctl.Object, KeyAscii)
            Case ctTabCtl: RaiseEvent TabControlKeyPress(ctl.Object, KeyAscii)
            Case ctTextBox: RaiseEvent TextBoxKeyPress(ctl.Object, KeyAscii)
            Case ctToggleButton: RaiseEvent ToggleButtonKeyPress(ctl.Object, KeyAscii)
            Case ctWebBrowserControl: RaiseEvent WebBrowserKeyPress(ctl.Object, KeyAscii)
        End Select
        
    #End If
    
End Sub

Friend Sub RaiseKeyUp(ByRef ctl As IControl, KeyCode As Integer, Shift As Integer)
    
    RaiseEvent onKeyUp(ctl, KeyCode, Shift)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case ctl.ControlType
            Case ctAttachment: RaiseEvent AttachmentKeyUp(ctl.Object, KeyCode, Shift)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameKeyUp(ctl.Object, KeyCode, Shift)
            Case ctCheckBox: RaiseEvent CheckboxKeyUp(ctl.Object, KeyCode, Shift)
            Case ctComboBox: RaiseEvent ComboBoxKeyUp(ctl.Object, KeyCode, Shift)
            Case ctCommandButton: RaiseEvent CommandButtonKeyUp(ctl.Object, KeyCode, Shift)
            Case ctListBox: RaiseEvent ListBoxKeyUp(ctl.Object, KeyCode, Shift)
            Case ctNavigationButton: RaiseEvent NavigationButtonKeyUp(ctl.Object, KeyCode, Shift)
            Case ctNavigationControl: RaiseEvent NavigationControlKeyUp(ctl.Object, KeyCode, Shift)
            Case ctOptionButton: RaiseEvent OptionButtonKeyUp(ctl.Object, KeyCode, Shift)
            Case ctTabCtl: RaiseEvent TabControlKeyUp(ctl.Object, KeyCode, Shift)
            Case ctTextBox: RaiseEvent TextBoxKeyUp(ctl.Object, KeyCode, Shift)
            Case ctToggleButton: RaiseEvent ToggleButtonKeyUp(ctl.Object, KeyCode, Shift)
            Case ctWebBrowserControl: RaiseEvent WebBrowserKeyUp(ctl.Object, KeyCode, Shift)
        End Select
    
    #End If
    
End Sub

Friend Sub RaiseLostFocus(ByRef ctl As IControl)
    
    RaiseEvent onLostFocus(ctl)
    
    #If IncludeSpecificEvents = True Then
        
        Select Case ctl.ControlType
            Case ctAttachment: RaiseEvent AttachmentLostFocus(ctl.Object)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameLostFocus(ctl.Object)
            Case ctChart: RaiseEvent ChartLostFocus(ctl.Object)
            Case ctCheckBox: RaiseEvent CheckboxLostFocus(ctl.Object)
            Case ctComboBox: RaiseEvent ComboBoxLostFocus(ctl.Object)
            Case ctCommandButton: RaiseEvent CommandButtonLostFocus(ctl.Object)
            Case ctListBox: RaiseEvent ListBoxLostFocus(ctl.Object)
            Case ctNavigationButton: RaiseEvent NavigationButtonLostFocus(ctl.Object)
            Case ctNavigationControl: RaiseEvent NavigationControlLostFocus(ctl.Object)
            Case ctObjectFrame: RaiseEvent ObjectFrameLostFocus(ctl.Object)
            Case ctOptionButton: RaiseEvent OptionButtonLostFocus(ctl.Object)
            Case ctTextBox: RaiseEvent TextBoxLostFocus(ctl.Object)
            Case ctToggleButton: RaiseEvent ToggleButtonLostFocus(ctl.Object)
            Case ctWebBrowserControl: RaiseEvent WebBrowserLostFocus(ctl.Object)
        End Select
    
    #End If
    
End Sub

Friend Sub RaiseMouseDown(ByRef ctl As IControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    RaiseEvent onMouseDown(ctl, Button, Shift, X, Y)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case ctl.ControlType
            Case ctAttachment: RaiseEvent AttachmentMouseDown(ctl.Object, Button, Shift, X, Y)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameMouseDown(ctl.Object, Button, Shift, X, Y)
            Case ctChart: RaiseEvent ChartMouseDown(ctl.Object, Button, Shift, X, Y)
            Case ctCheckBox: RaiseEvent CheckboxMouseDown(ctl.Object, Button, Shift, X, Y)
            Case ctComboBox: RaiseEvent ComboBoxMouseDown(ctl.Object, Button, Shift, X, Y)
            Case ctCommandButton: RaiseEvent CommandButtonMouseDown(ctl.Object, Button, Shift, X, Y)
            Case ctImage: RaiseEvent ImageMouseDown(ctl.Object, Button, Shift, X, Y)
            Case ctLabel: RaiseEvent LabelMouseDown(ctl.Object, Button, Shift, X, Y)
            Case ctListBox: RaiseEvent ListBoxMouseDown(ctl.Object, Button, Shift, X, Y)
            Case ctNavigationButton: RaiseEvent NavigationButtonMouseDown(ctl.Object, Button, Shift, X, Y)
            Case ctNavigationControl: RaiseEvent NavigationControlMouseDown(ctl.Object, Button, Shift, X, Y)
            Case ctObjectFrame: RaiseEvent ObjectFrameMouseDown(ctl.Object, Button, Shift, X, Y)
            Case ctOptionButton: RaiseEvent OptionButtonMouseDown(ctl.Object, Button, Shift, X, Y)
            Case ctOptionGroup: RaiseEvent OptionGroupMouseDown(ctl.Object, Button, Shift, X, Y)
            Case ctPage: RaiseEvent PageMouseDown(ctl.Object, Button, Shift, X, Y)
            Case ctRectangle: RaiseEvent RectangleMouseDown(ctl.Object, Button, Shift, X, Y)
            Case ctTabCtl: RaiseEvent TabControlMouseDown(ctl.Object, Button, Shift, X, Y)
            Case ctTextBox: RaiseEvent TextBoxMouseDown(ctl.Object, Button, Shift, X, Y)
            Case ctToggleButton: RaiseEvent ToggleButtonMouseDown(ctl.Object, Button, Shift, X, Y)
            Case ctWebBrowserControl: RaiseEvent WebBrowserMouseDown(ctl.Object, Button, Shift, X, Y)
        End Select
    
    #End If
    
End Sub

Friend Sub RaiseMouseMove(ByRef ctl As IControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    RaiseEvent onMouseMove(ctl, Button, Shift, X, Y)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case ctl.ControlType
            Case ctAttachment: RaiseEvent AttachmentMouseMove(ctl.Object, Button, Shift, X, Y)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameMouseMove(ctl.Object, Button, Shift, X, Y)
            Case ctChart: RaiseEvent ChartMouseMove(ctl.Object, Button, Shift, X, Y)
            Case ctCheckBox: RaiseEvent CheckboxMouseMove(ctl.Object, Button, Shift, X, Y)
            Case ctComboBox: RaiseEvent ComboBoxMouseMove(ctl.Object, Button, Shift, X, Y)
            Case ctCommandButton: RaiseEvent CommandButtonMouseMove(ctl.Object, Button, Shift, X, Y)
            Case ctImage: RaiseEvent ImageMouseMove(ctl.Object, Button, Shift, X, Y)
            Case ctLabel: RaiseEvent LabelMouseMove(ctl.Object, Button, Shift, X, Y)
            Case ctListBox: RaiseEvent ListBoxMouseMove(ctl.Object, Button, Shift, X, Y)
            Case ctNavigationButton: RaiseEvent NavigationButtonMouseMove(ctl.Object, Button, Shift, X, Y)
            Case ctNavigationControl: RaiseEvent NavigationControlMouseMove(ctl.Object, Button, Shift, X, Y)
            Case ctObjectFrame: RaiseEvent ObjectFrameMouseMove(ctl.Object, Button, Shift, X, Y)
            Case ctOptionButton: RaiseEvent OptionButtonMouseMove(ctl.Object, Button, Shift, X, Y)
            Case ctOptionGroup: RaiseEvent OptionGroupMouseMove(ctl.Object, Button, Shift, X, Y)
            Case ctPage: RaiseEvent PageMouseMove(ctl.Object, Button, Shift, X, Y)
            Case ctRectangle: RaiseEvent RectangleMouseMove(ctl.Object, Button, Shift, X, Y)
            Case ctTabCtl: RaiseEvent TabControlMouseMove(ctl.Object, Button, Shift, X, Y)
            Case ctTextBox: RaiseEvent TextBoxMouseMove(ctl.Object, Button, Shift, X, Y)
            Case ctToggleButton: RaiseEvent ToggleButtonMouseMove(ctl.Object, Button, Shift, X, Y)
            Case ctWebBrowserControl: RaiseEvent WebBrowserMouseMove(ctl.Object, Button, Shift, X, Y)
        End Select
    
    #End If
    
End Sub

Friend Sub RaiseMouseUp(ByRef ctl As IControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    RaiseEvent onMouseUp(ctl, Button, Shift, X, Y)
    
    #If IncludeSpecificEvents = True Then
        
        Select Case ctl.ControlType
            Case ctAttachment: RaiseEvent AttachmentMouseUp(ctl.Object, Button, Shift, X, Y)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameMouseUp(ctl.Object, Button, Shift, X, Y)
            Case ctChart: RaiseEvent ChartMouseUp(ctl.Object, Button, Shift, X, Y)
            Case ctCheckBox: RaiseEvent CheckboxMouseUp(ctl.Object, Button, Shift, X, Y)
            Case ctComboBox: RaiseEvent ComboBoxMouseUp(ctl.Object, Button, Shift, X, Y)
            Case ctCommandButton: RaiseEvent CommandButtonMouseUp(ctl.Object, Button, Shift, X, Y)
            Case ctImage: RaiseEvent ImageMouseUp(ctl.Object, Button, Shift, X, Y)
            Case ctLabel: RaiseEvent LabelMouseUp(ctl.Object, Button, Shift, X, Y)
            Case ctListBox: RaiseEvent ListBoxMouseUp(ctl.Object, Button, Shift, X, Y)
            Case ctNavigationButton: RaiseEvent NavigationButtonMouseUp(ctl.Object, Button, Shift, X, Y)
            Case ctNavigationControl: RaiseEvent NavigationControlMouseUp(ctl.Object, Button, Shift, X, Y)
            Case ctObjectFrame: RaiseEvent ObjectFrameMouseUp(ctl.Object, Button, Shift, X, Y)
            Case ctOptionButton: RaiseEvent OptionButtonMouseUp(ctl.Object, Button, Shift, X, Y)
            Case ctOptionGroup: RaiseEvent OptionGroupMouseUp(ctl.Object, Button, Shift, X, Y)
            Case ctPage: RaiseEvent PageMouseUp(ctl.Object, Button, Shift, X, Y)
            Case ctRectangle: RaiseEvent RectangleMouseUp(ctl.Object, Button, Shift, X, Y)
            Case ctTabCtl: RaiseEvent TabControlMouseUp(ctl.Object, Button, Shift, X, Y)
            Case ctTextBox: RaiseEvent TextBoxMouseUp(ctl.Object, Button, Shift, X, Y)
            Case ctToggleButton: RaiseEvent ToggleButtonMouseUp(ctl.Object, Button, Shift, X, Y)
            Case ctWebBrowserControl: RaiseEvent WebBrowserMouseUp(ctl.Object, Button, Shift, X, Y)
        End Select
        
    #End If
    
End Sub

Friend Sub RaiseNavigateError(ByRef ctl As IControl, ByVal pDisp As Object, URL As Variant, TargetFrameName As Variant, StatusCode As Variant, Cancel As Boolean)
    
    RaiseEvent onNavigateError(ctl, pDisp, URL, TargetFrameName, StatusCode, Cancel)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case ctl.ControlType
            Case ctWebBrowserControl: RaiseEvent WebBrowserNavigateError(ctl.Object, pDisp, URL, TargetFrameName, StatusCode, Cancel)
        End Select
    
    #End If
    
End Sub

Friend Sub RaiseNotInList(ByRef ctl As IControl, NewData As String, Response As Integer)
    
    RaiseEvent onNotInList(ctl, NewData, Response)

    #If IncludeSpecificEvents = True Then
    
        Select Case ctl.ControlType
            Case ctComboBox: RaiseEvent ComboBoxNotInList(ctl.Object, NewData, Response)
        End Select
        
    #End If
    
End Sub

Friend Sub RaiseProgressChange(ByRef ctl As IControl, ByVal Progress As Long, ByVal ProgressMax As Long)
    
    RaiseEvent onProgressChange(ctl, Progress, ProgressMax)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case ctl.ControlType
            Case ctWebBrowserControl: RaiseEvent WebBrowserProgressChange(ctl.Object, Progress, ProgressMax)
        End Select

    #End If
    
End Sub

Friend Sub RaiseUndo(ByRef ctl As IControl, Cancel As Integer)
    
    RaiseEvent onUndo(ctl, Cancel)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case ctl.ControlType
            Case ctComboBox: RaiseEvent ComboBoxUndo(ctl.Object, Cancel)
            Case ctNavigationControl: RaiseEvent NavigationControlUndo(ctl.Object, Cancel)
            Case ctTextBox: RaiseEvent TextBoxUndo(ctl.Object, Cancel)
        End Select
    
    #End If
    
End Sub

Friend Sub RaiseUpdated(ByRef ctl As IControl, Code As Integer)
    
    RaiseEvent onUpdated(ctl, Code)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case ctl.ControlType
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameUpdated(ctl.Object, Code)
            Case ctChart: RaiseEvent ChartUpdated(ctl.Object, Code)
            Case ctObjectFrame: RaiseEvent ObjectFrameUpdated(ctl.Object, Code)
            Case ctWebBrowserControl: RaiseEvent WebBrowserUpdated(ctl.Object, Code)
        End Select
    
    #End If
    
End Sub

'Fix for VBIDE bug that does not correct case on enumerations
#If False Then
    
    #Const ctAll = -1
    #Const ctAttachment = acAttachment
    #Const ctBoundObjectFrame = acBoundObjectFrame
    #Const ctChart = 113
    #Const ctCheckBox = acCheckBox
    #Const ctComboBox = acComboBox
    #Const ctCommandButton = acCommandButton
    #Const ctImage = acImage
    #Const ctLabel = acLabel
    #Const ctListBox = acListBox
    #Const ctNavigationButton = acNavigationButton
    #Const ctNavigationControl = acNavigationControl
    #Const ctObjectFrame = acObjectFrame
    #Const ctOptionButton = acOptionButton
    #Const ctOptionGroup = acOptionGroup
    #Const ctPage = acPage
    #Const ctRectangle = acRectangle
    #Const ctSubForm = acSubform
    #Const ctTabCtl = acTabCtl
    #Const ctTextBox = acTextBox
    #Const ctToggleButton = acToggleButton
    #Const ctWebBrowserControl = acWebBrowser
    
#End If
