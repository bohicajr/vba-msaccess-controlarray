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

#Const IncludeSpecificEvents = True 'toggle if you want to see strongly typed events raised

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

Public Event onAfterUpdate(ByRef objControl As IControl)
Public Event onAttachmentCurrent(ByRef objControl As IControl)
Public Event onBeforeUpdate(ByRef objControl As IControl, Cancel As Integer)
Public Event onBeforeNavigate2(ByRef objControl As IControl, ByVal pDisp As Object, ByRef URL As Variant, ByRef flags As Variant, ByRef TargetFrameName As Variant, ByRef PostData As Variant, ByRef Headers As Variant, ByRef Cancel As Boolean)
Public Event onChange(ByRef objControl As IControl)
Public Event onClick(ByRef objControl As IControl)
Public Event onDblClick(ByRef objControl As IControl, Cancel As Integer)
Public Event onDirty(ByRef objControl As IControl, Cancel As Integer)
Public Event onDocumentComplete(ByRef objControl As IControl, ByVal pDisp As Object, URL As Variant)
Public Event onEnter(ByRef objControl As IControl)
Public Event onExit(ByRef objControl As IControl, Cancel As Integer)
Public Event onGotFocus(ByRef objControl As IControl)
Public Event onKeyDown(ByRef objControl As IControl, KeyCode As Integer, Shift As Integer)
Public Event onKeyPress(ByRef objControl As IControl, KeyAscii As Integer)
Public Event onKeyUp(ByRef objControl As IControl, KeyCode As Integer, Shift As Integer)
Public Event onLostFocus(ByRef objControl As IControl)
Public Event onMouseDown(ByRef objControl As IControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event onMouseMove(ByRef objControl As IControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event onMouseUp(ByRef objControl As IControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event onNavigateError(ByRef objControl As IControl, ByVal pDisp As Object, URL As Variant, TargetFrameName As Variant, StatusCode As Variant, Cancel As Boolean)
Public Event onNotInList(ByRef objControl As IControl, NewData As String, Response As Integer)
Public Event onProgressChange(ByRef objControl As IControl, ByVal Progress As Long, ByVal ProgressMax As Long)
Public Event onUndo(ByRef objControl As IControl, Cancel As Integer)
Public Event onUpdated(ByRef objControl As IControl, Code As Integer)

#If IncludeSpecificEvents = True Then
    
    'strongly typed events
    Public Event AttachmentAfterUpdate(ByRef objAttachment As Access.Attachment)
    Public Event AttachmentCurrent(ByRef objAttachment As Access.Attachment)
    Public Event AttachmentBeforeUpdate(ByRef objAttachment As Access.Attachment, Cancel As Integer)
    Public Event AttachmentChange(ByRef objAttachment As Access.Attachment)
    Public Event AttachmentClick(ByRef objAttachment As Access.Attachment)
    Public Event AttachmentDblClick(ByRef objAttachment As Access.Attachment, Cancel As Integer)
    Public Event AttachmentDirty(ByRef objAttachment As Access.Attachment, Cancel As Integer)
    Public Event AttachmentEnter(ByRef objAttachment As Access.Attachment)
    Public Event AttachmentExit(ByRef objAttachment As Access.Attachment, Cancel As Integer)
    Public Event AttachmentGotFocus(ByRef objAttachment As Access.Attachment)
    Public Event AttachmentKeyDown(ByRef objAttachment As Access.Attachment, KeyCode As Integer, Shift As Integer)
    Public Event AttachmentKeyPress(ByRef objAttachment As Access.Attachment, KeyAscii As Integer)
    Public Event AttachmentKeyUp(ByRef objAttachment As Access.Attachment, KeyCode As Integer, Shift As Integer)
    Public Event AttachmentLostFocus(ByRef objAttachment As Access.Attachment)
    Public Event AttachmentMouseDown(ByRef objAttachment As Access.Attachment, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event AttachmentMouseMove(ByRef objAttachment As Access.Attachment, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event AttachmentMouseUp(ByRef objAttachment As Access.Attachment, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Public Event BoundObjectFrameAfterUpdate(ByRef objBoundObjectFrame As Access.BoundObjectFrame)
    Public Event BoundObjectFrameBeforeUpdate(ByRef objBoundObjectFrame As Access.BoundObjectFrame, Cancel As Integer)
    Public Event BoundObjectFrameClick(ByRef objBoundObjectFrame As Access.BoundObjectFrame)
    Public Event BoundObjectFrameDblClick(ByRef objBoundObjectFrame As Access.BoundObjectFrame, Cancel As Integer)
    Public Event BoundObjectFrameEnter(ByRef objBoundObjectFrame As Access.BoundObjectFrame)
    Public Event BoundObjectFrameExit(ByRef objBoundObjectFrame As Access.BoundObjectFrame, Cancel As Integer)
    Public Event BoundObjectFrameGotFocus(ByRef objBoundObjectFrame As Access.BoundObjectFrame)
    Public Event BoundObjectFrameKeyDown(ByRef objBoundObjectFrame As Access.BoundObjectFrame, KeyCode As Integer, Shift As Integer)
    Public Event BoundObjectFrameKeyPress(ByRef objBoundObjectFrame As Access.BoundObjectFrame, KeyAscii As Integer)
    Public Event BoundObjectFrameKeyUp(ByRef objBoundObjectFrame As Access.BoundObjectFrame, KeyCode As Integer, Shift As Integer)
    Public Event BoundObjectFrameLostFocus(ByRef objBoundObjectFrame As Access.BoundObjectFrame)
    Public Event BoundObjectFrameMouseDown(ByRef objBoundObjectFrame As Access.BoundObjectFrame, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event BoundObjectFrameMouseMove(ByRef objBoundObjectFrame As Access.BoundObjectFrame, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event BoundObjectFrameMouseUp(ByRef objBoundObjectFrame As Access.BoundObjectFrame, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event BoundObjectFrameUpdated(ByRef objBoundObjectFrame As Access.BoundObjectFrame, Code As Integer)
    
    Public Event ChartClick(ByRef objObjectFrame As Access.ObjectFrame)
    Public Event ChartDblClick(ByRef objObjectFrame As Access.ObjectFrame, Cancel As Integer)
    Public Event ChartEnter(ByRef objObjectFrame As Access.ObjectFrame)
    Public Event ChartExit(ByRef objObjectFrame As Access.ObjectFrame, Cancel As Integer)
    Public Event ChartGotFocus(ByRef objObjectFrame As Access.ObjectFrame)
    Public Event ChartLostFocus(ByRef objObjectFrame As Access.ObjectFrame)
    Public Event ChartMouseDown(ByRef objObjectFrame As Access.ObjectFrame, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ChartMouseMove(ByRef objObjectFrame As Access.ObjectFrame, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ChartMouseUp(ByRef objObjectFrame As Access.ObjectFrame, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ChartUpdated(ByRef objObjectFrame As Access.ObjectFrame, Code As Integer)

    Public Event CheckboxAfterUpdate(ByRef objCheckbox As Access.CheckBox)
    Public Event CheckboxBeforeUpdate(ByRef objCheckbox As Access.CheckBox, Cancel As Integer)
    Public Event CheckboxClick(ByRef objCheckbox As Access.CheckBox)
    Public Event CheckboxDblClick(ByRef objCheckbox As Access.CheckBox, Cancel As Integer)
    Public Event CheckboxEnter(ByRef objCheckbox As Access.CheckBox)
    Public Event CheckboxExit(ByRef objCheckbox As Access.CheckBox, Cancel As Integer)
    Public Event CheckboxGotFocus(ByRef objCheckbox As Access.CheckBox)
    Public Event CheckboxKeyDown(ByRef objCheckbox As Access.CheckBox, KeyCode As Integer, Shift As Integer)
    Public Event CheckboxKeyPress(ByRef objCheckbox As Access.CheckBox, KeyAscii As Integer)
    Public Event CheckboxKeyUp(ByRef objCheckbox As Access.CheckBox, KeyCode As Integer, Shift As Integer)
    Public Event CheckboxLostFocus(ByRef objCheckbox As Access.CheckBox)
    Public Event CheckboxMouseDown(ByRef objCheckbox As Access.CheckBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event CheckboxMouseMove(ByRef objCheckbox As Access.CheckBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event CheckboxMouseUp(ByRef objCheckbox As Access.CheckBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Public Event ComboBoxAfterUpdate(ByRef objComboBox As Access.ComboBox)
    Public Event ComboBoxBeforeUpdate(ByRef objComboBox As Access.ComboBox, Cancel As Integer)
    Public Event ComboBoxChange(ByRef objComboBox As Access.ComboBox)
    Public Event ComboBoxClick(ByRef objComboBox As Access.ComboBox)
    Public Event ComboBoxDblClick(ByRef objComboBox As Access.ComboBox, Cancel As Integer)
    Public Event ComboBoxDirty(ByRef objComboBox As Access.ComboBox, Cancel As Integer)
    Public Event ComboBoxEnter(ByRef objComboBox As Access.ComboBox)
    Public Event ComboBoxExit(ByRef objComboBox As Access.ComboBox, Cancel As Integer)
    Public Event ComboBoxGotFocus(ByRef objComboBox As Access.ComboBox)
    Public Event ComboBoxKeyDown(ByRef objComboBox As Access.ComboBox, KeyCode As Integer, Shift As Integer)
    Public Event ComboBoxKeyPress(ByRef objComboBox As Access.ComboBox, KeyAscii As Integer)
    Public Event ComboBoxKeyUp(ByRef objComboBox As Access.ComboBox, KeyCode As Integer, Shift As Integer)
    Public Event ComboBoxLostFocus(ByRef objComboBox As Access.ComboBox)
    Public Event ComboBoxMouseDown(ByRef objComboBox As Access.ComboBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ComboBoxMouseMove(ByRef objComboBox As Access.ComboBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ComboBoxMouseUp(ByRef objComboBox As Access.ComboBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ComboBoxNotInList(ByRef objComboBox As Access.ComboBox, NewData As String, Response As Integer)
    Public Event ComboBoxUndo(ByRef objComboBox As Access.ComboBox, Cancel As Integer)
    
    Public Event CommandButtonClick(ByRef objCommandButton As Access.CommandButton)
    Public Event CommandButtonDblClick(ByRef objCommandButton As Access.CommandButton, Cancel As Integer)
    Public Event CommandButtonEnter(ByRef objCommandButton As Access.CommandButton)
    Public Event CommandButtonExit(ByRef objCommandButton As Access.CommandButton, Cancel As Integer)
    Public Event CommandButtonGotFocus(ByRef objCommandButton As Access.CommandButton)
    Public Event CommandButtonKeyDown(ByRef objCommandButton As Access.CommandButton, KeyCode As Integer, Shift As Integer)
    Public Event CommandButtonKeyPress(ByRef objCommandButton As Access.CommandButton, KeyAscii As Integer)
    Public Event CommandButtonKeyUp(ByRef objCommandButton As Access.CommandButton, KeyCode As Integer, Shift As Integer)
    Public Event CommandButtonLostFocus(ByRef objCommandButton As Access.CommandButton)
    Public Event CommandButtonMouseDown(ByRef objCommandButton As Access.CommandButton, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event CommandButtonMouseMove(ByRef objCommandButton As Access.CommandButton, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event CommandButtonMouseUp(ByRef objCommandButton As Access.CommandButton, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Public Event ImageClick(ByRef objImage As Access.Image)
    Public Event ImageDblClick(ByRef objImage As Access.Image, Cancel As Integer)
    Public Event ImageMouseDown(ByRef objImage As Access.Image, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ImageMouseMove(ByRef objImage As Access.Image, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ImageMouseUp(ByRef objImage As Access.Image, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Public Event LabelClick(ByRef objLabel As Access.Label)
    Public Event LabelDblClick(ByRef objLabel As Access.Label, Cancel As Integer)
    Public Event LabelMouseDown(ByRef objLabel As Access.Label, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event LabelMouseMove(ByRef objLabel As Access.Label, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event LabelMouseUp(ByRef objLabel As Access.Label, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Public Event ListBoxAfterUpdate(ByRef objListBox As Access.ListBox)
    Public Event ListBoxBeforeUpdate(ByRef objListBox As Access.ListBox, Cancel As Integer)
    Public Event ListBoxClick(ByRef objListBox As Access.ListBox)
    Public Event ListBoxDblClick(ByRef objListBox As Access.ListBox, Cancel As Integer)
    Public Event ListBoxEnter(ByRef objListBox As Access.ListBox)
    Public Event ListBoxExit(ByRef objListBox As Access.ListBox, Cancel As Integer)
    Public Event ListBoxGotFocus(ByRef objListBox As Access.ListBox)
    Public Event ListBoxKeyDown(ByRef objListBox As Access.ListBox, KeyCode As Integer, Shift As Integer)
    Public Event ListBoxKeyPress(ByRef objListBox As Access.ListBox, KeyAscii As Integer)
    Public Event ListBoxKeyUp(ByRef objListBox As Access.ListBox, KeyCode As Integer, Shift As Integer)
    Public Event ListBoxLostFocus(ByRef objListBox As Access.ListBox)
    Public Event ListBoxMouseDown(ByRef objListBox As Access.ListBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ListBoxMouseMove(ByRef objListBox As Access.ListBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ListBoxMouseUp(ByRef objListBox As Access.ListBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
 
    Public Event NavigationButtonClick(ByRef objNavigationButton As Access.NavigationButton)
    Public Event NavigationButtonDblClick(ByRef objNavigationButton As Access.NavigationButton, Cancel As Integer)
    Public Event NavigationButtonEnter(ByRef objNavigationButton As Access.NavigationButton)
    Public Event NavigationButtonExit(ByRef objNavigationButton As Access.NavigationButton, Cancel As Integer)
    Public Event NavigationButtonGotFocus(ByRef objNavigationButton As Access.NavigationButton)
    Public Event NavigationButtonKeyDown(ByRef objNavigationButton As Access.NavigationButton, KeyCode As Integer, Shift As Integer)
    Public Event NavigationButtonKeyPress(ByRef objNavigationButton As Access.NavigationButton, KeyAscii As Integer)
    Public Event NavigationButtonKeyUp(ByRef objNavigationButton As Access.NavigationButton, KeyCode As Integer, Shift As Integer)
    Public Event NavigationButtonLostFocus(ByRef objNavigationButton As Access.NavigationButton)
    Public Event NavigationButtonMouseDown(ByRef objNavigationButton As Access.NavigationButton, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event NavigationButtonMouseMove(ByRef objNavigationButton As Access.NavigationButton, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event NavigationButtonMouseUp(ByRef objNavigationButton As Access.NavigationButton, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Public Event NavigationControlAfterUpdate(ByRef objNavigationControl As Access.NavigationControl)
    Public Event NavigationControlBeforeUpdate(ByRef objNavigationControl As Access.NavigationControl, Cancel As Integer)
    Public Event NavigationControlChange(ByRef objNavigationControl As Access.NavigationControl)
    Public Event NavigationControlClick(ByRef objNavigationControl As Access.NavigationControl)
    Public Event NavigationControlDblClick(ByRef objNavigationControl As Access.NavigationControl, Cancel As Integer)
    Public Event NavigationControlDirty(ByRef objNavigationControl As Access.NavigationControl, Cancel As Integer)
    Public Event NavigationControlEnter(ByRef objNavigationControl As Access.NavigationControl)
    Public Event NavigationControlExit(ByRef objNavigationControl As Access.NavigationControl, Cancel As Integer)
    Public Event NavigationControlGotFocus(ByRef objNavigationControl As Access.NavigationControl)
    Public Event NavigationControlKeyDown(ByRef objNavigationControl As Access.NavigationControl, KeyCode As Integer, Shift As Integer)
    Public Event NavigationControlKeyPress(ByRef objNavigationControl As Access.NavigationControl, KeyAscii As Integer)
    Public Event NavigationControlKeyUp(ByRef objNavigationControl As Access.NavigationControl, KeyCode As Integer, Shift As Integer)
    Public Event NavigationControlLostFocus(ByRef objNavigationControl As Access.NavigationControl)
    Public Event NavigationControlMouseDown(ByRef objNavigationControl As Access.NavigationControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event NavigationControlMouseMove(ByRef objNavigationControl As Access.NavigationControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event NavigationControlMouseUp(ByRef objNavigationControl As Access.NavigationControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event NavigationControlUndo(ByRef objNavigationControl As Access.NavigationControl, Cancel As Integer)
    
    Public Event ObjectFrameClick(ByRef objObjectFrame As Access.ObjectFrame)
    Public Event ObjectFrameDblClick(ByRef objObjectFrame As Access.ObjectFrame, Cancel As Integer)
    Public Event ObjectFrameEnter(ByRef objObjectFrame As Access.ObjectFrame)
    Public Event ObjectFrameExit(ByRef objObjectFrame As Access.ObjectFrame, Cancel As Integer)
    Public Event ObjectFrameGotFocus(ByRef objObjectFrame As Access.ObjectFrame)
    Public Event ObjectFrameLostFocus(ByRef objObjectFrame As Access.ObjectFrame)
    Public Event ObjectFrameMouseDown(ByRef objObjectFrame As Access.ObjectFrame, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ObjectFrameMouseMove(ByRef objObjectFrame As Access.ObjectFrame, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ObjectFrameMouseUp(ByRef objObjectFrame As Access.ObjectFrame, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ObjectFrameUpdated(ByRef objObjectFrame As Access.ObjectFrame, Code As Integer)
    
    Public Event OptionButtonAfterUpdate(ByRef objOptionButton As Access.OptionButton)
    Public Event OptionButtonBeforeUpdate(ByRef objOptionButton As Access.OptionButton, Cancel As Integer)
    Public Event OptionButtonClick(ByRef objOptionButton As Access.OptionButton)
    Public Event OptionButtonDblClick(ByRef objOptionButton As Access.OptionButton, Cancel As Integer)
    Public Event OptionButtonEnter(ByRef objOptionButton As Access.OptionButton)
    Public Event OptionButtonExit(ByRef objOptionButton As Access.OptionButton, Cancel As Integer)
    Public Event OptionButtonGotFocus(ByRef objOptionButton As Access.OptionButton)
    Public Event OptionButtonKeyDown(ByRef objOptionButton As Access.OptionButton, KeyCode As Integer, Shift As Integer)
    Public Event OptionButtonKeyPress(ByRef objOptionButton As Access.OptionButton, KeyAscii As Integer)
    Public Event OptionButtonKeyUp(ByRef objOptionButton As Access.OptionButton, KeyCode As Integer, Shift As Integer)
    Public Event OptionButtonLostFocus(ByRef objOptionButton As Access.OptionButton)
    Public Event OptionButtonMouseDown(ByRef objOptionButton As Access.OptionButton, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event OptionButtonMouseMove(ByRef objOptionButton As Access.OptionButton, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event OptionButtonMouseUp(ByRef objOptionButton As Access.OptionButton, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Public Event OptionGroupAfterUpdate(ByRef objOptionGroup As Access.OptionGroup)
    Public Event OptionGroupBeforeUpdate(ByRef objOptionGroup As Access.OptionGroup, Cancel As Integer)
    Public Event OptionGroupClick(ByRef objOptionGroup As Access.OptionGroup)
    Public Event OptionGroupDblClick(ByRef objOptionGroup As Access.OptionGroup, Cancel As Integer)
    Public Event OptionGroupEnter(ByRef objOptionGroup As Access.OptionGroup)
    Public Event OptionGroupExit(ByRef objOptionGroup As Access.OptionGroup, Cancel As Integer)
    Public Event OptionGroupMouseDown(ByRef objOptionGroup As Access.OptionGroup, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event OptionGroupMouseMove(ByRef objOptionGroup As Access.OptionGroup, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event OptionGroupMouseUp(ByRef objOptionGroup As Access.OptionGroup, Button As Integer, Shift As Integer, X As Single, Y As Single)
   
    Public Event PageClick(ByRef objPage As Access.Page)
    Public Event PageDblClick(ByRef objPage As Access.Page, Cancel As Integer)
    Public Event PageMouseDown(ByRef objPage As Access.Page, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event PageMouseMove(ByRef objPage As Access.Page, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event PageMouseUp(ByRef objPage As Access.Page, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Public Event RectangleClick(ByRef objRectangle As Access.Rectangle)
    Public Event RectangleDblClick(ByRef objRectangle As Access.Rectangle, Cancel As Integer)
    Public Event RectangleMouseDown(ByRef objRectangle As Access.Rectangle, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event RectangleMouseMove(ByRef objRectangle As Access.Rectangle, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event RectangleMouseUp(ByRef objRectangle As Access.Rectangle, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Public Event SubFormEnter(ByRef objSubForm As Access.SubForm)
    Public Event SubFormExit(ByRef objSubForm As Access.SubForm, Cancel As Integer)
    
    Public Event TabControlChange(ByRef objTabControl As Access.TabControl)
    Public Event TabControlClick(ByRef objTabControl As Access.TabControl)
    Public Event TabControlDblClick(ByRef objTabControl As Access.TabControl, Cancel As Integer)
    Public Event TabControlKeyDown(ByRef objTabControl As Access.TabControl, KeyCode As Integer, Shift As Integer)
    Public Event TabControlKeyPress(ByRef objTabControl As Access.TabControl, KeyAscii As Integer)
    Public Event TabControlKeyUp(ByRef objTabControl As Access.TabControl, KeyCode As Integer, Shift As Integer)
    Public Event TabControlMouseDown(ByRef objTabControl As Access.TabControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event TabControlMouseMove(ByRef objTabControl As Access.TabControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event TabControlMouseUp(ByRef objTabControl As Access.TabControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Public Event TextBoxAfterUpdate(ByRef objTextBox As Access.TextBox)
    Public Event TextBoxBeforeUpdate(ByRef objTextBox As Access.TextBox, Cancel As Integer)
    Public Event TextBoxChange(ByRef objTextBox As Access.TextBox)
    Public Event TextBoxClick(ByRef objTextBox As Access.TextBox)
    Public Event TextBoxDblClick(ByRef objTextBox As Access.TextBox, Cancel As Integer)
    Public Event TextBoxDirty(ByRef objTextBox As Access.TextBox, Cancel As Integer)
    Public Event TextBoxEnter(ByRef objTextBox As Access.TextBox)
    Public Event TextBoxExit(ByRef objTextBox As Access.TextBox, Cancel As Integer)
    Public Event TextBoxGotFocus(ByRef objTextBox As Access.TextBox)
    Public Event TextBoxKeyDown(ByRef objTextBox As Access.TextBox, KeyCode As Integer, Shift As Integer)
    Public Event TextBoxKeyPress(ByRef objTextBox As Access.TextBox, KeyAscii As Integer)
    Public Event TextBoxKeyUp(ByRef objTextBox As Access.TextBox, KeyCode As Integer, Shift As Integer)
    Public Event TextBoxLostFocus(ByRef objTextBox As Access.TextBox)
    Public Event TextBoxMouseDown(ByRef objTextBox As Access.TextBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event TextBoxMouseMove(ByRef objTextBox As Access.TextBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event TextBoxMouseUp(ByRef objTextBox As Access.TextBox, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event TextBoxUndo(ByRef objTextBox As Access.TextBox, Cancel As Integer)
    
    Public Event ToggleButtonAfterUpdate(ByRef objToggleButton As Access.ToggleButton)
    Public Event ToggleButtonBeforeUpdate(ByRef objToggleButton As Access.ToggleButton, Cancel As Integer)
    Public Event ToggleButtonClick(ByRef objToggleButton As Access.ToggleButton)
    Public Event ToggleButtonDblClick(ByRef objToggleButton As Access.ToggleButton, Cancel As Integer)
    Public Event ToggleButtonEnter(ByRef objToggleButton As Access.ToggleButton)
    Public Event ToggleButtonExit(ByRef objToggleButton As Access.ToggleButton, Cancel As Integer)
    Public Event ToggleButtonGotFocus(ByRef objToggleButton As Access.ToggleButton)
    Public Event ToggleButtonKeyDown(ByRef objToggleButton As Access.ToggleButton, KeyCode As Integer, Shift As Integer)
    Public Event ToggleButtonKeyPress(ByRef objToggleButton As Access.ToggleButton, KeyAscii As Integer)
    Public Event ToggleButtonKeyUp(ByRef objToggleButton As Access.ToggleButton, KeyCode As Integer, Shift As Integer)
    Public Event ToggleButtonLostFocus(ByRef objToggleButton As Access.ToggleButton)
    Public Event ToggleButtonMouseDown(ByRef objToggleButton As Access.ToggleButton, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ToggleButtonMouseMove(ByRef objToggleButton As Access.ToggleButton, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event ToggleButtonMouseUp(ByRef objToggleButton As Access.ToggleButton, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Public Event WebBrowserAfterUpdate(ByRef objBrowser As Access.WebBrowserControl)
    Public Event WebBrowserBeforeUpdate(ByRef objBrowser As Access.WebBrowserControl, Cancel As Integer)
    Public Event WebBrowserBeforeNavigate2(ByRef objBrowser As Access.WebBrowserControl, ByVal pDisp As Object, ByRef URL As Variant, ByRef flags As Variant, ByRef TargetFrameName As Variant, ByRef PostData As Variant, ByRef Headers As Variant, ByRef Cancel As Boolean)
    Public Event WebBrowserChange(ByRef objBrowser As Access.WebBrowserControl)
    Public Event WebBrowserClick(ByRef objBrowser As Access.WebBrowserControl)
    Public Event WebBrowserDblClick(ByRef objBrowser As Access.WebBrowserControl, Cancel As Integer)
    Public Event WebBrowserDirty(ByRef objBrowser As Access.WebBrowserControl, Cancel As Integer)
    Public Event WebBrowserDocumentComplete(ByRef objBrowser As Access.WebBrowserControl, ByVal pDisp As Object, URL As Variant)
    Public Event WebBrowserEnter(ByRef objBrowser As Access.WebBrowserControl)
    Public Event WebBrowserExit(ByRef objBrowser As Access.WebBrowserControl, Cancel As Integer)
    Public Event WebBrowserGotFocus(ByRef objBrowser As Access.WebBrowserControl)
    Public Event WebBrowserKeyDown(ByRef objBrowser As Access.WebBrowserControl, KeyCode As Integer, Shift As Integer)
    Public Event WebBrowserKeyPress(ByRef objBrowser As Access.WebBrowserControl, KeyAscii As Integer)
    Public Event WebBrowserKeyUp(ByRef objBrowser As Access.WebBrowserControl, KeyCode As Integer, Shift As Integer)
    Public Event WebBrowserLostFocus(ByRef objBrowser As Access.WebBrowserControl)
    Public Event WebBrowserMouseDown(ByRef objBrowser As Access.WebBrowserControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event WebBrowserMouseMove(ByRef objBrowser As Access.WebBrowserControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event WebBrowserMouseUp(ByRef objBrowser As Access.WebBrowserControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Public Event WebBrowserNavigateError(ByRef objBrowser As Access.WebBrowserControl, ByVal pDisp As Object, URL As Variant, TargetFrameName As Variant, StatusCode As Variant, Cancel As Boolean)
    Public Event WebBrowserProgressChange(ByRef objBrowser As Access.WebBrowserControl, ByVal Progress As Long, ByVal ProgressMax As Long)
    Public Event WebBrowserUpdated(ByRef objBrowser As Access.WebBrowserControl, Code As Integer)

#End If

Public Function Create(Optional ByRef objControls As Access.Controls, Optional ByVal eType As eControlTypes = ctAll, Optional ByVal strTagFilter As String) As ControlArray

'    On Error GoTo errHandler

    Dim ctl As Access.Control

    With New ControlArray
        
        Set This.Col = New Collection
    
        If objControls Is Nothing = False Then

            For Each ctl In objControls

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

Friend Sub RaiseAfterUpdate(ByRef objControl As IControl)
    
    RaiseEvent onAfterUpdate(objControl)

    #If IncludeSpecificEvents = True Then
    
        Select Case objControl.ControlType
            Case ctAttachment: RaiseEvent AttachmentAfterUpdate(objControl.Object)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameAfterUpdate(objControl.Object)
            Case ctCheckBox: RaiseEvent CheckboxAfterUpdate(objControl.Object)
            Case ctComboBox: RaiseEvent ComboBoxAfterUpdate(objControl.Object)
            Case ctListBox: RaiseEvent ListBoxAfterUpdate(objControl.Object)
            Case ctNavigationControl: RaiseEvent NavigationControlAfterUpdate(objControl.Object)
            Case ctOptionButton: RaiseEvent OptionButtonAfterUpdate(objControl.Object)
            Case ctOptionGroup: RaiseEvent OptionGroupAfterUpdate(objControl.Object)
            Case ctTextBox: RaiseEvent TextBoxAfterUpdate(objControl.Object)
            Case ctToggleButton: RaiseEvent ToggleButtonAfterUpdate(objControl.Object)
            Case ctWebBrowserControl: RaiseEvent WebBrowserAfterUpdate(objControl.Object)
        End Select
        
    #End If
    
End Sub

Friend Sub RaiseAttachmentCurrent(ByRef objControl As IControl)

    RaiseEvent onAttachmentCurrent(objControl)
    
    #If IncludeSpecificEvents = True Then
        
        Select Case objControl.ControlType
            Case ctAttachment: RaiseEvent AttachmentCurrent(objControl.Object)
        End Select
    
    #End If
    
End Sub

Friend Sub RaiseBeforeNavigate2(ByRef objControl As IControl, ByVal pDisp As Object, ByRef URL As Variant, ByRef flags As Variant, ByRef TargetFrameName As Variant, ByRef PostData As Variant, ByRef Headers As Variant, ByRef Cancel As Boolean)
    
    RaiseEvent onBeforeNavigate2(objControl, pDisp, URL, flags, TargetFrameName, PostData, Headers, Cancel)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case objControl.ControlType
            Case ctWebBrowserControl: RaiseEvent WebBrowserBeforeNavigate2(objControl.Object, pDisp, URL, flags, TargetFrameName, PostData, Headers, Cancel)
        End Select
        
    #End If
    
End Sub

Friend Sub RaiseBeforeUpdate(ByRef objControl As IControl, Cancel As Integer)
    
    RaiseEvent onBeforeUpdate(objControl, Cancel)
        
    #If IncludeSpecificEvents = True Then
        
        Select Case objControl.ControlType
            Case ctAttachment: RaiseEvent AttachmentBeforeUpdate(objControl.Object, Cancel)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameBeforeUpdate(objControl.Object, Cancel)
            Case ctCheckBox: RaiseEvent CheckboxBeforeUpdate(objControl.Object, Cancel)
            Case ctComboBox: RaiseEvent ComboBoxBeforeUpdate(objControl.Object, Cancel)
            Case ctListBox: RaiseEvent ListBoxBeforeUpdate(objControl.Object, Cancel)
            Case ctNavigationControl: RaiseEvent NavigationControlBeforeUpdate(objControl.Object, Cancel)
            Case ctOptionButton: RaiseEvent OptionButtonBeforeUpdate(objControl.Object, Cancel)
            Case ctOptionGroup: RaiseEvent OptionGroupBeforeUpdate(objControl.Object, Cancel)
            Case ctTextBox: RaiseEvent TextBoxBeforeUpdate(objControl.Object, Cancel)
            Case ctToggleButton: RaiseEvent ToggleButtonBeforeUpdate(objControl.Object, Cancel)
            Case ctWebBrowserControl: RaiseEvent WebBrowserBeforeUpdate(objControl.Object, Cancel)
        End Select
        
    #End If
    
End Sub

Friend Sub RaiseChange(ByRef objControl As IControl)
    
    RaiseEvent onChange(objControl)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case objControl.ControlType
            Case ctAttachment: RaiseEvent AttachmentChange(objControl.Object)
            Case ctComboBox: RaiseEvent ComboBoxChange(objControl.Object)
            Case ctNavigationControl: RaiseEvent NavigationControlChange(objControl.Object)
            Case ctTabCtl: RaiseEvent TabControlChange(objControl.Object)
            Case ctTextBox: RaiseEvent TextBoxChange(objControl.Object)
            Case ctWebBrowserControl: RaiseEvent WebBrowserChange(objControl.Object)
        End Select
    
    #End If
    
End Sub

Friend Sub RaiseClick(ByRef objControl As IControl)
    
    RaiseEvent onClick(objControl)

    #If IncludeSpecificEvents = True Then
    
        Select Case objControl.ControlType
            Case ctAttachment: RaiseEvent AttachmentClick(objControl.Object)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameClick(objControl.Object)
            Case ctChart: RaiseEvent ChartClick(objControl.Object)
            Case ctCheckBox: RaiseEvent CheckboxClick(objControl.Object)
            Case ctComboBox: RaiseEvent ComboBoxClick(objControl.Object)
            Case ctCommandButton: RaiseEvent CommandButtonClick(objControl.Object)
            Case ctImage: RaiseEvent ImageClick(objControl.Object)
            Case ctLabel: RaiseEvent LabelClick(objControl.Object)
            Case ctListBox: RaiseEvent ListBoxClick(objControl.Object)
            Case ctNavigationButton: RaiseEvent NavigationButtonClick(objControl.Object)
            Case ctNavigationControl: RaiseEvent NavigationControlClick(objControl.Object)
            Case ctObjectFrame: RaiseEvent ObjectFrameClick(objControl.Object)
            Case ctOptionButton: RaiseEvent OptionButtonClick(objControl.Object)
            Case ctOptionGroup: RaiseEvent OptionGroupClick(objControl.Object)
            Case ctPage: RaiseEvent PageClick(objControl.Object)
            Case ctRectangle: RaiseEvent RectangleClick(objControl.Object)
            Case ctTabCtl: RaiseEvent TabControlClick(objControl.Object)
            Case ctTextBox: RaiseEvent TextBoxClick(objControl.Object)
            Case ctToggleButton: RaiseEvent ToggleButtonClick(objControl.Object)
            Case ctWebBrowserControl: RaiseEvent WebBrowserClick(objControl.Object)
        End Select
        
    #End If
        
End Sub

Friend Sub RaiseDblClick(ByRef objControl As IControl, Cancel As Integer)
    
    RaiseEvent onDblClick(objControl, Cancel)
    
    #If IncludeSpecificEvents = True Then
        
        Select Case objControl.ControlType
            Case ctAttachment: RaiseEvent AttachmentDblClick(objControl.Object, Cancel)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameDblClick(objControl.Object, Cancel)
            Case ctChart: RaiseEvent ChartDblClick(objControl.Object, Cancel)
            Case ctCheckBox: RaiseEvent CheckboxDblClick(objControl.Object, Cancel)
            Case ctComboBox: RaiseEvent ComboBoxDblClick(objControl.Object, Cancel)
            Case ctCommandButton: RaiseEvent CommandButtonDblClick(objControl.Object, Cancel)
            Case ctImage: RaiseEvent ImageDblClick(objControl.Object, Cancel)
            Case ctLabel: RaiseEvent LabelDblClick(objControl.Object, Cancel)
            Case ctListBox: RaiseEvent ListBoxDblClick(objControl.Object, Cancel)
            Case ctNavigationButton: RaiseEvent NavigationButtonDblClick(objControl.Object, Cancel)
            Case ctNavigationControl: RaiseEvent NavigationControlDblClick(objControl.Object, Cancel)
            Case ctObjectFrame: RaiseEvent ObjectFrameDblClick(objControl.Object, Cancel)
            Case ctOptionButton: RaiseEvent OptionButtonDblClick(objControl.Object, Cancel)
            Case ctOptionGroup: RaiseEvent OptionGroupDblClick(objControl.Object, Cancel)
            Case ctPage: RaiseEvent PageDblClick(objControl.Object, Cancel)
            Case ctRectangle: RaiseEvent RectangleDblClick(objControl.Object, Cancel)
            Case ctTabCtl: RaiseEvent TabControlDblClick(objControl.Object, Cancel)
            Case ctTextBox: RaiseEvent TextBoxDblClick(objControl.Object, Cancel)
            Case ctToggleButton: RaiseEvent ToggleButtonDblClick(objControl.Object, Cancel)
            Case ctWebBrowserControl: RaiseEvent WebBrowserDblClick(objControl.Object, Cancel)
        End Select
            
    #End If
    
End Sub

Friend Sub RaiseDirty(ByRef objControl As IControl, Cancel As Integer)
    
    RaiseEvent onDirty(objControl, Cancel)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case objControl.ControlType
            Case ctAttachment: RaiseEvent AttachmentDirty(objControl.Object, Cancel)
            Case ctComboBox: RaiseEvent ComboBoxDirty(objControl.Object, Cancel)
            Case ctNavigationControl: RaiseEvent NavigationControlDirty(objControl.Object, Cancel)
            Case ctTextBox: RaiseEvent TextBoxDirty(objControl.Object, Cancel)
            Case ctWebBrowserControl: RaiseEvent WebBrowserDirty(objControl.Object, Cancel)
        End Select
            
    #End If
    
End Sub

Friend Sub RaiseDocumentComplete(ByRef objControl As IControl, ByVal pDisp As Object, URL As Variant)
    
    RaiseEvent onDocumentComplete(objControl, pDisp, URL)

    #If IncludeSpecificEvents = True Then

        Select Case objControl.ControlType
            Case ctWebBrowserControl: RaiseEvent WebBrowserDocumentComplete(objControl.Object, pDisp, URL)
        End Select
        
    #End If
    
End Sub

Friend Sub RaiseEnter(ByRef objControl As IControl)
    
    RaiseEvent onEnter(objControl)
    
    #If IncludeSpecificEvents = True Then
        
        Select Case objControl.ControlType
            Case ctAttachment: RaiseEvent AttachmentEnter(objControl.Object)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameEnter(objControl.Object)
            Case ctChart: RaiseEvent ChartEnter(objControl.Object)
            Case ctCheckBox: RaiseEvent CheckboxEnter(objControl.Object)
            Case ctComboBox: RaiseEvent ComboBoxEnter(objControl.Object)
            Case ctCommandButton: RaiseEvent CommandButtonEnter(objControl.Object)
            Case ctListBox: RaiseEvent ListBoxEnter(objControl.Object)
            Case ctNavigationButton: RaiseEvent NavigationButtonEnter(objControl.Object)
            Case ctNavigationControl: RaiseEvent NavigationControlEnter(objControl.Object)
            Case ctObjectFrame: RaiseEvent ObjectFrameEnter(objControl.Object)
            Case ctOptionButton: RaiseEvent OptionButtonEnter(objControl.Object)
            Case ctOptionGroup: RaiseEvent OptionGroupEnter(objControl.Object)
            Case ctSubForm: RaiseEvent SubFormEnter(objControl.Object)
            Case ctTextBox: RaiseEvent TextBoxEnter(objControl.Object)
            Case ctToggleButton: RaiseEvent ToggleButtonEnter(objControl.Object)
            Case ctWebBrowserControl: RaiseEvent WebBrowserEnter(objControl.Object)
        End Select
    
    #End If
    
End Sub

Friend Sub RaiseExit(ByRef objControl As IControl, Cancel As Integer)
    
    RaiseEvent onExit(objControl, Cancel)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case objControl.ControlType
            Case ctAttachment: RaiseEvent AttachmentExit(objControl.Object, Cancel)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameExit(objControl.Object, Cancel)
            Case ctChart: RaiseEvent ChartExit(objControl.Object, Cancel)
            Case ctCheckBox: RaiseEvent CheckboxExit(objControl.Object, Cancel)
            Case ctComboBox: RaiseEvent ComboBoxExit(objControl.Object, Cancel)
            Case ctCommandButton: RaiseEvent CommandButtonExit(objControl.Object, Cancel)
            Case ctListBox: RaiseEvent ListBoxExit(objControl.Object, Cancel)
            Case ctNavigationButton: RaiseEvent NavigationButtonExit(objControl.Object, Cancel)
            Case ctNavigationControl: RaiseEvent NavigationControlExit(objControl.Object, Cancel)
            Case ctObjectFrame: RaiseEvent ObjectFrameExit(objControl.Object, Cancel)
            Case ctOptionButton: RaiseEvent OptionButtonExit(objControl.Object, Cancel)
            Case ctOptionGroup: RaiseEvent OptionGroupExit(objControl.Object, Cancel)
            Case ctSubForm: RaiseEvent SubFormExit(objControl.Object, Cancel)
            Case ctTextBox: RaiseEvent TextBoxExit(objControl.Object, Cancel)
            Case ctToggleButton: RaiseEvent ToggleButtonExit(objControl.Object, Cancel)
            Case ctWebBrowserControl: RaiseEvent WebBrowserExit(objControl.Object, Cancel)
        End Select
        
    #End If
    
End Sub

Friend Sub RaiseGotFocus(ByRef objControl As IControl)
    
    RaiseEvent onGotFocus(objControl)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case objControl.ControlType
            Case ctAttachment: RaiseEvent AttachmentGotFocus(objControl.Object)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameGotFocus(objControl.Object)
            Case ctChart: RaiseEvent ChartGotFocus(objControl.Object)
            Case ctCheckBox: RaiseEvent CheckboxGotFocus(objControl.Object)
            Case ctComboBox: RaiseEvent ComboBoxGotFocus(objControl.Object)
            Case ctCommandButton: RaiseEvent CommandButtonGotFocus(objControl.Object)
            Case ctListBox: RaiseEvent ListBoxGotFocus(objControl.Object)
            Case ctNavigationButton: RaiseEvent NavigationButtonGotFocus(objControl.Object)
            Case ctNavigationControl: RaiseEvent NavigationControlGotFocus(objControl.Object)
            Case ctObjectFrame: RaiseEvent ObjectFrameGotFocus(objControl.Object)
            Case ctOptionButton: RaiseEvent OptionButtonGotFocus(objControl.Object)
            Case ctTextBox: RaiseEvent TextBoxGotFocus(objControl.Object)
            Case ctToggleButton: RaiseEvent ToggleButtonGotFocus(objControl.Object)
            Case ctWebBrowserControl: RaiseEvent WebBrowserGotFocus(objControl.Object)
        End Select
        
    #End If
    
End Sub

Friend Sub RaiseKeyDown(ByRef objControl As IControl, KeyCode As Integer, Shift As Integer)
    
    RaiseEvent onKeyDown(objControl, KeyCode, Shift)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case objControl.ControlType
            Case ctAttachment: RaiseEvent AttachmentKeyDown(objControl.Object, KeyCode, Shift)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameKeyDown(objControl.Object, KeyCode, Shift)
            Case ctCheckBox: RaiseEvent CheckboxKeyDown(objControl.Object, KeyCode, Shift)
            Case ctComboBox: RaiseEvent ComboBoxKeyDown(objControl.Object, KeyCode, Shift)
            Case ctCommandButton: RaiseEvent CommandButtonKeyDown(objControl.Object, KeyCode, Shift)
            Case ctListBox: RaiseEvent ListBoxKeyDown(objControl.Object, KeyCode, Shift)
            Case ctNavigationButton: RaiseEvent NavigationButtonKeyDown(objControl.Object, KeyCode, Shift)
            Case ctNavigationControl: RaiseEvent NavigationControlKeyDown(objControl.Object, KeyCode, Shift)
            Case ctOptionButton: RaiseEvent OptionButtonKeyDown(objControl.Object, KeyCode, Shift)
            Case ctTabCtl: RaiseEvent TabControlKeyDown(objControl.Object, KeyCode, Shift)
            Case ctTextBox: RaiseEvent TextBoxKeyDown(objControl.Object, KeyCode, Shift)
            Case ctToggleButton: RaiseEvent ToggleButtonKeyDown(objControl.Object, KeyCode, Shift)
            Case ctWebBrowserControl: RaiseEvent WebBrowserKeyDown(objControl.Object, KeyCode, Shift)
        End Select
        
    #End If
    
End Sub

Friend Sub RaiseKeyPress(ByRef objControl As IControl, KeyAscii As Integer)
    
    RaiseEvent onKeyPress(objControl, KeyAscii)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case objControl.ControlType
            Case ctAttachment: RaiseEvent AttachmentKeyPress(objControl.Object, KeyAscii)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameKeyPress(objControl.Object, KeyAscii)
            Case ctCheckBox: RaiseEvent CheckboxKeyPress(objControl.Object, KeyAscii)
            Case ctComboBox: RaiseEvent ComboBoxKeyPress(objControl.Object, KeyAscii)
            Case ctCommandButton: RaiseEvent CommandButtonKeyPress(objControl.Object, KeyAscii)
            Case ctListBox: RaiseEvent ListBoxKeyPress(objControl.Object, KeyAscii)
            Case ctNavigationButton: RaiseEvent NavigationButtonKeyPress(objControl.Object, KeyAscii)
            Case ctNavigationControl: RaiseEvent NavigationControlKeyPress(objControl.Object, KeyAscii)
            Case ctOptionButton: RaiseEvent OptionButtonKeyPress(objControl.Object, KeyAscii)
            Case ctTabCtl: RaiseEvent TabControlKeyPress(objControl.Object, KeyAscii)
            Case ctTextBox: RaiseEvent TextBoxKeyPress(objControl.Object, KeyAscii)
            Case ctToggleButton: RaiseEvent ToggleButtonKeyPress(objControl.Object, KeyAscii)
            Case ctWebBrowserControl: RaiseEvent WebBrowserKeyPress(objControl.Object, KeyAscii)
        End Select
        
    #End If
    
End Sub

Friend Sub RaiseKeyUp(ByRef objControl As IControl, KeyCode As Integer, Shift As Integer)
    
    RaiseEvent onKeyUp(objControl, KeyCode, Shift)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case objControl.ControlType
            Case ctAttachment: RaiseEvent AttachmentKeyUp(objControl.Object, KeyCode, Shift)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameKeyUp(objControl.Object, KeyCode, Shift)
            Case ctCheckBox: RaiseEvent CheckboxKeyUp(objControl.Object, KeyCode, Shift)
            Case ctComboBox: RaiseEvent ComboBoxKeyUp(objControl.Object, KeyCode, Shift)
            Case ctCommandButton: RaiseEvent CommandButtonKeyUp(objControl.Object, KeyCode, Shift)
            Case ctListBox: RaiseEvent ListBoxKeyUp(objControl.Object, KeyCode, Shift)
            Case ctNavigationButton: RaiseEvent NavigationButtonKeyUp(objControl.Object, KeyCode, Shift)
            Case ctNavigationControl: RaiseEvent NavigationControlKeyUp(objControl.Object, KeyCode, Shift)
            Case ctOptionButton: RaiseEvent OptionButtonKeyUp(objControl.Object, KeyCode, Shift)
            Case ctTabCtl: RaiseEvent TabControlKeyUp(objControl.Object, KeyCode, Shift)
            Case ctTextBox: RaiseEvent TextBoxKeyUp(objControl.Object, KeyCode, Shift)
            Case ctToggleButton: RaiseEvent ToggleButtonKeyUp(objControl.Object, KeyCode, Shift)
            Case ctWebBrowserControl: RaiseEvent WebBrowserKeyUp(objControl.Object, KeyCode, Shift)
        End Select
    
    #End If
    
End Sub

Friend Sub RaiseLostFocus(ByRef objControl As IControl)
    
    RaiseEvent onLostFocus(objControl)
    
    #If IncludeSpecificEvents = True Then
        
        Select Case objControl.ControlType
            Case ctAttachment: RaiseEvent AttachmentLostFocus(objControl.Object)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameLostFocus(objControl.Object)
            Case ctChart: RaiseEvent ChartLostFocus(objControl.Object)
            Case ctCheckBox: RaiseEvent CheckboxLostFocus(objControl.Object)
            Case ctComboBox: RaiseEvent ComboBoxLostFocus(objControl.Object)
            Case ctCommandButton: RaiseEvent CommandButtonLostFocus(objControl.Object)
            Case ctListBox: RaiseEvent ListBoxLostFocus(objControl.Object)
            Case ctNavigationButton: RaiseEvent NavigationButtonLostFocus(objControl.Object)
            Case ctNavigationControl: RaiseEvent NavigationControlLostFocus(objControl.Object)
            Case ctObjectFrame: RaiseEvent ObjectFrameLostFocus(objControl.Object)
            Case ctOptionButton: RaiseEvent OptionButtonLostFocus(objControl.Object)
            Case ctTextBox: RaiseEvent TextBoxLostFocus(objControl.Object)
            Case ctToggleButton: RaiseEvent ToggleButtonLostFocus(objControl.Object)
            Case ctWebBrowserControl: RaiseEvent WebBrowserLostFocus(objControl.Object)
        End Select
    
    #End If
    
End Sub

Friend Sub RaiseMouseDown(ByRef objControl As IControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    RaiseEvent onMouseDown(objControl, Button, Shift, X, Y)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case objControl.ControlType
            Case ctAttachment: RaiseEvent AttachmentMouseDown(objControl.Object, Button, Shift, X, Y)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameMouseDown(objControl.Object, Button, Shift, X, Y)
            Case ctChart: RaiseEvent ChartMouseDown(objControl.Object, Button, Shift, X, Y)
            Case ctCheckBox: RaiseEvent CheckboxMouseDown(objControl.Object, Button, Shift, X, Y)
            Case ctComboBox: RaiseEvent ComboBoxMouseDown(objControl.Object, Button, Shift, X, Y)
            Case ctCommandButton: RaiseEvent CommandButtonMouseDown(objControl.Object, Button, Shift, X, Y)
            Case ctImage: RaiseEvent ImageMouseDown(objControl.Object, Button, Shift, X, Y)
            Case ctLabel: RaiseEvent LabelMouseDown(objControl.Object, Button, Shift, X, Y)
            Case ctListBox: RaiseEvent ListBoxMouseDown(objControl.Object, Button, Shift, X, Y)
            Case ctNavigationButton: RaiseEvent NavigationButtonMouseDown(objControl.Object, Button, Shift, X, Y)
            Case ctNavigationControl: RaiseEvent NavigationControlMouseDown(objControl.Object, Button, Shift, X, Y)
            Case ctObjectFrame: RaiseEvent ObjectFrameMouseDown(objControl.Object, Button, Shift, X, Y)
            Case ctOptionButton: RaiseEvent OptionButtonMouseDown(objControl.Object, Button, Shift, X, Y)
            Case ctOptionGroup: RaiseEvent OptionGroupMouseDown(objControl.Object, Button, Shift, X, Y)
            Case ctPage: RaiseEvent PageMouseDown(objControl.Object, Button, Shift, X, Y)
            Case ctRectangle: RaiseEvent RectangleMouseDown(objControl.Object, Button, Shift, X, Y)
            Case ctTabCtl: RaiseEvent TabControlMouseDown(objControl.Object, Button, Shift, X, Y)
            Case ctTextBox: RaiseEvent TextBoxMouseDown(objControl.Object, Button, Shift, X, Y)
            Case ctToggleButton: RaiseEvent ToggleButtonMouseDown(objControl.Object, Button, Shift, X, Y)
            Case ctWebBrowserControl: RaiseEvent WebBrowserMouseDown(objControl.Object, Button, Shift, X, Y)
        End Select
    
    #End If
    
End Sub

Friend Sub RaiseMouseMove(ByRef objControl As IControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    RaiseEvent onMouseMove(objControl, Button, Shift, X, Y)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case objControl.ControlType
            Case ctAttachment: RaiseEvent AttachmentMouseMove(objControl.Object, Button, Shift, X, Y)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameMouseMove(objControl.Object, Button, Shift, X, Y)
            Case ctChart: RaiseEvent ChartMouseMove(objControl.Object, Button, Shift, X, Y)
            Case ctCheckBox: RaiseEvent CheckboxMouseMove(objControl.Object, Button, Shift, X, Y)
            Case ctComboBox: RaiseEvent ComboBoxMouseMove(objControl.Object, Button, Shift, X, Y)
            Case ctCommandButton: RaiseEvent CommandButtonMouseMove(objControl.Object, Button, Shift, X, Y)
            Case ctImage: RaiseEvent ImageMouseMove(objControl.Object, Button, Shift, X, Y)
            Case ctLabel: RaiseEvent LabelMouseMove(objControl.Object, Button, Shift, X, Y)
            Case ctListBox: RaiseEvent ListBoxMouseMove(objControl.Object, Button, Shift, X, Y)
            Case ctNavigationButton: RaiseEvent NavigationButtonMouseMove(objControl.Object, Button, Shift, X, Y)
            Case ctNavigationControl: RaiseEvent NavigationControlMouseMove(objControl.Object, Button, Shift, X, Y)
            Case ctObjectFrame: RaiseEvent ObjectFrameMouseMove(objControl.Object, Button, Shift, X, Y)
            Case ctOptionButton: RaiseEvent OptionButtonMouseMove(objControl.Object, Button, Shift, X, Y)
            Case ctOptionGroup: RaiseEvent OptionGroupMouseMove(objControl.Object, Button, Shift, X, Y)
            Case ctPage: RaiseEvent PageMouseMove(objControl.Object, Button, Shift, X, Y)
            Case ctRectangle: RaiseEvent RectangleMouseMove(objControl.Object, Button, Shift, X, Y)
            Case ctTabCtl: RaiseEvent TabControlMouseMove(objControl.Object, Button, Shift, X, Y)
            Case ctTextBox: RaiseEvent TextBoxMouseMove(objControl.Object, Button, Shift, X, Y)
            Case ctToggleButton: RaiseEvent ToggleButtonMouseMove(objControl.Object, Button, Shift, X, Y)
            Case ctWebBrowserControl: RaiseEvent WebBrowserMouseMove(objControl.Object, Button, Shift, X, Y)
        End Select
    
    #End If
    
End Sub

Friend Sub RaiseMouseUp(ByRef objControl As IControl, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    RaiseEvent onMouseUp(objControl, Button, Shift, X, Y)
    
    #If IncludeSpecificEvents = True Then
        
        Select Case objControl.ControlType
            Case ctAttachment: RaiseEvent AttachmentMouseUp(objControl.Object, Button, Shift, X, Y)
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameMouseUp(objControl.Object, Button, Shift, X, Y)
            Case ctChart: RaiseEvent ChartMouseUp(objControl.Object, Button, Shift, X, Y)
            Case ctCheckBox: RaiseEvent CheckboxMouseUp(objControl.Object, Button, Shift, X, Y)
            Case ctComboBox: RaiseEvent ComboBoxMouseUp(objControl.Object, Button, Shift, X, Y)
            Case ctCommandButton: RaiseEvent CommandButtonMouseUp(objControl.Object, Button, Shift, X, Y)
            Case ctImage: RaiseEvent ImageMouseUp(objControl.Object, Button, Shift, X, Y)
            Case ctLabel: RaiseEvent LabelMouseUp(objControl.Object, Button, Shift, X, Y)
            Case ctListBox: RaiseEvent ListBoxMouseUp(objControl.Object, Button, Shift, X, Y)
            Case ctNavigationButton: RaiseEvent NavigationButtonMouseUp(objControl.Object, Button, Shift, X, Y)
            Case ctNavigationControl: RaiseEvent NavigationControlMouseUp(objControl.Object, Button, Shift, X, Y)
            Case ctObjectFrame: RaiseEvent ObjectFrameMouseUp(objControl.Object, Button, Shift, X, Y)
            Case ctOptionButton: RaiseEvent OptionButtonMouseUp(objControl.Object, Button, Shift, X, Y)
            Case ctOptionGroup: RaiseEvent OptionGroupMouseUp(objControl.Object, Button, Shift, X, Y)
            Case ctPage: RaiseEvent PageMouseUp(objControl.Object, Button, Shift, X, Y)
            Case ctRectangle: RaiseEvent RectangleMouseUp(objControl.Object, Button, Shift, X, Y)
            Case ctTabCtl: RaiseEvent TabControlMouseUp(objControl.Object, Button, Shift, X, Y)
            Case ctTextBox: RaiseEvent TextBoxMouseUp(objControl.Object, Button, Shift, X, Y)
            Case ctToggleButton: RaiseEvent ToggleButtonMouseUp(objControl.Object, Button, Shift, X, Y)
            Case ctWebBrowserControl: RaiseEvent WebBrowserMouseUp(objControl.Object, Button, Shift, X, Y)
        End Select
        
    #End If
    
End Sub

Friend Sub RaiseNavigateError(ByRef objControl As IControl, ByVal pDisp As Object, URL As Variant, TargetFrameName As Variant, StatusCode As Variant, Cancel As Boolean)
    
    RaiseEvent onNavigateError(objControl, pDisp, URL, TargetFrameName, StatusCode, Cancel)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case objControl.ControlType
            Case ctWebBrowserControl: RaiseEvent WebBrowserNavigateError(objControl.Object, pDisp, URL, TargetFrameName, StatusCode, Cancel)
        End Select
    
    #End If
    
End Sub

Friend Sub RaiseNotInList(ByRef objControl As IControl, NewData As String, Response As Integer)
    
    RaiseEvent onNotInList(objControl, NewData, Response)

    #If IncludeSpecificEvents = True Then
    
        Select Case objControl.ControlType
            Case ctComboBox: RaiseEvent ComboBoxNotInList(objControl.Object, NewData, Response)
        End Select
        
    #End If
    
End Sub

Friend Sub RaiseProgressChange(ByRef objControl As IControl, ByVal Progress As Long, ByVal ProgressMax As Long)
    
    RaiseEvent onProgressChange(objControl, Progress, ProgressMax)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case objControl.ControlType
            Case ctWebBrowserControl: RaiseEvent WebBrowserProgressChange(objControl.Object, Progress, ProgressMax)
        End Select

    #End If
    
End Sub

Friend Sub RaiseUndo(ByRef objControl As IControl, Cancel As Integer)
    
    RaiseEvent onUndo(objControl, Cancel)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case objControl.ControlType
            Case ctComboBox: RaiseEvent ComboBoxUndo(objControl.Object, Cancel)
            Case ctNavigationControl: RaiseEvent NavigationControlUndo(objControl.Object, Cancel)
            Case ctTextBox: RaiseEvent TextBoxUndo(objControl.Object, Cancel)
        End Select
    
    #End If
    
End Sub

Friend Sub RaiseUpdated(ByRef objControl As IControl, Code As Integer)
    
    RaiseEvent onUpdated(objControl, Code)
    
    #If IncludeSpecificEvents = True Then
    
        Select Case objControl.ControlType
            Case ctBoundObjectFrame: RaiseEvent BoundObjectFrameUpdated(objControl.Object, Code)
            Case ctChart: RaiseEvent ChartUpdated(objControl.Object, Code)
            Case ctObjectFrame: RaiseEvent ObjectFrameUpdated(objControl.Object, Code)
            Case ctWebBrowserControl: RaiseEvent WebBrowserUpdated(objControl.Object, Code)
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
