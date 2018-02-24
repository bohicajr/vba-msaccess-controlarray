# VBA Control Array For Microsoft Access

Add the power of control arrays to VBA and simplify event handling.  

Here is a list of the controls that can be added to the control array collection.
```
    Attachment
    BoundObjectFrame
    Chart
    CheckBox
    ComboBox
    CommandButton
    Image
    Label
    ListBox
    NavigationButton
    NavigationControl
    ObjectFrame
    OptionButton
    OptionGroup
    Page
    Rectangle
    SubForm
    TabCtl
    TextBox
    ToggleButton
    WebBrowserControl
```
### Prerequisites

Microsoft Access.  Tested to work on Microsoft Access 2010-2016. This version will not work on any other office product, see [vba-controlarray](https//github.com/bohicajr/vba-controlarray) for code compatible with the rest of Microsoft Office.

### Manual Install

Download all the .cls files from github and import the files into your VBA project.
**Copy and paste will not work!**

Here is a list of the class files you must import into the VBA IDE.

```
CAAttachment.cls
CABoundObjectFrame.cls
CAChart.cls
CACheckBox.cls
CAComboBox.cls
CACommandButton.cls
CACustomControl.cls
CAImage.cls
CALabel.cls
CAListBox.cls
CANavigationButton.cls
CANavigationControl.cls
CAObjectFrame.cls
CAOptionButton.cls
CAOptionGroup.cls
CAPage.cls
CARectangle.cls
CASubForm.cls
CATabControl.cls
CATextBox.cls
CAToggleButton.cls
CAWebBrowserControl.cls
ControlArray.cls
IControl.cls
```

### Basic Example
Add a form to your Access database then and add as many command buttons as you want, then add the following code behind the form.

```VBA
Option Explicit

Private WithEvents Buttons as ControlArray

Private Sub UserForm_Initialize()
    
    Set Buttons = ControlArray.Create(Me.Controls)

End Sub

Private Sub Buttons_onClick(ctl As IControl)
    
    MsgBox "You clicked the button " & ctl.Name
    
End Sub
```

note that after you add the line that has WithEvents, that in the editor you now can choose the Buttons object, and select from it's many events to handle.  Every event also sends back a reference to the object that raised it!

### Inspiration
[Rubberduck VBA: Factories](https://rubberduckvba.wordpress.com/2016/07/05/oop-vba-pt-2-factories-and-cheap-hotels/) influenced the style that I used in writing this code.  Strongly recommend reading this article.
