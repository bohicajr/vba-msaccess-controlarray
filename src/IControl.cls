VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
''
' VBA-ControlArray v1.0.0
' (c) J Martinez - https://github.com/bohicajr/vba-msaccess-controlarray
'
'
''----------------------------------------------------------------------------''
Option Compare Binary
Option Explicit

Public Property Get Name() As String: End Property
Public Property Get ControlType() As AcControlType: End Property
Public Property Get Object() As Access.Control: End Property
