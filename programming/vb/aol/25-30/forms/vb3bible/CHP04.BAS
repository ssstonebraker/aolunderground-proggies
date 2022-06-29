Global Document() As New formChild        'MDI Child forms
Global DocState() As Integer              'Tracks if child form is modified

Global Const CASCADE = 0                  'These arrange child windows
Global Const TILE_HORIZONTAL = 1
Global Const TILE_VERTICAL = 2
Global Const ARRANGE_ICONS = 3

Global Const MODAL = 1                    'Used to display dialog boxes
Global Const MODELESS = 0

Global Const NORMAL = 0                   'These are for windowstates
Global Const MINIMIZE = 1
Global Const MAXIMIZE = 2

Global Const RIGHT_BUTTON = 2             'Mouse button

Sub DisplayModal ()
     formModal.Show MODAL
     MsgBox "This line immediately follows the Show method"
End Sub

Sub DisplayModeless ()
     formModeless.Show MODELESS
     MsgBox "This line immediately follows the Show method"
End Sub

Sub DisplayNewChild ()
     Dim docCount As Integer
     docCount = UBound(Document) 'count # of open documents
     'make space in the tracking array and create a new document
     ReDim Preserve Document(docCount + 1)  'state of the new document
     ReDim Preserve DocState(docCount + 1) 'make the new child document
     Document(docCount + 1).Tag = docCount + 1
     Document(docCount + 1).Show
End Sub

Sub Quit ()
     Unload MDIForm1
End Sub

