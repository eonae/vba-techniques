VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsController"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'//////////////////////////////////////////////////////////////////////////////////
'/////////////// Controller-class, to manage multiple form-controls ///////////////
'//////////////////////////////////////////////////////////////////////////////////

Private pElements As Collection '//

Public Sub PassControls(ParamArray Controls() As Variant)
'// Gatheres control inside Controller instance.
'// This method can be used once to initialize Controller
    
    Dim Ctrl As Variant
    Dim i%
    Dim Wrapper As clsWrapper
    
    If Not pElements Is Nothing Then
        Debug.Print "Controls are already passed"
        Exit Sub '// Exiting sub if Controller-instance is already initialized.
    End If
    
    Set pElements = New Collection

    For Each Ctrl In Controls
    
        Set Wrapper = New clsWrapper '// Initializing new instance of Wrapper-object
        Set Wrapper.Control = Ctrl   '// Wrapping up control
        Set Wrapper.Controller = Me  '// Setting up Controller-object witch contains event handlers
        
        pElements.Add Wrapper '// Adding wrapped up control to collection
        
        Set Wrapper = Nothing
    Next
    
End Sub

'/////////////////////////////////////////////////////////
'//////////////////// Event hanlders /////////////////////
'/////////////////////////////////////////////////////////

'// This functions can take any kind of parameter. The only important thing is that this functions must be
'// correctly linked with functions in Wrapper-class

Public Sub MyEvent1(ByRef W As clsWrapper)
    MsgBox (W.Control.name & " fired MyEvent1!")
End Sub

Public Sub MyEvent2(ByRef W As clsWrapper)
    MsgBox (W.Control.name & " fired MyEvent2!")
End Sub

Public Sub MyEvent3(ByRef W As clsWrapper)
    MsgBox (W.Control.name & " fired MyEvent3!")
End Sub

Public Sub MyEvent4(ByRef W As clsWrapper)
    MsgBox (W.Control.name & " fired MyEvent4!")
End Sub
