VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cbosearchprivate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'local variable(s) to hold property value(s)
Private WithEvents mvaroCBO As ComboBox 'local copy
Attribute mvaroCBO.VB_VarHelpID = -1


Private Const CB_SETDROPPEDWIDTH = &H160  ' Constant for sendmessage to cboboxes to auto Size
Private Const LB_ITEMFROMPOINT = &H1A9 'For Getting listindex of listbox
 
Private Const LB_FINDSTRING As Long = &H18F
Private Const LB_FINDSTRINGEXACT As Long = &H1A2
Private Const CB_ERR As Long = (-1)
Private Const LB_ERR As Long = (-1)

Private Const WM_USER As Long = &H400
Private Const CB_FINDSTRING As Long = &H14C
Private Const CB_SHOWDROPDOWN As Long = &H14F


'local variable(s) to hold property value(s)
Private mvarAutosearch As Boolean 'local copy
Private mvarAutodrop As Boolean 'local copy
Private mvarAutoDropWidth As Integer 'local copy

Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" ( _
    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
ByVal lParam As Long) As Long


Private Declare Function SendMessageStr Lib _
    "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     ByVal lParam As String) As Long


Public Property Let AutoDropWidth(ByVal vData As Integer)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.AutoDropWidth = 5
    mvarAutoDropWidth = vData
End Property


Public Property Get AutoDropWidth() As Integer
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.AutoDropWidth
    AutoDropWidth = mvarAutoDropWidth
End Property



Public Property Let Autodrop(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Autodrop = 5
    mvarAutodrop = vData
End Property


Public Property Get Autodrop() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Autodrop
    Autodrop = mvarAutodrop
End Property



Public Property Let Autosearch(ByVal vData As Boolean)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.Autosearch = 5
    mvarAutosearch = vData
End Property


Public Property Get Autosearch() As Boolean
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.Autosearch
    Autosearch = mvarAutosearch
End Property




'Here is the error generator
'This code is legal in a private class module - Classes within an EXE project are always private to the EXE

Public Property Let oCBOX(ByRef vData As ComboBox)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.oCBO = Form1

    Set mvaroCBO = vData

End Property



Public Property Let ptCBOX(ByVal vData As Long)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.oCBO = Form1
If vData = 0 Then
    Set mvaroCBO = Nothing
    Exit Property
End If

   Set mvaroCBO = ObjectFromPtr(vData)
End Property



Private Sub Class_Initialize()
mvarAutodrop = False
mvarAutosearch = False
End Sub



Private Sub Class_Terminate()
Set mvaroCBO = Nothing
End Sub

Private Sub mvaroCBO_GotFocus()
If mvarAutodrop Then
    SendMessage mvaroCBO.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&
End If
End Sub

Private Sub mvaroCBO_KeyDown(KeyCode As Integer, Shift As Integer)
    If mvarAutosearch Then
        If KeyCode > 33 And KeyCode < 41 Then KeyCode = 0
    End If
    
End Sub

Private Sub mvaroCBO_KeyPress(KeyAscii As Integer)
If mvarAutosearch Then
    search KeyAscii
End If



End Sub


Private Sub search(ByRef intKey As Integer)
'This is the code that could be placed on a your form to handle the search

'Dont place multiple code fragments such as this in your forms and expect to be able to maintain code in large projects - It does not work
' You would need to add the constants and declarations on the form etc, and if you want to change the code will you remember where all code occurs/requires changing

'An option is place the code in a bas  module - again dont do it . Why ? It still ends up being a night mare to maintain if you use the code in multiple VB projects

' Maintain Code ReUse by putting all the GUI code in an ActiveX control or a class (preferably a class in a single easily maintainable activeX DLL) - See cbosearch.cls

' Doing it this way will let you encapsulate GUI functions in a class/s

' If you do the right thing - Have business logic and GUI functions in classes your form code will be minimal and easy to maintain.



Dim lngIdx As Long
Dim FindString As String

If (intKey < 32 Or intKey > 127) And _
   (Not (intKey = 13 Or intKey = 8)) Then Exit Sub

If Not intKey = 13 Or intKey = 8 Then
    If Len(mvaroCBO.Text) = 0 Then
        FindString = Chr$(intKey)
    Else
        FindString = Left$(mvaroCBO.Text, mvaroCBO.SelStart) & Chr$(intKey)
    End If
End If

If intKey = 8 Then
   If Len(mvaroCBO.Text) = 0 Then Exit Sub
   Dim numChars As Integer
   numChars = mvaroCBO.SelStart - 1
   'FindString = Left(str, numChars)
   If numChars > 0 Then FindString = Left(mvaroCBO.Text, numChars)
End If

        If intKey = 13 Then
          Call SendMessageStr(mvaroCBO.hwnd, _
             CB_SHOWDROPDOWN, True, 0&)
          Exit Sub
        End If
    lngIdx = SendMessageStr(mvaroCBO.hwnd, _
       CB_FINDSTRING, -1, FindString)

 
If lngIdx <> -1 Then
        mvaroCBO.ListIndex = lngIdx
        mvaroCBO.SelStart = Len(FindString)
        mvaroCBO.SelLength = Len(mvaroCBO.Text) - mvaroCBO.SelStart
End If
intKey = 0



End Sub


Private Function ObjectFromPtr(ByVal lPtr As Long) As Object
Dim objT As Object
   If Not (lPtr = 0) Then
      ' Turn the pointer into an illegal, uncounted interface
      CopyMemory objT, lPtr, 4
      ' Do NOT hit the End button here! You will crash!
      ' Assign to legal reference
      Set ObjectFromPtr = objT
      ' Still do NOT hit the End button here! You will still crash!
      ' Destroy the illegal reference
      CopyMemory objT, 0&, 4
   End If
End Function


