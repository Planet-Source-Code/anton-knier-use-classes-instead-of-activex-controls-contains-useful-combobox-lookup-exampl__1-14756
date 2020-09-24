VERSION 5.00
Begin VB.UserControl UserControl1 
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4650
   ScaleHeight     =   345
   ScaleWidth      =   4650
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   0
      Width           =   4665
   End
End
Attribute VB_Name = "UserControl1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'Default Property Values:
Const m_def_AutoDrop = False
Const m_def_Sorted = 0

Const m_def_ToolTipText = ""
Const m_def_AllowAutoSearch = False
Const m_def_AutoSearchDelay = 50
'Property Variables:
Dim m_AutoDrop As Boolean
Dim m_Sorted As Boolean
Dim m_ToolTipText As String
Dim m_AllowAutoSearch As Boolean
Dim m_AutoSearchDelay As Integer
'Event Declarations:
Event Change() 'MappingInfo=Combo1,Combo1,-1,Change
Attribute Change.VB_Description = "Occurs when the contents of a control have changed."
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
Event Click() 'MappingInfo=Combo1,Combo1,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."
Event DblClick() 'MappingInfo=Combo1,Combo1,-1,DblClick
Attribute DblClick.VB_Description = "Occurs when the user presses and releases a mouse button and then presses and releases it again over an object."
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=Combo1,Combo1,-1,KeyDown
Attribute KeyDown.VB_Description = "Occurs when the user presses a key while an object has the focus."
Event KeyPress(KeyAscii As Integer) 'MappingInfo=Combo1,Combo1,-1,KeyPress
Attribute KeyPress.VB_Description = "Occurs when the user presses and releases an ANSI key."
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=Combo1,Combo1,-1,KeyUp
Attribute KeyUp.VB_Description = "Occurs when the user releases a key while an object has the focus."
'Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
'Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Private Const CB_SETDROPPEDWIDTH = &H160  ' Constant for sendmessage to cboboxes to auto Size
Private Const LB_ITEMFROMPOINT = &H1A9 'For Getting listindex of listbox
 
Private Const LB_FINDSTRING As Long = &H18F
Private Const LB_FINDSTRINGEXACT As Long = &H1A2
Private Const CB_ERR As Long = (-1)
Private Const LB_ERR As Long = (-1)

Private Const WM_USER As Long = &H400
Private Const CB_FINDSTRING As Long = &H14C
Private Const CB_SHOWDROPDOWN As Long = &H14F












 Private Declare Function SendMessageStr Lib _
    "user32" Alias "SendMessageA" _
    (ByVal hwnd As Long, _
     ByVal wMsg As Long, _
     ByVal wParam As Long, _
     ByVal lParam As String) As Long
     
     
     
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
ByVal lParam As Long) As Long







'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = Combo1.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    Combo1.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,ForeColor
Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = Combo1.ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    Combo1.ForeColor() = New_ForeColor
    PropertyChanged "ForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,Enabled
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = Combo1.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    Combo1.Enabled() = New_Enabled
    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = Combo1.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set Combo1.Font = New_Font
    PropertyChanged "Font"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BorderStyle
Public Property Get BorderStyle() As Integer
Attribute BorderStyle.VB_Description = "Returns/sets the border style for an object."
    BorderStyle = UserControl.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As Integer)
    UserControl.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,Refresh
Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    Combo1.Refresh
End Sub

Private Sub Combo1_Click()
    RaiseEvent Click
End Sub

Private Sub Combo1_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub Combo1_GotFocus()
If m_AutoDrop Then
    SendMessage Combo1.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&
End If
End Sub

Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
    If m_AllowAutoSearch Then
        If KeyCode > 33 And KeyCode < 41 Then KeyCode = 0
        Exit Sub
    End If
    RaiseEvent KeyDown(KeyCode, Shift)

End Sub

Private Sub Combo1_KeyPress(KeyAscii As Integer)
If m_AllowAutoSearch Then
    search KeyAscii
End If



'If m_AllowAutoSearch Then
'    cbox.FindIndexStr Combo1, Combo1.Text, KeyAscii
'    KeyAscii = 0
'    Combo1.SelLength = 0
'End If



RaiseEvent KeyPress(KeyAscii)
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
    If Len(Combo1.Text) = 0 Then
        FindString = Chr$(intKey)
    Else
        FindString = Left$(Combo1.Text, Combo1.SelStart) & Chr$(intKey)
    End If
End If

If intKey = 8 Then
   If Len(Combo1.Text) = 0 Then Exit Sub
   Dim numChars As Integer
   numChars = Combo1.SelStart - 1
   'FindString = Left(str, numChars)
   If numChars > 0 Then FindString = Left(Combo1.Text, numChars)
End If

        If intKey = 13 Then
          Call SendMessageStr(Combo1.hwnd, _
             CB_SHOWDROPDOWN, True, 0&)
          Exit Sub
        End If
    lngIdx = SendMessageStr(Combo1.hwnd, _
       CB_FINDSTRING, -1, FindString)

 
If lngIdx <> -1 Then
        Combo1.ListIndex = lngIdx
        Combo1.SelStart = Len(FindString)
        Combo1.SelLength = Len(Combo1.Text) - Combo1.SelStart
End If
intKey = 0



End Sub















Private Sub Combo1_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,FontBold
Public Property Get FontBold() As Boolean
Attribute FontBold.VB_Description = "Returns/sets bold font styles."
    FontBold = Combo1.FontBold
End Property

Public Property Let FontBold(ByVal New_FontBold As Boolean)
    Combo1.FontBold() = New_FontBold
    PropertyChanged "FontBold"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,FontItalic
Public Property Get FontItalic() As Boolean
Attribute FontItalic.VB_Description = "Returns/sets italic font styles."
    FontItalic = Combo1.FontItalic
End Property

Public Property Let FontItalic(ByVal New_FontItalic As Boolean)
    Combo1.FontItalic() = New_FontItalic
    PropertyChanged "FontItalic"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,FontName
Public Property Get FontName() As String
Attribute FontName.VB_Description = "Specifies the name of the font that appears in each row for the given level."
    FontName = Combo1.FontName
End Property

Public Property Let FontName(ByVal New_FontName As String)
    Combo1.FontName() = New_FontName
    PropertyChanged "FontName"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,FontSize
Public Property Get FontSize() As Single
Attribute FontSize.VB_Description = "Specifies the size (in points) of the font that appears in each row for the given level."
    FontSize = Combo1.FontSize
End Property

Public Property Let FontSize(ByVal New_FontSize As Single)
    Combo1.FontSize() = New_FontSize
    PropertyChanged "FontSize"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,FontStrikethru
Public Property Get FontStrikethru() As Boolean
Attribute FontStrikethru.VB_Description = "Returns/sets strikethrough font styles."
    FontStrikethru = Combo1.FontStrikethru
End Property

Public Property Let FontStrikethru(ByVal New_FontStrikethru As Boolean)
    Combo1.FontStrikethru() = New_FontStrikethru
    PropertyChanged "FontStrikethru"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,FontUnderline
Public Property Get FontUnderline() As Boolean
Attribute FontUnderline.VB_Description = "Returns/sets underline font styles."
    FontUnderline = Combo1.FontUnderline
End Property

Public Property Let FontUnderline(ByVal New_FontUnderline As Boolean)
    Combo1.FontUnderline() = New_FontUnderline
    PropertyChanged "FontUnderline"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,ListCount
Public Property Get ListCount() As Integer
Attribute ListCount.VB_Description = "Returns the number of items in the list portion of a control."
    ListCount = Combo1.ListCount
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,ListIndex
Public Property Get ListIndex() As Integer
Attribute ListIndex.VB_Description = "Returns/sets the index of the currently selected item in the control."
    ListIndex = Combo1.ListIndex
End Property

Public Property Let ListIndex(ByVal New_ListIndex As Integer)
    Combo1.ListIndex() = New_ListIndex
    PropertyChanged "ListIndex"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,List
Public Property Get List(ByVal Index As Integer) As String
Attribute List.VB_Description = "Returns/sets the items contained in a control's list portion."
    List = Combo1.List(Index)
End Property

Public Property Let List(ByVal Index As Integer, ByVal New_List As String)
    Combo1.List(Index) = New_List
    PropertyChanged "List"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,Locked
Public Property Get Locked() As Boolean
Attribute Locked.VB_Description = "Determines whether a control can be edited."
    Locked = Combo1.Locked
End Property

Public Property Let Locked(ByVal New_Locked As Boolean)
    Combo1.Locked() = New_Locked
    PropertyChanged "Locked"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,SelLength
Public Property Get SelLength() As Long
Attribute SelLength.VB_Description = "Returns/sets the number of characters selected."
    SelLength = Combo1.SelLength
End Property

Public Property Let SelLength(ByVal New_SelLength As Long)
    Combo1.SelLength() = New_SelLength
    PropertyChanged "SelLength"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,SelStart
Public Property Get SelStart() As Long
Attribute SelStart.VB_Description = "Returns/sets the starting point of text selected."
    SelStart = Combo1.SelStart
End Property

Public Property Let SelStart(ByVal New_SelStart As Long)
    Combo1.SelStart() = New_SelStart
    PropertyChanged "SelStart"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,SelText
Public Property Get SelText() As String
Attribute SelText.VB_Description = "Returns/sets the string containing the currently selected text."
    SelText = Combo1.SelText
End Property

Public Property Let SelText(ByVal New_SelText As String)
    Combo1.SelText() = New_SelText
    PropertyChanged "SelText"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,Text
Public Property Get Text() As String
Attribute Text.VB_Description = "Returns/sets the text contained in the control."
    Text = Combo1.Text
End Property

Public Property Let Text(ByVal New_Text As String)
    Combo1.Text() = New_Text
    PropertyChanged "Text"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get AllowAutoSearch() As Boolean
    AllowAutoSearch = m_AllowAutoSearch
End Property

Public Property Let AllowAutoSearch(ByVal New_AllowAutoSearch As Boolean)
    m_AllowAutoSearch = New_AllowAutoSearch
    PropertyChanged "AllowAutoSearch"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,1000
Public Property Get AutoSearchDelay() As Integer
    AutoSearchDelay = m_AutoSearchDelay
End Property

Public Property Let AutoSearchDelay(ByVal New_AutoSearchDelay As Integer)
    m_AutoSearchDelay = New_AutoSearchDelay
    PropertyChanged "AutoSearchDelay"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_AllowAutoSearch = m_def_AllowAutoSearch
    m_AutoSearchDelay = m_def_AutoSearchDelay
    m_Sorted = m_def_Sorted
    m_ToolTipText = m_def_ToolTipText
    m_AutoDrop = m_def_AutoDrop
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
Dim Index As Integer

    Combo1.BackColor = PropBag.ReadProperty("BackColor", &H80000005)
    Combo1.ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    Combo1.Enabled = PropBag.ReadProperty("Enabled", True)
    Set Combo1.Font = PropBag.ReadProperty("Font", Ambient.Font)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 1)
    UserControl.BorderStyle = PropBag.ReadProperty("BorderStyle", 0)
    Combo1.FontBold = PropBag.ReadProperty("FontBold", 0)
    Combo1.FontItalic = PropBag.ReadProperty("FontItalic", 0)
    Combo1.FontName = PropBag.ReadProperty("FontName", "MS Sans Serif")
    Combo1.FontSize = PropBag.ReadProperty("FontSize", 9)
    Combo1.FontStrikethru = PropBag.ReadProperty("FontStrikethru", 0)
    Combo1.FontUnderline = PropBag.ReadProperty("FontUnderline", 0)
    Combo1.ListIndex = PropBag.ReadProperty("ListIndex", -1)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    Combo1.List(Index) = PropBag.ReadProperty("List" & Index, "")
    Combo1.Locked = PropBag.ReadProperty("Locked", False)
    Combo1.SelLength = PropBag.ReadProperty("SelLength", 0)
    Combo1.SelStart = PropBag.ReadProperty("SelStart", 0)
    Combo1.SelText = PropBag.ReadProperty("SelText", "")
    Combo1.Text = PropBag.ReadProperty("Text", "Combo1")
    m_AllowAutoSearch = PropBag.ReadProperty("AllowAutoSearch", m_def_AllowAutoSearch)
    m_AutoSearchDelay = PropBag.ReadProperty("AutoSearchDelay", m_def_AutoSearchDelay)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    Combo1.ItemData(Index) = PropBag.ReadProperty("ItemData" & Index, 0)
    Combo1.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    m_Sorted = PropBag.ReadProperty("Sorted", m_def_Sorted)
    m_ToolTipText = PropBag.ReadProperty("ToolTipText", m_def_ToolTipText)
    UserControl.DrawStyle = PropBag.ReadProperty("DrawStyle", 0)
    m_AutoDrop = PropBag.ReadProperty("AutoDrop", m_def_AutoDrop)
End Sub

Private Sub UserControl_Resize()
Combo1.Top = 0
Combo1.Left = 0
Combo1.Width = UserControl.ScaleWidth

End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
Dim Index As Integer

    Call PropBag.WriteProperty("BackColor", Combo1.BackColor, &H80000005)
    Call PropBag.WriteProperty("ForeColor", Combo1.ForeColor, &H80000012)
    Call PropBag.WriteProperty("Enabled", Combo1.Enabled, True)
    Call PropBag.WriteProperty("Font", Combo1.Font, Ambient.Font)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 1)
    Call PropBag.WriteProperty("BorderStyle", UserControl.BorderStyle, 0)
    Call PropBag.WriteProperty("FontBold", Combo1.FontBold, 0)
    Call PropBag.WriteProperty("FontItalic", Combo1.FontItalic, 0)
    Call PropBag.WriteProperty("FontName", Combo1.FontName, "")
    Call PropBag.WriteProperty("FontSize", Combo1.FontSize, 0)
    Call PropBag.WriteProperty("FontStrikethru", Combo1.FontStrikethru, 0)
    Call PropBag.WriteProperty("FontUnderline", Combo1.FontUnderline, 0)
    Call PropBag.WriteProperty("ListIndex", Combo1.ListIndex, 0)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    Call PropBag.WriteProperty("List" & Index, Combo1.List(Index), "")
    Call PropBag.WriteProperty("Locked", Combo1.Locked, False)
    Call PropBag.WriteProperty("SelLength", Combo1.SelLength, 0)
    Call PropBag.WriteProperty("SelStart", Combo1.SelStart, 0)
    Call PropBag.WriteProperty("SelText", Combo1.SelText, "")
    Call PropBag.WriteProperty("Text", Combo1.Text, "Combo1")
    Call PropBag.WriteProperty("AllowAutoSearch", m_AllowAutoSearch, m_def_AllowAutoSearch)
    Call PropBag.WriteProperty("AutoSearchDelay", m_AutoSearchDelay, m_def_AutoSearchDelay)
'TO DO: The member you have mapped to contains an array of data.
'   You must supply the code to persist the array.  A prototype
'   line is shown next:
    Call PropBag.WriteProperty("ItemData" & Index, Combo1.ItemData(Index), 0)
    Call PropBag.WriteProperty("MousePointer", Combo1.MousePointer, 0)
    Call PropBag.WriteProperty("Sorted", m_Sorted, m_def_Sorted)
    Call PropBag.WriteProperty("ToolTipText", m_ToolTipText, m_def_ToolTipText)
    Call PropBag.WriteProperty("DrawStyle", UserControl.DrawStyle, 0)
    Call PropBag.WriteProperty("AutoDrop", m_AutoDrop, m_def_AutoDrop)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,hWnd
Public Property Get hwnd() As Long
Attribute hwnd.VB_Description = "Returns a handle (from Microsoft Windows) to an object's window."
    hwnd = Combo1.hwnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,ItemData
Public Property Get ItemData(ByVal Index As Integer) As Long
Attribute ItemData.VB_Description = "Returns/sets a specific number for each item in a ComboBox or ListBox control."
    ItemData = Combo1.ItemData(Index)
End Property

Public Property Let ItemData(ByVal Index As Integer, ByVal New_ItemData As Long)
    Combo1.ItemData(Index) = New_ItemData
    PropertyChanged "ItemData"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = Combo1.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    Combo1.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get Sorted() As Boolean
Attribute Sorted.VB_Description = "Indicates whether the elements of a control are automatically sorted alphabetically."
    Sorted = m_Sorted
End Property

Public Property Let Sorted(ByVal New_Sorted As Boolean)
    m_Sorted = New_Sorted
    PropertyChanged "Sorted"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=13,0,0,
Public Property Get ToolTipText() As String
Attribute ToolTipText.VB_Description = "Returns/sets the text displayed when the mouse is paused over the control."
    ToolTipText = m_ToolTipText
End Property

Public Property Let ToolTipText(ByVal New_ToolTipText As String)
    m_ToolTipText = New_ToolTipText
    PropertyChanged "ToolTipText"
End Property

Private Sub Combo1_Change()
    RaiseEvent Change
    
    
    
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,Clear
Public Sub Clear()
Attribute Clear.VB_Description = "Clears the contents of a control or the system Clipboard."
    Combo1.Clear
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Combo1,Combo1,-1,AddItem
Public Sub AddItem(ByVal Item As String, Optional ByVal Index As Variant)
Attribute AddItem.VB_Description = "Adds an item to a Listbox or ComboBox control or a row to a Grid control."
    Combo1.AddItem Item, Index
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,DrawStyle
Public Property Get DrawStyle() As Integer
Attribute DrawStyle.VB_Description = "Determines the line style for output from graphics methods."
    DrawStyle = UserControl.DrawStyle
End Property

Public Property Let DrawStyle(ByVal New_DrawStyle As Integer)
    UserControl.DrawStyle() = New_DrawStyle
    PropertyChanged "DrawStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,0
Public Property Get AutoDrop() As Boolean
    AutoDrop = m_AutoDrop
End Property

Public Property Let AutoDrop(ByVal New_AutoDrop As Boolean)
    m_AutoDrop = New_AutoDrop
    PropertyChanged "AutoDrop"
End Property

