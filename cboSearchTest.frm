VERSION 5.00
Object = "*\ACBOActiveSrchControl.vbp"
Begin VB.Form frmComboTesr 
   Caption         =   "Test Combo"
   ClientHeight    =   4635
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6525
   LinkTopic       =   "Form1"
   ScaleHeight     =   4635
   ScaleWidth      =   6525
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   0
      TabIndex        =   6
      Text            =   "Project Class - Quick and easy and relatively Maintainable"
      Top             =   3960
      Width           =   6495
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      TabIndex        =   3
      Text            =   "A Very Classy Control in AXCTIVEX DLL - Ultimate in code reuse"
      Top             =   1320
      Width           =   6465
   End
   Begin VB.CheckBox chkAutoDrop 
      Caption         =   "AutoDrop"
      Height          =   315
      Left            =   3960
      TabIndex        =   2
      Top             =   0
      Width           =   1665
   End
   Begin VB.CheckBox chkAutoSrch 
      Caption         =   "AutoSearch"
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   0
      Width           =   1365
   End
   Begin combocontrol.UserControl1 UserControl11 
      Height          =   345
      Left            =   0
      TabIndex        =   0
      Top             =   2400
      Width           =   6465
      _ExtentX        =   11404
      _ExtentY        =   609
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FontName        =   "MS Sans Serif"
      FontSize        =   8.25
      ListIndex       =   -1
      Text            =   "ActiveX - Also Good Code reUse"
   End
   Begin VB.Label Label3 
      Caption         =   "Combo Box Referenced by Class cbosearchprivate in project class"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   3600
      Width           =   4935
   End
   Begin VB.Label Label2 
      Caption         =   "Combo Box Referenced by Class cbosearch in AxctiveX DLL"
      Height          =   255
      Left            =   840
      TabIndex        =   5
      Top             =   960
      Width           =   5415
   End
   Begin VB.Label Label1 
      Caption         =   "ActiveX Control - UserControl1"
      Height          =   255
      Left            =   1080
      TabIndex        =   4
      Top             =   2040
      Width           =   3615
   End
End
Attribute VB_Name = "frmComboTesr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private cbosrch As cbosearch
Private cbosrchpriv As cbosearchprivate

'Note that this form contains very little code

' It just instantiates the class objects , sets their control reference properties
'           and loads the combobox lists with random data

'THERE IS NO CODE IN THIS FORM THAT PERFORMS ANY OF THE ACTUAL METHODS/PROPERTIES/EVENTS USED IN THE LOOKUP/DROPDOWN FUNCTIONS


Private Sub chkAutoSrch_Click()
'Auto Search or No Autosearch
UserControl11.AllowAutoSearch = CBool(chkAutoSrch.Value)
cbosrch.Autosearch = CBool(chkAutoSrch.Value)
cbosrchpriv.Autosearch = CBool(chkAutoSrch.Value)
End Sub

Private Sub chkAutoDrop_Click()
'Auto Drop or No AutoDrop
 UserControl11.Autodrop = CBool(chkAutoDrop.Value)
 cbosrch.Autodrop = CBool(chkAutoDrop.Value)
 cbosrchpriv.Autodrop = CBool(chkAutoDrop.Value)
 
 
End Sub



Private Sub Form_Load()
UserControl11.Clear
Combo1.Clear
Combo2.Clear
Dim aloop As Integer
Dim iloop As Integer
Dim s As String



'FILL THE COMBOLISTS WITH RANDOM LETTERS
For iloop = 1 To 10000
    s = ""
    For aloop = 1 To 10
        s = s & Chr(65 + (Rnd * 26))
    Next
    UserControl11.AddItem (s)
    Combo1.AddItem (s)
    Combo2.AddItem (s)
Next



'SET UP THE INITIAL PROPERTIES OF THE AXCTIVEX CONTROL IF YOU NEED TO
'NOTE THE PLETHORA OF OTHER PROPERTIES / METHODS / EVENTS THAT THIS ACTIVE X CONTROL IMPLEMENTS AS PASSTHROUGHS TO ITS CONTAINED COMBOBOX
UserControl11.BackColor = &H80000005
'ETC






'INSTANTIATES THE CLASS OBJECT AND SET THE CONTROL REFERENCE PROPERTY OF OUR ACTIVE X DLL CLASS
Set cbosrch = New cbosearch

'cbosrch.oCBOX =Combo1
'The above causes an error because Public classes cannot reference controls directly in their Let procedures
'"Compile Error - Private object models cannot be used in public object models as parameters or return types for public procedures, as public data members , or as fields of public user defined types."

cbosrch.ptCBOX = ObjPtr(Combo1)


'NB. CLASS ONLY EXPOSES A FEW PROPERTIES AND IS MUCH LESS CLUTTERED IMHO - CF USERCONTROL1
'Properties of the combobox can be set in the form here or in the Class eg
Combo1.BackColor = &H80000005





'INSTANTIATES THE CLASS OBJECT AND SET THE CONTROL REFERENCE PROPERTY OF OUR PROJECT (PRIVATE) CLASS
'NOTE That it is a bit easier here as we can leave out the pointer hack

Set cbosrchpriv = New cbosearchprivate
cbosrchpriv.oCBOX = Combo2
'cbosrchpriv.ptCBOX = ObjPtr(Combo2)






End Sub
'THATS IT FOR THIS SIMPLE FORM





'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX
'CODE BELOW IS WHAT YOU WOULD HAVE --- IF YOU HAD NOT SEGREGATED
'THE GUI Functions into a class,  ACTIVEX DLL or ACTIVEX CONTROL.
'AND THIS IS JUST FOR SETTING UP ONE CONTROL

'This is the code that could be placed on a your form to handle the search

'Dont place multiple code fragments such as this in your forms and expect to be able to maintain code in large projects - It does not work
 
'You would need to add the constants and declarations on the form etc, and if you want to change the code will you remember where all code occurs/requires changing

'An option is place the code in a bas  module - again dont do it . Why ? It still ends up being a night mare to maintain if you use the code in multiple VB projects

' Maintain Code ReUse by putting all the GUI code in an ActiveX control or a class (preferably a class in a single easily maintainable activeX DLL) - See cbosearch.cls

' Doing it this way will let you encapsulate GUI functions in a class/s

' If you do the right thing - Have business logic and GUI functions in classes your form code will be minimal and easy to maintain.



'XXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXXX

'local variable(s) to hold property value(s)
'
'Private Const CB_SETDROPPEDWIDTH = &H160  ' Constant for sendmessage to cboboxes to auto Size
'Private Const LB_ITEMFROMPOINT = &H1A9 'For Getting listindex of listbox
'
'Private Const LB_FINDSTRING As Long = &H18F
'Private Const LB_FINDSTRINGEXACT As Long = &H1A2
'Private Const CB_ERR As Long = (-1)
'Private Const LB_ERR As Long = (-1)
'
'Private Const WM_USER As Long = &H400
'Private Const CB_FINDSTRING As Long = &H14C
'Private Const CB_SHOWDROPDOWN As Long = &H14F
'
'
'Private Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" ( _
'    lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
'    Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" _
'(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, _
'ByVal lParam As Long) As Long
'
'
'Private Declare Function SendMessageStr Lib _
'    "user32" Alias "SendMessageA" _
'    (ByVal hwnd As Long, _
'     ByVal wMsg As Long, _
'     ByVal wParam As Long, _
'     ByVal lParam As String) As Long
'
'
'
'Private Sub Combo1_GotFocus()
'If mvarAutodrop Then
'    SendMessage Combo1.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0&
'End If
'End Sub
'
'Private Sub Combo1_KeyDown(KeyCode As Integer, Shift As Integer)
'    If mvarAutosearch Then
'        If KeyCode > 33 And KeyCode < 41 Then KeyCode = 0
'    End If
'
'End Sub
'
'Private Sub Combo1_KeyPress(KeyAscii As Integer)
'If mvarAutosearch Then
'    search KeyAscii
'End If
'
'
'
'End Sub
'
'
'Private Sub search(ByRef intKey As Integer)


'
''Put the code in a class
'
'Dim lngIdx As Long
'Dim FindString As String
'
'If (intKey < 32 Or intKey > 127) And _
'   (Not (intKey = 13 Or intKey = 8)) Then Exit Sub
'
'If Not intKey = 13 Or intKey = 8 Then
'    If Len(Combo1.Text) = 0 Then
'        FindString = Chr$(intKey)
'    Else
'        FindString = Left$(Combo1.Text, Combo1.SelStart) & Chr$(intKey)
'    End If
'End If
'
'If intKey = 8 Then
'   If Len(Combo1.Text) = 0 Then Exit Sub
'   Dim numChars As Integer
'   numChars = Combo1.SelStart - 1
'   'FindString = Left(str, numChars)
'   If numChars > 0 Then FindString = Left(Combo1.Text, numChars)
'End If
'
'        If intKey = 13 Then
'          Call SendMessageStr(Combo1.hwnd, _
'             CB_SHOWDROPDOWN, True, 0&)
'          Exit Sub
'        End If
'    lngIdx = SendMessageStr(Combo1.hwnd, _
'       CB_FINDSTRING, -1, FindString)
'
'
'If lngIdx <> -1 Then
'        Combo1.ListIndex = lngIdx
'        Combo1.SelStart = Len(FindString)
'        Combo1.SelLength = Len(Combo1.Text) - Combo1.SelStart
'End If
'intKey = 0
'
'
'
'End Sub
'


