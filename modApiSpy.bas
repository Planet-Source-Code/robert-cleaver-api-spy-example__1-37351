Attribute VB_Name = "modApiSpy"
'///////////////////////////////////////////////////
'////////// By Robert Cleaver /////////////////
'//////////////////////////////////////////////////

Option Explicit


Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long

Public Type POINTAPI
        X As Long
        Y As Long
End Type


'// Constants //
Public Const WM_ACTIVATE = &H6
Public Const WM_ACTIVATEAPP = &H1C
Public Const WM_ASKCBFORMATNAME = &H30C
Public Const WM_CANCELJOURNAL = &H4B
Public Const WM_CANCELMODE = &H1F
Public Const WM_CHANGECBCHAIN = &H30D
Public Const WM_CHAR = &H102
Public Const WM_CHARTOITEM = &H2F
Public Const WM_CHILDACTIVATE = &H22
Public Const WM_CLEAR = &H303
Public Const WM_CLOSE = &H10
Public Const WM_COMMAND = &H111
Public Const WM_COMPACTING = &H41
Public Const WM_COMPAREITEM = &H39
Public Const WM_COPY = &H301
Public Const WM_COPYDATA = &H4A
Public Const WM_CREATE = &H1
Public Const WM_CTLCOLORBTN = &H135
Public Const WM_CTLCOLORDLG = &H136
Public Const WM_CTLCOLOREDIT = &H133
Public Const WM_CTLCOLORLISTBOX = &H134
Public Const WM_CTLCOLORMSGBOX = &H132
Public Const WM_CTLCOLORSCROLLBAR = &H137
Public Const WM_CTLCOLORSTATIC = &H138
Public Const WM_CUT = &H300
Public Const WM_DEADCHAR = &H103
Public Const WM_DELETEITEM = &H2D
Public Const WM_DESTROY = &H2
Public Const WM_DESTROYCLIPBOARD = &H307
Public Const WM_DEVMODECHANGE = &H1B
Public Const WM_DRAWCLIPBOARD = &H308
Public Const WM_DRAWITEM = &H2B
Public Const WM_DROPFILES = &H233
Public Const WM_ENABLE = &HA
Public Const WM_ENDSESSION = &H16
Public Const WM_ENTERIDLE = &H121
Public Const WM_ENTERMENULOOP = &H211
Public Const WM_ERASEBKGND = &H14
Public Const WM_EXITMENULOOP = &H212
Public Const WM_FONTCHANGE = &H1D
Public Const WM_GETDLGCODE = &H87
Public Const WM_GETFONT = &H31
Public Const WM_GETHOTKEY = &H33
Public Const WM_GETMINMAXINFO = &H24
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_HOTKEY = &H312
Public Const WM_HSCROLL = &H114
Public Const WM_HSCROLLCLIPBOARD = &H30E
Public Const WM_ICONERASEBKGND = &H27
Public Const WM_INITDIALOG = &H110
Public Const WM_INITMENU = &H116
Public Const WM_INITMENUPOPUP = &H117
Public Const WM_KEYDOWN = &H100
Public Const WM_KEYFIRST = &H100
Public Const WM_KEYLAST = &H108
Public Const WM_KEYUP = &H101
Public Const WM_KILLFOCUS = &H8
Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_LBUTTONUP = &H202
Public Const WM_MBUTTONDBLCLK = &H209
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MDIACTIVATE = &H222
Public Const WM_MDICASCADE = &H227
Public Const WM_MDICREATE = &H220
Public Const WM_MDIDESTROY = &H221
Public Const WM_MDIGETACTIVE = &H229
Public Const WM_MDIICONARRANGE = &H228
Public Const WM_MDIMAXIMIZE = &H225
Public Const WM_MDINEXT = &H224
Public Const WM_MDIREFRESHMENU = &H234
Public Const WM_MDIRESTORE = &H223
Public Const WM_MDISETMENU = &H230
Public Const WM_MDITILE = &H226
Public Const WM_MEASUREITEM = &H2C
Public Const WM_MENUCHAR = &H120
Public Const WM_MENUSELECT = &H11F
Public Const WM_MOUSEACTIVATE = &H21
Public Const WM_MOUSEFIRST = &H200
Public Const WM_MOUSELAST = &H209
Public Const WM_MOUSEMOVE = &H200
Public Const WM_MOVE = &H3
Public Const WM_NCACTIVATE = &H86
Public Const WM_NCCALCSIZE = &H83
Public Const WM_NCCREATE = &H81
Public Const WM_NCDESTROY = &H82
Public Const WM_NCHITTEST = &H84
Public Const WM_NCLBUTTONDBLCLK = &HA3
Public Const WM_NCLBUTTONDOWN = &HA1
Public Const WM_NCLBUTTONUP = &HA2
Public Const WM_NCMBUTTONDBLCLK = &HA9
Public Const WM_NCMBUTTONDOWN = &HA7
Public Const WM_NCMBUTTONUP = &HA8
Public Const WM_NCMOUSEMOVE = &HA0
Public Const WM_NCPAINT = &H85
Public Const WM_NCRBUTTONDOWN = &HA4
Public Const WM_NCRBUTTONUP = &HA5
Public Const WM_NEXTDLGCTL = &H28
Public Const WM_NULL = &H0
Public Const WM_PAINT = &HF
Public Const WM_PAINTCLIPBOARD = &H309
Public Const WM_PAINTICON = &H26
Public Const WM_PALETTECHANGED = &H311
Public Const WM_PALETTEISCHANGING = &H310
Public Const WM_PARENTNOTIFY = &H210
Public Const WM_PASTE = &H302
Public Const WM_PENWINFIRST = &H380
Public Const WM_PENWINLAST = &H38F
Public Const WM_POWER = &H48
Public Const WM_QUERYDRAGICON = &H37
Public Const WM_QUERYENDSESSION = &H11
Public Const WM_QUERYNEWPALETTE = &H30F
Public Const WM_QUERYOPEN = &H13
Public Const WM_QUEUESYNC = &H23
Public Const WM_QUIT = &H12
Public Const WM_RBUTTONDBLCLK = &H206
Public Const WM_RBUTTONDOWN = &H204
Public Const WM_RBUTTONUP = &H205
Public Const WM_RENDERALLFORMATS = &H306
Public Const WM_RENDERFORMAT = &H305
Public Const WM_SETCURSOR = &H20
Public Const WM_SETFOCUS = &H7
Public Const WM_SETFONT = &H30
Public Const WM_SETHOTKEY = &H32
Public Const WM_SETREDRAW = &HB
Public Const WM_SETTEXT = &HC
Public Const WM_SHOWWINDOW = &H18
Public Const WM_SIZE = &H5
Public Const WM_SIZECLIPBOARD = &H30B
Public Const WM_SPOOLERSTATUS = &H2A
Public Const WM_SYSCHAR = &H106
Public Const WM_SYSCOLORCHANGE = &H15
Public Const WM_SYSCOMMAND = &H112
Public Const WM_SYSDEADCHAR = &H107
Public Const WM_SYSKEYDOWN = &H104
Public Const WM_SYSKEYUP = &H105
Public Const WM_TIMECHANGE = &H1E
Public Const WM_TIMER = &H113
Public Const WM_UNDO = &H304
Public Const WM_USER = &H400
Public Const WM_VKEYTOITEM = &H2E
Public Const WM_VSCROLL = &H115
Public Const WM_VSCROLLCLIPBOARD = &H30A
Public Const WM_WINDOWPOSCHANGED = &H47
Public Const WM_WINDOWPOSCHANGING = &H46
Public Const WM_WININICHANGE = &H1A
Public Const WPF_RESTORETOMAXIMIZED = &H2
Public Const WPF_SETMINPOSITION = &H1

'// Booleans //
    Global InformationNow As Boolean





Function GetWindowInformation(WindowHandle2&, WindowClassName2$, WindowText2$, ParentList As ListBox)
    Dim CursorPos As POINTAPI
    Dim BufferAll&, WindowHandle&, TextLength&, PrevHandle&
    Dim WindowClassName$, WindowText$
    Call GetCursorPos(CursorPos)
    WindowHandle& = WindowFromPoint(CursorPos.X, CursorPos.Y)
    WindowClassName$ = String(100, Chr(0))
    BufferAll& = GetClassName(WindowHandle&, WindowClassName$, 100)
    WindowClassName$ = Left(WindowClassName$, BufferAll&)
    WindowText$ = String(100, Chr(0))
    BufferAll& = GetWindowTextLength(WindowHandle&)
    BufferAll& = GetWindowText(WindowHandle&, WindowText$, BufferAll& + 1)
    WindowText$ = Left(WindowText$, BufferAll&)
    WindowHandle2& = WindowHandle&
    WindowClassName2$ = WindowClassName$
    WindowText2$ = WindowText$
    PrevHandle& = GetParent(WindowHandle&)
    ParentList.Clear
    Do While PrevHandle& <> 0
        PrevHandle& = GetParent(WindowHandle&)
        WindowHandle& = PrevHandle&
        WindowClassName$ = String(100, Chr(0))
        BufferAll& = GetClassName(WindowHandle&, WindowClassName$, 100)
        WindowClassName$ = Left(WindowClassName$, BufferAll&)
        ParentList.AddItem (WindowClassName$)
    Loop
End Function
