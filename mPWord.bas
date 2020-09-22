Attribute VB_Name = "mPWord"
Option Explicit


' all credit for the api creation and subclassing of the windows
' goes to Joseph Huntley:
'
'
'-------------------------------------------
'Author APIForm: Joseph Huntley
'Level:          Advanced
'VB Version:     6.0
'IRC:            Joseph (IRC.xnet.org)
'ICQ:            55449964
'E -mail:        joseph_huntley@ email.com
'Webpage:        http://joseph.vr9.com
'-------------------------------------------

' thanks Joseph for showing the way, to possibly create more
' creative, efficient VB applications.

' Password Module Idea and Implentation by gh0ul.

'*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*
'
' Study this module carefully... do not make any changes until you are
' certain how these procedures work.
'
'*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*!*


Private Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Private Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Private Declare Function DefWindowProc Lib "user32" Alias "DefWindowProcA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As Msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function TranslateMessage Lib "user32" (lpMsg As Msg) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As Msg) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Any) As Long
Private Declare Function LoadIcon Lib "user32" Alias "LoadIconA" (ByVal hInstance As Long, ByVal lpIconName As String) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Declare Sub PostQuitMessage Lib "user32" (ByVal nExitCode As Long)
Private Declare Function SetWindowText Lib "user32.dll" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal Msg As Long, wParam As Any, lParam As Any) As Long
Private Declare Function GetWindowText Lib "user32.dll" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowTextLength Lib "user32.dll" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long


Private Type WNDCLASS
    style As Long
    lpfnwndproc As Long
    cbClsextra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type


Private Type POINTAPI
    x As Long
    y As Long
End Type


Private Type Msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Type PWord_Options
    Save As Integer
    Encoded As Boolean
    Path As String
    AppName As String
    Section As String
    Key As String
End Type

Public pw As PWord_Options


' A P I   D E F I N E D   C O N S T A N T S
 Const CS_VREDRAW = &H1
 Const CS_HREDRAW = &H2

 Const CW_USEDEFAULT = &H80000000

 Const ES_MULTILINE = &H4&
 Const ES_READONLY = &H800&

 Const WS_BORDER = &H800000
 Const WS_CHILD = &H40000000
 Const WS_OVERLAPPED = &H0&
 Const WS_CAPTION = &HC00000 ' WS_BORDER Or WS_DLGFRAME
 Const WS_SYSMENU = &H80000
 Const WS_THICKFRAME = &H40000
 Const WS_MINIMIZEBOX = &H20000
 Const WS_MAXIMIZEBOX = &H10000
 Const WS_OVERLAPPEDWINDOW = (WS_OVERLAPPED Or WS_CAPTION Or WS_SYSMENU Or WS_THICKFRAME)

 Const WS_EX_CLIENTEDGE = &H200&

 Const COLOR_WINDOW = 5

 Const WM_CLOSE = &H10
 Const WM_DESTROY = &H2
 Const WM_LBUTTONDOWN = &H201
 Const WM_LBUTTONUP = &H202

 Const IDC_ARROW = 32512&

 Const IDI_APPLICATION = 32512&
 Const IDI_HAND = 32513&
 Const IDI_QUESTION = 32514&
 Const IDI_EXCLAMATION = 32515&
 Const IDI_ASTERISK = 32516&

 Const GWL_WNDPROC = (-4)

 Const SW_SHOWNORMAL = 1
 Const SW_HIDE = 0

 Const MB_OK = &H0&
 Const MB_ICONEXCLAMATION = &H30&

' edit control constants..
 Const ES_PASSWORD = &H20&
 Const EN_SETFOCUS = &H100

 Const GWL_HINSTANCE = (-6)
 Const GWL_HWNDPARENT = (-8)
 Const GWL_STYLE = (-16)
 Const GWL_EXSTYLE = (-20)
 Const GWL_USERDATA = (-21)
 Const GWL_ID = (-12)

 Const gClassName = "WinClass"
 Const gAppName = "Enter Password"

 Const gClassName2 = "WinClass2"
 Const gAppName2 = "New Password"




Private m_ButOldProc As Long ''Will hold address of the old window proc for the button
Private m_OldProcND As Long ' Will hold address of the old window proc for the default button on the New Pasword Dialog

' Handles for the windows on the PWord Dialog Box
Private m_Hwnd As Long               ' Main Window
Private m_ButtonHwnd1 As Long        ' OK Button
Private m_ButtonHwnd2 As Long        ' Cancel Button
Private m_EditHwnd As Long           ' Edit Control

' Handles for Thw windows on the New Password Dialog
Private m_NDHwnd As Long             ' Main Window
Private m_NDButtHwndOK As Long       ' OK Button
Private m_NDButtHwndCancel As Long   ' Cancel Button
Private m_NDEditPWordHwnd As Long    ' Pword Edit
Private m_NDEditConfirmHwnd As Long  ' Confirm Edit
Private m_NDEditLbl1Hwnd As Long     ' Label for password
Private m_NDEditLbl2Hwnd As Long     ' Label for Confirm password


' Public Properties
Public sInput As String ' the users input in the Password textbox
Public NewPassWord As String
Public ConfirmPassWord As String




' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '' CODE TO CREATE AND SUBCLASS THE PASSWORD DIALOGBOX EVENTS
' ''
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Sub GetPW_Dialog()

   Dim wMsg As Msg
   
   ''Call procedure to register window classname. If false, then exit.
   If RegisterWindowClass = False Then Exit Sub
    
      ''Create window
      If CreateWindows Then
      
         ''Loop will exit when WM_QUIT is sent to the window.
         Do While GetMessage(wMsg, 0&, 0&, 0&)
            ''TranslateMessage takes keyboard messages and converts
            ''them to WM_CHAR for easier processing.
            Call TranslateMessage(wMsg)
            ''Dispatchmessage calls the default window procedure
            ''to process the window message. (WndProc)
            Call DispatchMessage(wMsg)
         Loop
      End If

    Call UnregisterClass(gClassName$, App.hInstance)


End Sub


 Function RegisterWindowClass() As Boolean

    Dim wc As WNDCLASS
    
    ''Registers our new window with windows so we
    ''can use our classname.
    
    wc.style = CS_HREDRAW Or CS_VREDRAW
    wc.lpfnwndproc = GetAddress(AddressOf WndProc) ''Address in memory of default window procedure.
    wc.hInstance = App.hInstance
    wc.hIcon = LoadIcon(0&, IDI_EXCLAMATION)  ''Default application icon
    wc.hCursor = LoadCursor(0&, IDC_ARROW) ''Default arrow
    wc.hbrBackground = COLOR_WINDOW ''Default a color for window.
    wc.lpszClassName = gClassName$

    RegisterWindowClass = RegisterClass(wc) <> 0
    
End Function


Private Function CreateWindows() As Boolean
    
    ''Create actual window.
    m_Hwnd& = CreateWindowEx(0&, gClassName$, gAppName$, WS_OVERLAPPEDWINDOW, 400, 400, 270, 115, 0&, 0&, App.hInstance, ByVal 0&)
    ''Create buttons
    m_ButtonHwnd1& = CreateWindowEx(0&, "Button", "Ok", WS_CHILD, 75, 48, 85, 25, m_Hwnd&, 0&, App.hInstance, 0&)
    m_ButtonHwnd2& = CreateWindowEx(0&, "Button", "&Cancel", WS_CHILD, 165, 48, 85, 25, m_Hwnd&, 0&, App.hInstance, 0&)
    ''Create textbox with a border (WS_EX_CLIENTEDGE) and make it multi-line (ES_MULTILINE)
    'Or ES_MULTILINE
    m_EditHwnd& = CreateWindowEx(WS_EX_CLIENTEDGE, "Edit", "", WS_CHILD Or ES_PASSWORD, 5, 10, 250, 25, m_Hwnd&, 0&, App.hInstance, 0&)

    
    ''Since windows are hidden, show them. You can use UpdateWindow to
    ''redraw the client area.
    Call ShowWindow(m_Hwnd&, SW_SHOWNORMAL)
    Call ShowWindow(m_ButtonHwnd1&, SW_SHOWNORMAL)
    Call ShowWindow(m_ButtonHwnd2&, SW_SHOWNORMAL)
    Call ShowWindow(m_EditHwnd&, SW_SHOWNORMAL)
    
    ''Get the memory address of the default window
    ''procedure for the button and store it in m_ButOldProc
    ''This will be used in OKWndProc to call the original
    ''window procedure for processing.
    m_ButOldProc& = GetWindowLong(m_ButtonHwnd1&, GWL_WNDPROC)
    
    
    ''Set default window procedure of button to OKWndProc. Different
    ''settings of windows is listed in the MSDN Library. We are using GWL_WNDPROC
    ''to set the address of the window procedure.
    Call SetWindowLong(m_ButtonHwnd1&, GWL_WNDPROC, GetAddress(AddressOf OKWndProc))
    Call SetWindowLong(m_ButtonHwnd2&, GWL_WNDPROC, GetAddress(AddressOf CancelWndProc))
    
    

    Call SetForegroundWindow(m_EditHwnd&)
    CreateWindows = (m_Hwnd& <> 0)
    
End Function




Private Function WndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  ''This our default window procedure for the window. It will handle all
  ''of our incoming window messages and we will write code based on the
  ''window message what the program should do.
  
  Dim strTemp As String
    
    Select Case uMsg&
       Case WM_DESTROY:
          ''Since DefWindowProc doesn't automatically call
          ''PostQuitMessage (WM_QUIT). We need to do it ourselves.
          ''You can use DestroyWindow to get rid of the window manually.
          Call PostQuitMessage(0&)
    End Select
    

  ''Let windows call the default window procedure since we're done.
  WndProc = DefWindowProc(hwnd&, uMsg&, wParam&, lParam&)

End Function

Private Function OKWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    
    Dim TextLen As Long, Length As String
    
    Select Case uMsg&
       Case WM_LBUTTONUP:
          
          ' **
          ' **
          ' ** One way to get the contents of the textbox. This
          ' ** method only seems effective once the window is closed.
          ' ** But works fine for the purpose of this  module.
          ' **
          ' **
          ' **
          TextLen& = GetWindowTextLength(m_EditHwnd&)
          ' make room in the buffer,
          ' allowing for the terminating null character
          sInput = Space(TextLen& + 1)
          ' read the text of the window
          Length$ = GetWindowText(m_EditHwnd&, sInput, TextLen& + 1)
          ' extract information from the buffer
          sInput = Left(sInput, Length$)
          
          If sInput = "" Then
             MsgBox "You must enter something...", vbCritical
             Call SetWindowText(m_EditHwnd&, "")
             GoTo FinishUp:
          End If
          
          If Mid(sInput$, 4, 1) = "." Then
             MsgBox "You may not use ""."" Try again.", vbCritical
             Call SetWindowText(m_EditHwnd&, "")
             GoTo FinishUp:
          End If
             
          CloseDialog
    End Select
    
FinishUp:
    
  ''Since in MyCreateWindow we made the default window proc
  ''this procedure, we have to call the old one using CallWindowProc
  OKWndProc = CallWindowProc(m_ButOldProc&, hwnd&, uMsg&, wParam&, lParam&)
   
End Function


Private Function CancelWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

    Select Case uMsg&
       Case WM_LBUTTONUP:
           sInput = "@"
           CloseDialog
    End Select
    
  ''Since in MyCreateWindow we made the default window proc
  ''this procedure, we have to call the old one using CallWindowProc
  CancelWndProc = CallWindowProc(m_ButOldProc&, hwnd&, uMsg&, wParam&, lParam&)
   
End Function


Private Function GetAddress(ByVal lngAddr As Long) As Long
    ''Used with AddressOf to return the address in memory of a procedure.

    GetAddress = lngAddr&
    
End Function

Private Sub CloseDialog()
    Call SendMessage(m_Hwnd&, WM_CLOSE, 0&, 0&)
End Sub














' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '' CODE TO CREATE AND SUBCLASS THE New PASSWORD DIALOGBOX EVENTS
' ''
' '' A VARITION OF THE ABOVE CODE.... THIS SHOWS HOW EASY IT IS TO
' '' TO CREATE WINDOWS FROM VB!
' '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Sub GetNewPWord()


    Dim wMsg As Msg
   
   ''Call procedure to register window classname. If false, then exit.
   If RegisterWindowClassND = False Then
      Exit Sub
   End If
      ''Create window
      If CreateNewPWWindows Then
      
         ''Loop will exit when WM_QUIT is sent to the window.
         Do While GetMessage(wMsg, 0&, 0&, 0&)
            ''TranslateMessage takes keyboard messages and converts
            ''them to WM_CHAR for easier processing.
            Call TranslateMessage(wMsg)
            ''Dispatchmessage calls the default window procedure
            ''to process the window message. (WndProc)
            Call DispatchMessage(wMsg)
         Loop
      End If

    Call UnregisterClass(gClassName2$, App.hInstance)


End Sub



 Function RegisterWindowClassND() As Boolean

   Dim wc As WNDCLASS
    
    ''Registers our new window with windows so we
    ''can use our classname.
    
    wc.style = CS_HREDRAW Or CS_VREDRAW
    wc.lpfnwndproc = GetAddress(AddressOf WndProc2) ''Address in memory of default window procedure.
    wc.hInstance = App.hInstance
    wc.hIcon = LoadIcon(0&, IDI_EXCLAMATION)  ''Default application icon
    wc.hCursor = LoadCursor(0&, IDC_ARROW) ''Default arrow
    wc.hbrBackground = COLOR_WINDOW ''Default a color for window.
    wc.lpszClassName = gClassName2$

    RegisterWindowClassND = RegisterClass(wc) <> 0
    
End Function

' Create the New Password Dialog
Private Function CreateNewPWWindows() As Boolean
    '
    ''Create actual window.
    m_NDHwnd& = CreateWindowEx(0&, gClassName2$, gAppName2$, WS_OVERLAPPEDWINDOW, 400, 400, 270, 180, 0&, 0&, App.hInstance, ByVal 0&)
    ''Create buttons
    m_NDButtHwndOK& = CreateWindowEx(0&, "Button", "Ok", WS_CHILD, 75, (48 + 70), 85, 25, m_NDHwnd&, 0&, App.hInstance, 0&)
    m_NDButtHwndCancel& = CreateWindowEx(0&, "Button", "&Cancel", WS_CHILD, 165, (48 + 70), 85, 25, m_NDHwnd&, 0&, App.hInstance, 0&)
    ''Create textbox with a border (WS_EX_CLIENTEDGE) and make it multi-line (ES_MULTILINE)
    'Or ES_MULTILINE
    m_NDEditPWordHwnd& = CreateWindowEx(WS_EX_CLIENTEDGE, "Edit", "", WS_CHILD Or ES_PASSWORD, 5, 30, 250, 25, m_NDHwnd&, 0&, App.hInstance, 0&)
    m_NDEditConfirmHwnd& = CreateWindowEx(WS_EX_CLIENTEDGE, "Edit", "", WS_CHILD Or ES_PASSWORD, 5, 80, 250, 25, m_NDHwnd&, 0&, App.hInstance, 0&)
    ' create a label to identify the text boxes
    m_NDEditLbl1Hwnd& = CreateWindowEx(0&, "Edit", "Enter a new Password:", WS_CHILD Or ES_READONLY, 5, 5, 250, 25, m_NDHwnd&, 0&, App.hInstance, 0&)
    m_NDEditLbl2Hwnd& = CreateWindowEx(0&, "Edit", "Confirm:", WS_CHILD Or ES_READONLY, 5, 60, 250, 20, m_NDHwnd&, 0&, App.hInstance, 0&)
    
    ' show windows.
    Call ShowWindow(m_NDHwnd&, SW_SHOWNORMAL)
    Call ShowWindow(m_NDButtHwndOK&, SW_SHOWNORMAL)
    Call ShowWindow(m_NDButtHwndCancel&, SW_SHOWNORMAL)
    Call ShowWindow(m_NDEditPWordHwnd&, SW_SHOWNORMAL)
    Call ShowWindow(m_NDEditConfirmHwnd&, SW_SHOWNORMAL)
    Call ShowWindow(m_NDEditLbl1Hwnd&, SW_SHOWNORMAL)
    Call ShowWindow(m_NDEditLbl2Hwnd&, SW_SHOWNORMAL)
    
    ''Get the memory address of the default window
    m_OldProcND = GetWindowLong(m_NDButtHwndOK&, GWL_WNDPROC)

    
    Call SetWindowLong(m_NDButtHwndOK&, GWL_WNDPROC, GetAddress(AddressOf ND_OKWndProc))
    Call SetWindowLong(m_NDButtHwndCancel&, GWL_WNDPROC, GetAddress(AddressOf ND_CancelWndProc))
    


    CreateNewPWWindows = (m_NDHwnd& <> 0)
    
End Function


Private Function WndProc2(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

  'default window procedure for the window.
  
  Dim strTemp As String
    
    Select Case uMsg&
       Case WM_DESTROY:
          Call PostQuitMessage(0&)
    End Select
    

  ''Let windows call the default window procedure since we're done.
  WndProc2 = DefWindowProc(hwnd&, uMsg&, wParam&, lParam&)

End Function


' NEW WIndow Proceduews
Private Function ND_OKWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

   Dim TextLen As Long, Length As String, RetVal As Integer
    
    Select Case uMsg&
       Case WM_LBUTTONUP:
          
          ' get the contents of the two Edit windows
          NewPassWord$ = GetText(1)
          ConfirmPassWord$ = GetText(2)
          
          ' compare.. if not the same STRINGS,...
          ' alert and try again
          RetVal% = StrComp(NewPassWord$, ConfirmPassWord$, vbTextCompare)
          
          If RetVal% <> 0 Then
              
              MsgBox "The two did not match... Try again.", vbCritical, "No Match"
              ' reset the text in the Edit controls
              Call SetWindowText(m_NDEditPWordHwnd&, "")
              Call SetWindowText(m_NDEditConfirmHwnd&, "")
              GoTo FinishUp:
              ' leave the window open for another try
          End If
          
          ' the password was correct, close and save it according
          ' to the variable values set by the programmer.
          MsgBox "Password Saved..."
          CloseDialog2
          SavePassWord NewPassWord$
    End Select
    
FinishUp:

  ND_OKWndProc = CallWindowProc(m_OldProcND&, hwnd&, uMsg&, wParam&, lParam&)
   
End Function


Private Function ND_CancelWndProc(ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long


    Select Case uMsg&
       Case WM_LBUTTONUP:
           CloseDialog2
    End Select
    
  ND_CancelWndProc = CallWindowProc(m_OldProcND&, hwnd&, uMsg&, wParam&, lParam&)
   
End Function



' close the NEW\CHANGE password dialog
Private Sub CloseDialog2()
    Call SendMessage(m_NDHwnd&, WM_CLOSE, 0&, 0&)
End Sub

Private Sub SavePassWord(PWord As String)
    Dim encPWord As String, File As String
    '
    ' first see if we need to encrypt the Password
    
    If pw.Encoded Then
       ' encrypt the password and return the encryption
       encPWord$ = Encrypt(PWord$)
       
      ' save in a file or in the registry?
      If pw.Save = 1 Then
          Dim Str As String
          ' ReGISTRY
          ' (DEMO SPECIFIC)
          ' if there is a PW.dat file present in the path dir
          ' delete...... else when the user trys to access the CheckPw
          ' Function they will read the file instead of the registry.
          ' if no file exists, the registry will be accessed. keep in
          ' mind only a pword can be saved in either\or the registry or
          ' a file. this could be altered.
          Str = Dir(pw.Path)
          If Str <> "" Then Kill pw.Path
          
          SaveSetting pw.AppName, pw.Section, pw.Key, encPWord$
      Else
          ' FILE
          File = Dir(pw.Path)
          If File <> "" Then Kill pw.Path
          Open pw.Path For Binary As #1
               Put #1, , encPWord$
          Close #1
        
      End If
    End If  ' end if Encoded
    
    If Not pw.Encoded Then
        If pw.Save = 1 Then
          ' ReGISTRY
          ' (DEMO SPECIFIC)
          ' if there is a PW.dat file present in the path dir
          ' delete...... else when the user trys to access the CheckPw
          ' Function they will read the file instead of the registry.
          ' if no file exists, the registry will be accessed. keep in
          ' mind only a pword can be saved in either\or the registry or
          ' a file. this could be altered.
          
          Str = Dir(pw.Path)
          If Str <> "" Then Kill pw.Path
          
          SaveSetting pw.AppName, pw.Section, pw.Key, PWord$
      Else
          ' FILE
          File = Dir(pw.Path)
          If File <> "" Then Kill pw.Path
          Open pw.Path For Binary As #1
               Put #1, , PWord$
          Close #1
        
      End If
    End If
    
End Sub


Private Function GetText(EditCtrl As Integer) As String
     Dim TextLen As Long, Length As String
     
     
     ' GET THE TEXT FROM THE PASSWORD AND CONFIRM
     ' EDIT CONTROLS
     Select Case EditCtrl%
         Case 1: ' New Password Text
            TextLen& = GetWindowTextLength(m_NDEditPWordHwnd&)
            ' make room in the buffer,
            ' allowing for the terminating null character
            GetText$ = Space(TextLen& + 1)
            ' read the text of the window
            Length$ = GetWindowText(m_NDEditPWordHwnd&, GetText$, TextLen& + 1)
            ' extract information from the buffer
            GetText$ = Left(GetText$, Length$)
         Case 2: ' confirm text
            TextLen& = GetWindowTextLength(m_NDEditConfirmHwnd&)
            ' make room in the buffer,
            ' allowing for the terminating null character
            GetText$ = Space(TextLen& + 1)
            ' read the text of the window
            Length$ = GetWindowText(m_NDEditConfirmHwnd&, GetText$, TextLen& + 1)
            ' extract information from the buffer
            GetText$ = Left(GetText$, Length$)
         
     End Select
End Function




' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' '' Boolean, CheckPW()
' ''
' ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function CheckPW(UserInput As String) As Boolean
     '
     
     ' This part of the code should be altered so that the
     ' stored password can be more easily and effeciently removed.
     ' Since this is a demo you have the option to save the password
     ' in the registry or in  a file. In a real world application
     ' it would be more benificial to pick one or the other.
     
     Dim PassWord As String, PWFile As String, tmpPWord As String
     
      'is there a PW.dat file in the dir
      PWFile$ = Dir(pw.Path)
      
      If PWFile$ = "" Then
         ' no dat file, must be in the registry"
         tmpPWord$ = GetSetting(pw.AppName, pw.Section, pw.Key)
      Else
         ' stored in a file.
         Open pw.Path For Binary As #1
           Do While Not EOF(1) ' Loop until end of file.
              tmpPWord$ = tmpPWord$ + Input(1, #1)
           Loop
         Close #1
      End If
      
      ' got the password...
      ' is it encrypted?
      
      ' the encryption Algorythmn I use, the first char will be
      ' reproduced with the Decimal 4 spaces from the left always
      ' so if it's there..... it's been encrypted.
      If Mid(tmpPWord$, 4, 1) = "." Then
        ' it is
         PassWord$ = Decrypt(tmpPWord$)
      Else
         PassWord$ = tmpPWord$
      End If
      
      ' check if correct
      
      Dim rv As Integer
      
      rv% = StrComp(PassWord$, UserInput$, vbTextCompare)
 'MsgBox "UserInput = " & UserInput
 'MsgBox "Password = " & PassWord
 'MsgBox "Len(UI) = " & Len(UserInput) & "  Len(PW) = " & Len(PassWord)
 'MsgBox "rv = " & rv
 
      If rv% = 0 Then
         CheckPW = True
      Else
         CheckPW = False
      End If
      
End Function



'THis function takes in a string, encodes it, then returns
' the encoded string to be saved.... etc

Public Function Encrypt(varPass As String) As String
    Dim i%

    Dim varEncrypt As String * 50, varEncStr$
    Dim varTmp As Double
    
    ' pure mathematics.... pretty simple, yet effective... huh?
    For i = 1 To Len(varPass$)
        varTmp = Asc(Mid$(varPass$, i, 1))
        varEncrypt$ = Str$(((((varTmp * 1.5) / 2.1113) * 1.111119) * i))
        varEncStr$ = varEncStr$ + varEncrypt$
    Next i
    
    Encrypt$ = varEncStr$
End Function


' it is a little more complicated to decode........ THis is a pretty good method
' but feel free to change the guts of these functions, as long as
' the parameters and returns are kept of the same value, no further
' changes to the module will be needed.


Private Function Decrypt(encPWord As String) As String
    Dim i As Integer
    Static cnt As Integer, LetCnt As Integer
    Static SpaceCnt As Integer
    
    Dim varConvert As Double, varTmpConvert As String
    Dim varFinalPass As String
    Dim strConvert As String
        
    For i% = 1 To Len(encPWord$)
      varTmpConvert$ = varTmpConvert$ & Mid$(encPWord$, i%, 1)
      ' count the characters
      cnt% = cnt% + 1
      
      ' reached the end
      If cnt% > 17 Then
         ' start counting the spaces
         SpaceCnt% = SpaceCnt% + 1
         ' reached the next # ?
         If SpaceCnt% = 34 Then
           cnt% = 0
           SpaceCnt% = 1
           varTmpConvert$ = ""
         End If
      End If
      
      ' Value found, Decode...
      If cnt% = 17 Then
        ' count the letters
        LetCnt% = LetCnt% + 1
        ' convert this value into a readable letter
        varConvert = Val(Trim(varTmpConvert$))
        varConvert = ((((varConvert / 1.5) * 2.1113) / 1.111119) / LetCnt)
        ' add 'em up
        varFinalPass$ = varFinalPass$ & Chr(varConvert)
      End If
    Next i%
    
    ' return decoded string
    Decrypt$ = varFinalPass$
    
End Function




