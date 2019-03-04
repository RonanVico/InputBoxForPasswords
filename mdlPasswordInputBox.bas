Attribute VB_Name = "mdlPasswordInputBox"
Option Explicit


Private Const EM_SETPASSWORDCHAR = &HCC
Private Const WH_CBT = 5
Private Const HCBT_ACTIVATE = 5
Private Const HC_ACTION = 0

#If VBA7 Then
    #If Win64 Then
        Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongLong
        Private Declare PtrSafe Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As LongPtr, ByVal hmod As LongPtr, ByVal dwThreadId As Long) As LongPtr
        Private Declare PtrSafe Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As LongPtr) As Long
        Private Declare PtrSafe Function SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageA" (ByVal hDlg As LongPtr, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As LongPtr, ByVal lParam As LongPtr) As LongPtr
        Private Declare PtrSafe Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As LongPtr, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
        Private Declare PtrSafe Function GetCurrentThreadId Lib "kernel32" () As Long
    #Else
        Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
        Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As LongPtr
        Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
        Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
        Private Declare Function SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
        Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
        Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
    #End If
#Else
    'API functions to be used
    Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
    Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
    Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
    Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long
    Private Declare Function SendDlgItemMessage Lib "user32" Alias "SendDlgItemMessageA" (ByVal hDlg As Long, ByVal nIDDlgItem As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
    Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long
#End If

#If VBA7 Then
    #If Win64 Then
        Private hHook As LongLong
    #Else
        Private hHook As LongPtr
    #End If
#Else
    Private hHook As Long
#End If


#If VBA7 Then
    #If Win64 Then
        Public Function NewProc(ByVal lngCode As Long, ByVal wParam As Long, ByVal lParam As Long) As LongLong
            Dim RetVal
            Dim strClassName As String, lngBuffer As Long
        
            If lngCode < HC_ACTION Then
                NewProc = CallNextHookEx(hHook, lngCode, wParam, lParam)
                Exit Function
            End If
        
            strClassName = String$(256, " ")
            lngBuffer = 255
        
            If lngCode = HCBT_ACTIVATE Then
        
                RetVal = GetClassName(wParam, strClassName, lngBuffer)
        
                If Left$(strClassName, RetVal) = "#32770" Then
        
                    SendDlgItemMessage wParam, &H1324, EM_SETPASSWORDCHAR, Asc("*"), &H0
                End If
            End If
        
            CallNextHookEx hHook, lngCode, wParam, lParam
        End Function
    #Else
        Public Function NewProc(ByVal lngCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
            Dim RetVal
            Dim strClassName As String, lngBuffer As Long
        
            If lngCode < HC_ACTION Then
                NewProc = CallNextHookEx(hHook, lngCode, wParam, lParam)
                Exit Function
            End If
        
            strClassName = String$(256, " ")
            lngBuffer = 255
        
            If lngCode = HCBT_ACTIVATE Then
        
                RetVal = GetClassName(wParam, strClassName, lngBuffer)
        
                If Left$(strClassName, RetVal) = "#32770" Then
        
                    SendDlgItemMessage wParam, &H1324, EM_SETPASSWORDCHAR, Asc("*"), &H0
                End If
            End If
        
            CallNextHookEx hHook, lngCode, wParam, lParam
        End Function
    #End If
#Else
        Public Function NewProc(ByVal lngCode As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
            Dim RetVal
            Dim strClassName As String, lngBuffer As Long
        
            If lngCode < HC_ACTION Then
                NewProc = CallNextHookEx(hHook, lngCode, wParam, lParam)
                Exit Function
            End If
        
            strClassName = String$(256, " ")
            lngBuffer = 255
        
            If lngCode = HCBT_ACTIVATE Then
        
                RetVal = GetClassName(wParam, strClassName, lngBuffer)
        
                If Left$(strClassName, RetVal) = "#32770" Then
        
                    SendDlgItemMessage wParam, &H1324, EM_SETPASSWORDCHAR, Asc("*"), &H0
                End If
            End If
        
            CallNextHookEx hHook, lngCode, wParam, lParam
        
        End Function
#End If
#If VBA7 Then
    #If Win64 Then
        Public Function InputBoxDK(Prompt, Optional Title, Optional Default, Optional XPos, _
                                   Optional YPos, Optional HelpFile, Optional Context) As String
            Dim lngModHwnd As LongLong, lngThreadID As Long
        
            lngThreadID = GetCurrentThreadId
            lngModHwnd = GetModuleHandle(vbNullString)
        
            hHook = SetWindowsHookEx(WH_CBT, AddressOf NewProc, lngModHwnd, lngThreadID)
        
            InputBoxDK = InputBox(Prompt, Title, Default, XPos, YPos, HelpFile, Context)
            UnhookWindowsHookEx hHook
        
        End Function
    #Else
        Public Function InputBoxDK(Prompt, Optional Title, Optional Default, Optional XPos, _
                                   Optional YPos, Optional HelpFile, Optional Context) As String
            Dim lngModHwnd As Long, lngThreadID As Long
        
            lngThreadID = GetCurrentThreadId
            lngModHwnd = GetModuleHandle(vbNullString)
        
            hHook = SetWindowsHookEx(WH_CBT, AddressOf NewProc, lngModHwnd, lngThreadID)
        
            InputBoxDK = InputBox(Prompt, Title, Default, XPos, YPos, HelpFile, Context)
            UnhookWindowsHookEx hHook
        
        End Function
    #End If
#Else
      'API functions to be used
  Public Function InputBoxDK(Prompt, Optional Title, Optional Default, Optional XPos, _
                             Optional YPos, Optional HelpFile, Optional Context) As String
      Dim lngModHwnd As Long, lngThreadID As Long
  
      lngThreadID = GetCurrentThreadId
      lngModHwnd = GetModuleHandle(vbNullString)
  
      hHook = SetWindowsHookEx(WH_CBT, AddressOf NewProc, lngModHwnd, lngThreadID)
  
      InputBoxDK = InputBox(Prompt, Title, Default, XPos, YPos, HelpFile, Context)
      UnhookWindowsHookEx hHook
  
  End Function
#End If



