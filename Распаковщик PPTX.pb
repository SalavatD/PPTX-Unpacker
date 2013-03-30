Enumeration
  #MAIN_WINDOW
EndEnumeration

Enumeration
  #MENU_BAR
EndEnumeration

Enumeration
  #ACTION_EXIT
  #ACTION_ABOUT
EndEnumeration

Enumeration
  #PPTX_UNPACK_BUTTON
EndEnumeration

Procedure OpenMainWindow()
  
  If OpenWindow(#MAIN_WINDOW, #PB_Ignore, #PB_Ignore, 270, 125, "����������� PPTX", #PB_Window_MinimizeGadget | #PB_Window_ScreenCentered)
    If CreateMenu(#MENU_BAR, WindowID(#MAIN_WINDOW)) 
      MenuTitle("&����")
      MenuItem(#ACTION_EXIT, "�&����")
      
      MenuTitle("&�������")
      MenuItem(#ACTION_ABOUT, "&� ���������")
    EndIf
    
    ButtonGadget(#PPTX_UNPACK_BUTTON, 70, 40, 130, 25, "����������� PPTX")
    
    SetWindowColor(#MAIN_WINDOW, $FFFFFF)
    
    SetGadgetColor(#PPTX_UNPACK_BUTTON, #PB_Gadget_BackColor, $FFFFFF)
  EndIf
EndProcedure

Procedure UnpackPPTX()
  archiveFilePath.s = OpenFileRequester("", "", "����������� PPTX (*.pptx)|*.pptx", 0)
  If archiveFilePath <> ""
    folder.s = ReplaceString(GetFilePart(archiveFilePath), ".pptx", "")
    If PureZIP_ExtractFiles(archiveFilePath, "*.*", GetPathPart(archiveFilePath) + folder, #True)
      MessageRequester("������", "����������� �����������", #MB_ICONINFORMATION)
    Else
      MessageRequester("������", "����������� �� �����������", #MB_ICONINFORMATION)
    EndIf
  EndIf
EndProcedure

OpenMainWindow()

Repeat
  event       = WaitWindowEvent()
  eventMenu   = EventMenu()
  eventGadget = EventGadget()
  eventWindow = EventWindow()
  eventType   = EventType()
  
  If eventWindow = #MAIN_WINDOW
    If event = #PB_Event_Menu
      If eventMenu = #ACTION_EXIT
        Break
      ElseIf eventMenu = #ACTION_ABOUT
        MessageRequester("� ���������", "����������� PPTX. ������ 1.0" + #CR$ + #CR$ + "�����: ������� ������" + #CR$ + #CR$ + "���� ��������: ���� 2013", #MB_ICONINFORMATION)
      EndIf
    EndIf
    
    If event = #PB_Event_Gadget
      If eventGadget = #PPTX_UNPACK_BUTTON
        UnpackPPTX()
      EndIf
    EndIf
  EndIf
Until event = #PB_Event_CloseWindow And eventWindow = #MAIN_WINDOW

; IDE Options = PureBasic 4.51 (Windows - x86)
; CursorPosition = 72
; FirstLine = 31
; Folding = -
; EnableXP
; UseIcon = Icon.ico
; Executable = ����������� PPTX.exe
; DisableDebugger
; IncludeVersionInfo
; VersionField0 = 1.0.0.0
; VersionField1 = 1.0.0.0
; VersionField2 = ������� ������
; VersionField3 = ����������� PPTX
; VersionField4 = 1.0
; VersionField5 = 1.0
; VersionField6 = ����������� PPTX
; VersionField7 = ����������� PPTX
; VersionField8 = ����������� PPTX.exe
; VersionField17 = 0419 Russian
