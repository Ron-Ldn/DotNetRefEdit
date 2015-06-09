# DotNetRefEdit
Examples of RefEdit like controls for Excel add-ins using C# and ExcelDna

NOTE
--------
UNDER CONSTRUCTION !

Overview
--------
This project is a proof of concept. This is not a library to be reused in other solutions. 

The purpose is to show how to build a .Net UI using WinForm or WPF, within an Excel add-in, where the user can select a range in Excel in order to see the range address appear in the UI control.

Several projects can be found on the internet which propose to implement the equivalent of the RefEdit control for .Net programs. But as far as I know, none of these projects show how to manage the window itself. Here is a list of issues I faced in the past:
- If the UI runs in the Excel thread, then it will freeze when Excel is busy.
- If the UI runs in the Excel thread, then in some conditions it may be impossible for the user to manually edit a control within the UI because Excel will activate and put the focus to the last selected cell. The conditions to reproduce this are quite unclear to me though.
- If the UI runs in its own thread, then the user will need to click twice into Excel in order to select a range. Actually, the first click will activate the window, and Excel will discard the message. The second click is to select the range into the activated window.

The best option, according to me, is to run the UI in its own thread. The "must-click-twice" issue described above can be resolved by hooking the WH_CALLWNDPROC messages: if the message is of type "WM_MOUSEACTIVATE" and if the handle is an Excel workbook window (in which case, the window name will be "EXCEL7") then it is possible to set the focus to that window before Excel processes the message. 

1. Hook
  ```C#
  _hHookCwp = SetWindowsHookEx(HookType.WH_CALLWNDPROC, _procCwp, (IntPtr)0, excelThreadId);
  ```

2. Check the class name
  ```C#
  GetClassNameW(cwpStruct.hwnd, cname, cname.Capacity);
  if (cname.ToString() == "EXCEL7")
  ```

3. Set the focus
  ```C#
  SetFocus(cwpStruct.hwnd);
  ```

Bonus: using the same hook, it is possible to notify the UI when the user clicks on the same cell as before. In fact, the UI is notified of a new selection by the Excel event "SheetSelectionChange", but this event is not triggered if the user takes the same cell again. By adding a special notification to the UI inside the hook method, it is possible to resolve that issue.
