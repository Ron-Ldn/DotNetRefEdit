# DotNetRefEdit
Examples of RefEdit like controls for Excel add-ins using C# and ExcelDna

Overview
--------
This project is a proof of concept. This is not a library to be reused in other solutions. 

The purpose is to show how to build a .Net UI using WinForm or WPF, within an Excel add-in, where the user can select a range in Excel in order to see the range address appear in the UI control.

Illustration
--------

- Step 1: open a form and focus to a "RefEdit" control

![Prepare Selection](https://raw.github.com/Ron-Ldn/DotNetRefEdit/master/Screenshots/RefEditUI.png)

- Step 2: select a range into Excel

![Select Range](https://raw.github.com/Ron-Ldn/DotNetRefEdit/master/Screenshots/ExcelSelection.png)

- Result: the range address is populated automatically into the "RefEdit" control into the form

![Populate Range Address](https://raw.github.com/Ron-Ldn/DotNetRefEdit/master/Screenshots/RefEditUI2.png)

Inventory
--------
Several projects can be found on the internet which propose to implement the equivalent of the RefEdit control for .Net programs. But as far as I know, none of these projects show how to manage the window itself. Here is a list of issues I faced in the past:
- If the UI runs in the Excel thread, then it will freeze when Excel is busy.
- If the UI runs in the Excel thread, then in some conditions it may be impossible for the user to manually edit a control within the UI because Excel will activate and put the focus to the last selected cell. The conditions to reproduce this are quite unclear to me though.
- If the UI runs in its own thread, then the user will need to click twice into Excel in order to select a range. Actually, the first click will activate the window and then Excel will discard the message. The second click is to select the range into the activated window.

Links:
- http://blogs.msdn.com/b/gabhan_berry/archive/2008/06/12/net-refedit-control.aspx
- http://www.codeproject.com/Articles/32805/RefEdit-Emulation-for-NET
- http://www.codeproject.com/Articles/34425/VS-NET-Excel-Addin-Refedit-Control

Solution
--------
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

Bonus: using the same hook, it is possible to notify the UI when the user clicks on the same cell as before. In fact, the UI is notified of a new selection by the Excel event "SheetSelectionChange", but this event is not triggered if the user points to the same cell again. By adding a special notification to the UI inside the hook method, it is possible to resolve that issue.

Code
--------
The solution available along with this project proposes 4 examples, all accessible from the Excel ribbon. These examples have been tested with Excel 2010 32bit, Excel 2013 32bit and Excel 2013 64bit, running on Windows 7 64bit. 

![Ribbon](https://raw.github.com/Ron-Ldn/DotNetRefEdit/master/Screenshots/Ribbon.png)

The "Excel Thread" buttons launch a WinForm and a WPF window running in the Excel main thread. The "Separate Thread" buttons launch the same UIs in their own threads.

When the UI is launched, it will subscribe to the "SheetSelectionChange" event and hook the WH_CALLWNDPROC messages. For more details on how to hook Windows messages, please refer to https://msdn.microsoft.com/en-us/library/windows/desktop/ms644959%28v=vs.85%29.aspx
