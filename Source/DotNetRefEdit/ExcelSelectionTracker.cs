using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace DotNetRefEdit
{
    public class RangeAddressEventArgs : EventArgs
    {
        public string Address { get; set; }    
    }

    public class ExcelSelectionTracker
    {
        private readonly Application _application;

        public event EventHandler<RangeAddressEventArgs> NewSelection;

        private readonly int _hHookCwp;
        private readonly WindowsInterop.HookProc _procCwp; // Note: do not make this delegate a local variable within the ExcelSelectionTracker constructor because it must not be collected by the GC before the unhook
        
        public ExcelSelectionTracker(int excelThreadId)
        {
            _application = (Application)ExcelDnaUtil.Application;
            _application.SheetSelectionChange += OnNewSelection;

            _procCwp = CwpProc;

            _hHookCwp = WindowsInterop.SetWindowsHookEx(HookType.WH_CALLWNDPROC, _procCwp, (IntPtr)0, excelThreadId);
            if (_hHookCwp == 0)
            {
                throw new Exception("Failed to hook WH_CALLWNDPROC");
            }
        }

        public void Stop()
        {
            _application.SheetSelectionChange -= OnNewSelection;

            if (!WindowsInterop.UnhookWindowsHookEx(_hHookCwp))
            {
                Debug.Print("Error: Failed to unhook WH_CALLWNDPROC");
            }
        }

        private int CwpProc(int nCode, IntPtr wParam, IntPtr lParam)
        {
            CwpStruct cwpStruct = (CwpStruct)Marshal.PtrToStructure(lParam, typeof(CwpStruct));

            if (nCode < 0)
            {
                return WindowsInterop.CallNextHookEx(_hHookCwp, nCode, wParam, lParam);
            }

            if (cwpStruct.message == WindowsInterop.WM_MOUSEACTIVATE)
            {
                // We got a WM_MOUSEACTIVATE message. Now we will check that the target handle is a workbook window.
                // Workbook windows have the name "EXCEL7".
                bool isWorkbookWindow = false;

                try
                {
                    StringBuilder cname = new StringBuilder(256);
                    WindowsInterop.GetClassNameW(cwpStruct.hwnd, cname, cname.Capacity);
                    if (cname.ToString() == "EXCEL7")
                    {
                        isWorkbookWindow = true;
                    }
                }
                catch (Exception e)
                {
                    Debug.Print("Could not get the window name: {0}", e);
                }

                if (isWorkbookWindow)
                {
                    // If the window is not activated, then Excel will activate it and then discard the message. That's why the user cannot select a range at the same time.
                    // The following statement will activate the window before Excel treats the message, thus it will not activate the window and it will keep proceeding the message. 
                    // In that way, it is possible to select the range.
                    try
                    {
                        WindowsInterop.SetFocus(cwpStruct.hwnd);
                    }
                    catch (Exception e)
                    {
                        Debug.Print("Failed to set the focus: {0}", e);
                    }

                    // If the user chooses a cell which was already selected, then the event SheetSelectionChange will not be raised.
                    // A workaround is to send the current selection when the Excel window gets the focus. 
                    // Note that if the user selects a different range, then 2 events will be raised: a first one with the current selection, 
                    // and a second one with the new selection.
                    try
                    {
                        OnNewSelection(null, (Range) _application.Selection);
                    }
                    catch
                    {
                    }
                }
            }

            return WindowsInterop.CallNextHookEx(_hHookCwp, nCode, wParam, lParam);
        }

        private void OnNewSelection(object sh, Range target)
        {
            try
            {
                var newSelection = NewSelection;
                if (newSelection != null)
                {
                    newSelection(this, new RangeAddressEventArgs { Address = target.Address[false, false, XlReferenceStyle.xlA1, true] });
                }
            }
            catch
            {
            }
        }
    }
}
