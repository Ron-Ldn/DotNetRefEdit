using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace DotNetRefEdit
{
    [ComVisible(true)]
    public class MyRibbon : ExcelRibbon
    {
        private readonly int _excelThreadId;

        private RefEditForm _refEditForm1;
        private RefEditWindow _refEditWindow1;
        private RefEditForm _refEditForm2;
        private RefEditWindow _refEditWindow2;

        public MyRibbon()
        {
            _excelThreadId = WindowsInterop.GetCurrentThreadId();
        }

        private bool CheckWorkbook()
        {
            try
            {
                Application app = (Application) ExcelDnaUtil.Application;
                if (app.Workbooks.Count == 0)
                {
                    MessageBox.Show("Please open a workbook before starting UI.", "Error");
                    return false;
                }

                return true;
            }
            catch (Exception e)
            {
                Debug.Print("Couldn't check workbook: {0}", e);
                return false;
            }
        }

        public void OpenWinFormInExcelThread(IRibbonControl control)
        {
            if (!CheckWorkbook())
            {
                return;
            }

            if (_refEditForm1 == null)
            {
                try
                {
                    _refEditForm1 = new RefEditForm(_excelThreadId);
                    _refEditForm1.Closed += delegate { _refEditForm1 = null; };
                    _refEditForm1.Show();
                }
                catch (Exception e)
                {
                    Debug.Print("Error: {0}", e);
                }
            }
            else
            {
                _refEditForm1.Activate();
            }
        }

        public void OpenWinFormInSeparateThread(IRibbonControl control)
        {
            if (!CheckWorkbook())
            {
                return;
            }

            if (_refEditForm2 == null)
            {
                Thread thread = new Thread(() =>
                {
                    try
                    {
                        _refEditForm2 = new RefEditForm(_excelThreadId);
                        _refEditForm2.Closed += delegate { _refEditForm2 = null; };
                        _refEditForm2.ShowDialog();
                    }
                    catch (Exception e)
                    {
                        Debug.Print("Error: {0}", e);
                    }
                });

                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
            }
            else
            {
                _refEditForm2.Invoke(new Action(() => _refEditForm2.Activate()));
            }
        }

        public void OpenWPFInExcelThread(IRibbonControl control)
        {
            if (!CheckWorkbook())
            {
                return;
            }

            if (_refEditWindow1 == null)
            {
                try
                {
                    _refEditWindow1 = new RefEditWindow(_excelThreadId);
                    _refEditWindow1.Closed += delegate { _refEditWindow1 = null; };
                    _refEditWindow1.Show();
                }
                catch (Exception e)
                {
                    Debug.Print("Error: {0}", e);
                }
            }
            else
            {
                _refEditWindow1.Activate();
            }
        }

        public void OpenWPFInSeparateThread(IRibbonControl control)
        {
            if (!CheckWorkbook())
            {
                return;
            }

            if (_refEditWindow2 == null)
            {
                Thread thread = new Thread(() =>
                {
                    try
                    {
                        _refEditWindow2 = new RefEditWindow(_excelThreadId);
                        _refEditWindow2.Closed += delegate { _refEditWindow2 = null; };
                        _refEditWindow2.ShowDialog();
                    }
                    catch (Exception e)
                    {
                        Debug.Print("Error: {0}", e);
                    }
                });

                thread.SetApartmentState(ApartmentState.STA);
                thread.Start();
            }
            else
            {
                _refEditWindow2.Dispatcher.Invoke(new Action(() => _refEditWindow2.Activate()));
            }
        }

        public override string GetCustomUI(string uiName)
        {
            return @"<customUI xmlns='http://schemas.microsoft.com/office/2006/01/customui'>
  <ribbon>
    <tabs>
      <tab id='CustomExcelAddInTab' label='DotNetRefEdit'>
        <group id='ExcelThreadGroup' label='Excel Thread'>
          <button id='Button1' label='WinForm' onAction='OpenWinFormInExcelThread'/>
          <button id='Button3' label='WPF' onAction='OpenWPFInExcelThread'/>
        </group>
        <group id='SeparateThreadGroup' label='Separate Thread'>
          <button id='Button2' label='WinForm' onAction='OpenWinFormInSeparateThread'/>
          <button id='Button4' label='WPF' onAction='OpenWPFInSeparateThread'/>
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>";
        }
    }
}
