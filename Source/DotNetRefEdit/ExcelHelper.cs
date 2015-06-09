using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace DotNetRefEdit
{
    /// <summary>
    /// Common functions used by the form and the WPF window
    /// </summary>
    static class ExcelHelper
    {
        /// <summary>
        /// Make Excel evaluate the formula.
        /// To be run in Excel thread.
        /// </summary>
        /// <param name="formula"></param>
        /// <param name="application"></param>
        /// <returns></returns>
        public static object EvaluateFormula(string formula, Application application)
        {
            try
            {
                object formulaResult = application.Evaluate(formula);

                // Check the Excel error codes
                if (formulaResult is int)
                {
                    switch ((int)formulaResult)
                    {
                        case -2146826288:
                        case -2146826281:
                        case -2146826265:
                        case -2146826259:
                        case -2146826252:
                        case -2146826246:
                        case -2146826273:
                            return "Could not evaluate function";
                    }
                }

                return formulaResult;
            }
            catch
            {
                return "Could not evaluate function";
            }
        }

        /// <summary>
        /// Insert formula into Excel range.
        /// To be run in Excel thread.
        /// </summary>
        /// <param name="formula"></param>
        /// <param name="application"></param>
        /// <param name="destination"></param>
        public static void InsertFormula(string formula, Application application, string destination)
        {
            Range rg = null;

            try
            {
                rg = application.Range[destination];
                rg.Formula = formula;
            }
            finally
            {
                if (rg != null)
                {
                    Marshal.ReleaseComObject(rg);
                }
            }
        }

        /// <summary>
        /// Try to switch the address format, following this sequence:
        /// 1. RowAbsolute=False, ColumnAbsolute=False
        /// 2. RowAbsolute=True, ColumnAbsolute=True
        /// 3. RowAbsolute=True, ColumnAbsolute=False
        /// 4. RowAbsolute=False, ColumnAbsolute=True
        /// This shall reproduce the behaviour of the Excel "Function Arguments" form when the user hits F4.
        /// To be run in Excel thread.
        /// </summary>
        /// <param name="text"></param>
        /// <param name="application"></param>
        /// <param name="newAddress"></param>
        /// <returns></returns>
        public static bool TryF4(string text, Application application, out string newAddress)
        {
            try
            {
                object formulaResult = application.Evaluate(text);

                if (formulaResult is Range)
                {
                    string relativePart = text;

                    if (text.Contains("!"))
                    {
                        relativePart = text.Substring(text.IndexOf("!") + 1, text.Length - text.IndexOf("!") - 1);
                    }

                    Range range = (Range) formulaResult;

                    List<string> addresses = new List<string>
                    {
                        range.Address[false, false, XlReferenceStyle.xlA1, false],
                        range.Address[true, true, XlReferenceStyle.xlA1, false],
                        range.Address[true, false, XlReferenceStyle.xlA1, false],
                        range.Address[false, true, XlReferenceStyle.xlA1, false]
                    };

                    bool found = false;
                    for (int i = 0; i < addresses.Count; i++)
                    {
                        if (addresses[i] == relativePart)
                        {
                            relativePart = addresses[i + 1 == addresses.Count ? 0 : i + 1];
                            found = true;
                            break;
                        }
                    }

                    if (!found)
                    {
                        newAddress = range.Address[false, false, XlReferenceStyle.xlA1, true];
                        return true;
                    }

                    newAddress = text.Contains("!")
                        ? string.Concat(text.Substring(0, text.IndexOf("!") + 1), relativePart)
                        : relativePart;

                    return true;
                }

                newAddress = null;
                return false;
            }
            catch
            {
                newAddress = null;
                return false;
            }
        }
    }
}
