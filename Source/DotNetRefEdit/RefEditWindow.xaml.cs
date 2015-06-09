using System;
using System.Windows;
using System.Windows.Input;
using ExcelDna.Integration;
using Application = Microsoft.Office.Interop.Excel.Application;
using TextBox = System.Windows.Controls.TextBox;

namespace DotNetRefEdit
{
    public partial class RefEditWindow
    {
        private readonly ExcelSelectionTracker _selectionTracker;
        private readonly Application _application;
        private TextBox _focusedBox;

        public RefEditWindow(int excelThreadId)
        {
            InitializeComponent();

            _selectionTracker = new ExcelSelectionTracker(excelThreadId);
            _application = (Application)ExcelDnaUtil.Application;

            Closed += delegate
            {
                _selectionTracker.NewSelection -= ChangeText;
                _selectionTracker.Stop();
            };

            _selectionTracker.NewSelection += ChangeText;

            Deactivated += CheckFocus;

            InputBox1.TextChanged += OnNewInput;
            InputBox2.TextChanged += OnNewInput;

            InputBox1.KeyDown += CheckF4;
            InputBox2.KeyDown += CheckF4;
            DestinationBox.KeyDown += CheckF4;
        }

        private void CheckFocus(object sender, EventArgs eventArgs)
        {
            if (InputBox1.IsFocused)
            {
                _focusedBox = InputBox1;
            }
            else if (InputBox2.IsFocused)
            {
                _focusedBox = InputBox2;
            }
            else if (DestinationBox.IsFocused)
            {
                _focusedBox = DestinationBox;
            }
            else
            {
                _focusedBox = null;
            }
        }

        private void ChangeText(object sender, RangeAddressEventArgs args)
        {
            Dispatcher.Invoke(new Action(() =>
            {
                if (_focusedBox != null)
                {
                    _focusedBox.Text = args.Address;
                    _focusedBox.CaretIndex = _focusedBox.Text.Length;
                }
            }));
        }

        private void CheckF4(object sender, KeyEventArgs e)
        {
            TextBox textBox = sender as TextBox;
            if (e.Key == Key.F4 && textBox != null)
            {
                string text = textBox.Text;

                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    string newAddress;
                    if (ExcelHelper.TryF4(text, _application, out newAddress))
                    {
                        Dispatcher.Invoke(new Action(() =>
                        {
                            textBox.Text = newAddress;
                            textBox.CaretIndex = textBox.Text.Length;
                        }));
                    }
                });
            }
        }

        /// <summary>
        /// Build final formula: to be run in UI thread
        /// </summary>
        /// <returns></returns>
        private string BuildFormula()
        {
            return string.Format("=sum({0},{1})", InputBox1.Text, InputBox2.Text);
        }

        /// <summary>
        /// Evaluate the formula: to be run in Excel thread
        /// </summary>
        private void UpdateEvaluation(string formula)
        {
            object formulaResult = ExcelHelper.EvaluateFormula(formula, _application);
            Dispatcher.Invoke(new Action(() => EvaluationBox.Text = (formulaResult ?? "").ToString()));
        }

        private void OnNewInput(object sender, EventArgs e)
        {
            string formula = BuildFormula();
            ExcelAsyncUtil.QueueAsMacro(() => UpdateEvaluation(formula));
        }

        private void InsertFormula(object sender, RoutedEventArgs e)
        {
            string formula = BuildFormula();
            string destination = DestinationBox.Text;

            if (formula != null && !string.IsNullOrEmpty(destination))
            {
                ExcelAsyncUtil.QueueAsMacro(() => ExcelHelper.InsertFormula(formula, _application, destination));
            }
        }
    }
}
