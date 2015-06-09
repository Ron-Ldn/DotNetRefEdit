using System;
using System.Windows.Forms;
using ExcelDna.Integration;
using Application = Microsoft.Office.Interop.Excel.Application;

namespace DotNetRefEdit
{
    public partial class RefEditForm : Form
    {
        private readonly ExcelSelectionTracker _selectionTracker;
        private readonly Application _application;
        private RichTextBox _focusedBox;

        public RefEditForm(int excelThreadId)
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

            Deactivate += CheckFocus;

            InputBox1.TextChanged += OnNewInput;
            InputBox2.TextChanged += OnNewInput;

            InputBox1.KeyDown += CheckF4;
            InputBox2.KeyDown += CheckF4;
            DestinationBox.KeyDown += CheckF4;
        }

        private void CheckFocus(object sender, EventArgs eventArgs)
        {
            if (InputBox1.Focused)
            {
                _focusedBox = InputBox1;
            }
            else if (InputBox2.Focused)
            {
                _focusedBox = InputBox2;
            }
            else if (DestinationBox.Focused)
            {
                _focusedBox = DestinationBox;
            }
            else
            {
                _focusedBox = null;
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
            Invoke(new Action(() => EvaluationBox.Text = (formulaResult ?? "").ToString()));
        }

        private void OnNewInput(object sender, EventArgs e)
        {
            string formula = BuildFormula();
            ExcelAsyncUtil.QueueAsMacro(() => UpdateEvaluation(formula));
        }

        private void ChangeText(object sender, RangeAddressEventArgs args)
        {
            Invoke(new Action(() =>
            {
                if (_focusedBox != null)
                {
                    _focusedBox.Text = args.Address;
                    _focusedBox.Select(_focusedBox.Text.Length, 0);
                }
            }));
        }

        private void CheckF4(object sender, KeyEventArgs e)
        {
            RichTextBox textBox = sender as RichTextBox;
            if (e.KeyCode == Keys.F4 && textBox != null)
            {
                string text = textBox.Text;

                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    string newAddress;
                    if (ExcelHelper.TryF4(text, _application, out newAddress))
                    {
                        Invoke(new Action(() =>
                        {
                            textBox.Text = newAddress;
                            textBox.Select(textBox.Text.Length, 0);
                        }));
                    }
                });
            }
        }

        private void InsertButton_Click(object sender, EventArgs e)
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
