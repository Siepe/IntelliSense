using ExcelDna.Integration;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.IO;
using System.Linq;
using System.Windows.Forms;

namespace ExcelDna.IntelliSense
{
    // CONSIDER: Maybe some ideas from here: http://codereview.stackexchange.com/questions/55916/lightweight-rich-link-label

    // TODO: Drop shadow: http://stackoverflow.com/questions/16493698/drop-shadow-on-a-borderless-winform
    class SelectBoxForm : Form
    {
        Win32Window _owner;
        ComboBox comboBox;
        // Help Link
        public SelectBoxForm(IntPtr hwndOwner)
        {
            Debug.Assert(hwndOwner != IntPtr.Zero);
            InitializeComponent();
            _owner = new Win32Window(hwndOwner);
            // CONSIDER: Maybe make a more general solution that lazy-loads as needed
        }

        void ShowToolTip()
        {
            try
            {
                Show(_owner);
            }
            catch (Exception e)
            {
                Debug.Write("ToolTipForm.Show error " + e);
            }
        }

        public void ShowToolTip(string text, string linePrefix, int left, int top, int topOffset, int? listLeft = null)
        {
            if (!Visible)
            {
                this.Size = new Size(200, 22);
                this.Top = top + topOffset;
                this.Left = left;
                var argDelim = ";ArgList - ";
                var argList = text.Substring(text.IndexOf(argDelim) + argDelim.Length);
                this.comboBox.Items.AddRange(argList.Split(','));
                ShowToolTip();
            }
            else
            {
                Invalidate();
            }
        }

        public IntPtr OwnerHandle
        {
            get
            {
                if (_owner == null)
                    return IntPtr.Zero;
                return _owner.Handle;
            }
            set
            {
                if (_owner == null || _owner.Handle != value)
                {
                    _owner = new Win32Window(value);
                    if (Visible)
                    {
                        // We want to change the owner.
                        // That's hard, so we hide and re-show.
                        Hide();
                        ShowToolTip();
                    }
                }
            }
        }

        protected override bool ShowWithoutActivation => true;

        void InitializeComponent()
        {
            this.MinimumSize = new Size(1, 1);
            //this.AutoSizeMode = AutoSizeMode.GrowAndShrink;
            this.FormBorderStyle = FormBorderStyle.None;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.ControlBox = false;
            this.DoubleBuffered = true;
            this.ShowInTaskbar = false;
            this.StartPosition = FormStartPosition.Manual;
            this.comboBox = new ComboBox();
            comboBox.DropDownWidth = 280;
            comboBox.Size = new Size(200, 24);
            comboBox.SelectionChangeCommitted += SelectedValueChanged;

            this.Controls.Add(comboBox);
        }

        private void SelectedValueChanged(object sender, EventArgs e)
        {
            var val = (sender as ComboBox).SelectedItem.ToString();
            SendKeys.Send("\"");
            SendKeys.Send(val);
            SendKeys.Send("\"");
            Dispose();
        }
    }
}
