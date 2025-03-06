using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing;

namespace ExcelTemplateManager
{
    public class ValidationOptionsForm : Form
    {
        private TextBox txtOption;
        private Button btnAdd;
        private Button btnRemove;
        private ListBox listOptions;
        private Button btnOK;
        private Button btnCancel;

        public List<string> Options { get; private set; }

        public ValidationOptionsForm(List<string> existingOptions = null)
        {
            Options = existingOptions != null ? new List<string>(existingOptions) : new List<string>();
            InitializeComponents();
            LoadExistingOptions();
        }

        private void InitializeComponents()
        {
            this.Text = "Opzioni di Validazione";
            this.Size = new Size(400, 400);
            this.StartPosition = FormStartPosition.CenterParent;
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            Label lblInstruction = new Label
            {
                Text = "Aggiungi opzioni per il menu a tendina nella colonna:",
                Location = new Point(10, 10),
                Size = new Size(360, 20)
            };

            txtOption = new TextBox
            {
                Location = new Point(10, 40),
                Size = new Size(260, 23)
            };

            btnAdd = new Button
            {
                Text = "Aggiungi",
                Location = new Point(280, 39),
                Size = new Size(90, 25)
            };
            btnAdd.Click += BtnAdd_Click;

            listOptions = new ListBox
            {
                Location = new Point(10, 80),
                Size = new Size(360, 220),
                SelectionMode = SelectionMode.One
            };

            btnRemove = new Button
            {
                Text = "Rimuovi",
                Location = new Point(10, 310),
                Size = new Size(90, 25)
            };
            btnRemove.Click += BtnRemove_Click;

            btnOK = new Button
            {
                Text = "OK",
                DialogResult = DialogResult.OK,
                Location = new Point(200, 310),
                Size = new Size(80, 25)
            };

            btnCancel = new Button
            {
                Text = "Annulla",
                DialogResult = DialogResult.Cancel,
                Location = new Point(290, 310),
                Size = new Size(80, 25)
            };

            this.Controls.AddRange(new Control[] {
                lblInstruction, txtOption, btnAdd, listOptions,
                btnRemove, btnOK, btnCancel
            });

            this.AcceptButton = btnAdd;
            this.CancelButton = btnCancel;
        }

        private void LoadExistingOptions()
        {
            foreach (var option in Options)
            {
                listOptions.Items.Add(option);
            }
        }

        private void BtnAdd_Click(object sender, EventArgs e)
        {
            string option = txtOption.Text.Trim();
            if (!string.IsNullOrEmpty(option) && !listOptions.Items.Contains(option))
            {
                listOptions.Items.Add(option);
                Options.Add(option);
                txtOption.Clear();
                txtOption.Focus();
            }
        }

        private void BtnRemove_Click(object sender, EventArgs e)
        {
            if (listOptions.SelectedIndex != -1)
            {
                Options.Remove(listOptions.SelectedItem.ToString());
                listOptions.Items.RemoveAt(listOptions.SelectedIndex);
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (DialogResult == DialogResult.OK)
            {
                Options = new List<string>();
                foreach (var item in listOptions.Items)
                {
                    Options.Add(item.ToString());
                }
            }
            base.OnFormClosing(e);
        }
    }
}
