using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using System.IO;
using System.Linq;

namespace ExcelTemplateManager
{
    public partial class MainForm : Form
    {
        private ExcelManager excelManager;
        private TemplateManager templateManager;
        private List<ColumnInfo> columns;
        private string? currentFilePath;
        private Dictionary<string, Template> savedTemplates;
        private const string TEMPLATES_FILE = "saved_templates.dat";

        public MainForm()
        {
            InitializeComponent();
            excelManager = new ExcelManager();
            templateManager = new TemplateManager();
            columns = new List<ColumnInfo>();
            savedTemplates = LoadSavedTemplates();
            SetupDragDrop();
            ApplyModernStyle();
            InitializeValidationContextMenu(); // Initialize context menu for validation
        }

        private Dictionary<string, Template> LoadSavedTemplates()
        {
            try
            {
                if (File.Exists(TEMPLATES_FILE))
                {
                    var templates = templateManager.LoadTemplatesFromFile(TEMPLATES_FILE);
                    MessageBox.Show($"Caricati {templates.Count} template dalla memoria.",
                        "Template caricati", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return templates;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Errore nel caricamento dei template salvati: {ex.Message}",
                    "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return new Dictionary<string, Template>();
        }

        private void SaveTemplatesState()
        {
            try
            {
                templateManager.SaveTemplatesToFile(savedTemplates, TEMPLATES_FILE);
                // Solo per debug, rimuovere in produzione
                MessageBox.Show($"Template salvati con successo. File: {Path.GetFullPath(TEMPLATES_FILE)}",
                    "Salvataggio completato", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Errore nel salvataggio dei template: {ex.Message}",
                    "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ApplyModernStyle()
        {
            // Stile del form
            Color backgroundColor = Color.FromArgb(240, 240, 240);
            this.BackColor = backgroundColor;
            this.Font = new Font("Segoe UI", 9F, FontStyle.Regular);

            // Configura i pulsanti principali
            Button[] mainButtons = { btnOpenFile, btnSaveTemplate, btnLoadTemplate, btnExport };
            foreach (var button in mainButtons)
            {
                button.Size = new Size(120, 30);
                button.BackColor = Color.FromArgb(240, 240, 240);
                button.ForeColor = Color.Black;
                button.FlatStyle = FlatStyle.Standard;
                button.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
                button.Cursor = Cursors.Hand;
            }

            // Configura i pulsanti delle azioni
            btnAddColumn.BackColor = Color.FromArgb(240, 240, 240);
            btnRemoveColumn.BackColor = Color.FromArgb(240, 240, 240);

            Button[] actionButtons = { btnAddColumn, btnRemoveColumn };
            foreach (var button in actionButtons)
            {
                button.Size = new Size(100, 25);
                button.ForeColor = Color.Black;
                button.FlatStyle = FlatStyle.Standard;
                button.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
                button.Cursor = Cursors.Hand;
            }

            btnAddColumn.Text = "Aggiungi";
            btnRemoveColumn.Text = "Rimuovi";

            // Stile della ListView
            listViewColumns.BackColor = Color.White;
            listViewColumns.BorderStyle = BorderStyle.FixedSingle;
            listViewColumns.Font = new Font("Segoe UI", 9F, FontStyle.Regular);
            listViewColumns.FullRowSelect = true;
            listViewColumns.GridLines = true;
            listViewColumns.View = View.Details;

            // Aggiungi dropdown per i template salvati
            ComboBox cmbTemplates = new ComboBox
            {
                Location = new Point(12, 85),
                Size = new Size(200, 23),
                DropDownStyle = ComboBoxStyle.DropDownList,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular)
            };
            cmbTemplates.SelectedIndexChanged += CmbTemplates_SelectedIndexChanged;
            Controls.Add(cmbTemplates);

            // Pulsante per salvare il template in memoria
            Button btnSaveToMemory = new Button
            {
                Text = "Salva Template in Memoria",
                Location = new Point(222, 85),
                Size = new Size(150, 23),
                BackColor = Color.FromArgb(240, 240, 240),
                ForeColor = Color.Black,
                FlatStyle = FlatStyle.Standard,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular)
            };
            btnSaveToMemory.Click += BtnSaveToMemory_Click;
            Controls.Add(btnSaveToMemory);

            // Pulsante per rimuovere il template
            Button btnRemoveTemplate = new Button
            {
                Text = "Rimuovi Template",
                Location = new Point(382, 85),
                Size = new Size(120, 23),
                BackColor = Color.FromArgb(240, 240, 240),
                ForeColor = Color.Black,
                FlatStyle = FlatStyle.Standard,
                Font = new Font("Segoe UI", 9F, FontStyle.Regular)
            };
            btnRemoveTemplate.Click += BtnRemoveTemplate_Click;
            Controls.Add(btnRemoveTemplate);


            // Aggiorna la posizione della ListView
            listViewColumns.Location = new Point(12, 120);
            listViewColumns.Size = new Size(760, 400);

            // Aggiungi label per i crediti
            Label lblCredits = new Label
            {
                Text = "Gestore Template Excel creato da Mirko Subri",
                Location = new Point(12, 530),
                Size = new Size(760, 20),
                Font = new Font("Segoe UI", 8F, FontStyle.Italic),
                ForeColor = Color.Gray,
                TextAlign = ContentAlignment.MiddleCenter
            };
            Controls.Add(lblCredits);

            // Aggiorna la lista dei template
            UpdateTemplatesList();
        }

        private void BtnRemoveTemplate_Click(object sender, EventArgs e)
        {
            var cmbTemplates = Controls.OfType<ComboBox>().FirstOrDefault();
            if (cmbTemplates?.SelectedItem == null)
            {
                MessageBox.Show("Seleziona un template da rimuovere.", "Selezione richiesta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string templateName = cmbTemplates.SelectedItem.ToString();
            if (MessageBox.Show($"Sei sicuro di voler rimuovere il template '{templateName}'?",
                "Conferma rimozione", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                savedTemplates.Remove(templateName);
                SaveTemplatesState();
                UpdateTemplatesList();
                MessageBox.Show("Template rimosso con successo!", "Operazione completata", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void CmbTemplates_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (sender is ComboBox cmb && cmb.SelectedItem != null)
            {
                string templateName = cmb.SelectedItem.ToString();
                if (savedTemplates.ContainsKey(templateName))
                {
                    // Crea una copia profonda delle colonne dal template
                    columns = savedTemplates[templateName].Columns.Select(tc => new ColumnInfo
                    {
                        Header = tc.Header,
                        OriginalHeader = tc.OriginalHeader,
                        OriginalIndex = tc.OriginalIndex,
                        ValidationOptions = new List<string>(tc.ValidationOptions),
                        HasValidation = tc.HasValidation
                    }).ToList();

                    UpdateColumnListView();
                }
            }
        }

        private void BtnSaveToMemory_Click(object sender, EventArgs e)
        {
            string templateName = Microsoft.VisualBasic.Interaction.InputBox(
                "Inserisci il nome del template:",
                "Salva Template in Memoria",
                "Template " + (savedTemplates.Count + 1));

            if (!string.IsNullOrWhiteSpace(templateName))
            {
                if (savedTemplates.ContainsKey(templateName))
                {
                    if (MessageBox.Show(
                        "Esiste già un template con questo nome. Vuoi sovrascriverlo?",
                        "Template esistente",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question) != DialogResult.Yes)
                    {
                        return;
                    }
                }

                // Assicurati che tutte le colonne abbiano un OriginalHeader prima del salvataggio
                foreach (var column in columns)
                {
                    if (string.IsNullOrEmpty(column.OriginalHeader))
                    {
                        column.OriginalHeader = column.Header;
                    }
                }

                // Crea una copia profonda delle colonne per il template
                var templateColumns = columns.Select(c => new ColumnInfo
                {
                    Header = c.Header,
                    OriginalHeader = c.OriginalHeader,
                    OriginalIndex = c.OriginalIndex,
                    ValidationOptions = new List<string>(c.ValidationOptions),
                    HasValidation = c.HasValidation
                }).ToList();

                savedTemplates[templateName] = new Template { Columns = templateColumns };
                SaveTemplatesState();
                UpdateTemplatesList();
                MessageBox.Show(
                    "Template salvato in memoria con successo!",
                    "Operazione completata",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        private void UpdateTemplatesList()
        {
            var cmbTemplates = Controls.OfType<ComboBox>().FirstOrDefault();
            if (cmbTemplates != null)
            {
                cmbTemplates.Items.Clear();
                cmbTemplates.Items.AddRange(savedTemplates.Keys.ToArray());
            }
        }

        private void SetupDragDrop()
        {
            listViewColumns.AllowDrop = true;
            listViewColumns.DragEnter += ListView_DragEnter;
            listViewColumns.DragDrop += ListView_DragDrop;
            listViewColumns.ItemDrag += ListView_ItemDrag;
        }

        private void btnOpenFile_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "File Excel|*.xlsx;*.xls";
                openFileDialog.Title = "Seleziona un file Excel";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        currentFilePath = openFileDialog.FileName;
                        LoadExcelFile(currentFilePath);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Errore durante l'apertura del file: {ex.Message}", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void LoadExcelFile(string filePath)
        {
            try
            {
                columns = excelManager.LoadExcelFile(filePath);

                // Salva le intestazioni originali per le nuove colonne
                foreach (var column in columns)
                {
                    if (string.IsNullOrEmpty(column.OriginalHeader))
                    {
                        column.OriginalHeader = column.Header;
                    }
                }

                // Cerca di applicare i nomi personalizzati dai template salvati
                ApplyCustomHeadersFromTemplates();

                UpdateColumnListView();
                MessageBox.Show("File Excel caricato con successo!", "Operazione completata", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Errore durante il caricamento del file: {ex.Message}", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ApplyCustomHeadersFromTemplates()
        {
            foreach (var template in savedTemplates.Values)
            {
                foreach (var templateColumn in template.Columns)
                {
                    // Cerca una colonna corrispondente nel file corrente
                    var matchingColumn = columns.FirstOrDefault(c =>
                        !string.IsNullOrEmpty(templateColumn.OriginalHeader) &&
                        c.OriginalHeader == templateColumn.OriginalHeader);

                    if (matchingColumn != null)
                    {
                        // Applica il nome personalizzato dal template
                        matchingColumn.Header = templateColumn.Header;
                        matchingColumn.ValidationOptions = templateColumn.ValidationOptions;
                        matchingColumn.HasValidation = templateColumn.HasValidation;
                    }
                }
            }
        }

        private void UpdateColumnListView()
        {
            listViewColumns.Items.Clear();
            foreach (var column in columns)
            {
                var item = new ListViewItem(column.Header);
                item.Tag = column;

                // Aggiungi il nome originale come seconda colonna
                item.SubItems.Add(column.OriginalHeader ?? "Nuova Colonna");

                // Aggiungi l'indicatore di validazione come terza colonna
                if (column.HasValidation)
                {
                    item.SubItems.Add("✓");
                }
                else
                {
                    item.SubItems.Add("");
                }

                if (column.OriginalIndex == -1)
                {
                    item.BackColor = Color.LightGreen; // Evidenzia le nuove colonne
                }
                listViewColumns.Items.Add(item);
            }

            // Configura le colonne della ListView
            if (listViewColumns.Columns.Count == 1)
            {
                listViewColumns.Columns[0].Text = "Nome Nuovo";
                listViewColumns.Columns[0].Width = 200;
                listViewColumns.Columns.Add("Nome Originale", 200);
                listViewColumns.Columns.Add("Validazione", 80);
            }
        }

        private void ListView_ItemDrag(object sender, ItemDragEventArgs e)
        {
            if (e.Item is ListViewItem item)
            {
                listViewColumns.DoDragDrop(item, DragDropEffects.Move);
            }
        }

        private void ListView_DragEnter(object sender, DragEventArgs e)
        {
            e.Effect = e.AllowedEffect;
        }

        private void ListView_DragDrop(object sender, DragEventArgs e)
        {
            Point targetPoint = listViewColumns.PointToClient(new Point(e.X, e.Y));
            ListViewItem targetItem = listViewColumns.GetItemAt(targetPoint.X, targetPoint.Y);
            ListViewItem draggedItem = (ListViewItem)e.Data.GetData(typeof(ListViewItem));

            if (draggedItem != null)
            {
                int targetIndex = targetItem != null ? targetItem.Index : listViewColumns.Items.Count;
                ColumnInfo draggedColumn = (ColumnInfo)draggedItem.Tag;

                columns.Remove(draggedColumn);
                columns.Insert(targetIndex, draggedColumn);
                UpdateColumnListView();
            }
        }

        private void btnAddColumn_Click(object sender, EventArgs e)
        {
            string columnName = txtNewColumn.Text.Trim();
            if (string.IsNullOrEmpty(columnName))
            {
                MessageBox.Show("Inserire un nome per la nuova colonna.", "Campo richiesto", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var newColumn = new ColumnInfo
            {
                Header = columnName,
                OriginalIndex = -1, // -1 indica una nuova colonna
                OriginalHeader = "" // Initialize OriginalHeader
            };

            using (var validationForm = new ValidationOptionsForm())
            {
                if (validationForm.ShowDialog() == DialogResult.OK)
                {
                    newColumn.ValidationOptions = validationForm.Options;
                    newColumn.HasValidation = validationForm.Options.Count > 0;
                }
            }

            columns.Add(newColumn);
            UpdateColumnListView();
            txtNewColumn.Clear();
        }

        private void btnRemoveColumn_Click(object sender, EventArgs e)
        {
            if (listViewColumns.SelectedItems.Count == 0)
            {
                MessageBox.Show("Selezionare una colonna da rimuovere.", "Selezione richiesta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (MessageBox.Show("Sei sicuro di voler rimuovere questa colonna?", "Conferma rimozione",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                var selectedItem = listViewColumns.SelectedItems[0];
                var columnToRemove = (ColumnInfo)selectedItem.Tag;
                columns.Remove(columnToRemove);
                UpdateColumnListView();
            }
        }

        private void btnSaveTemplate_Click(object sender, EventArgs e)
        {
            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "File Template|*.tmpl";
                saveFileDialog.Title = "Salva Template";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        Template template = new Template { Columns = new List<ColumnInfo>(columns) };
                        templateManager.SaveTemplate(template, saveFileDialog.FileName);
                        MessageBox.Show("Template salvato con successo!", "Operazione completata", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Errore durante il salvataggio del template: {ex.Message}", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnLoadTemplate_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Filter = "File Template|*.tmpl";
                openFileDialog.Title = "Carica Template";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        Template template = templateManager.LoadTemplate(openFileDialog.FileName);
                        columns = new List<ColumnInfo>(template.Columns);
                        UpdateColumnListView();
                        MessageBox.Show("Template caricato con successo!", "Operazione completata", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Errore durante il caricamento del template: {ex.Message}", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(currentFilePath))
            {
                MessageBox.Show("Aprire prima un file Excel.", "Operazione richiesta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (SaveFileDialog saveFileDialog = new SaveFileDialog())
            {
                saveFileDialog.Filter = "File Excel|*.xlsx;*.xls";
                saveFileDialog.Title = "Esporta file Excel";

                if (saveFileDialog.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        excelManager.ExportExcelFile(currentFilePath, saveFileDialog.FileName, columns);
                        MessageBox.Show("File esportato con successo!", "Operazione completata", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Errore durante l'esportazione del file: {ex.Message}", "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }


        private void InitializeValidationContextMenu()
        {
            var contextMenu = new ContextMenuStrip();

            var editValidationItem = new ToolStripMenuItem("Modifica opzioni di validazione");
            editValidationItem.Click += EditValidation_Click;
            contextMenu.Items.Add(editValidationItem);

            contextMenu.Items.Add(new ToolStripSeparator());

            var renameColumnItem = new ToolStripMenuItem("Rinomina colonna");
            renameColumnItem.Click += RenameColumn_Click;
            contextMenu.Items.Add(renameColumnItem);

            listViewColumns.ContextMenuStrip = contextMenu;
        }

        private void EditValidation_Click(object sender, EventArgs e)
        {
            if (listViewColumns.SelectedItems.Count == 0)
            {
                MessageBox.Show("Selezionare una colonna per modificare le opzioni di validazione.",
                    "Selezione richiesta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var selectedItem = listViewColumns.SelectedItems[0];
            var columnInfo = (ColumnInfo)selectedItem.Tag;

            using (var validationForm = new ValidationOptionsForm(columnInfo.ValidationOptions))
            {
                if (validationForm.ShowDialog() == DialogResult.OK)
                {
                    columnInfo.ValidationOptions = validationForm.Options;
                    columnInfo.HasValidation = validationForm.Options.Count > 0;
                    UpdateColumnListView();
                }
            }
        }

        private void RenameColumn_Click(object sender, EventArgs e)
        {
            if (listViewColumns.SelectedItems.Count == 0)
            {
                MessageBox.Show("Selezionare una colonna da rinominare.",
                    "Selezione richiesta", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var selectedItem = listViewColumns.SelectedItems[0];
            var columnInfo = (ColumnInfo)selectedItem.Tag;

            string newName = Microsoft.VisualBasic.Interaction.InputBox(
                "Inserisci il nuovo nome per la colonna:",
                "Rinomina Colonna",
                columnInfo.Header);

            if (!string.IsNullOrWhiteSpace(newName))
            {
                // Se è una nuova colonna, salva l'intestazione originale
                if (string.IsNullOrEmpty(columnInfo.OriginalHeader))
                {
                    columnInfo.OriginalHeader = columnInfo.Header;
                }
                columnInfo.Header = newName;
                UpdateColumnListView();
            }
        }
    }
}