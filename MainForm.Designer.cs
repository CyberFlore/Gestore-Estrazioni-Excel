namespace ExcelTemplateManager
{
    partial class MainForm
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.btnOpenFile = new System.Windows.Forms.Button();
            this.btnSaveTemplate = new System.Windows.Forms.Button();
            this.btnLoadTemplate = new System.Windows.Forms.Button();
            this.btnExport = new System.Windows.Forms.Button();
            this.listViewColumns = new System.Windows.Forms.ListView();
            this.btnAddColumn = new System.Windows.Forms.Button();
            this.btnRemoveColumn = new System.Windows.Forms.Button();
            this.txtNewColumn = new System.Windows.Forms.TextBox();

            // Impostazioni del form
            this.Text = "Gestore Template Excel";
            this.Size = new System.Drawing.Size(800, 600);
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;

            // Pulsante: Apri File
            this.btnOpenFile.Location = new System.Drawing.Point(12, 12);
            this.btnOpenFile.Size = new System.Drawing.Size(120, 30);
            this.btnOpenFile.Text = "Apri File Excel";
            this.btnOpenFile.Click += new System.EventHandler(this.btnOpenFile_Click);

            // Pulsante: Salva Template
            this.btnSaveTemplate.Location = new System.Drawing.Point(142, 12);
            this.btnSaveTemplate.Size = new System.Drawing.Size(120, 30);
            this.btnSaveTemplate.Text = "Salva Template";
            this.btnSaveTemplate.Click += new System.EventHandler(this.btnSaveTemplate_Click);

            // Pulsante: Carica Template
            this.btnLoadTemplate.Location = new System.Drawing.Point(272, 12);
            this.btnLoadTemplate.Size = new System.Drawing.Size(120, 30);
            this.btnLoadTemplate.Text = "Carica Template";
            this.btnLoadTemplate.Click += new System.EventHandler(this.btnLoadTemplate_Click);

            // Pulsante: Esporta
            this.btnExport.Location = new System.Drawing.Point(402, 12);
            this.btnExport.Size = new System.Drawing.Size(120, 30);
            this.btnExport.Text = "Esporta Excel";
            this.btnExport.Click += new System.EventHandler(this.btnExport_Click);

            // TextBox: Nuova Colonna
            this.txtNewColumn.Location = new System.Drawing.Point(12, 50);
            this.txtNewColumn.Size = new System.Drawing.Size(200, 23);
            this.txtNewColumn.PlaceholderText = "Nome nuova colonna...";

            // Pulsante: Aggiungi Colonna
            this.btnAddColumn.Location = new System.Drawing.Point(222, 50);
            this.btnAddColumn.Size = new System.Drawing.Size(120, 23);
            this.btnAddColumn.Text = "Aggiungi Colonna";
            this.btnAddColumn.Click += new System.EventHandler(this.btnAddColumn_Click);

            // Pulsante: Rimuovi Colonna
            this.btnRemoveColumn.Location = new System.Drawing.Point(352, 50);
            this.btnRemoveColumn.Size = new System.Drawing.Size(120, 23);
            this.btnRemoveColumn.Text = "Rimuovi Colonna";
            this.btnRemoveColumn.Click += new System.EventHandler(this.btnRemoveColumn_Click);

            // ListView: Colonne
            this.listViewColumns.Location = new System.Drawing.Point(12, 85);
            this.listViewColumns.Size = new System.Drawing.Size(760, 465);
            this.listViewColumns.View = System.Windows.Forms.View.Details;
            this.listViewColumns.FullRowSelect = true;
            this.listViewColumns.GridLines = true;
            this.listViewColumns.Columns.Add("Intestazioni Colonne", 740);

            // Aggiunge i controlli al form
            this.Controls.AddRange(new System.Windows.Forms.Control[] {
                this.btnOpenFile,
                this.btnSaveTemplate,
                this.btnLoadTemplate,
                this.btnExport,
                this.txtNewColumn,
                this.btnAddColumn,
                this.btnRemoveColumn,
                this.listViewColumns
            });
        }

        private System.Windows.Forms.Button btnOpenFile;
        private System.Windows.Forms.Button btnSaveTemplate;
        private System.Windows.Forms.Button btnLoadTemplate;
        private System.Windows.Forms.Button btnExport;
        private System.Windows.Forms.Button btnAddColumn;
        private System.Windows.Forms.Button btnRemoveColumn;
        private System.Windows.Forms.TextBox txtNewColumn;
        private System.Windows.Forms.ListView listViewColumns;
    }
}
