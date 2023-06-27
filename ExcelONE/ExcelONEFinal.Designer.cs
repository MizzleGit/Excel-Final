namespace ExcelONE
{
    partial class ExcelONEFinal
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            components = new System.ComponentModel.Container();
            btnBrowse = new Button();
            openExcel = new OpenFileDialog();
            btnModify = new Button();
            pbarMain = new ProgressBar();
            lblDebug = new Label();
            btnGlobal = new Button();
            tipBtnBrowse = new ToolTip(components);
            SuspendLayout();
            // 
            // btnBrowse
            // 
            btnBrowse.Location = new Point(12, 12);
            btnBrowse.Name = "btnBrowse";
            btnBrowse.Size = new Size(326, 168);
            btnBrowse.TabIndex = 0;
            btnBrowse.Text = "1. Ouvrir fichiers";
            tipBtnBrowse.SetToolTip(btnBrowse, "Cliquez ici pour sélectionner les fichiers que vous souhaitez modifier!");
            btnBrowse.UseVisualStyleBackColor = true;
            btnBrowse.Click += btnBrowse_Click;
            // 
            // openExcel
            // 
            openExcel.FileName = "Open File";
            // 
            // btnModify
            // 
            btnModify.Location = new Point(357, 12);
            btnModify.Name = "btnModify";
            btnModify.Size = new Size(328, 168);
            btnModify.TabIndex = 1;
            btnModify.Text = "2. Modifier les fichiers";
            btnModify.UseVisualStyleBackColor = true;
            btnModify.Click += btnModify_Click;
            // 
            // pbarMain
            // 
            pbarMain.Location = new Point(12, 430);
            pbarMain.Name = "pbarMain";
            pbarMain.Size = new Size(776, 34);
            pbarMain.TabIndex = 2;
            // 
            // lblDebug
            // 
            lblDebug.AutoSize = true;
            lblDebug.Font = new Font("Segoe UI", 8F, FontStyle.Regular, GraphicsUnit.Point);
            lblDebug.Location = new Point(12, 395);
            lblDebug.Name = "lblDebug";
            lblDebug.Size = new Size(91, 21);
            lblDebug.TabIndex = 3;
            lblDebug.Text = "debugLabel";
            // 
            // btnGlobal
            // 
            btnGlobal.Location = new Point(12, 217);
            btnGlobal.Name = "btnGlobal";
            btnGlobal.Size = new Size(326, 168);
            btnGlobal.TabIndex = 4;
            btnGlobal.Text = "3. Créer un fichier global";
            btnGlobal.UseVisualStyleBackColor = true;
            // 
            // ExcelONEFinal
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 475);
            Controls.Add(btnGlobal);
            Controls.Add(lblDebug);
            Controls.Add(pbarMain);
            Controls.Add(btnModify);
            Controls.Add(btnBrowse);
            Name = "ExcelONEFinal";
            Text = "ONE Recouvrement";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private Button btnBrowse;
        private OpenFileDialog openExcel;
        private Button btnModify;
        private ProgressBar pbarMain;
        private Label lblDebug;
        private Button btnGlobal;
        private ToolTip tipBtnBrowse;
    }
}