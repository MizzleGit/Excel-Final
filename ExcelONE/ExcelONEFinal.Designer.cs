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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ExcelONEFinal));
            btnBrowse = new Button();
            openExcel = new OpenFileDialog();
            btnModify = new Button();
            pbarMain = new ProgressBar();
            lblDebug = new Label();
            btnGlobal = new Button();
            tipBtnBrowse = new ToolTip(components);
            tipBtnModify = new ToolTip(components);
            tipBtnGlobal = new ToolTip(components);
            btnDestination = new Button();
            tipBtnDestination = new ToolTip(components);
            openDestination = new OpenFileDialog();
            lblWait = new Label();
            SuspendLayout();
            // 
            // btnBrowse
            // 
            btnBrowse.Location = new Point(12, 12);
            btnBrowse.Name = "btnBrowse";
            btnBrowse.Size = new Size(326, 168);
            btnBrowse.TabIndex = 0;
            btnBrowse.Text = "1. Ouvrir fichiers";
            tipBtnBrowse.SetToolTip(btnBrowse, "Cliquez ici pour sélectionner les fichiers que vous souhaitez modifier.");
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
            tipBtnModify.SetToolTip(btnModify, "Cliquez ici si vous souhaitez appliquer des modifications aux fichiers.");
            btnModify.UseVisualStyleBackColor = true;
            btnModify.Click += btnModify_Click;
            // 
            // pbarMain
            // 
            pbarMain.Location = new Point(12, 430);
            pbarMain.Name = "pbarMain";
            pbarMain.Size = new Size(673, 34);
            pbarMain.TabIndex = 2;
            // 
            // lblDebug
            // 
            lblDebug.AutoSize = true;
            lblDebug.Font = new Font("Segoe UI", 8F, FontStyle.Regular, GraphicsUnit.Point);
            lblDebug.Location = new Point(12, 193);
            lblDebug.Name = "lblDebug";
            lblDebug.Size = new Size(91, 21);
            lblDebug.TabIndex = 3;
            lblDebug.Text = "debugLabel";
            lblDebug.Visible = false;
            // 
            // btnGlobal
            // 
            btnGlobal.Location = new Point(12, 217);
            btnGlobal.Name = "btnGlobal";
            btnGlobal.Size = new Size(326, 168);
            btnGlobal.TabIndex = 4;
            btnGlobal.Text = "3. Créer un fichier global";
            tipBtnGlobal.SetToolTip(btnGlobal, "Cliquez ici pour créer un fichier global. Assurez-vous d'avoir sélectionné exactement 8 fichiers.");
            btnGlobal.UseVisualStyleBackColor = true;
            btnGlobal.Click += btnGlobal_Click;
            // 
            // btnDestination
            // 
            btnDestination.Location = new Point(357, 217);
            btnDestination.Name = "btnDestination";
            btnDestination.Size = new Size(328, 168);
            btnDestination.TabIndex = 5;
            btnDestination.Text = "4. Modifiez le fichier de destination.";
            tipBtnDestination.SetToolTip(btnDestination, "Utilisez la fonction VLOOKUP sur globa.xlsx pour appliquer les modifications sur les données TR.");
            btnDestination.UseVisualStyleBackColor = true;
            btnDestination.Click += btnDestination_Click;
            // 
            // openDestination
            // 
            openDestination.FileName = "Open destination file";
            // 
            // lblWait
            // 
            lblWait.AutoSize = true;
            lblWait.Location = new Point(12, 402);
            lblWait.Name = "lblWait";
            lblWait.Size = new Size(484, 25);
            lblWait.TabIndex = 6;
            lblWait.Text = "Ce processus prendra quelques secondes, veuillez patienter.";
            lblWait.Visible = false;
            // 
            // ExcelONEFinal
            // 
            AutoScaleDimensions = new SizeF(10F, 25F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(697, 475);
            Controls.Add(lblWait);
            Controls.Add(btnDestination);
            Controls.Add(btnGlobal);
            Controls.Add(lblDebug);
            Controls.Add(pbarMain);
            Controls.Add(btnModify);
            Controls.Add(btnBrowse);
            Icon = (Icon)resources.GetObject("$this.Icon");
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
        private ToolTip tipBtnModify;
        private ToolTip tipBtnGlobal;
        private Button btnDestination;
        private ToolTip tipBtnDestination;
        private OpenFileDialog openDestination;
        private Label lblWait;
    }
}