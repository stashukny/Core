namespace BillingProcessor
{
    partial class Form1
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnAddScheduleA = new System.Windows.Forms.Button();
            this.btnAddFormulas = new System.Windows.Forms.Button();
            this.btnListFiles = new System.Windows.Forms.Button();
            this.btnRunAll = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnAddScheduleA
            // 
            this.btnAddScheduleA.Location = new System.Drawing.Point(73, 65);
            this.btnAddScheduleA.Name = "btnAddScheduleA";
            this.btnAddScheduleA.Size = new System.Drawing.Size(168, 23);
            this.btnAddScheduleA.TabIndex = 0;
            this.btnAddScheduleA.Text = "Add to Schedule A";
            this.btnAddScheduleA.UseVisualStyleBackColor = true;
            this.btnAddScheduleA.Click += new System.EventHandler(this.btnAddScheduleA_Click);
            // 
            // btnAddFormulas
            // 
            this.btnAddFormulas.Location = new System.Drawing.Point(73, 117);
            this.btnAddFormulas.Name = "btnAddFormulas";
            this.btnAddFormulas.Size = new System.Drawing.Size(168, 23);
            this.btnAddFormulas.TabIndex = 1;
            this.btnAddFormulas.Text = "Add to Formulas";
            this.btnAddFormulas.UseVisualStyleBackColor = true;
            this.btnAddFormulas.Click += new System.EventHandler(this.btnAddFormulas_Click);
            // 
            // btnListFiles
            // 
            this.btnListFiles.Location = new System.Drawing.Point(73, 21);
            this.btnListFiles.Name = "btnListFiles";
            this.btnListFiles.Size = new System.Drawing.Size(168, 23);
            this.btnListFiles.TabIndex = 2;
            this.btnListFiles.Text = "List files";
            this.btnListFiles.UseVisualStyleBackColor = true;
            this.btnListFiles.Click += new System.EventHandler(this.btnListFiles_Click);
            // 
            // btnRunAll
            // 
            this.btnRunAll.Location = new System.Drawing.Point(73, 183);
            this.btnRunAll.Name = "btnRunAll";
            this.btnRunAll.Size = new System.Drawing.Size(168, 23);
            this.btnRunAll.TabIndex = 3;
            this.btnRunAll.Text = "Run all changes";
            this.btnRunAll.UseVisualStyleBackColor = true;
            this.btnRunAll.Click += new System.EventHandler(this.btnRunAll_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(282, 255);
            this.Controls.Add(this.btnRunAll);
            this.Controls.Add(this.btnListFiles);
            this.Controls.Add(this.btnAddFormulas);
            this.Controls.Add(this.btnAddScheduleA);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button btnAddScheduleA;
        private System.Windows.Forms.Button btnAddFormulas;
        private System.Windows.Forms.Button btnListFiles;
        private System.Windows.Forms.Button btnRunAll;
    }
}

