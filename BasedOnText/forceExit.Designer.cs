namespace BasedOnText
{
    partial class ForceExit
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(ForceExit));
            this.btnForceExit = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.waningMessageLabel = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnForceExit
            // 
            this.btnForceExit.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnForceExit.Location = new System.Drawing.Point(185, 66);
            this.btnForceExit.Name = "btnForceExit";
            this.btnForceExit.Size = new System.Drawing.Size(75, 25);
            this.btnForceExit.TabIndex = 0;
            this.btnForceExit.Text = "확인";
            this.btnForceExit.UseVisualStyleBackColor = true;
            this.btnForceExit.Click += new System.EventHandler(this.btnForceExit_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Font = new System.Drawing.Font("BatangChe", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCancel.Location = new System.Drawing.Point(287, 66);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 25);
            this.btnCancel.TabIndex = 1;
            this.btnCancel.Text = "취소";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // waningMessageLabel
            // 
            this.waningMessageLabel.AutoSize = true;
            this.waningMessageLabel.Font = new System.Drawing.Font("BatangChe", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(129)));
            this.waningMessageLabel.Location = new System.Drawing.Point(33, 20);
            this.waningMessageLabel.Name = "waningMessageLabel";
            this.waningMessageLabel.Size = new System.Drawing.Size(0, 13);
            this.waningMessageLabel.TabIndex = 2;
            // 
            // ForceExit
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(402, 113);
            this.Controls.Add(this.waningMessageLabel);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnForceExit);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "ForceExit";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "경고!";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnForceExit;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Label waningMessageLabel;
    }
}