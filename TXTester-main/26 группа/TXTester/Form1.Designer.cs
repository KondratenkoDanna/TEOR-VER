namespace TXTester
{
    partial class f_TXTexter
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            this.tb_Title = new System.Windows.Forms.TextBox();
            this.tb_Description = new System.Windows.Forms.TextBox();
            this.b_Start = new System.Windows.Forms.Button();
            this.flp_QuestionPanel = new System.Windows.Forms.FlowLayoutPanel();
            this.b_Previous = new System.Windows.Forms.Button();
            this.b_Next = new System.Windows.Forms.Button();
            this.b_toDocx = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // tb_Title
            // 
            this.tb_Title.CausesValidation = false;
            this.tb_Title.Location = new System.Drawing.Point(13, 13);
            this.tb_Title.Name = "tb_Title";
            this.tb_Title.ReadOnly = true;
            this.tb_Title.Size = new System.Drawing.Size(775, 20);
            this.tb_Title.TabIndex = 0;
            this.tb_Title.TabStop = false;
            this.tb_Title.Text = "Title";
            // 
            // tb_Description
            // 
            this.tb_Description.AcceptsReturn = true;
            this.tb_Description.AcceptsTab = true;
            this.tb_Description.CausesValidation = false;
            this.tb_Description.Location = new System.Drawing.Point(13, 40);
            this.tb_Description.Multiline = true;
            this.tb_Description.Name = "tb_Description";
            this.tb_Description.ReadOnly = true;
            this.tb_Description.Size = new System.Drawing.Size(775, 55);
            this.tb_Description.TabIndex = 1;
            this.tb_Description.TabStop = false;
            this.tb_Description.Text = "Description";
            // 
            // b_Start
            // 
            this.b_Start.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.b_Start.Location = new System.Drawing.Point(269, 451);
            this.b_Start.Name = "b_Start";
            this.b_Start.Size = new System.Drawing.Size(264, 23);
            this.b_Start.TabIndex = 2;
            this.b_Start.Text = "Start";
            this.b_Start.UseVisualStyleBackColor = true;
            this.b_Start.Click += new System.EventHandler(this.b_Start_Click);
            // 
            // flp_QuestionPanel
            // 
            this.flp_QuestionPanel.AutoSize = true;
            this.flp_QuestionPanel.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink;
            this.flp_QuestionPanel.FlowDirection = System.Windows.Forms.FlowDirection.TopDown;
            this.flp_QuestionPanel.Location = new System.Drawing.Point(13, 13);
            this.flp_QuestionPanel.Name = "flp_QuestionPanel";
            this.flp_QuestionPanel.Size = new System.Drawing.Size(0, 0);
            this.flp_QuestionPanel.TabIndex = 3;
            this.flp_QuestionPanel.Visible = false;
            // 
            // b_Previous
            // 
            this.b_Previous.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.b_Previous.Enabled = false;
            this.b_Previous.Location = new System.Drawing.Point(12, 451);
            this.b_Previous.Name = "b_Previous";
            this.b_Previous.Size = new System.Drawing.Size(130, 23);
            this.b_Previous.TabIndex = 0;
            this.b_Previous.Text = "Previous";
            this.b_Previous.UseVisualStyleBackColor = true;
            this.b_Previous.Visible = false;
            this.b_Previous.Click += new System.EventHandler(this.b_Previous_Click);
            // 
            // b_Next
            // 
            this.b_Next.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.b_Next.Enabled = false;
            this.b_Next.Location = new System.Drawing.Point(655, 451);
            this.b_Next.Name = "b_Next";
            this.b_Next.Size = new System.Drawing.Size(133, 23);
            this.b_Next.TabIndex = 4;
            this.b_Next.Text = "Next";
            this.b_Next.UseVisualStyleBackColor = true;
            this.b_Next.Visible = false;
            this.b_Next.Click += new System.EventHandler(this.b_Next_Click);
            // 
            // b_toDocx
            // 
            this.b_toDocx.Location = new System.Drawing.Point(169, 451);
            this.b_toDocx.Name = "b_toDocx";
            this.b_toDocx.Size = new System.Drawing.Size(75, 23);
            this.b_toDocx.TabIndex = 5;
            this.b_toDocx.Text = "To .docx";
            this.b_toDocx.UseVisualStyleBackColor = true;
            this.b_toDocx.Click += new System.EventHandler(this.b_toDoc_Click);
            // 
            // f_TXTexter
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoScroll = true;
            this.ClientSize = new System.Drawing.Size(800, 486);
            this.Controls.Add(this.b_toDocx);
            this.Controls.Add(this.b_Next);
            this.Controls.Add(this.b_Previous);
            this.Controls.Add(this.b_Start);
            this.Controls.Add(this.flp_QuestionPanel);
            this.Controls.Add(this.tb_Description);
            this.Controls.Add(this.tb_Title);
            this.Name = "f_TXTexter";
            this.Text = "TXTester";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tb_Title;
        private System.Windows.Forms.TextBox tb_Description;
        private System.Windows.Forms.Button b_Start;
        private System.Windows.Forms.FlowLayoutPanel flp_QuestionPanel;
        private System.Windows.Forms.Button b_Previous;
        private System.Windows.Forms.Button b_Next;
        private System.Windows.Forms.Button b_toDocx;
    }
}