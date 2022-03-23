namespace TestovoeZadanieRPA
{
    partial class Form1
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
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.update = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.excel = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.rateTable = new System.Windows.Forms.DataGridView();
            this.tableLayoutPanel2 = new System.Windows.Forms.TableLayoutPanel();
            this.label8 = new System.Windows.Forms.Label();
            this.emailTo = new System.Windows.Forms.TextBox();
            this.password = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.email = new System.Windows.Forms.Button();
            this.label3 = new System.Windows.Forms.Label();
            this.login = new System.Windows.Forms.TextBox();
            this.saveFile = new System.Windows.Forms.SaveFileDialog();
            this.tableLayoutPanel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.rateTable)).BeginInit();
            this.tableLayoutPanel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.ColumnCount = 2;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 70F));
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 30F));
            this.tableLayoutPanel1.Controls.Add(this.update, 0, 2);
            this.tableLayoutPanel1.Controls.Add(this.label1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.excel, 1, 2);
            this.tableLayoutPanel1.Controls.Add(this.label2, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.rateTable, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.tableLayoutPanel2, 1, 1);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 7F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 83F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(800, 450);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // update
            // 
            this.update.Dock = System.Windows.Forms.DockStyle.Fill;
            this.update.Font = new System.Drawing.Font("Arial", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.update.Location = new System.Drawing.Point(7, 411);
            this.update.Margin = new System.Windows.Forms.Padding(7);
            this.update.Name = "update";
            this.update.Size = new System.Drawing.Size(546, 32);
            this.update.TabIndex = 7;
            this.update.Text = "Обновить";
            this.update.UseVisualStyleBackColor = true;
            this.update.Click += new System.EventHandler(this.update_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label1.Font = new System.Drawing.Font("Arial", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label1.Location = new System.Drawing.Point(563, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(234, 31);
            this.label1.TabIndex = 5;
            this.label1.Text = "Экспорт данных";
            this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // excel
            // 
            this.excel.Dock = System.Windows.Forms.DockStyle.Fill;
            this.excel.Font = new System.Drawing.Font("Arial", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.excel.Location = new System.Drawing.Point(567, 411);
            this.excel.Margin = new System.Windows.Forms.Padding(7);
            this.excel.Name = "excel";
            this.excel.Size = new System.Drawing.Size(226, 32);
            this.excel.TabIndex = 4;
            this.excel.Text = "Excel";
            this.excel.UseVisualStyleBackColor = true;
            this.excel.Click += new System.EventHandler(this.excel_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label2.Font = new System.Drawing.Font("Arial", 14.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label2.Location = new System.Drawing.Point(3, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(554, 31);
            this.label2.TabIndex = 3;
            this.label2.Text = "Курс валют";
            this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // rateTable
            // 
            this.rateTable.AllowUserToAddRows = false;
            this.rateTable.AllowUserToDeleteRows = false;
            this.rateTable.AllowUserToResizeColumns = false;
            this.rateTable.AllowUserToResizeRows = false;
            this.rateTable.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.rateTable.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.rateTable.Dock = System.Windows.Forms.DockStyle.Fill;
            this.rateTable.Location = new System.Drawing.Point(3, 34);
            this.rateTable.Name = "rateTable";
            this.rateTable.ReadOnly = true;
            this.rateTable.RowTemplate.Height = 25;
            this.rateTable.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.rateTable.Size = new System.Drawing.Size(554, 367);
            this.rateTable.TabIndex = 1;
            // 
            // tableLayoutPanel2
            // 
            this.tableLayoutPanel2.ColumnCount = 1;
            this.tableLayoutPanel2.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel2.Controls.Add(this.label8, 0, 8);
            this.tableLayoutPanel2.Controls.Add(this.emailTo, 0, 9);
            this.tableLayoutPanel2.Controls.Add(this.password, 0, 4);
            this.tableLayoutPanel2.Controls.Add(this.label7, 0, 0);
            this.tableLayoutPanel2.Controls.Add(this.label6, 0, 7);
            this.tableLayoutPanel2.Controls.Add(this.label5, 0, 5);
            this.tableLayoutPanel2.Controls.Add(this.label4, 0, 3);
            this.tableLayoutPanel2.Controls.Add(this.email, 0, 10);
            this.tableLayoutPanel2.Controls.Add(this.label3, 0, 1);
            this.tableLayoutPanel2.Controls.Add(this.login, 0, 2);
            this.tableLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel2.Location = new System.Drawing.Point(563, 34);
            this.tableLayoutPanel2.Name = "tableLayoutPanel2";
            this.tableLayoutPanel2.RowCount = 11;
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 5F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 5F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 15F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 5F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tableLayoutPanel2.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 10F));
            this.tableLayoutPanel2.Size = new System.Drawing.Size(234, 367);
            this.tableLayoutPanel2.TabIndex = 6;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label8.Font = new System.Drawing.Font("Arial", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label8.Location = new System.Drawing.Point(3, 253);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(228, 36);
            this.label8.TabIndex = 15;
            this.label8.Text = "Логин";
            this.label8.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // emailTo
            // 
            this.emailTo.Dock = System.Windows.Forms.DockStyle.Fill;
            this.emailTo.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.emailTo.Location = new System.Drawing.Point(3, 292);
            this.emailTo.Name = "emailTo";
            this.emailTo.Size = new System.Drawing.Size(228, 22);
            this.emailTo.TabIndex = 14;
            // 
            // password
            // 
            this.password.Dock = System.Windows.Forms.DockStyle.Fill;
            this.password.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.password.Location = new System.Drawing.Point(3, 129);
            this.password.Name = "password";
            this.password.PasswordChar = '*';
            this.password.Size = new System.Drawing.Size(228, 22);
            this.password.TabIndex = 13;
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label7.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label7.Location = new System.Drawing.Point(3, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(228, 18);
            this.label7.TabIndex = 11;
            this.label7.Text = "Почта с котрой отправится письмо:";
            this.label7.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label6.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label6.Location = new System.Drawing.Point(3, 235);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(228, 18);
            this.label6.TabIndex = 10;
            this.label6.Text = "Почта на которую прийдет письмо:";
            this.label6.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label5.Font = new System.Drawing.Font("Arial", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label5.Location = new System.Drawing.Point(3, 162);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(228, 18);
            this.label5.TabIndex = 9;
            this.label5.Text = "Сервер: smtp.gmail.com Порт: 587";
            this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label4.Font = new System.Drawing.Font("Arial", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label4.Location = new System.Drawing.Point(3, 90);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(228, 36);
            this.label4.TabIndex = 8;
            this.label4.Text = "Пароль";
            this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // email
            // 
            this.email.Dock = System.Windows.Forms.DockStyle.Fill;
            this.email.Font = new System.Drawing.Font("Arial", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.email.Location = new System.Drawing.Point(7, 332);
            this.email.Margin = new System.Windows.Forms.Padding(7);
            this.email.Name = "email";
            this.email.Size = new System.Drawing.Size(220, 28);
            this.email.TabIndex = 7;
            this.email.Text = "Отправить письмо";
            this.email.UseVisualStyleBackColor = true;
            this.email.Click += new System.EventHandler(this.email_Click);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.label3.Font = new System.Drawing.Font("Arial", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.label3.Location = new System.Drawing.Point(3, 18);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(228, 36);
            this.label3.TabIndex = 6;
            this.label3.Text = "Логин";
            this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            // 
            // login
            // 
            this.login.Dock = System.Windows.Forms.DockStyle.Fill;
            this.login.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point);
            this.login.Location = new System.Drawing.Point(3, 57);
            this.login.Name = "login";
            this.login.Size = new System.Drawing.Size(228, 22);
            this.login.TabIndex = 12;
            this.login.Text = "s3arch.mov13@gmail.com";
            // 
            // saveFile
            // 
            this.saveFile.FileOk += new System.ComponentModel.CancelEventHandler(this.saveFileDialog1_FileOk);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.tableLayoutPanel1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.rateTable)).EndInit();
            this.tableLayoutPanel2.ResumeLayout(false);
            this.tableLayoutPanel2.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.DataGridView rateTable;
        private System.Windows.Forms.Button excel;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.SaveFileDialog saveFile;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button email;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox emailTo;
        private System.Windows.Forms.TextBox password;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.TextBox login;
        private System.Windows.Forms.Button update;
    }
}
