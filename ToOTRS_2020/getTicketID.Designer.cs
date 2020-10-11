namespace TestOutlookAddIn
{
    partial class getTicketID
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
            this.labelTicketID = new System.Windows.Forms.Label();
            this.tbox_TicketID = new System.Windows.Forms.TextBox();
            this.btn_resetdstfolder = new System.Windows.Forms.Button();
            this.lbl_info = new System.Windows.Forms.Label();
            this.btn_ok = new System.Windows.Forms.Button();
            this.lbl_Subject = new System.Windows.Forms.Label();
            this.txtbox_Subject = new System.Windows.Forms.TextBox();
            this.checkNewTicket = new System.Windows.Forms.CheckBox();
            this.helpProvider1 = new System.Windows.Forms.HelpProvider();
            this.label_info = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // labelTicketID
            // 
            this.labelTicketID.AutoSize = true;
            this.labelTicketID.Location = new System.Drawing.Point(21, 80);
            this.labelTicketID.Name = "labelTicketID";
            this.labelTicketID.Size = new System.Drawing.Size(101, 13);
            this.labelTicketID.TabIndex = 0;
            this.labelTicketID.Text = "TicketID eingeben: ";
            // 
            // tbox_TicketID
            // 
            this.tbox_TicketID.Location = new System.Drawing.Point(131, 73);
            this.tbox_TicketID.Name = "tbox_TicketID";
            this.tbox_TicketID.Size = new System.Drawing.Size(219, 20);
            this.tbox_TicketID.TabIndex = 1;
            this.tbox_TicketID.KeyDown += new System.Windows.Forms.KeyEventHandler(this.tbox_KeyDown);
            // 
            // btn_resetdstfolder
            // 
            this.btn_resetdstfolder.Location = new System.Drawing.Point(12, 3);
            this.btn_resetdstfolder.Name = "btn_resetdstfolder";
            this.btn_resetdstfolder.Size = new System.Drawing.Size(97, 24);
            this.btn_resetdstfolder.TabIndex = 3;
            this.btn_resetdstfolder.Text = "Zielordner Reset";
            this.btn_resetdstfolder.UseVisualStyleBackColor = true;
            this.btn_resetdstfolder.Click += new System.EventHandler(this.btn_resetdstfolder_Click);
            // 
            // lbl_info
            // 
            this.lbl_info.AutoSize = true;
            this.lbl_info.Location = new System.Drawing.Point(486, 72);
            this.lbl_info.Name = "lbl_info";
            this.lbl_info.Size = new System.Drawing.Size(40, 13);
            this.lbl_info.TabIndex = 4;
            this.lbl_info.Text = "lbl_info";
            // 
            // btn_ok
            // 
            this.btn_ok.BackColor = System.Drawing.Color.Lime;
            this.btn_ok.Location = new System.Drawing.Point(370, 72);
            this.btn_ok.Name = "btn_ok";
            this.btn_ok.Size = new System.Drawing.Size(89, 20);
            this.btn_ok.TabIndex = 5;
            this.btn_ok.Text = "Absenden";
            this.btn_ok.UseVisualStyleBackColor = false;
            this.btn_ok.Click += new System.EventHandler(this.btn_ok_Click);
            // 
            // lbl_Subject
            // 
            this.lbl_Subject.AutoSize = true;
            this.lbl_Subject.Location = new System.Drawing.Point(20, 44);
            this.lbl_Subject.Name = "lbl_Subject";
            this.lbl_Subject.Size = new System.Drawing.Size(94, 13);
            this.lbl_Subject.TabIndex = 6;
            this.lbl_Subject.Text = "Betreff bearbeiten:";
            // 
            // txtbox_Subject
            // 
            this.txtbox_Subject.Location = new System.Drawing.Point(131, 39);
            this.txtbox_Subject.Name = "txtbox_Subject";
            this.txtbox_Subject.Size = new System.Drawing.Size(834, 20);
            this.txtbox_Subject.TabIndex = 7;
            // 
            // checkNewTicket
            // 
            this.checkNewTicket.AutoSize = true;
            this.checkNewTicket.Location = new System.Drawing.Point(23, 112);
            this.checkNewTicket.Name = "checkNewTicket";
            this.checkNewTicket.Size = new System.Drawing.Size(132, 17);
            this.checkNewTicket.TabIndex = 8;
            this.checkNewTicket.Text = "Neues Ticket erstellen";
            this.checkNewTicket.UseVisualStyleBackColor = true;
            this.checkNewTicket.CheckedChanged += new System.EventHandler(this.checkNewTicket_CheckedChanged);
            // 
            // label_info
            // 
            this.label_info.AutoSize = true;
            this.label_info.Location = new System.Drawing.Point(822, 9);
            this.label_info.Name = "label_info";
            this.label_info.Size = new System.Drawing.Size(169, 13);
            this.label_info.TabIndex = 9;
            this.label_info.Text = "ToOTRS © 2017 JoHe - vX.X.X.X";
            // 
            // getTicketID
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1003, 142);
            this.Controls.Add(this.label_info);
            this.Controls.Add(this.checkNewTicket);
            this.Controls.Add(this.txtbox_Subject);
            this.Controls.Add(this.lbl_Subject);
            this.Controls.Add(this.btn_ok);
            this.Controls.Add(this.lbl_info);
            this.Controls.Add(this.btn_resetdstfolder);
            this.Controls.Add(this.tbox_TicketID);
            this.Controls.Add(this.labelTicketID);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "getTicketID";
            this.Text = "ToOTRS";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label labelTicketID;
        private System.Windows.Forms.TextBox tbox_TicketID;
        private System.Windows.Forms.Button btn_resetdstfolder;
        private System.Windows.Forms.Label lbl_info;
        private System.Windows.Forms.Button btn_ok;
        private System.Windows.Forms.Label lbl_Subject;
        private System.Windows.Forms.TextBox txtbox_Subject;
        private System.Windows.Forms.CheckBox checkNewTicket;
        private System.Windows.Forms.HelpProvider helpProvider1;
        private System.Windows.Forms.Label label_info;
    }
}