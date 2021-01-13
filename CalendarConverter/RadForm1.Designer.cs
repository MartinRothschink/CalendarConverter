
namespace CalendarConverter
{
   partial class RadForm1
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
         this.BtnLoad = new Telerik.WinControls.UI.RadButton();
         this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
         this.listBox1 = new System.Windows.Forms.ListBox();
         this.panel1 = new System.Windows.Forms.Panel();
         this.BtnSave = new Telerik.WinControls.UI.RadButton();
         this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
         ((System.ComponentModel.ISupportInitialize)(this.BtnLoad)).BeginInit();
         this.panel1.SuspendLayout();
         ((System.ComponentModel.ISupportInitialize)(this.BtnSave)).BeginInit();
         ((System.ComponentModel.ISupportInitialize)(this)).BeginInit();
         this.SuspendLayout();
         // 
         // BtnLoad
         // 
         this.BtnLoad.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
         this.BtnLoad.Location = new System.Drawing.Point(12, 12);
         this.BtnLoad.Name = "BtnLoad";
         this.BtnLoad.Size = new System.Drawing.Size(110, 24);
         this.BtnLoad.TabIndex = 0;
         this.BtnLoad.Text = "Load Word docx";
         this.BtnLoad.Click += new System.EventHandler(this.radButton1_Click);
         // 
         // openFileDialog1
         // 
         this.openFileDialog1.FileName = "openFileDialog1";
         this.openFileDialog1.Filter = "Word|*.docx";
         // 
         // listBox1
         // 
         this.listBox1.Dock = System.Windows.Forms.DockStyle.Fill;
         this.listBox1.Font = new System.Drawing.Font("Lucida Console", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
         this.listBox1.FormattingEnabled = true;
         this.listBox1.ItemHeight = 15;
         this.listBox1.Location = new System.Drawing.Point(0, 54);
         this.listBox1.Name = "listBox1";
         this.listBox1.Size = new System.Drawing.Size(734, 516);
         this.listBox1.TabIndex = 1;
         // 
         // panel1
         // 
         this.panel1.Controls.Add(this.BtnSave);
         this.panel1.Controls.Add(this.BtnLoad);
         this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
         this.panel1.Location = new System.Drawing.Point(0, 0);
         this.panel1.Name = "panel1";
         this.panel1.Size = new System.Drawing.Size(734, 54);
         this.panel1.TabIndex = 2;
         // 
         // BtnSave
         // 
         this.BtnSave.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
         this.BtnSave.Location = new System.Drawing.Point(128, 12);
         this.BtnSave.Name = "BtnSave";
         this.BtnSave.Size = new System.Drawing.Size(110, 24);
         this.BtnSave.TabIndex = 1;
         this.BtnSave.Text = "Save CSV";
         this.BtnSave.Click += new System.EventHandler(this.BtnSaveClick);
         // 
         // RadForm1
         // 
         this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
         this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
         this.ClientSize = new System.Drawing.Size(734, 570);
         this.Controls.Add(this.listBox1);
         this.Controls.Add(this.panel1);
         this.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
         this.Name = "RadForm1";
         // 
         // 
         // 
         this.RootElement.ApplyShapeToControl = true;
         this.Text = "Calendar Converter";
         ((System.ComponentModel.ISupportInitialize)(this.BtnLoad)).EndInit();
         this.panel1.ResumeLayout(false);
         ((System.ComponentModel.ISupportInitialize)(this.BtnSave)).EndInit();
         ((System.ComponentModel.ISupportInitialize)(this)).EndInit();
         this.ResumeLayout(false);

      }

      #endregion

      private Telerik.WinControls.UI.RadButton BtnLoad;
      private System.Windows.Forms.OpenFileDialog openFileDialog1;
      private System.Windows.Forms.ListBox listBox1;
      private System.Windows.Forms.Panel panel1;
      private Telerik.WinControls.UI.RadButton BtnSave;
      private System.Windows.Forms.SaveFileDialog saveFileDialog1;
   }
}