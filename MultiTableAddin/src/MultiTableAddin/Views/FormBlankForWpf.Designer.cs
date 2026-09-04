namespace MultiTableAddin.Views;

partial class FormBlankForWpf
{
    private System.ComponentModel.IContainer components = null;

    protected override void Dispose(bool disposing)
    {
        if (disposing && components != null)
        {
            components.Dispose();
        }

        base.Dispose(disposing);
    }

    private void InitializeComponent()
    {
        SuspendLayout();
        AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
        AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
        BackColor = System.Drawing.Color.White;
        ClientSize = new System.Drawing.Size(1200, 860);
        Name = "FormBlankForWpf";
        ShowIcon = false;
        StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
        FormClosing += FormBlankForWpf_FormClosing;
        Load += FormBlankForWpf_Load;
        Shown += FormBlankForWpf_Shown;
        ResumeLayout(false);
    }
}
