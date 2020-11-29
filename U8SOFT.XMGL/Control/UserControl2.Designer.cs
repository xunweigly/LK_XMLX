namespace U8SOFT.XMRZ
{
    partial class UserControl2
    {
        /// <summary> 
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 组件设计器生成的代码

        /// <summary> 
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.粘贴 = new System.Windows.Forms.ToolStripMenuItem();
            this.清除 = new System.Windows.Forms.ToolStripMenuItem();
            this.另存为 = new System.Windows.Forms.ToolStripMenuItem();
            this.contextMenuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this)).BeginInit();
            this.SuspendLayout();
            // 
            // contextMenuStrip1
            // 
            this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.粘贴,
            this.清除,
            this.另存为});
            this.contextMenuStrip1.Name = "contextMenuStrip1";
            this.contextMenuStrip1.Size = new System.Drawing.Size(113, 70);
            // 
            // 粘贴
            // 
            this.粘贴.Name = "粘贴";
            this.粘贴.Size = new System.Drawing.Size(112, 22);
            this.粘贴.Text = "粘贴";
            this.粘贴.Click += new System.EventHandler(this.粘贴ToolStripMenuItem_Click);
            // 
            // 清除
            // 
            this.清除.Name = "清除";
            this.清除.Size = new System.Drawing.Size(112, 22);
            this.清除.Text = "清除";
            this.清除.Click += new System.EventHandler(this.清除ToolStripMenuItem_Click);
            // 
            // 另存为
            // 
            this.另存为.Name = "另存为";
            this.另存为.Size = new System.Drawing.Size(112, 22);
            this.另存为.Text = "另存为";
            this.另存为.Click += new System.EventHandler(this.另存为_Click);
            // 
            // UserControl2
            // 
            this.ContextMenuStrip = this.contextMenuStrip1;
            this.contextMenuStrip1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.ContextMenuStrip contextMenuStrip1;
        private System.Windows.Forms.ToolStripMenuItem 粘贴;
        private System.Windows.Forms.ToolStripMenuItem 清除;
        private System.Windows.Forms.ToolStripMenuItem 另存为;

    }
}
