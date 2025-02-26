namespace PresPio
    {
    partial class TaskPanelController
        {
        /// <summary> 
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

       

        #region 组件设计器生成的代码

        /// <summary> 
        /// 设计器支持所需的方法 - 不要修改
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
            {
            this.contentWebView = new Microsoft.Web.WebView2.WinForms.WebView2();
            ((System.ComponentModel.ISupportInitialize)(this.contentWebView)).BeginInit();
            this.SuspendLayout();
            // 
            // contentWebView
            // 
            this.contentWebView.AllowExternalDrop = true;
            this.contentWebView.CreationProperties = null;
            this.contentWebView.DefaultBackgroundColor = System.Drawing.Color.White;
            this.contentWebView.Dock = System.Windows.Forms.DockStyle.Fill;
            this.contentWebView.Location = new System.Drawing.Point(0, 0);
            this.contentWebView.Name = "contentWebView";
            this.contentWebView.Size = new System.Drawing.Size(341, 750);
            this.contentWebView.TabIndex = 0;
            this.contentWebView.ZoomFactor = 1D;
            this.contentWebView.VisibleChanged += new System.EventHandler(this.contentWebView_VisibleChanged);
            // 
            // TaskPanelController
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.contentWebView);
            this.Name = "TaskPanelController";
            this.Size = new System.Drawing.Size(341, 750);
            ((System.ComponentModel.ISupportInitialize)(this.contentWebView)).EndInit();
            this.ResumeLayout(false);

            }

        #endregion
        private Microsoft.Web.WebView2.WinForms.WebView2 contentWebView;
        }
    }
