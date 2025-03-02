using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using HandyControl.Controls;
using Microsoft.Web.WebView2.Core;

namespace PresPio
    {
    public partial class TaskPanelController : UserControl
        {
        private CoreWebView2DownloadOperation _downloadOperation;
        private bool _isWebViewInitialized = false;
        private TaskCompletionSource<bool> _initializationCompletionSource;

        public TaskPanelController()
            {
            InitializeComponent();
            _initializationCompletionSource = new TaskCompletionSource<bool>();
            contentWebView.VisibleChanged += contentWebView_VisibleChanged; // 绑定事件
            }

        /// <summary>
        /// 缓存文件夹地址
        /// </summary>
        public static string CachePath { get; } = System.IO.Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
            "PresPioCache"
        );

        private async Task EnsureInitializedAsync()
            {
            if (_isWebViewInitialized)
                {
                return;
                }

            try
                {
                if (contentWebView == null)
                    {
                    throw new InvalidOperationException("WebView2 control is not properly initialized");
                    }

                var options = new CoreWebView2EnvironmentOptions("--allow-file-access-from-files");
                var env = await CoreWebView2Environment.CreateAsync(null, CachePath, options);
                await contentWebView.EnsureCoreWebView2Async(env);

                ConfigureWebView();
                _isWebViewInitialized = true;
                _initializationCompletionSource.TrySetResult(true);
                }
            catch (Exception ex)
                {
                _initializationCompletionSource.TrySetException(ex);
                Growl.WarningGlobal($"WebView2初始化失败: {ex.Message}");
                throw;
                }
            }

        private void ConfigureWebView()
            {
            if (contentWebView?.CoreWebView2 == null)
                {
                return;
                }

            contentWebView.CoreWebView2.Settings.IsStatusBarEnabled = false;
            contentWebView.CoreWebView2.Settings.AreDevToolsEnabled = false;
            contentWebView.CoreWebView2.Settings.IsZoomControlEnabled = false;
            contentWebView.CoreWebView2.Settings.AreDefaultContextMenusEnabled = false;

            contentWebView.CoreWebView2.NewWindowRequested += CoreWebView2_NewWindowRequested;
            contentWebView.CoreWebView2.DownloadStarting += CoreWebView2_DownloadStarting;
            contentWebView.CoreWebView2.NavigationCompleted += CoreWebView2_NavigationCompleted;
            }

        private void CoreWebView2_NavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
            {
            if (!e.IsSuccess)
                {
                //  Growl.WarningGlobal($"页面加载失败: {e.WebErrorStatus}");
                }
            }

        public async Task NavigateToUrlAsync(string url)
            {
            if (string.IsNullOrEmpty(url))
                {
                throw new ArgumentNullException(nameof(url));
                }

            try
                {
                await EnsureInitializedAsync();

                if (contentWebView?.CoreWebView2 != null)
                    {
                    contentWebView.CoreWebView2.NavigateToString("");
                    contentWebView.CoreWebView2.Navigate(url);
                    }
                }
            catch (Exception ex)
                {
                Growl.WarningGlobal($"导航失败: {ex.Message}");
                throw;
                }
            }

        public void NavigateToUrl(string url)
            {
            _ = NavigateToUrlAsync(url);
            }

        private void CoreWebView2_NewWindowRequested(object sender, CoreWebView2NewWindowRequestedEventArgs e)
            {
            if (contentWebView?.CoreWebView2 != null)
                {
                e.NewWindow = contentWebView.CoreWebView2;
                e.Handled = true;
                }
            }

        private void CoreWebView2_DownloadStarting(object sender, CoreWebView2DownloadStartingEventArgs e)
            {
            _downloadOperation = e.DownloadOperation;
            }

        public async Task ClearCacheAsync()
            {
            try
                {
                if (contentWebView?.CoreWebView2?.Profile != null)
                    {
                    await contentWebView.CoreWebView2.Profile.ClearBrowsingDataAsync();
                    }
                }
            catch (Exception ex)
                {
                Growl.WarningGlobal($"清除缓存失败: {ex.Message}");
                }
            }

        protected override void Dispose(bool disposing)
            {
            if (disposing)
                {
                if (contentWebView != null)
                    {
                    contentWebView.CoreWebView2?.Stop();
                    contentWebView.Dispose();
                    }

                if (_downloadOperation != null)
                    {
                    _downloadOperation.Cancel();
                    _downloadOperation = null;
                    }
                }
            base.Dispose(disposing);
            }

        private void contentWebView_VisibleChanged(object sender, EventArgs e)
            {
            // 检测窗体是否不可见
            if (!this.Visible)
                {
                // 调用 ThisAddIn 类的公共方法关闭自定义窗格
                Globals.ThisAddIn.CloseCustomTaskPane();
                }
            }
        }
    }