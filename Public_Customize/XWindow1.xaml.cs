using Microsoft.Web.WebView2.Core;
using System;
using System.ComponentModel;
using System.Windows;

namespace PresPio
    {
    public partial class XWindow1 : Window
        {
        private static string tempUrl = "http://www.baidu.com";
        private bool isWebViewInitialized = false;

        public XWindow1()
            {
            InitializeComponent();
            contentWebView.CoreWebView2InitializationCompleted += WebView_CoreWebView2InitializationCompleted;
            contentWebView.NavigationCompleted += ContentWebView_NavigationCompleted;
            }

        private async void Window_Loaded(object sender, RoutedEventArgs e)
            {
            // 设置 WebView2 的 UserDataFolder
            var env = await CoreWebView2Environment.CreateAsync(null, "path/to/user/data/folder");
            await contentWebView.EnsureCoreWebView2Async(env);
            }

        private void WebView_CoreWebView2InitializationCompleted(object sender, CoreWebView2InitializationCompletedEventArgs e)
            {
            if (e.IsSuccess)
                {
                isWebViewInitialized = true;
                SetMobileUserAgent();
                if (!string.IsNullOrEmpty(tempUrl))
                    {
                    contentWebView.CoreWebView2.Navigate(tempUrl);
                    }
                }
            else
                {
                MessageBox.Show($"WebView2 初始化失败: {e.InitializationException}");
                }
            }

        private void SetMobileUserAgent()
            {
            contentWebView.CoreWebView2.Settings.UserAgent = "Mozilla/5.0 (iPhone; CPU iPhone OS 13_3 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/13.0 Mobile/15E148 Safari/604.1";
            }

        private void ContentWebView_NavigationCompleted(object sender, CoreWebView2NavigationCompletedEventArgs e)
            {
            }

        private async void SetFocusToInputElement()
            {
            try
                {
                await contentWebView.CoreWebView2.ExecuteScriptAsync("document.querySelector('input').focus();");
                }
            catch (Exception ex)
                {
                MessageBox.Show($"设置焦点到输入框失败: {ex.Message}");
                }
            }

        public void OpenWindow(string url, int width = 800, int height = 450)
            {
            this.Width = width;
            this.Height = height;
            if (isWebViewInitialized)
                {
                contentWebView.CoreWebView2.Navigate(url);
                }
            else
                {
                tempUrl = url;
                Show();
                }
            CenterWindowOnScreen();
            this.Activate();
            }

        private void CenterWindowOnScreen()
            {
            double screenWidth = SystemParameters.PrimaryScreenWidth;
            double screenHeight = SystemParameters.PrimaryScreenHeight;
            double windowWidth = this.Width;
            double windowHeight = this.Height;

            this.Left = (screenWidth / 2) - (windowWidth / 2);
            this.Top = (screenHeight / 2) - (windowHeight / 2);
            }

        protected override void OnClosing(CancelEventArgs e)
            {
            e.Cancel = true;
            this.Left -= 10000;
            ShowInTaskbar = false;
            }

        private void OpenUrlButton_Click(object sender, RoutedEventArgs e)
            {
            string url = "http://www.example.com"; // 修正 URL 字符串
            OpenWindow(url);
            }

        private void xWindow_Closed(object sender, EventArgs e)
            {
            this.Close();
            }
        }
    }