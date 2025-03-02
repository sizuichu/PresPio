using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Speech.Synthesis;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Threading;
using HandyControl.Controls;
using Microsoft.Win32;
using TextBox = System.Windows.Controls.TextBox;

namespace PresPio.Public_Wpf
    {
    public partial class Wpf_VoiceAssistant : INotifyPropertyChanged
        {
        private SpeechSynthesizer synthesizer;
        private ObservableCollection<InstalledVoice> windowsVoices;
        private bool isPlaying;
        private DispatcherTimer playbackTimer;
        private TimeSpan currentPlaybackPosition;
        private string currentPlaybackTime;
        private TextBox textContent;
        private ProgressBar playbackProgressBar;
        private Button playPauseButton;
        private Slider speechSpeedSlider;
        private Slider speechVolumeSlider;
        private Slider speechPitchSlider;

        public event PropertyChangedEventHandler PropertyChanged;

        public ObservableCollection<InstalledVoice> WindowsVoices
            {
            get => windowsVoices;
            set
                {
                windowsVoices = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(WindowsVoices)));
                }
            }

        public string CurrentPlaybackTime
            {
            get => currentPlaybackTime;
            set
                {
                currentPlaybackTime = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(CurrentPlaybackTime)));
                }
            }

        public Wpf_VoiceAssistant()
            {
            InitializeComponent();
            InitializeControls();
            InitializeSpeechSynthesizer();
            InitializePlaybackTimer();
            DataContext = this;
            }

        private void InitializeControls()
            {
            textContent = (TextBox)FindName("TextContent");
            playbackProgressBar = (ProgressBar)FindName("PlaybackProgressBar");
            playPauseButton = (Button)FindName("PlayPauseButton");
            speechSpeedSlider = (Slider)FindName("SpeechSpeedSlider");
            speechVolumeSlider = (Slider)FindName("SpeechVolumeSlider");
            speechPitchSlider = (Slider)FindName("SpeechPitchSlider");
            }

        private void InitializeSpeechSynthesizer()
            {
            synthesizer = new SpeechSynthesizer();
            WindowsVoices = new ObservableCollection<InstalledVoice>(synthesizer.GetInstalledVoices());

            if (WindowsVoices.Count > 0)
                {
                synthesizer.SelectVoice(WindowsVoices[0].VoiceInfo.Name);
                }

            synthesizer.SpeakCompleted += Synthesizer_SpeakCompleted;
            synthesizer.SpeakProgress += Synthesizer_SpeakProgress;
            }

        private void InitializePlaybackTimer()
            {
            playbackTimer = new DispatcherTimer();
            playbackTimer.Interval = TimeSpan.FromSeconds(1);
            playbackTimer.Tick += PlaybackTimer_Tick;
            CurrentPlaybackTime = "00:00:00";
            }

        private void PlaybackTimer_Tick(object sender, EventArgs e)
            {
            if (isPlaying)
                {
                currentPlaybackPosition = currentPlaybackPosition.Add(TimeSpan.FromSeconds(1));
                CurrentPlaybackTime = currentPlaybackPosition.ToString(@"hh\:mm\:ss");
                }
            }

        private void Synthesizer_SpeakProgress(object sender, SpeakProgressEventArgs e)
            {
            Dispatcher.Invoke(() =>
            {
                double progress = (double)e.CharacterPosition / e.CharacterCount * 100;
                playbackProgressBar.Value = progress;
            });
            }

        private void OnPlayButtonClick(object sender, RoutedEventArgs e)
            {
            try
                {
                if (string.IsNullOrWhiteSpace(textContent.Text))
                    {
                    Growl.Warning("请输入要转换的文字");
                    return;
                    }

                if (!isPlaying)
                    {
                    StartPlayback();
                    }
                else
                    {
                    StopPlayback();
                    }
                }
            catch (Exception ex)
                {
                Growl.Error($"播放操作失败: {ex.Message}");
                }
            }

        private void StartPlayback()
            {
            try
                {
                isPlaying = true;
                currentPlaybackPosition = TimeSpan.Zero;
                playbackTimer.Start();
                playbackProgressBar.Value = 0;
                synthesizer.SpeakAsync(textContent.Text);
                UpdatePlayButton(true);
                }
            catch (Exception ex)
                {
                Growl.Error($"开始播放失败: {ex.Message}");
                StopPlayback();
                }
            }

        private void StopPlayback()
            {
            try
                {
                if (synthesizer != null)
                    {
                    synthesizer.SpeakAsyncCancelAll();
                    isPlaying = false;
                    playbackTimer.Stop();
                    playbackProgressBar.Value = 0;
                    UpdatePlayButton(false);
                    }
                }
            catch (Exception ex)
                {
                Growl.Error($"停止播放失败: {ex.Message}");
                }
            }

        private void UpdatePlayButton(bool playing)
            {
            var path = new System.Windows.Shapes.Path
                {
                Data = (Geometry)FindResource(playing ? "PauseGeometry" : "PlayGeometry"),
                Fill = Brushes.White,
                Width = 24,
                Height = 24,
                Stretch = Stretch.Uniform
                };
            playPauseButton.Content = path;
            }

        private void Synthesizer_SpeakCompleted(object sender, SpeakCompletedEventArgs e)
            {
            isPlaying = false;
            playbackTimer.Stop();
            Dispatcher.Invoke(() =>
            {
                UpdatePlayButton(false);
                CurrentPlaybackTime = "00:00:00";
                currentPlaybackPosition = TimeSpan.Zero;
                playbackProgressBar.Value = 0;
            });
            }

        private void OnSpeedChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
            {
            if (synthesizer != null)
                {
                try
                    {
                    // 将0-3的值映射到-10到10的范围
                    int rate = (int)((e.NewValue - 1.5) * 13.33);
                    synthesizer.Rate = Math.Max(-10, Math.Min(10, rate));

                    // 如果正在播放，则重新开始播放以应用新的语速
                    if (isPlaying)
                        {
                        string currentText = textContent.Text;
                        StopPlayback();
                        textContent.Text = currentText;
                        StartPlayback();
                        }
                    }
                catch (Exception ex)
                    {
                    Growl.Error($"设置语速失败: {ex.Message}");
                    }
                }
            }

        private void OnVolumeChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
            {
            if (synthesizer != null)
                {
                try
                    {
                    synthesizer.Volume = (int)e.NewValue;
                    }
                catch (Exception ex)
                    {
                    Growl.Error($"设置音量失败: {ex.Message}");
                    }
                }
            }

        private void OnPitchChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
            {
            if (synthesizer != null && !string.IsNullOrEmpty(textContent.Text))
                {
                try
                    {
                    // 停止当前播放
                    if (isPlaying)
                        {
                        StopPlayback();
                        }

                    // 使用SSML设置音高
                    string ssml = $@"<speak version='1.0' xmlns='http://www.w3.org/2001/10/synthesis' xml:lang='zh-CN'>
                        <prosody pitch='{e.NewValue}st'>{textContent.Text}</prosody>
                    </speak>";

                    // 如果之前在播放，则继续播放
                    if (isPlaying)
                        {
                        synthesizer.SpeakSsmlAsync(ssml);
                        StartPlayback();
                        }
                    }
                catch (Exception ex)
                    {
                    Growl.Error($"设置音高失败: {ex.Message}");
                    }
                }
            }

        private void OnSaveButtonClick(object sender, RoutedEventArgs e)
            {
            if (string.IsNullOrWhiteSpace(textContent.Text))
                {
                Growl.Warning("请先输入要保存的文字");
                return;
                }

            SaveFileDialog saveFileDialog = new SaveFileDialog
                {
                Filter = "音频文件 (*.wav)|*.wav",
                DefaultExt = ".wav",
                AddExtension = true
                };

            if (saveFileDialog.ShowDialog() == true)
                {
                try
                    {
                    synthesizer.SetOutputToWaveFile(saveFileDialog.FileName);
                    synthesizer.Speak(textContent.Text);
                    synthesizer.SetOutputToDefaultAudioDevice();
                    Growl.Success("文件保存成功");
                    }
                catch (Exception ex)
                    {
                    Growl.Error($"保存失败: {ex.Message}");
                    }
                }
            }

        private void OnEffectButtonClick(object sender, RoutedEventArgs e)
            {
            DrawerEffects.IsOpen = true;
            }

        private void OnApplyEffectsClick(object sender, RoutedEventArgs e)
            {
            // 应用音频效果
            if (EchoEffect.IsChecked == true)
                {
                // 添加回声效果
                }
            if (ReverbEffect.IsChecked == true)
                {
                // 添加混响效果
                }
            if (ChorusEffect.IsChecked == true)
                {
                // 添加合唱效果
                }
            DrawerEffects.IsOpen = false;
            Growl.Success("已应用音频效果");
            }

        private void OnVoiceEffectButtonClick(object sender, RoutedEventArgs e)
            {
            if (sender is Button button)
                {
                switch (button.Tag.ToString())
                    {
                    case "Normal":
                        speechSpeedSlider.Value = 1;
                        speechVolumeSlider.Value = 100;
                        speechPitchSlider.Value = 0;
                        break;

                    case "Deep":
                        speechSpeedSlider.Value = 0.8;
                        speechVolumeSlider.Value = 100;
                        speechPitchSlider.Value = -5;
                        break;

                    case "High":
                        speechSpeedSlider.Value = 1.2;
                        speechVolumeSlider.Value = 90;
                        speechPitchSlider.Value = 5;
                        break;

                    case "Fast":
                        speechSpeedSlider.Value = 2;
                        speechVolumeSlider.Value = 100;
                        speechPitchSlider.Value = 0;
                        break;

                    case "Slow":
                        speechSpeedSlider.Value = 0.5;
                        speechVolumeSlider.Value = 100;
                        speechPitchSlider.Value = 0;
                        break;
                    }
                }
            }

        private void OnImportButtonClick(object sender, RoutedEventArgs e)
            {
            OpenFileDialog openFileDialog = new OpenFileDialog
                {
                Filter = "文本文件 (*.txt)|*.txt|所有文件 (*.*)|*.*"
                };

            if (openFileDialog.ShowDialog() == true)
                {
                try
                    {
                    textContent.Text = File.ReadAllText(openFileDialog.FileName);
                    Growl.Success("文本导入成功");
                    }
                catch (Exception ex)
                    {
                    Growl.Error($"导入失败: {ex.Message}");
                    }
                }
            }

        private void OnBatchConvertButtonClick(object sender, RoutedEventArgs e)
            {
            if (string.IsNullOrWhiteSpace(textContent.Text))
                {
                Growl.Warning("请先输入要转换的文字");
                return;
                }

            var folderDialog = new System.Windows.Forms.FolderBrowserDialog();
            if (folderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                try
                    {
                    string[] paragraphs = textContent.Text.Split(new[] { "\r\n", "\r", "\n" },
                        StringSplitOptions.RemoveEmptyEntries);

                    for (int i = 0 ; i < paragraphs.Length ; i++)
                        {
                        string outputPath = System.IO.Path.Combine(folderDialog.SelectedPath, $"audio_{i + 1}.wav");
                        synthesizer.SetOutputToWaveFile(outputPath);
                        synthesizer.Speak(paragraphs[i]);
                        synthesizer.SetOutputToDefaultAudioDevice();
                        }

                    Growl.Success("批量转换完成");
                    }
                catch (Exception ex)
                    {
                    Growl.Error($"批量转换失败: {ex.Message}");
                    }
                }
            }

        protected override void OnClosing(CancelEventArgs e)
            {
            base.OnClosing(e);
            if (synthesizer != null)
                {
                StopPlayback();
                synthesizer.Dispose();
                }
            if (playbackTimer != null)
                {
                playbackTimer.Stop();
                }
            }

        private void OnVoiceSelectionChanged(object sender, SelectionChangedEventArgs e)
            {
            try
                {
                if (e.AddedItems.Count > 0 && e.AddedItems[0] is InstalledVoice selectedVoice)
                    {
                    synthesizer.SelectVoice(selectedVoice.VoiceInfo.Name);

                    // 如果正在播放，则重新开始播放以应用新的语音
                    if (isPlaying)
                        {
                        string currentText = textContent.Text;
                        StopPlayback();
                        textContent.Text = currentText;
                        StartPlayback();
                        }
                    }
                }
            catch (Exception ex)
                {
                Growl.Error($"切换语音失败: {ex.Message}");
                }
            }

        private void OnSettingButtonClick(object sender, RoutedEventArgs e)
            {
            DrawerSettings.IsOpen = true;
            }

        private void OnVoiceRadioButtonChecked(object sender, RoutedEventArgs e)
            {
            try
                {
                if (sender is RadioButton radioButton && radioButton.DataContext is InstalledVoice selectedVoice)
                    {
                    synthesizer.SelectVoice(selectedVoice.VoiceInfo.Name);

                    // 如果正在播放，则重新开始播放以应用新的语音
                    if (isPlaying)
                        {
                        string currentText = textContent.Text;
                        StopPlayback();
                        textContent.Text = currentText;
                        StartPlayback();
                        }
                    }
                }
            catch (Exception ex)
                {
                Growl.Error($"切换语音失败: {ex.Message}");
                }
            }
        }
    }