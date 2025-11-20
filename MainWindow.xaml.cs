using System;
using System.Windows;
using System.Windows.Threading;

namespace WordHighlighter
{
    public partial class MainWindow : Window
    {
        private WordInteropService? _wordService;
        private DispatcherTimer? _clickDetectionTimer;

        public MainWindow()
        {
            InitializeComponent();
            LogMessage("🚀 Application started. Ready to connect to Microsoft Word.");
        }

        private void BtnConnectWord_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LogMessage("Attempting to connect to Microsoft Word...");
                txtStatusBar.Text = "Connecting to Word...";

                _wordService = new WordInteropService();
                
                if (_wordService.ConnectToWord())
                {
                    txtConnectionStatus.Text = "Connected";
                    statusIndicator.Fill = new System.Windows.Media.SolidColorBrush(
                        System.Windows.Media.Color.FromRgb(21, 195, 154)); // Green
                    
                    var parentBorder = statusIndicator.Parent as StackPanel;
                    if (parentBorder?.Parent is Border badge)
                    {
                        badge.Background = new System.Windows.Media.SolidColorBrush(
                            System.Windows.Media.Color.FromRgb(220, 252, 231)); // Light green
                    }
                    
                    btnHighlightAll.IsEnabled = true;
                    btnClearHighlights.IsEnabled = true;
                    btnConnectWord.IsEnabled = false;

                    LogMessage("✅ Successfully connected to Microsoft Word.");
                    txtStatusBar.Text = "Connected to Word • Ready to highlight";

                    // Start monitoring for clicks
                    StartClickDetection();
                }
                else
                {
                    MessageBox.Show("Could not connect to Word. Please make sure Microsoft Word is open with a document.", 
                        "Connection Failed", MessageBoxButton.OK, MessageBoxImage.Warning);
                    LogMessage("❌ Failed to connect. Please ensure Word is running.");
                    txtStatusBar.Text = "Connection failed • Please open Word";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error connecting to Word: {ex.Message}", 
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                LogMessage($"⚠️ Error: {ex.Message}");
                txtStatusBar.Text = "Error occurred • Please try again";
            }
        }

        private void BtnHighlightAll_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_wordService == null) return;

                LogMessage("🎨 Highlighting all words in the document...");
                txtStatusBar.Text = "Processing • Highlighting words...";

                int wordCount = _wordService.HighlightAllWords();
                
                LogMessage($"✨ Successfully highlighted {wordCount:N0} words.");
                txtStatusBar.Text = $"{wordCount:N0} words highlighted • Click any word to edit";

                MessageBox.Show($"Highlighted {wordCount:N0} words in the document.\n\nClick on any highlighted word in Word to replace it.", 
                    "Highlighting Complete", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error highlighting words: {ex.Message}", 
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                LogMessage($"⚠️ Error: {ex.Message}");
                txtStatusBar.Text = "Error occurred • Please try again";
            }
        }

        private void BtnClearHighlights_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (_wordService == null) return;

                LogMessage("🧹 Clearing all highlights...");
                txtStatusBar.Text = "Processing • Clearing highlights...";

                _wordService.ClearAllHighlights();
                
                LogMessage("✓ All highlights cleared.");
                txtStatusBar.Text = "Highlights cleared • Document restored";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error clearing highlights: {ex.Message}", 
                    "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                LogMessage($"⚠️ Error: {ex.Message}");
                txtStatusBar.Text = "Error occurred • Please try again";
            }
        }

        private void StartClickDetection()
        {
            _clickDetectionTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(500)
            };
            _clickDetectionTimer.Tick += CheckForWordClick;
            _clickDetectionTimer.Start();
            LogMessage("👆 Click detection started. You can now click highlighted words in Word.");
        }

        private void CheckForWordClick(object? sender, EventArgs e)
        {
            try
            {
                if (_wordService == null) return;

                var selectedWord = _wordService.GetSelectedWord();
                if (!string.IsNullOrEmpty(selectedWord))
                {
                    // Show replacement widget
                    var widget = new WordReplacementWidget(selectedWord, _wordService);
                    widget.Owner = this;
                    
                    if (widget.ShowDialog() == true)
                    {
                        LogMessage($"📝 Replaced '{selectedWord}' → '{widget.ReplacementWord}'");
                        txtStatusBar.Text = $"Word replaced: {selectedWord} → {widget.ReplacementWord}";
                    }
                }
            }
            catch (Exception ex)
            {
                // Silently log errors to avoid interrupting the user
                LogMessage($"Detection error: {ex.Message}");
            }
        }

        private void LogMessage(string message)
        {
            string timestamp = DateTime.Now.ToString("HH:mm:ss");
            txtLog.Text += $"[{timestamp}] {message}\n";
        }

        protected override void OnClosed(EventArgs e)
        {
            base.OnClosed(e);
            
            _clickDetectionTimer?.Stop();
            _wordService?.Dispose();
            
            LogMessage("Application closing...");
        }
    }
}
