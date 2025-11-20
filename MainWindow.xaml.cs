using System;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace WordHighlighter
{
    public partial class MainWindow : Window
    {
        private WordInteropService? _wordService;
        private DispatcherTimer? _clickDetectionTimer;
        private DispatcherTimer? _scrollDetectionTimer;
        private WordReplacementWidget? _currentWidget = null;
        private string _lastSelectedWord = string.Empty;
        private System.Collections.Generic.List<OverlayBox> _overlayBoxes = new();
        private int _lastScrollPosition = 0;

        public MainWindow()
        {
            InitializeComponent();
            LogMessage("🚀 Application started. Ready to connect to Microsoft Word.");
        }

        private void BtnConnectWord_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LogMessage("🔍 Attempting to connect to Microsoft Word...");
                txtStatusBar.Text = "Connecting to Word...";

                _wordService = new WordInteropService();

                if (_wordService.ConnectToWord())
                {
                    txtConnectionStatus.Text = "Connected";
                    statusIndicator.Fill = new SolidColorBrush(
                        Color.FromRgb(21, 195, 154)); // Green

                    var parentBorder = statusIndicator.Parent as StackPanel;
                    if (parentBorder?.Parent is Border badge)
                    {
                        badge.Background = new SolidColorBrush(
                            Color.FromRgb(220, 252, 231)); // Light green
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
            catch (InvalidOperationException ex) when (ex.Message.Contains("not installed"))
            {
                string errorMessage = "Microsoft Word is not installed on this system.\n\n" +
                                       "This application requires Microsoft Office Word to be installed.\n\n" +
                                       "Please install Microsoft Office and try again.";
                MessageBox.Show(errorMessage, "Microsoft Word Not Found", MessageBoxButton.OK, MessageBoxImage.Error);
                LogMessage($"⚠️ Error: {ex.Message}");
                txtStatusBar.Text = "Error • Word not installed";
            }
            catch (System.IO.FileNotFoundException ex) when (ex.Message.Contains("office"))
            {
                string errorMessage = "Microsoft Office assemblies could not be found.\n\n" +
                                       "Possible solutions:\n" +
                                       "1. Repair your Office installation\n" +
                                       "2. Ensure Office is properly registered\n" +
                                       "3. Run this application as Administrator\n\n" +
                                       $"Technical details: {ex.Message}";
                MessageBox.Show(errorMessage, "Office Components Missing", MessageBoxButton.OK, MessageBoxImage.Error);
                LogMessage($"⚠️ Assembly Error: {ex.Message}");
                txtStatusBar.Text = "Error • Office assemblies missing";
            }
            catch (Exception ex)
            {
                string errorMessage = $"Error connecting to Word:\n\n{ex.Message}\n\n";

                if (ex.InnerException != null)
                {
                    errorMessage += $"Details: {ex.InnerException.Message}";
                }

                MessageBox.Show(errorMessage, "Connection Error", MessageBoxButton.OK, MessageBoxImage.Error);
                LogMessage($"⚠️ Error: {ex.Message}");
                txtStatusBar.Text = "Error occurred • Check details";
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

                // Create overlay boxes for each word position
                var positions = _wordService.GetWordPositions();
                foreach (var pos in positions)
                {
                    var overlayBox = new OverlayBox(pos.Left, pos.Top, pos.Width, pos.Height, pos.Text);
                    overlayBox.OverlayClicked += OverlayBox_Clicked;
                    overlayBox.Show();
                    _overlayBoxes.Add(overlayBox);
                }

                LogMessage($"✨ Successfully highlighted {wordCount:N0} words with overlay boxes.");
                txtStatusBar.Text = $"{wordCount:N0} words highlighted • Click any word to edit";

                // Start scroll detection
                StartScrollDetection();

                MessageBox.Show($"Added overlay boxes to {wordCount:N0} words in the document.\n\nClick on any highlighted word in Word to replace it.",
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

                // Stop scroll detection
                _scrollDetectionTimer?.Stop();

                // Close all overlay boxes
                foreach (var box in _overlayBoxes)
                {
                    box.Close();
                }
                _overlayBoxes.Clear();

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

        private void StartScrollDetection()
        {
            _scrollDetectionTimer = new DispatcherTimer
            {
                Interval = TimeSpan.FromMilliseconds(300)
            };
            _scrollDetectionTimer.Tick += CheckForScroll;
            _scrollDetectionTimer.Start();
        }

        private void CheckForScroll(object? sender, EventArgs e)
        {
            try
            {
                if (_wordService == null) return;

                // Get current scroll position from Word
                var currentScroll = _wordService.GetScrollPosition();

                if (currentScroll != _lastScrollPosition)
                {
                    _lastScrollPosition = currentScroll;

                    // Redraw overlay boxes
                    RedrawOverlays();
                }
            }
            catch
            {
                // Ignore errors
            }
        }

        private void RedrawOverlays()
        {
            try
            {
                if (_wordService == null) return;

                // Recalculate positions
                _wordService.HighlightAllWords();
                var positions = _wordService.GetWordPositions();

                // Update existing overlay boxes
                for (int i = 0; i < Math.Min(_overlayBoxes.Count, positions.Count); i++)
                {
                    _overlayBoxes[i].UpdatePosition(
                        positions[i].Left,
                        positions[i].Top,
                        positions[i].Width,
                        positions[i].Height
                    );
                }
            }
            catch
            {
                // Ignore errors
            }
        }

        private void OverlayBox_Clicked(object? sender, string wordText)
        {
            try
            {
                if (_wordService == null) return;

                var overlayBox = sender as OverlayBox;
                if (overlayBox == null) return;

                // Prevent opening widget for same word immediately
                if (wordText == _lastSelectedWord) return;

                // Close existing widget if any
                if (_currentWidget != null)
                {
                    _currentWidget.Close();
                    _currentWidget = null;
                }

                _lastSelectedWord = wordText;

                // Show replacement widget positioned near the overlay box
                _currentWidget = new WordReplacementWidget(wordText, _wordService);
                _currentWidget.WindowStartupLocation = WindowStartupLocation.Manual;

                // Position widget near the overlay box
                _currentWidget.Left = overlayBox.Left + 20;
                _currentWidget.Top = overlayBox.Top - 100; // Above the overlay

                // Make sure widget stays on screen
                if (_currentWidget.Left + _currentWidget.Width > SystemParameters.VirtualScreenWidth)
                    _currentWidget.Left = SystemParameters.VirtualScreenWidth - _currentWidget.Width - 20;
                if (_currentWidget.Top < 0)
                    _currentWidget.Top = overlayBox.Top + overlayBox.Height + 5; // Below overlay if above doesn't fit

                // Handle word replacement
                _currentWidget.WordReplaced += (s, replacementWord) =>
                {
                    LogMessage($"📝 Replaced '{wordText}' → '{replacementWord}'");
                    txtStatusBar.Text = $"Word replaced: {wordText} → {replacementWord}";

                    // Remove overlay box for replaced word
                    var boxToRemove = _overlayBoxes.FirstOrDefault(b => b.WordText == wordText);
                    if (boxToRemove != null)
                    {
                        boxToRemove.Close();
                        _overlayBoxes.Remove(boxToRemove);
                    }
                };

                // Handle widget closing
                _currentWidget.Closed += (s, args) =>
                {
                    _currentWidget = null;
                    // Reset last selected word after delay so popup doesn't reopen immediately
                    System.Threading.Tasks.Task.Delay(500).ContinueWith(_ =>
                    {
                        Dispatcher.Invoke(() => _lastSelectedWord = string.Empty);
                    });
                };

                _currentWidget.Show();
            }
            catch (Exception ex)
            {
                LogMessage($"Overlay click error: {ex.Message}");
            }
        }

        private void CheckForWordClick(object? sender, EventArgs e)
        {
            // This method is no longer needed since we're using overlay click events
            // Keeping it empty to avoid breaking the timer
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
            _scrollDetectionTimer?.Stop();

            // Close any open widget
            if (_currentWidget != null)
            {
                _currentWidget.Close();
                _currentWidget = null;
            }

            // Close all overlay boxes
            foreach (var box in _overlayBoxes)
            {
                box.Close();
            }
            _overlayBoxes.Clear();

            _wordService?.Dispose();

            LogMessage("Application closing...");
        }
    }
}
