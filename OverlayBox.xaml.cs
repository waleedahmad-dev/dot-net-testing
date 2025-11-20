using System;
using System.Windows;
using System.Windows.Input;

namespace WordHighlighter
{
    public partial class OverlayBox : Window
    {
        public string WordText { get; set; }
        public event EventHandler<string>? OverlayClicked;

        public OverlayBox(double left, double top, double width, double height, string wordText)
        {
            InitializeComponent();

            WordText = wordText;

            // Calculate proper width based on word length
            double calculatedWidth = Math.Max(width, wordText.Length * 8.5);
            double calculatedHeight = Math.Max(height, 18);

            // Position the overlay (adjust top to align better)
            this.Left = left;
            this.Top = top + 2; // Shift down slightly for better alignment
            this.Width = calculatedWidth;
            this.Height = calculatedHeight;
        }

        public void UpdatePosition(double left, double top, double width, double height)
        {
            this.Left = left;
            this.Top = top + 2;
            this.Width = Math.Max(width, WordText.Length * 8.5);
            this.Height = Math.Max(height, 18);
        }

        private void OverlayBorder_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            // Raise event when overlay is clicked
            OverlayClicked?.Invoke(this, WordText);
        }
    }
}
