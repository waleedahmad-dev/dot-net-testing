using System;
using System.Globalization;
using System.Windows;
using System.Windows.Input;

namespace WordHighlighter
{
    public partial class WordReplacementWidget : Window
    {
        private readonly string _originalWord;
        private readonly WordInteropService _wordService;

        public string ReplacementWord { get; private set; } = string.Empty;
        public event EventHandler<string>? WordReplaced;

        public WordReplacementWidget(string originalWord, WordInteropService wordService)
        {
            InitializeComponent();

            // Make window topmost to appear above Word
            Topmost = true;

            _originalWord = originalWord;
            _wordService = wordService;

            txtOriginalWord.Text = originalWord;
            txtReplacement.Text = originalWord;
            txtReplacement.Focus();
            txtReplacement.SelectAll();

            // Select the word in the document when widget opens
            _wordService.SelectWordInDocument(originalWord);
        }

        private void BtnReplace_Click(object sender, RoutedEventArgs e)
        {
            PerformReplacement();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void TxtReplacement_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                PerformReplacement();
            }
            else if (e.Key == Key.Escape)
            {
                Close();
            }
        }

        private void BtnUppercase_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(txtReplacement.Text))
            {
                txtReplacement.Text = txtReplacement.Text.ToUpper();
            }
        }

        private void BtnLowercase_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(txtReplacement.Text))
            {
                txtReplacement.Text = txtReplacement.Text.ToLower();
            }
        }

        private void BtnCapitalize_Click(object sender, RoutedEventArgs e)
        {
            if (!string.IsNullOrEmpty(txtReplacement.Text))
            {
                var textInfo = CultureInfo.CurrentCulture.TextInfo;
                txtReplacement.Text = textInfo.ToTitleCase(txtReplacement.Text.ToLower());
            }
        }

        private void PerformReplacement()
        {
            string replacement = txtReplacement.Text.Trim();

            if (string.IsNullOrWhiteSpace(replacement))
            {
                txtError.Text = "⚠️ Please enter a replacement word.";
                errorBorder.Visibility = Visibility.Visible;
                return;
            }

            try
            {
                _wordService.ReplaceSelectedWord(replacement);
                ReplacementWord = replacement;
                WordReplaced?.Invoke(this, replacement);
                Close();
            }
            catch (Exception ex)
            {
                txtError.Text = $"⚠️ Error: {ex.Message}";
                errorBorder.Visibility = Visibility.Visible;
            }
        }
    }
}
