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

        public WordReplacementWidget(string originalWord, WordInteropService wordService)
        {
            InitializeComponent();
            
            _originalWord = originalWord;
            _wordService = wordService;
            
            txtOriginalWord.Text = originalWord;
            txtReplacement.Text = originalWord;
            txtReplacement.Focus();
            txtReplacement.SelectAll();
        }

        private void BtnReplace_Click(object sender, RoutedEventArgs e)
        {
            PerformReplacement();
        }

        private void BtnCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
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
                DialogResult = false;
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
                DialogResult = true;
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
