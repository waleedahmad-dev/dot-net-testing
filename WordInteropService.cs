using System;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Word;

namespace WordHighlighter
{
    public class WordInteropService : IDisposable
    {
        private Application? _wordApp;
        private Document? _activeDocument;
        private bool _disposed = false;

        public bool ConnectToWord()
        {
            try
            {
                // Try to get existing Word instance
                _wordApp = (Application)Marshal.GetActiveObject("Word.Application");
                
                if (_wordApp.Documents.Count == 0)
                {
                    return false;
                }

                _activeDocument = _wordApp.ActiveDocument;
                _wordApp.Visible = true;
                
                return true;
            }
            catch (COMException)
            {
                // Word is not running, try to start it
                try
                {
                    _wordApp = new Application();
                    _wordApp.Visible = true;
                    
                    // Create a new document
                    _activeDocument = _wordApp.Documents.Add();
                    
                    return true;
                }
                catch
                {
                    return false;
                }
            }
            catch
            {
                return false;
            }
        }

        public int HighlightAllWords()
        {
            if (_activeDocument == null)
                throw new InvalidOperationException("No active document connected.");

            int wordCount = 0;

            try
            {
                // Get all words in the document
                Range documentRange = _activeDocument.Range();
                Words words = documentRange.Words;

                foreach (Range wordRange in words)
                {
                    string text = wordRange.Text.Trim();
                    
                    // Skip empty strings, punctuation, and whitespace
                    if (!string.IsNullOrWhiteSpace(text) && 
                        text.Length > 1 && 
                        !IsPunctuation(text))
                    {
                        // Highlight the word with yellow background
                        wordRange.HighlightColorIndex = WdColorIndex.wdYellow;
                        wordCount++;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to highlight words: {ex.Message}", ex);
            }

            return wordCount;
        }

        public void ClearAllHighlights()
        {
            if (_activeDocument == null)
                throw new InvalidOperationException("No active document connected.");

            try
            {
                Range documentRange = _activeDocument.Range();
                documentRange.HighlightColorIndex = WdColorIndex.wdNoHighlight;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to clear highlights: {ex.Message}", ex);
            }
        }

        public string GetSelectedWord()
        {
            if (_wordApp == null || _activeDocument == null)
                return string.Empty;

            try
            {
                Selection selection = _wordApp.Selection;
                
                // Check if something is selected and highlighted
                if (selection != null && selection.Range.HighlightColorIndex == WdColorIndex.wdYellow)
                {
                    string selectedText = selection.Text?.Trim() ?? string.Empty;
                    
                    // Return the word if it's valid
                    if (!string.IsNullOrWhiteSpace(selectedText) && !IsPunctuation(selectedText))
                    {
                        return selectedText;
                    }
                }
            }
            catch
            {
                // Return empty string on error
            }

            return string.Empty;
        }

        public void ReplaceSelectedWord(string newWord)
        {
            if (_wordApp == null || _activeDocument == null)
                throw new InvalidOperationException("No active document connected.");

            try
            {
                Selection selection = _wordApp.Selection;
                
                if (selection != null && !string.IsNullOrEmpty(selection.Text))
                {
                    // Store the original highlight color
                    WdColorIndex originalHighlight = selection.Range.HighlightColorIndex;
                    
                    // Replace the text
                    selection.Text = newWord;
                    
                    // Restore highlight if it was highlighted before
                    if (originalHighlight == WdColorIndex.wdYellow)
                    {
                        selection.Range.HighlightColorIndex = WdColorIndex.wdYellow;
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to replace word: {ex.Message}", ex);
            }
        }

        private bool IsPunctuation(string text)
        {
            if (string.IsNullOrEmpty(text))
                return true;

            foreach (char c in text)
            {
                if (char.IsLetterOrDigit(c))
                    return false;
            }

            return true;
        }

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!_disposed)
            {
                if (disposing)
                {
                    // Release COM objects
                    if (_activeDocument != null)
                    {
                        Marshal.ReleaseComObject(_activeDocument);
                        _activeDocument = null;
                    }

                    if (_wordApp != null)
                    {
                        // Don't quit Word, just release the reference
                        Marshal.ReleaseComObject(_wordApp);
                        _wordApp = null;
                    }
                }

                _disposed = true;
            }
        }

        ~WordInteropService()
        {
            Dispose(false);
        }
    }
}
