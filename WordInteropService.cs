using System;
using System.Runtime.InteropServices;
using System.Reflection;
using System.Linq;

#if NET5_0_OR_GREATER
using System.Runtime.Versioning;
#endif

namespace WordHighlighter
{
#if NET5_0_OR_GREATER
    [SupportedOSPlatform("windows")]
#endif
    public class WordInteropService : IDisposable
    {
        private dynamic? _wordApp;
        private dynamic? _activeDocument;
        private bool _disposed = false;
        private System.Collections.Generic.List<WordPosition> _wordPositions = new();

        public struct WordPosition
        {
            public string Text { get; set; }
            public double Left { get; set; }
            public double Top { get; set; }
            public double Width { get; set; }
            public double Height { get; set; }
        }

        // Shape and color constants
        private const int msoShapeRectangle = 1;
        private const int msoTextOrientationHorizontal = 1;

        [DllImport("oleaut32.dll", PreserveSig = false)]
        private static extern void GetActiveObject(ref Guid rclsid, IntPtr pvReserved, [MarshalAs(UnmanagedType.IUnknown)] out object ppunk);

        [DllImport("user32.dll")]
        private static extern bool GetCursorPos(out POINT lpPoint);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        public bool ConnectToWord()
        {
            // First check if Office is installed by trying to create an instance
            Type? wordType = Type.GetTypeFromProgID("Word.Application");
            if (wordType == null)
            {
                throw new InvalidOperationException("Microsoft Word is not installed on this system. Please install Microsoft Office to use this application.");
            }

            try
            {
                // Try to get existing Word instance
                object? wordApp = null;
                try
                {
                    // Word Application CLSID
                    Guid clsid = new Guid("000209FF-0000-0000-C000-000000000046");
                    GetActiveObject(ref clsid, IntPtr.Zero, out wordApp);
                }
                catch (COMException)
                {
                    // Word is not running
                }
                catch (Exception)
                {
                    // GetActiveObject failed, Word might not be running
                }

                if (wordApp != null)
                {
                    _wordApp = wordApp;

                    int docCount = _wordApp.Documents.Count;
                    if (docCount == 0)
                    {
                        // No documents open, create a new one
                        _activeDocument = _wordApp.Documents.Add();
                    }
                    else
                    {
                        _activeDocument = _wordApp.ActiveDocument;
                    }

                    _wordApp.Visible = true;
                    return true;
                }
                else
                {
                    // Word is not running, try to start it
                    _wordApp = Activator.CreateInstance(wordType)!;
                    _wordApp.Visible = true;

                    // Create a new document
                    _activeDocument = _wordApp.Documents.Add();

                    return true;
                }
            }
            catch (Exception ex)
            {
                // Re-throw with more context
                throw new InvalidOperationException($"Failed to connect to Word: {ex.Message}", ex);
            }
        }

        public int HighlightAllWords()
        {
            if (_activeDocument == null)
                throw new InvalidOperationException("No active document connected.");

            int wordCount = 0;

            try
            {
                // Clear existing positions
                _wordPositions.Clear();

                // Get all words in the document
                dynamic documentRange = _activeDocument.Range();
                dynamic words = documentRange.Words;

                foreach (dynamic wordRange in words)
                {
                    string text = wordRange.Text.ToString().Trim();

                    // Skip empty strings, punctuation, and whitespace
                    if (!string.IsNullOrWhiteSpace(text) &&
                        text.Length > 1 &&
                        !IsPunctuation(text))
                    {
                        try
                        {
                            // Get bounding rectangle of the word
                            dynamic activeWindow = _wordApp.ActiveWindow;

                            // Get screen position using GetPoint
                            int left = 0, top = 0, width = 0, height = 0;

                            // Use GetPoint to get screen coordinates
                            activeWindow.GetPoint(out left, out top, out width, out height, wordRange);

                            if (width > 0 && height > 0)
                            {
                                var position = new WordPosition
                                {
                                    Text = text,
                                    Left = left,
                                    Top = top,
                                    Width = width,
                                    Height = height
                                };

                                _wordPositions.Add(position);
                                wordCount++;
                            }
                        }
                        catch (Exception ex)
                        {
                            // Try alternative method with RangeFromPoint
                            try
                            {
                                // Fallback: estimate position based on range info
                                double left = 100; // Default position
                                double top = 100 + (wordCount * 20);
                                double width = text.Length * 8;
                                double height = 16;

                                var position = new WordPosition
                                {
                                    Text = text,
                                    Left = left,
                                    Top = top,
                                    Width = width,
                                    Height = height
                                };

                                _wordPositions.Add(position);
                                wordCount++;
                            }
                            catch
                            {
                                // Skip this word
                                continue;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to highlight words: {ex.Message}", ex);
            }

            return wordCount;
        }

        public System.Collections.Generic.List<WordPosition> GetWordPositions()
        {
            return _wordPositions;
        }

        public int GetScrollPosition()
        {
            try
            {
                if (_wordApp == null) return 0;

                dynamic activeWindow = _wordApp.ActiveWindow;
                int scrollTop = activeWindow.VerticalPercentScrolled;
                return scrollTop;
            }
            catch
            {
                return 0;
            }
        }

        public bool IsWordWindowFocused()
        {
            try
            {
                if (_wordApp == null) return false;

                // Get the handle of the active Word window
                dynamic activeWindow = _wordApp.ActiveWindow;
                int wordHwnd = activeWindow.Hwnd;

                // Get the currently focused window
                IntPtr foregroundWindow = GetForegroundWindow();

                // Check if Word window is the focused window
                return foregroundWindow.ToInt32() == wordHwnd;
            }
            catch
            {
                return false;
            }
        }

        private void ClearAllOverlays()
        {
            _wordPositions.Clear();
        }

        public void ClearAllHighlights()
        {
            if (_activeDocument == null)
                throw new InvalidOperationException("No active document connected.");

            try
            {
                ClearAllOverlays();
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
                dynamic selection = _wordApp.Selection;

                if (selection != null)
                {
                    string selectedText = selection.Text?.ToString().Trim() ?? string.Empty;

                    // Check if this word is in our positions list
                    if (!string.IsNullOrWhiteSpace(selectedText) &&
                        !IsPunctuation(selectedText) &&
                        _wordPositions.Any(w => w.Text == selectedText))
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
        public POINT GetCursorPosition()
        {
            GetCursorPos(out POINT point);
            return point;
        }

        public System.Drawing.Rectangle GetWordWindowBounds()
        {
            if (_wordApp == null)
                return System.Drawing.Rectangle.Empty;

            try
            {
                int left = _wordApp.Left;
                int top = _wordApp.Top;
                int width = _wordApp.Width;
                int height = _wordApp.Height;
                return new System.Drawing.Rectangle(left, top, width, height);
            }
            catch
            {
                return System.Drawing.Rectangle.Empty;
            }
        }

        public void ReplaceSelectedWord(string newWord)
        {
            if (_wordApp == null || _activeDocument == null)
                throw new InvalidOperationException("No active document connected.");

            try
            {
                dynamic selection = _wordApp.Selection;
                
                // Get the currently selected text
                string oldWord = selection.Text.ToString().Trim();

                // Replace the selected text with the new word
                selection.Text = newWord;

                // Remove from positions list
                _wordPositions.RemoveAll(w => w.Text == oldWord);
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to replace word: {ex.Message}", ex);
            }
        }

        public void SelectWordInDocument(string wordToSelect)
        {
            if (_wordApp == null || _activeDocument == null)
                return;

            try
            {
                // Find and select the word in the document
                dynamic find = _wordApp.Selection.Find;
                find.ClearFormatting();
                find.Text = wordToSelect;
                find.MatchWholeWord = true;
                find.MatchCase = false;
                find.Forward = true;
                
                // Execute the find
                find.Execute();
            }
            catch
            {
                // Ignore errors - word might not be found
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
