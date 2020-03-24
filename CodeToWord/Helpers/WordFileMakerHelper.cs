using Word = Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CodeToWord.Helpers
{
    public static class WordFileMakerHelper
    {
        public static void CreateAndSaveWordDocument(string codeText, string saveFilePath, IProgress<(int Percentage, string Status, bool IsError)> progress)
        {
            Word.Application application = null;
            try
            {
                progress.Report((0,"Opening Word Application.", false));

                application = new Word.Application();

                progress.Report((10, "Opening Word Template.", false));

                Word.Document document = application.Documents.Open($@"{AppDomain.CurrentDomain.BaseDirectory}Template.docx");

                progress.Report((15, "Looking for Table.", false));

                Word.Table codeTable = document.Tables[1];

                progress.Report((20, "Processing code. Removing unnecessary spacing.", false));

                
                string[] codeLines = codeText.Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                double onePercent = 100/(codeLines.Length * 2);

                int smallestAmountOfSpaces = codeLines[0].TakeWhile(char.IsWhiteSpace).Count();
                IEnumerable<char> spaceToRemove = null;
                for (int i = 0; i < codeLines.Length; i++)
                {
                    IEnumerable<char> lineStartBlankSpaces = codeLines[i].TakeWhile(char.IsWhiteSpace);
                    int whiteSpaceStartOfLine = lineStartBlankSpaces.Count();
                    if (spaceToRemove == null)
                    {
                        spaceToRemove = lineStartBlankSpaces;
                    }
                    if(whiteSpaceStartOfLine < smallestAmountOfSpaces)
                    {
                        smallestAmountOfSpaces = whiteSpaceStartOfLine;
                        spaceToRemove = lineStartBlankSpaces; ;
                    }
                    progress.Report((20 + Convert.ToInt32((onePercent * (i + 1))), $"Processing code. Finding shortest spacing. Line: {i+1}/{codeLines.Length}", false));
                }


                int charsToRemoveIndex = 0;
                foreach (var c in spaceToRemove) charsToRemoveIndex++;

                for(int i = 0; i < codeLines.Length; i++)
                {
                    codeLines[i] = codeLines[i].Remove(0, charsToRemoveIndex);
                    progress.Report((20 + Convert.ToInt32((onePercent * (i + 1 + codeLines.Length))), $"Processing code. Finding shortest spacing. Line: {i + 1}/{codeLines.Length}", false));
                }

                progress.Report((60, "Adding code lines to table.", false));

                codeTable.Rows[1].Cells[3].Range.Text = codeLines[0];
                for (int i = 1; i < codeLines.Length;i++)
                {
                    codeTable.Rows.Add();
                    codeTable.Rows[i+1].Cells[3].Range.Text = codeLines[i];
                }

                progress.Report((90, "Saving new Word Document.", false));

                document.SaveAs2(saveFilePath);
                document.Close(false);
                application.Quit(false);
                progress.Report((100, "Done.", false));

            }
            catch
            {
                application?.Quit(false);
                progress.Report((-1, "Creating Word Document Failed", true));
            }
        }
    }
}
