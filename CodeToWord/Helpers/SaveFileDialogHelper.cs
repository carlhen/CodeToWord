using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CodeToWord.Helpers
{
    public static class SaveFileDialogHelper
    {

        public static string SelectNewSaveFileLocation(string filter, string previouslySelectedFile = null)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = filter;
            bool previousFileNotNull = !string.IsNullOrEmpty(previouslySelectedFile);
            if (previousFileNotNull)
            {
                try
                {
                    var directory = Path.GetDirectoryName(previouslySelectedFile);
                    saveFileDialog.InitialDirectory = directory;
                }
                catch { }
            }

            // Show save dialog
            if (saveFileDialog.ShowDialog() == true)
            {
                return saveFileDialog.FileName;
            }
            else
            {
                return previouslySelectedFile;
            }
                
        }
    }
}
