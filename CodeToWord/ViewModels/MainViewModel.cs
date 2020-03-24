using Caliburn.Micro;
using CodeToWord.Helpers;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CodeToWord.ViewModels
{
	class MainViewModel : Screen
    {

		private string _codeText;
		private string _saveLocation;
		private int _saveProgress;
		private string _saveProgressStatus;
		private const string Filter = "Word Document (*.docx)|*.docx";

		public MainViewModel()
		{
			SaveProgressStatus = "Ready.";
		}

		public string CodeText
		{
			get => _codeText;
			set 
			{ 
				_codeText = value;
				NotifyOfPropertyChange(() => CodeText);
				NotifyOfPropertyChange(() => CanSave);
			}
		}

		public string SaveLocation
		{
			get =>_saveLocation;
			set {
				_saveLocation = value;
				NotifyOfPropertyChange(() => SaveLocation);
				NotifyOfPropertyChange(() => CanSave);
			}
		}

		
		public int SaveProgress
		{
			get => _saveProgress;
			set { 
				_saveProgress = value;
				NotifyOfPropertyChange(() => SaveProgress);
			}
		}

		
		public string SaveProgressStatus
		{
			get => _saveProgressStatus;
			set {
				_saveProgressStatus = value;
				NotifyOfPropertyChange(() => SaveProgressStatus);
			}
		}

		private bool _saveProgressStatusIsError;

		public bool SaveProgressStatusIsError
		{
			get =>_saveProgressStatusIsError;
			set {
				_saveProgressStatusIsError = value;
				NotifyOfPropertyChange(() => SaveProgressStatusIsError);
			}
		}



		public void SelectSaveLocation()
		{
			SaveLocation = SaveFileDialogHelper.SelectNewSaveFileLocation(Filter, SaveLocation);
		}

		public bool CanSave
		{
			get => !string.IsNullOrEmpty(SaveLocation) && !string.IsNullOrEmpty(CodeText);
		}

		public async Task Save()
		{
			var progress = new Progress<(int Percentage, string Status, bool IsError)>(x => {
				if (x.Percentage >= 0)
				{
					SaveProgress = x.Percentage;
				}
				SaveProgressStatus= x.Status;
				SaveProgressStatusIsError = x.IsError;
			});
			await Task.Run(() => WordFileMakerHelper.CreateAndSaveWordDocument(CodeText, SaveLocation, progress));
		}


	}
}
