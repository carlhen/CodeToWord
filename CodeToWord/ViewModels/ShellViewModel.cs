using Caliburn.Micro;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CodeToWord.ViewModels
{
    class ShellViewModel : Conductor<IScreen>
    {
        private IScreen _screen;
        public ShellViewModel(MainViewModel mainView)
        {
            _screen = mainView;
            ActivateItem(_screen);
        }


    }
}
