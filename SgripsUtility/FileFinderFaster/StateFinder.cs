using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

namespace BRT.Framework.FileFinderFaster.UI
{

    [DataObjectAttribute]
    public class StateFinder : INotifyPropertyChanged
    {

        int _dirCount;
        int _dirProgess;
        int _fileCount;
        int _fileSearchCount;

        public int FileSearchCount
        {
            get
            {
                return _fileSearchCount;
            }

            set
            {
                _fileSearchCount = value;
                NotifyPropertyChanged();
            }
        }

        public int DirCount
        {
            get
            {
                return _dirCount;
            }

            set
            {
                _dirCount = value;
                NotifyPropertyChanged();
            }
        }


        public int DirProgess
        {
            get
            {
                return _dirProgess;
            }

            set
            {
                _dirProgess = value;
                NotifyPropertyChanged();
            }
        }

        public int FileCount
        {
            get
            {
                return _fileCount;
            }

            set
            {
                _fileCount = value;
                NotifyPropertyChanged();
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        // This method is called by the Set accessor of each property.
        // The CallerMemberName attribute that is applied to the optional propertyName
        // parameter causes the property name of the caller to be substituted as an argument.
        private void NotifyPropertyChanged([CallerMemberName] String propertyName = "")
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(this, new PropertyChangedEventArgs(propertyName));
            }
        }
    }
}
