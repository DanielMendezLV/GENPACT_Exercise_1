using System;
using System.Collections.Generic;
using System.Text;

namespace DMFolderWatcher.Class
{
    public class CFolder
    {
        private string _path;
        private string _parent;
        private string _processed_folder;
        private string _notaplicable_folder;
        private string _master_file;

        public string Path
        {
            get
            {
                return _path;
            }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    _path = value;
                }
            }
        }


        public string Parent
        {
            get
            {
                return _parent;
            }
            set
            {
                if (!string.IsNullOrEmpty(value))
                {
                    _parent = value;
                    _processed_folder = _parent + "\\Processed";
                    _notaplicable_folder = _parent + "\\Not Applicable";
                    _master_file = _parent + "\\Master.xlsx";
                }
            }
        }

        public string Processed_folder
        {
            get
            {
                return _processed_folder;
            }

        }


        public string NotApplicable_folder
        {
            get
            {
                return _notaplicable_folder;
            }

        }


        public string Master_file
        {
            get
            {
                return _master_file;
            }

        }


    }
}
