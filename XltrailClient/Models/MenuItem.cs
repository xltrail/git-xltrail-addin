using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Xltrail.Client.Models
{
    public class MenuItem
    {
        public Repository Repository { get; private set; }
        public string Folder = "";
        public Workbook Workbook;

        public MenuItem()
        {
            Repository = null;
        }

        public MenuItem(Repository repository)
        {
            Repository = repository;
        }

        public MenuItem(Repository repository, Workbook workbook)
        {
            Repository = repository;
            Workbook = workbook;
        }

        public MenuItem(Repository repository, string folder)
        {
            Repository = repository;
            Folder = folder;
        }

        public bool IsRoot()
        {
            return Repository == null;
        }

        public bool IsFolder()
        {
            return Folder != "";
        }

        public bool IsWorkbook()
        {
            return Workbook != null;
        }

    }
}
