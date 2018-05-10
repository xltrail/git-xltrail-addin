using System.Collections.Generic;

namespace Xltrail.Client.Models
{

    public class Folder
    {
        public IList<Workbook> Workbooks { get; private set; }
        public IList<string> Folders { get; private set; }
        public Repository Repository { get; private set; }

        public Folder(Repository repository, string path)
        {

            Folders = repository.GetFolders(path);
            Workbooks = repository.GetWorkbooks(path);
        }
    }
}