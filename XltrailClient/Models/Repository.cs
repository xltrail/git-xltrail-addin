using System.Collections.Generic;
using LibGit2Sharp;
using System.Linq;

namespace Xltrail.Client.Models
{

    public class Folder
    {
        public IList<Workbook> Workbooks;
        public HashSet<string> Folders;

        public Folder()
        {
            Workbooks = new List<Workbook>();
            Folders = new HashSet<string>();
        }
    }


    public class Repository
    {
        public LibGit2Sharp.Repository GitRepository { get; private set; }
        private IDictionary<string, IList<string>> workbooks;
        public string Path { get; private set; }

        public Repository(string path)
        {
            Path = path;
            GitRepository = new LibGit2Sharp.Repository(path);
            workbooks = GetWorkbooks();
        }

        public static bool IsValid(string path)
        {
            return LibGit2Sharp.Repository.IsValid(path);

        }

        private IDictionary<string, IList<string>> GetWorkbooks()
        {
            var workbooks = new Dictionary<string, IList<string>>();
            foreach(var branch in GitRepository.Branches)
            {
                Commands.Checkout(GitRepository, branch);
                foreach (var file in GitRepository.Index)
                {
                    if (IsWorkbook(file.Path))
                    {
                        if (!workbooks.ContainsKey(file.Path))
                            workbooks[file.Path] = new List<string>();
                        workbooks[file.Path].Add(branch.FriendlyName);
                    }
                }
            }
            return workbooks;
        }

            private bool IsWorkbook(string path)
        {
            return (
                path.EndsWith(".xls") ||
                path.EndsWith(".xlsb") ||
                path.EndsWith(".xlsm") ||
                path.EndsWith(".xlsx") ||
                path.EndsWith(".xla") ||
                path.EndsWith(".xlam"));
        }

        public Folder Workbooks(string path)
        {
            var folder = new Folder();

            foreach (var workbook in workbooks.Keys)
            {
                if (workbook.StartsWith(path))
                {
                    if(System.IO.Path.GetDirectoryName(workbook) == path)
                    {
                        folder.Workbooks.Add(new Workbook(
                            workbook,
                            workbooks[workbook].Select(b => b.Replace("origin/", "")).Where(b => b != "HEAD").Distinct().ToList()));
                    }
                    else
                    {
                        folder.Folders.Add(System.IO.Path.GetDirectoryName(workbook));
                    }
                }
            }
            return folder;
        }

    }
}
