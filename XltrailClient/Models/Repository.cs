using System.Collections.Generic;
using LibGit2Sharp;
using System.Linq;
using System;

namespace Xltrail.Client.Models
{
    public class Repository
    {
        public LibGit2Sharp.Repository GitRepository { get; private set; }
        public string Id { get; private set; }
        public string Path { get; private set; }
        public string Name { get; private set; }
        public IList<Workbook> Workbooks { get; private set; }
        public Dictionary<string, string> Folders { get; private set; }

        public Repository(string path)
        {
            Id = CreateId();
            Path = path;
            Name = System.IO.Path.GetFileName(path);
            GitRepository = new LibGit2Sharp.Repository(path);
            InitialiseWorkbooks();
        }

        private string CreateId()
        {
            return "id-" + Convert.ToBase64String(Guid.NewGuid().ToByteArray()).Replace("+", "").Replace("/", "_").Replace(" ", "").Replace("=", "");
        }

        public static bool IsValid(string path)
        {
            return LibGit2Sharp.Repository.IsValid(path);
        }

        public Workbook GetWorkbookById(string id)
        {
            return Workbooks.Where(workbook => workbook.Id == id).FirstOrDefault();
        }

        public Branch GetWorkbookBranchById(string id)
        {
            foreach(var workbook in Workbooks)
            {
                foreach(var branch in workbook.Branches)
                {
                    if (branch.Id == id)
                        return branch;
                }
            }
            return null;
        }

        public static List<string> TraverseFolders(string path)
        {
            var separator = '/';
            var parts = path.Split(separator);
            var folders = new List<string>();

            foreach (var part in parts)
            {
                if(folders.Count == 0)
                {
                    folders.Add(part);
                }
                else
                {
                    folders.Add(folders.Last() + separator + part);
                }
            }

            if (folders[0] != "")
                folders.Insert(0, "");

            return folders;
        }

        private void InitialiseWorkbooks()
        {
            Folders = new Dictionary<string, string>();
            Workbooks = new List<Workbook>();
            foreach (var branch in GitRepository.Branches.Where(branch => !branch.IsRemote))
            {
                Commands.Checkout(GitRepository, branch);
                foreach (var file in GitRepository.Index)
                {
                    if (IsWorkbook(file.Path))
                    {
                        var workbook = Workbooks.Where(x => x.Path == file.Path).FirstOrDefault();
                        if (workbook == null)
                        {
                            workbook = new Workbook(this, file.Path);
                            Workbooks.Add(workbook);
                            foreach(var folder in TraverseFolders(workbook.Folder))
                            {
                                if (!Folders.ContainsKey(folder))
                                {
                                    Folders[folder] = CreateId();
                                }
                            }
                        }
                    }
                }
            }
        }

        private bool IsWorkbook(string path)
        {
            return (
                path.EndsWith(".xls") ||
                path.EndsWith(".xlsx") ||
                path.EndsWith(".xlsb") ||
                path.EndsWith(".xlsm") ||
                path.EndsWith(".xla") ||
                path.EndsWith(".xlam"));
        }

        public IList<Workbook> GetWorkbooks(string path)
        {
            var files = new List<Workbook>();
            if (path == null)
                return files;

            foreach (var workbook in Workbooks)
            {
                if (workbook.Path.StartsWith(path))
                {
                    if (path == System.IO.Path.GetDirectoryName(workbook.Path).Replace("\\", "/"))
                    {
                        files.Add(workbook);
                    }
                }
            }
            return files;
        }

        public IList<string> GetFolders(string path)
        {
            var folders = new List<string>();
            if (path == null)
                return folders;

            foreach (var workbook in Workbooks)
            {
                if (workbook.Path.StartsWith(path))
                {
                    var p = System.IO.Path.GetDirectoryName(workbook.Path).Replace("\\", "/");
                    if (p != path)
                    {
                        var subPath = workbook.Path.Substring(path.Length);
                        if (subPath.StartsWith("/"))
                            subPath = subPath.Substring(1);
                        var f = subPath.Split('/').First();
                        if (f != null)
                            folders.Add(f);
                    }
                }
            }
            return folders.Distinct().OrderBy(x => x).ToList();
        }
    }
}
