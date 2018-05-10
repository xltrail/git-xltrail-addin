using log4net;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Collections;

namespace Xltrail.Client.Models
{
    public class Repositories : IEnumerable<Repository>
    {
        private string path;
        private IList<Repository> repositories;
        private IEnumerable<string> configuredRepositories;
        private static readonly ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);


        public Repositories(string path, IEnumerable<string> configuredRepositories)
        {
            this.path = path;
            this.configuredRepositories = configuredRepositories;
            InitialiseRepositories();
        }

        private void InitialiseRepositories()
        {
            repositories = new List<Repository>();
            foreach (var path in Directory.GetDirectories(path))
            {
                var repoName = Path.GetFileName(path);
                if (configuredRepositories.Contains(repoName) && Repository.IsValid(path))
                {
                    logger.InfoFormat("Synchronise Git workbook repository from remote: {0}", path);
                    repositories.Add(new Repository(path));
                }
            }
        }

        public (Repository, string) GetRepositoryFromId(string id)
        {
            foreach (var repository in repositories)
            {
                if (repository.Id == id)
                    return (repository, "");

                var folderId = repository.Folders.Values.Where(v => v == id).FirstOrDefault();
                if (folderId != null)
                    return (repository, folderId);
            }
            return (null, null);
        }

        public Workbook GetWorkbookFromId(string id)
        {
            foreach (var repository in repositories)
            {
                foreach (var workbook in repository.Workbooks)
                {
                    if (workbook.Id == id)
                        return workbook;
                }
            }
            return null;
        }

        public Branch GetWorkbookVersionFromId(string id)
        {
            foreach (var repository in repositories)
            {
                foreach (var workbook in repository.Workbooks)
                {
                    foreach (var branch in workbook.Branches)
                    {
                        if (branch.Id == id)
                            return branch;
                    }
                }
            }
            return null;
        }

        public Branch GetWorkbookVersionFromPath(string path)
        {
            var repoName = path.Split('/').First();
            var workbookBranchPath = String.Join("/", path.Split('/').Skip(1));

            var repository = repositories.Where(r => r.Name == repoName).FirstOrDefault();
            if (repository == null)
                return null;

            foreach (var workbook in repository.Workbooks)
            {
                var branch = workbook.Branches.Where(b => b.Path == workbookBranchPath).FirstOrDefault();
                if (branch != null)
                    return branch;
            }
            return null;
        }

        public IEnumerator<Repository> GetEnumerator()
        {
            return repositories.GetEnumerator();
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }
    }
}
