using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Xltrail.Client.Models
{
    public class Workbook
    {
        public string Path { get; private set; }
        public string Folder { get; private set; }
        public string Id { get; private set; }
        public Repository Repository { get; private set; }
        public List<Branch> Branches { get; private set; }

        public Workbook(Repository repository, string path)
        {
            Path = path.Replace("\\", "/");
            Id = "id-" + Convert.ToBase64String(Guid.NewGuid().ToByteArray()).Replace("+", "").Replace("/", "_").Replace(" ", "").Replace("=", "");
            Folder = System.IO.Path.GetDirectoryName(Path.Replace("/", "\\")).Replace("\\", "/");
            Repository = repository;
            InitialiseBranches();
        }

        public void InitialiseBranches()
        {
            Branches = new List<Branch>();
            var branches = Repository
                .GitRepository
                .Branches
                .Where(branch => !branch.IsRemote)
                .Where(branch => branch.FriendlyName != "HEAD")
                .ToList();

            foreach (var branch in branches)
            {
                //check if workbook exists in current branch
                var treeEntry = Repository.GitRepository.Branches[branch.FriendlyName][Path];
                if (treeEntry != null)
                    Branches.Add(new Branch(this, branch.FriendlyName));
            }
        }

        public Branch GetBranch(string branchName)
        {
            return Branches.Where(branch => branch.Name == branchName).FirstOrDefault();
        }
    }
}
