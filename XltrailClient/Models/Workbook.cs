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
            Id = "id-" + Convert.ToBase64String(Guid.NewGuid().ToByteArray()).Replace("+", "").Replace("/", "_").Replace(" ", "").Replace("=", "");
            Path = path;
            Folder = System.IO.Path.GetDirectoryName(Path).Replace("\\", "/");
            Repository = repository;
            InitialiseBranches();
        }

        public void InitialiseBranches()
        {
            Branches = new List<Branch>();
            foreach(var branch in Repository
                .GitRepository
                .Branches
                .Select(branch => branch.FriendlyName.Replace("origin/", ""))
                .Where(name => name != "HEAD")
                .Distinct())
            {
                Branches.Add(new Branch(this, branch));
            }
        }
    }
}
