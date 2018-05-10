using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Xltrail.Client.Models
{
    public class Branch
    {
        public Workbook Workbook { get; private set; }
        public string Id { get; private set; }
        public string Name { get; private set; }
        public string Path { get; private set; }
        public string LocalFileName { get; private set; }
        public string DisplayName { get; private set; }

        public Branch(Workbook workbook, string branch)
        {
            Id = "id-" + Convert.ToBase64String(Guid.NewGuid().ToByteArray()).Replace("+", "").Replace("/", "_").Replace(" ", "").Replace("=", "");
            Workbook = workbook;
            Name = branch;
            DisplayName = Name;// Name.Replace("_local", "");
            LocalFileName = System.IO.Path.GetFileNameWithoutExtension(Workbook.Path) + "_" + Name + System.IO.Path.GetExtension(Workbook.Path);
            Path = System.IO.Path.Combine(
                Workbook.Repository.Path.Split('\\').Last(),
                Workbook.Path,
                LocalFileName);
        }


        public string Head
        {
            get
            {
                return Workbook
                    .Repository
                    .GitRepository
                    .Commits
                    .QueryBy(Workbook.Path, new LibGit2Sharp.CommitFilter { IncludeReachableFrom = Name })
                    .First().Commit.Id.Sha;
            }
            
        }
    }
}
