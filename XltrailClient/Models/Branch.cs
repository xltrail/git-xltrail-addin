using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Xltrail.Client.Models
{
    public class Branch
    {
        public static string StagingBranchSuffix = "_local";

        public Workbook Workbook { get; private set; }
        public string Id { get; private set; }
        public string Name { get; private set; }
        public string Path { get; private set; }
        public string FileName { get; private set; }
        public bool IsStagingBranch { get; private set; }

        public Branch(Workbook workbook, string branch)
        {
            Id = "id-" + Convert.ToBase64String(Guid.NewGuid().ToByteArray()).Replace("+", "").Replace("/", "_").Replace(" ", "").Replace("=", "");
            Workbook = workbook;
            Name = branch;
            FileName = System.IO.Path.GetFileNameWithoutExtension(Workbook.Path) + "_" + Name + System.IO.Path.GetExtension(Workbook.Path);
            Path = System.IO.Path.Combine(Workbook.Path, FileName).Replace("\\", "/");
            IsStagingBranch = Name.EndsWith(StagingBranchSuffix) && Regex.Matches(Name, StagingBranchSuffix).Count % 2 != 0;
        }

        public void Commit(string path, string userName, string email)
        {
            var fileRepositoryPath = System.IO.Path.Combine(Workbook.Repository.Path.Replace("/", "\\"), Workbook.Path.Replace("/", "\\"));

            //checkout branch
            LibGit2Sharp.Commands.Checkout(Workbook.Repository.GitRepository, Name);

            //copy file
            File.Copy(path, fileRepositoryPath, true);

            //stage
            LibGit2Sharp.Commands.Stage(Workbook.Repository.GitRepository, fileRepositoryPath);

            //prepare commit details
            var author = new LibGit2Sharp.Signature(userName, email, DateTime.Now);
            var commitOptions = new LibGit2Sharp.CommitOptions();
            var commitMessage = "Updated " + Workbook.Path;
            var committer = author;

            //commit
            Workbook.Repository.GitRepository.Commit(commitMessage, author, committer, commitOptions);
        }

        public string DisplayName
        {
            get
            {
                if (!IsStagingBranch)
                {
                    var headCommit = GetHeadCommit();
                    return Name + " [origin, created by " + headCommit.Author.Name + ", " + headCommit.Author.When.ToLocalTime().DateTime.ToString("MMMM dd, yyyy, HH:MM:ss") + "]";
                }


                var displayName = Name.Substring(0, Name.Length - StagingBranchSuffix.Length);

                //find corresponding branch
                var branch = Workbook.Branches.Where(b => b.Name == displayName).FirstOrDefault();
                var branchHeadCommit = branch.GetHeadCommit();

                //check if heads match
                if(branchHeadCommit.Sha != Head)
                {
                    displayName += " [your version, last modified " + branchHeadCommit.Author.When.ToLocalTime().DateTime.ToString("MMMM dd, yyyy, HH:MM:ss") + "]";
                }

                return displayName;
            }
        }

        public bool HasConflict
        {
            get
            {
                if (IsStagingBranch)
                    return false;

                //find corresponding staging branch
                var stagingBranch = Workbook.Branches.Where(b => b.Name == Name + StagingBranchSuffix).FirstOrDefault();

                //circuit breaker
                if (stagingBranch == null)
                    return false;

                var stagingHead = stagingBranch.Head;
                var head = Head;

                return stagingHead != head;
            }
        }


        public LibGit2Sharp.Commit GetHeadCommit()
        {
            return Workbook
             .Repository
             .GitRepository
             .Commits
             .QueryBy(Workbook.Path, new LibGit2Sharp.CommitFilter { IncludeReachableFrom = Name })
             .First().Commit;
        }


        public string Head
        {
            get
            {
                return GetHeadCommit().Sha;
            }
            
        }
    }
}
