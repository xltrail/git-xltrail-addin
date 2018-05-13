using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace Xltrail.Client.Models
{
    public class Branch
    {
        public static string StagingBranchPrefix = "staging__";

        public Workbook Workbook { get; private set; }
        public string Id { get; private set; }
        public string Name { get; private set; }
        public string Path { get; private set; }
        public string FileName { get; private set; }

        public Branch(Workbook workbook, string branch)
        {
            Id = "id-" + Convert.ToBase64String(Guid.NewGuid().ToByteArray()).Replace("+", "").Replace("/", "_").Replace(" ", "").Replace("=", "");
            Workbook = workbook;
            Name = branch;
            FileName = System.IO.Path.GetFileNameWithoutExtension(Workbook.Path) + "_" + DisplayName + (IsStagingBranch? "_local" : "") + System.IO.Path.GetExtension(Workbook.Path);
            Path = System.IO.Path.Combine(Workbook.Path, FileName).Replace("\\", "/");
        }

        public bool IsStagingBranch
        {
            get
            {
                return Name.StartsWith(StagingBranchPrefix) && Regex.Matches(Name, StagingBranchPrefix).Count % 2 != 0;
            }
        }

        public string BranchId
        {
            get
            {
                var history = GetHistory().ToList();
                return history.Last().Commit.Sha;
            }
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

        public void Discard()
        {
            if (!IsStagingBranch)
                throw new Exception("Can only discard staging branch");

            //checkout branch
            LibGit2Sharp.Commands.Checkout(Workbook.Repository.GitRepository, OtherBranchName);

            //delete Git staging branch
            Workbook.Repository.GitRepository.Branches.Remove(Name);

            //delete branch from workbook
            Workbook.Branches.Remove(this);
        }

        public Branch Checkout()
        {
            if (IsStagingBranch)
                throw new Exception("Cannot check out staging branch");

            var stagingBranchName = StagingBranchPrefix + BranchId + "__" + Name;
            var libGit2Branch = Workbook.Repository.GitRepository.Branches[Name];

            //create staging branch if not exists
            if (Workbook.Repository.GitRepository.Branches[stagingBranchName] == null)
            {
                Workbook.Repository.GitRepository.Branches.Update(
                    Workbook.Repository.GitRepository.Branches.Add(stagingBranchName, libGit2Branch.Tip),
                    b => b.TrackedBranch = libGit2Branch.CanonicalName);

                Workbook.Branches.Add(new Branch(Workbook, stagingBranchName));
            }

            //checkout staging branch
            LibGit2Sharp.Commands.Checkout(Workbook.Repository.GitRepository, stagingBranchName);
            return Workbook.Branches.Where(branch => branch.Name == stagingBranchName).FirstOrDefault();
        }

        public string DisplayName
        {
            get
            {
                if (!IsStagingBranch)
                    return Name;

                return Name.Substring(StagingBranchPrefix.Length + BranchId.Length + 2);
            }
        }

        public string Description
        {
            get
            {
                var head = GetHeadCommit();
                if (!IsStagingBranch)
                {
                    return DisplayName + " [origin, created by " + head.Author.Name + ", " + head.Author.When.ToLocalTime().DateTime.ToString("MMMM dd, yyyy, HH:MM:ss") + "]";
                }
                return DisplayName + " [" + head.Author.Name + ", " + head.Author.When.ToLocalTime().DateTime.ToString("MMMM dd, yyyy, HH:MM:ss") + "]" + (HasConflict ? "*" : "");
            }
        }


        private string OtherBranchName
        {
            get
            {
                if (!IsStagingBranch)
                    return Name;

                return Name.Substring(StagingBranchPrefix.Length + BranchId.Length + 2);
            }
        }

    public Branch OtherBranch
        {
            get
            {
                //find corresponding `other` branch
                var otherBranch = Workbook.Branches.Where(b => b.Name == OtherBranchName).FirstOrDefault();

                return otherBranch;
            }
        }

        public bool HasConflict
        {
            get
            {
                //find corresponding `other` branch
                var otherBranch = OtherBranch;

                //circuit breaker
                if (otherBranch == null)
                    return false;

                return otherBranch.Head != Head;
            }
        }

        private IEnumerable<LibGit2Sharp.LogEntry> GetHistory()
        {
            return Workbook
             .Repository
             .GitRepository
             .Commits
             .QueryBy(Workbook.Path, new LibGit2Sharp.CommitFilter { IncludeReachableFrom = Name });
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
