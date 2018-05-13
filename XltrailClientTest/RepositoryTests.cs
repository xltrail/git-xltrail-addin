using System;
using NUnit.Framework;
using System.IO;
using System.Linq;
using Xltrail.Client.Models;
using System.Collections.Generic;

namespace Xltrail.Client.Test
{
    class RepositoryTests
    {
        private string path;

        private void CopyWorkbook(string source, string target)
        {
            var path = Path.GetDirectoryName(target);
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            File.Copy(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Workbooks", source), target, true);
        }

        private void UpdateFileAttributes(DirectoryInfo dInfo)
        {
            dInfo.Attributes = FileAttributes.Normal;
            foreach (FileInfo file in dInfo.GetFiles())
                file.Attributes = FileAttributes.Normal;
            foreach (DirectoryInfo subDir in dInfo.GetDirectories())
                UpdateFileAttributes(subDir);
        }

        [SetUp]
        public void CreateRepository()
        {
            path = Path.Combine(Path.GetTempPath(), "test-" + Guid.NewGuid().ToString());
            if (Directory.Exists(path))
            {
                UpdateFileAttributes(new DirectoryInfo(path));
                Directory.Delete(path, true);
            }

            Directory.CreateDirectory(path);
            LibGit2Sharp.Repository.Init(path);

            var author = new LibGit2Sharp.Signature("Bjoern", "bjoern.stiel@zoomeranalytics.com", DateTime.Now);
            var commitOptions = new LibGit2Sharp.CommitOptions();

            CopyWorkbook("Book1_v1.xlsx", Path.Combine(path, "xlwings", "tests", "test book.xlsx"));
            CopyWorkbook("Book2_v1.xlsx", Path.Combine(path, "Book2.xlsx"));
            using (var gitRepo = new LibGit2Sharp.Repository(path))
            {
                LibGit2Sharp.Commands.Stage(gitRepo, "*");
                gitRepo.Commit("Added Book1 and Book2", author, author, new LibGit2Sharp.CommitOptions());

                gitRepo.Branches.Add("dev", "HEAD");
                LibGit2Sharp.Commands.Checkout(gitRepo, "dev");

                File.Delete(Path.Combine(path, "Book2.xlsx")); //delete Book2.xlsx
                CopyWorkbook("Book1_v2.xlsx", Path.Combine(path, "xlwings", "tests", "test book.xlsx")); //new version
                LibGit2Sharp.Commands.Stage(gitRepo, "*");
                gitRepo.Commit("Modified Book1, Deleted Book2", author, author, new LibGit2Sharp.CommitOptions());

            }
        }

        [Test]
        public void TestCanTraverseFolders()
        {
            var folders = Repository.TraverseFolders("xlwings/tests");
            Assert.AreEqual(new List<string>() { "", "xlwings", "xlwings/tests" }, folders);
        }

        [Test]
        public void TestCanTraverseRootFolder()
        {
            var folders = Repository.TraverseFolders("");
            Assert.AreEqual(new List<string>() { "" }, folders);
        }

        [Test]
        public void TestWorkbooks()
        {
            var repository = new Repository(path);

            //workbooks
            Assert.AreEqual(2, repository.Workbooks.Count);

            var workbook = repository.GetWorkbooks("xlwings/tests").First();
            Assert.AreEqual("xlwings/tests/test book.xlsx", workbook.Path);
            Assert.AreEqual(new List<string>() { "xlwings/tests/test book.xlsx/test book_dev.xlsx", "xlwings/tests/test book.xlsx/test book_master.xlsx" }, workbook.Branches.Select(x => x.Path));
        }

        [Test]
        public void TestCanGetFoldersAndWorkbooks()
        {
            var repository = new Repository(path);

            //folders
            Assert.AreEqual(new List<string>() { "", "xlwings", "xlwings/tests" }, repository.Folders.Keys);

            //root folder
            Assert.AreEqual(new List<string>() { "xlwings" }, repository.GetFolders(""));
            Assert.AreEqual(new List<string>() { "Book2.xlsx" }, repository.GetWorkbooks("").Select(x => x.Path).ToList());

            //xlwings folder
            Assert.AreEqual(new List<string>() { "tests" }, repository.GetFolders("xlwings"));
            Assert.AreEqual(new List<string>(), repository.GetWorkbooks("xlwings").ToList());

            //xlwings/tests folder
            Assert.AreEqual(new List<string>(), repository.GetFolders("xlwings/tests"));
            Assert.AreEqual(new List<string>() { "xlwings/tests/test book.xlsx" }, repository.GetWorkbooks("xlwings/tests").Select(x => x.Path).ToList());

        }

        [Test]
        public void TestCanCheckout()
        {
            var repository = new Repository(path);
            var workbook = repository.GetWorkbooks("xlwings/tests").First();

            Assert.AreEqual(2, workbook.Branches.Count);

            var branch = workbook.GetBranch("dev");
            Assert.IsFalse(branch.IsStagingBranch);

            var stagingBranch = branch.Checkout();
            Assert.IsTrue(stagingBranch.IsStagingBranch);
            Assert.AreEqual(branch.BranchId, stagingBranch.BranchId);

            Assert.AreEqual(3, workbook.Branches.Count);

            //get branch from stagingBranch
            branch = stagingBranch.OtherBranch;
            Assert.AreEqual(branch.Name, "dev");
            Assert.IsFalse(branch.IsStagingBranch);

            branch.Checkout();
            Assert.AreEqual(3, workbook.Branches.Count);
        }

        [Test]
        public void TestCanDeleteCheckout()
        {
            var repository = new Repository(path);
            var workbook = repository.GetWorkbooks("xlwings/tests").First();

            Assert.AreEqual(2, workbook.Branches.Count);

            var branch = workbook.GetBranch("dev");
            var stagingBranch = branch.Checkout();
            Assert.AreEqual(3, workbook.Branches.Count);

            stagingBranch.Discard();
            Assert.AreEqual(2, workbook.Branches.Count);
        }
    }
}
