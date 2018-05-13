using System.Collections.Generic;
using System.Linq;
using NUnit.Framework;
using Xltrail.Client.Providers;
using System.IO;
using System;

namespace TestXltrailClient
{
    public class ConfigTest
    {
        private string configFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Resources", "config.yaml");

        [Test]
        public void TestCanLoadConfiguredRepositories()
        {
            var configuredRepositories = ConfigProvider.LoadRepositoryConfigsFromFile(configFile);
            Assert.AreEqual(new List<string>() { "https://github.com/ZoomerAnalytics/git-xltrail-workbooks.git" }, configuredRepositories.Select(r => r.Url));
        }
    }
}
