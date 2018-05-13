using log4net;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Config=Xltrail.Client.Models.Config;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;
using Xltrail.Client.Models;

namespace Xltrail.Client.Providers
{
    public class ConfigProvider
    {
        private static readonly ILog logger = LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        public static void PullConfigFromRemote(string url, string path, Config.Credentials credentials)
        {
            if (url == null)
            {
                logger.Info("Skip pull config yaml from remote, no url defined");
                return;
            }
            logger.Info("Pull config repository from remote");
            try
            {
                GitProvider.PullFromRemote(url, path, credentials);
            }
            catch (Exception ex)
            {
                logger.WarnFormat(ex.Message);
            }
        }


        public static List<Config.Repository> LoadRepositoryConfigsFromFile(string path)
        {
            logger.InfoFormat("Load configured repositories from file: {0}", path);
            var yaml = File.ReadAllText(path);
            logger.Info(yaml);
            var deserializer = new DeserializerBuilder()
                .WithNamingConvention(new CamelCaseNamingConvention())
                .Build();
            var repositoryConfigs = deserializer.Deserialize<List<Config.Repository>>(yaml);
            logger.InfoFormat("Configured repositories: {0}", repositoryConfigs.Count);
            return repositoryConfigs;
        }

        public static Repositories LoadRepositoriesFromFilesystem(string repositoriesPath, IList<Config.Repository> repositoryConfigs)
        {
            var repositoryNames = repositoryConfigs.Select(r => r.Alias);
            return new Repositories(repositoriesPath, repositoryNames);
        }



        /*
        public static void PullRepositories()
        {
            foreach (var repositoryConfig in configuredRepositories.Repositories)
            {
                var repositoryPath = Path.Combine(RepositoriesPath, repositoryConfig.Alias);
                var credentials = Config.Credentials.Where(c => repository.Url.StartsWith(c.Url)).FirstOrDefault();
                GitProvider.PullFromRemote(repository.Url, repositoryPath, credentials);
                GitProvider.EnsureStagingBranches(repositoryPath);
            }
        }*/

    }
}
