using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YamlDotNet.Serialization;
using YamlDotNet.Serialization.NamingConventions;


namespace Xltrail.Client.Models.Config
{
    /*
    public class Repositories
    {
        [YamlMember(Alias = "repositories", ApplyNamingConventions = false)]
        public List<Repository> Repositories { get; set; }

    }
    */

    public class Repository
    {
        private string alias;

        [YamlMember(Alias = "url", ApplyNamingConventions = false)]
        public string Url { get; set; }

        [YamlMember(Alias = "credentials", ApplyNamingConventions = false)]
        public string Credentials { get; set; }

        [YamlMember(Alias = "alias", ApplyNamingConventions = false)]
        public string Alias
        {
            get
            {
                if (alias != null)
                    return alias;

                return Url.Split('/').Last().Replace(".git", "");
            }
            set
            {
                alias = value;
            }
        }

        [YamlMember(Alias = "username", ApplyNamingConventions = false)]
        public string Username { get; set; }

        [YamlMember(Alias = "password", ApplyNamingConventions = false)]
        public string Password { get; set; }
    }
}
