using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using YamlDotNet.Serialization;

namespace Xltrail.Client.Models.Config
{
    public class Credentials
    {
        [YamlMember(Alias = "url", ApplyNamingConventions = false)]
        public string Url { get; set; }

        [YamlMember(Alias = "username", ApplyNamingConventions = false)]
        public string Username { get; set; }

        [YamlMember(Alias = "email", ApplyNamingConventions = false)]
        public string Email { get; set; }

        [YamlMember(Alias = "password", ApplyNamingConventions = false)]
        public string Password { get; set; }
    }


    public class Config
    {
        [YamlMember(Alias = "credentials", ApplyNamingConventions = false)]
        public List<Credentials> Credentials { get; set; }

        [YamlMember(Alias = "repositories", ApplyNamingConventions = false)]
        public string Repositories { get; set; }

        public Config()
        {
            Credentials = new List<Credentials>();
            Repositories = null;
        }
    }
}
