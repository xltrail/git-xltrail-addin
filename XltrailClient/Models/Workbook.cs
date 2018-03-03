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
        public IList<string> Branches;

        public Workbook(string path, IList<string> branches)
        {
            Path = path;
            Branches = branches;
        }

        public Workbook(string path, string branch)
        {
            Path = path;
            Branches = new List<string>() { branch };
        }
    }
}
