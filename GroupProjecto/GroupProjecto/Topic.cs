using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GroupProjecto
{
    class Topic
    {
        public string Name { get; set; }
        public int Days { get; set; }
        public string Notes { get; set; }


        public Topic(string name, int days, string notes)
        {
            Name = name;
            Days = days;
            Notes = notes;
        }
    }
}
