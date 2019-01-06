using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DCTonam_cnf
{
    public class dulieu
    {
        public string masp { get; set; }
        public int soluong { get; set; }

        public dulieu(string masp, int soluong)
        {
            this.masp = masp;
            this.soluong = soluong;
        }
    }
}
