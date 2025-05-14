using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sgrips.WishList.Model
{
    
    public class Wish
    {

        public String Name { get; set; }

        public List<WishItem> Items { get; set; }

    }
}
