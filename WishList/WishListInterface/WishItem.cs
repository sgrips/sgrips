using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sgrips.WishList.Model
{
    public class WishItem
    {
        public string Id { get; set; }
        public string Description { get; set; }
        public string OriginalPrice { get; set; }
        public string NewPrice { get; set; }
        public string DataTime { get; set; }
        public bool Deleted { get; set; }

        public bool InStock { get; set; }

        public byte[] Image { get; set; }

        public double ReducedPrice { get; set; }

    }
}
