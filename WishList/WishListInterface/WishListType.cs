using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sgrips.WishList.Model
{
   public  class WishListType
    {
        public string IdType { get; set; }
        public string Url { get; set; }
        public List<WishListCode>  Codes { get; set; }
    }
}
