using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sgrips.Utility.IO.SharedResourceMapCheck.Engine.Model
{
    public class NetViewModel:  IEquatable<NetViewModel>

    {
        public string SharedName { get; set; }
        public string ShareType { get; set; }
        public string ShareComment { get; set; }


        public bool Equals(NetViewModel other)
        {
            return this != null && other!=null && 
                   this.SharedName == other.SharedName && 
                   this.ShareType == other.ShareType;
        }

    }

    public class NetViewModelComparer : IEqualityComparer<NetViewModel>
    {

        public NetViewModelComparer ()            {}

        public bool Equals(NetViewModel x, NetViewModel y)
        {
            //Check whether the objects are the same object. 
            if (Object.ReferenceEquals(x, y)) return true;

            return x.Equals(y);

        }

        public int GetHashCode(NetViewModel obj)
        {
            //Get hash code for the Name field if it is not null. 
            int hashProductName = obj.SharedName == null ? 0 : obj.SharedName.GetHashCode();

            //Get hash code for the Code field. 
            int hashProductCode = obj.ShareType.GetHashCode();

            return hashProductName ^ hashProductCode;
        }
    }
}
