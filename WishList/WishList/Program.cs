using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace WishList
{
    class Program
    {

       
        static void Main(string[] args)
        {


             
            Sgrips.WishList.Amazon.DabaseHelper.Update("3RNP0TNJ8L32A",null);    

            Console.WriteLine("Hit ENTER to exit...");

            Console.ReadLine();
        }



    }



}
