using HtmlAgilityPack;
using Sgrips.Framework.FileSystemDatabase;
using Sgrips.WishList.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading.Tasks;

namespace Sgrips.WishList.Amazon
{


    public class DabaseHelper
    {

        static private JSONRepository dbWishListTypes;
        static private JSONRepository dbWish;
        static private string currentIdWish;

        public delegate void LoadedHandler(object sender, EventArgs e);

        public LoadedHandler OnLoadedHandler { get; set; }

        public DabaseHelper()
        {

        }

        static private JSONRepository DBWishListTypes
        {
            get
            {
                if (dbWishListTypes == null) dbWishListTypes = new JSONRepository("WishListTypes.json", "ID", false);
                return dbWishListTypes;
            }
        }

        static private JSONRepository GetDBWish(string idWish)
        {
            if (dbWish == null || currentIdWish != idWish)
            {
                currentIdWish = idWish;
                dbWish = new JSONRepository(string.Format("{0}.json", idWish), "Name", false);
            }
            return dbWish;
        }

        public static IList<WishListType> GetWishListTypesFromDB()
        {
            try
            {
                //original ROWS
                return DBWishListTypes.GetAll<WishListType>() as IList<WishListType>;
            }
            catch (Exception ex)
            {

                throw;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        public static void CreateDB()
        {
            DBWishListTypes.Save<WishListType>(new List<WishListType>{ (new WishListType()
            {
                IdType = "Amazon",
                Url = "https://www.amazon.it/hz/wishlist/ls/{0}",
                Codes = new List<WishListCode>()
            })});


        }



        public static void UpdateCode(IList<WishListType> codes)
        {
            try
            {
                //original ROWS
                DBWishListTypes.Save(codes);
            }
            catch (Exception ex)
            {

                throw;
            }
        }


        public static async void Update(string idWish, LoadedHandler callbackFunction)
        {
            //original ROWS
            Wish wish = null;
            wish = GetDBWish(idWish).Get<Wish>(idWish);
            //URL
            string url = DabaseHelper.GetURLfromIdWish(idWish);
            //new ROWS
            List<WishItem> listWishNew = await DabaseHelper.GetFromURL(url, idWish) as List<WishItem>;

            List<WishItem> listWishSave = null;
            if (wish != null && wish.Items != null)
            {
                var qry = (from e in wish.Items
                           join newElem in listWishNew on e.Id equals newElem.Id into joinResult
                           from newElemValues in joinResult.DefaultIfEmpty()
                           select new { e, newElemValues }).Select(f =>
                           {
                               f.e.NewPrice = f.newElemValues != null && f.newElemValues.Id != null ? f.newElemValues.OriginalPrice : f.e.OriginalPrice;
                               f.e.InStock = f.newElemValues != null && f.newElemValues.Id != null ? f.newElemValues.InStock : f.e.InStock;
                               f.e.Image = f.newElemValues != null && f.newElemValues.Id != null ? f.newElemValues.Image : f.e.Image;
                               f.e.Deleted = f.newElemValues != null && f.newElemValues.Id != null ? false : true;
                               f.e.ReducedPrice = f.newElemValues != null && f.newElemValues.Id != null ?
                                            CalculateReducedPrice(f.e.OriginalPrice, f.newElemValues.OriginalPrice) : f.e.ReducedPrice;
                               return f.e;
                           }).Where(f => f.Deleted == false).Union(
                           from e in listWishNew
                           join oldElem in wish.Items on e.Id equals oldElem.Id into joinResult
                           from oldElemValues in joinResult.DefaultIfEmpty()
                           where oldElemValues == null
                           select e);

                listWishSave = qry.ToList();
            }
            else
                listWishSave = listWishNew;

            GetDBWish(idWish).Modify(new Wish { Name = idWish, Items = listWishSave });

            if (callbackFunction != null) callbackFunction.Invoke(null, null);
        }


        private static double CalculateReducedPrice(string oldPriceS, string newPriceS)
        {
            try
            {

                double newPrice = 0;
                if (newPriceS != null) newPrice = Double.Parse(newPriceS.Replace("EUR", "").Replace("€", "").Trim());

                double oldPrice = Double.Parse(oldPriceS.Replace("EUR", "").Replace("€", "").Trim());
                double perc = 100 * ((oldPrice - newPrice) / oldPrice);
                if (perc == 100)
                    return 0;
                else
                    return Math.Round(perc, 2);
            }
            catch (Exception)
            {

                return 0;
            }

        }


        private static string GetURLfromIdWish(string idWishCode)
        {
            WishListType element = GetWishListTypesFromDB().Where(f => f.Codes.Exists(ff => ff.Id == idWishCode)).FirstOrDefault();
            return element.Url;
        }


        public static async Task<string> GetWishNameFromURL(string idWish)
        {
            List<WishItem> wishItems = new List<WishItem>();
            try
            {
                HttpClient httpClient = new HttpClient();
                HttpResponseMessage response;

                response = await httpClient.GetAsync(
                    string.Format("https://www.amazon.it/hz/wishlist/ls/{0}", idWish));
                response.EnsureSuccessStatusCode();

                Stream stream = await response.Content.ReadAsStreamAsync();

                HtmlDocument htmlDoc = new HtmlDocument();
                htmlDoc.Load(stream);

                HtmlNode htmlNode = htmlDoc.DocumentNode.SelectSingleNode("//span[@id='left-nav-rid_" + idWish + "']");

                return htmlNode.InnerText == null || htmlNode.InnerText == "Amazon.it" ? null : htmlNode.InnerText.Trim();


            }
            catch (Exception ex)
            {

                throw;
            }
            finally
            {


            }



        }



        public static async Task<IList<WishItem>> GetFromURL(string url, string idWish)
        {
            List<WishItem> wishItems = new List<WishItem>();
            try
            {
                HttpClient httpClient = new HttpClient();
                HttpResponseMessage response;

                response = await httpClient.GetAsync(
                    string.Format(url, idWish));
                response.EnsureSuccessStatusCode();

                Stream stream = await response.Content.ReadAsStreamAsync();

                HtmlDocument htmlDoc = new HtmlDocument();


                htmlDoc.Load(stream, true);



                foreach (var item in htmlDoc.DocumentNode.SelectNodes("//div[contains(@id,'itemMain_')]"))
                {
                    string idItem = item.Attributes["id"].Value.Replace("itemMain", "");
                    string price = null;
                    string title = null;
                    string img = null;
                    byte[] imgByte = null;

                    HtmlNode htmlNode = htmlDoc.DocumentNode.SelectSingleNode("//a[@id='itemName" + idItem + "']");
                    if (htmlNode != null)
                    {
                        title = htmlNode.Attributes["title"].Value;
                        htmlNode = htmlDoc.DocumentNode.SelectSingleNode("//span[@id='itemPrice" + idItem + "']");
                        if (htmlNode != null)
                            price = htmlNode.FirstChild.InnerText.Replace("\n", "").Trim();
                    }

                    htmlNode = htmlDoc.DocumentNode.SelectSingleNode("//div[@id='itemImage" + idItem + "']//img");
                    if (htmlNode != null)
                    {
                        img = htmlNode.Attributes["src"].Value;
                        response = await httpClient.GetAsync(img);
                        imgByte = await response.Content.ReadAsByteArrayAsync();
                    }

                    wishItems.Add(new WishItem { Id = idItem, Description = title, Image = imgByte, OriginalPrice = price, InStock = (price != null && price.ToLower() != "non disponibile") });
                    Console.WriteLine("item:{0} des:{1} price:{2}", item.Attributes["id"].Value, title, price);

                }
            }
            catch (Exception ex)
            {

                throw;
            }


            return wishItems;

        }

        public static Wish GetFromDB(string idWish)
        {
            Wish wish = null;
            try
            {

                //original ROWS
                wish = GetDBWish(idWish).Get<Wish>(idWish);
            }
            catch (Exception ex)
            {
                GetDBWish(idWish).Save<Wish>(new List<Wish>() { new Wish() { Name = idWish, Items = new List<WishItem>() } });
            }


            return wish;

        }

    }
}
