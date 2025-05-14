using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;

namespace Sgrips.WishList.WPF
{
    class ValidateWishListCode : ValidationRule
    {
        public override ValidationResult Validate(object value, CultureInfo cultureInfo)
        {
            string codeWish = value as string;
            bool err = false;

            if (codeWish != null && (codeWish.Length == 13))
            {



                Task<string> task = Task.Run<string>(async () => await GetWishName(codeWish));
                task.Wait(); 


                err = task.Result == null;

            }



            if (err)
                return new ValidationResult(false, "Sorry, WishCode is not valid");
            else
                return new ValidationResult(true, string.Empty);

        }



        private async Task<string> GetWishName(string codeWish)
        {
            try
            {
                var result = await Amazon.DabaseHelper.GetWishNameFromURL(codeWish);
                return result;

            }
            catch (Exception)
            {

                return null; 
            }
        }

    }
}
