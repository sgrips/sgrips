using Sgrips.WishList.Amazon;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Markup;
using System.Windows.Media;
using System.Windows.Media.Imaging;

namespace Sgrips.WishList.WPF
{

    /// <summary>
    /// 
    /// </summary>
    /// <remarks></remarks>
    [ValueConversion(typeof(ReadOnlyObservableCollection<ValidationError>), typeof(string))]
    public class ConverterStringError : MarkupExtension, IValueConverter
    {

        public override object ProvideValue(System.IServiceProvider serviceProvider)
        {
            return new ConverterStringError();
        }


        public ConverterStringError()
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="value"></param>
        /// <param name="targetType"></param>
        /// <param name="parameter"></param>
        /// <param name="culture"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {

            throw new NotImplementedException();

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="value"></param>
        /// <param name="targetType"></param>
        /// <param name="parameter"></param>
        /// <param name="culture"></param>
        /// <returns></returns>
        /// <remarks></remarks>
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            ValidationError s = value as ValidationError;

            if (s != null)
            {
                if (!String.IsNullOrWhiteSpace(s.ErrorContent.ToString()))
                    return s.ErrorContent;
            }
            else if (value is ICollection)
            {

                ReadOnlyObservableCollection<System.Windows.Controls.ValidationError> errors = value as System.Collections.ObjectModel.ReadOnlyObservableCollection<System.Windows.Controls.ValidationError>;
                if (errors.Count > 0 && !string.IsNullOrEmpty(errors[0].ErrorContent.ToString()))
                    return errors[0].ErrorContent;
            }

            return "NO ERRORI";
        }
    }



    public class Base64ImageConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            try
            {
                BitmapImage bi = new BitmapImage();
                string s = null;
                byte[] image = null;

                if (value.GetType() == typeof(String))
                {
                    s = value as string;
                    image = System.Convert.FromBase64String(s);
                }
                else
                    image = value as byte[];

                if (image == null)
                    return null;


                //bi.BeginInit();
                Bitmap img = new Bitmap(new MemoryStream(image));
                System.Drawing.Color backColor = img.GetPixel(1, 1);
                img.MakeTransparent(backColor);


                MemoryStream ms = new MemoryStream();

                img.Save(ms, ImageFormat.Png);

                bi.BeginInit();
                bi.StreamSource = ms;
                bi.EndInit();


                return bi;
            }
            catch (Exception)
            {

                return null;
            }

        }



        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }

    public class RatingConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            try
            {
                double valueDouble =  Math.Abs((double)value);

                if (valueDouble > 0 && valueDouble <= 5.0)
                    return 0.1;
                if (valueDouble > 5.0 && valueDouble <= 10.0)
                    return 0.2;
                else if (valueDouble > 10.0 && valueDouble <= 25.0)
                    return 0.3;
                else if (valueDouble > 25.0 && valueDouble <= 35.0 )
                    return 0.4;
                else if (valueDouble > 35.0 && valueDouble <= 60.0)
                    return 0.8;
                else if (valueDouble > 60.0 )
                    return 1;

                return 0.0;

            }
            catch (Exception)
            {

                return null;
            }

        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }


    public class RatingColorConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            try
            {
                if (Math.Sign((double)value) == -1)
                    return Colors.Red;
                else
                    return Colors.Yellow;

            }
            catch (Exception)
            {

                return null;
            }

        }

        public object ConvertBack(object value, Type targetType, object parameter, System.Globalization.CultureInfo culture)
        {
            throw new NotImplementedException();
        }
    }
}
