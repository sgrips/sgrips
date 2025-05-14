using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Sgrips.Controls
{
    /// <summary>
    /// Interaction logic for UserControl1.xaml
    /// </summary>
    public partial class SgripsRating : UserControl
    {
        public static DependencyProperty RateValueProperty = 
            DependencyProperty.Register("RateValue", typeof(double), typeof(SgripsRating), 
                new PropertyMetadata(0.5, OnRateValuePropertyChanged, OnCoerceRateValueProperty));
        public static DependencyProperty RateColorProperty =
            DependencyProperty.Register("RateColor", typeof(Color), typeof(SgripsRating), 
                new PropertyMetadata(Colors.Yellow, OnRateColorPropertyChanged, OnCoerceRateColorProperty));


        public SgripsRating()
        {
            InitializeComponent();
        }


        private static void OnRateValuePropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs baseValue)
        {
           
        }

        private static  object OnCoerceRateValueProperty(DependencyObject source, object baseValue)
        {
            return (double)baseValue;

        }

        public  double RateValue
        {
            get { return (double)GetValue(RateValueProperty); }
            set { SetValue(RateValueProperty, value); }
        }

        public Color RateColor
        {
            get { return (Color)GetValue(RateColorProperty); }
            set { SetValue(RateColorProperty, value); }
        }

        private static void OnRateColorPropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs baseValue)
        {

        }

        private static object OnCoerceRateColorProperty(DependencyObject source, object baseValue)
        {
            return (Color)baseValue;

        }

    }
}
