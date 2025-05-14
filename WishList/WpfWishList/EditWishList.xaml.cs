using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using Sgrips.WishList.Model;

namespace Sgrips.WishList.WPF
{

    /// <summary>
    /// Interaction logic for EditWishList.xaml
    /// </summary>
    public partial class EditWishList : Window
    {


        public ObservableCollection<Model.WishListCode> WishCodes { get; private set; }

        public static RoutedCommand AddWishCode = new RoutedCommand();

        private void AddWishCodeCmdCanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            if (WishCodeTextBox != null)
                e.CanExecute = !Validation.GetHasError(WishCodeTextBox) && WishCodeTextBox.Text.Length == 13;
        }



        private async void AddWishCodeCmdExecuted(object sender, ExecutedRoutedEventArgs e)
        {
            string wishOwner = await Amazon.DabaseHelper.GetWishNameFromURL(WishCodeTextBox.Text);
            WishCodes.Add(new Model.WishListCode { Id = WishCodeTextBox.Text, Description = wishOwner });
        }


        public EditWishList()
        {
            InitializeComponent();

            var Binding = BindingOperations.GetBinding(WishCodeTextBox, TextBox.TextProperty);
            Binding.ValidationRules.Add(new ValidateWishListCode());

        }



        private void cmdChiudi_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void cmdSalva_Click(object sender, RoutedEventArgs e)
        {
            IList<WishListType> types = Amazon.DabaseHelper.GetWishListTypesFromDB();
            foreach (var item in types)
            {
                if (item.IdType == "Amazon")
                {
                    item.Codes.Clear ();
                    item.Codes.AddRange(WishCodes);
                }

            }
            Amazon.DabaseHelper.UpdateCode(types);
            this.Close();
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="wishCodes"></param>
        internal void Show(ref ObservableCollection<WishListCode> wishCodes)
        {
            WishCodes = wishCodes;

            CollectionViewSource col = Resources["WishListCodes"] as CollectionViewSource;
            col.Source = WishCodes;

            this.ShowDialog();
        }
    }
}
