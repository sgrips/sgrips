using Sgrips.WishList.Model;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;

namespace Sgrips.WishList.WPF
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private ObservableCollection<Model.WishListCode> _listCodes = new ObservableCollection<Model.WishListCode>();
        private ObservableCollection<FrameworkElement> _listCodesMenu = new ObservableCollection<FrameworkElement>();

        private string _currentWishCode;
        private MenuItem _lastCheckMenuItem = null;

        public static RoutedCommand WishChangeCmd = new RoutedCommand();
        private void WishChangeCmd_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            if (_lastCheckMenuItem != null) _lastCheckMenuItem.IsChecked = false;
            _lastCheckMenuItem = e.OriginalSource as MenuItem;
            _currentWishCode = e.Parameter as string;
            LoadWishItems(_currentWishCode);
        }

        private void WishChangeCmd_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = true;
        }

        public static RoutedCommand WishRefreshCmd = new RoutedCommand();

        private void WishRefreshCmd_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        {
            e.CanExecute = _lastCheckMenuItem != null;
        }

        private void WishRefreshCmd_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            var oldCur = this.Cursor;
            this.Cursor = Cursors.Wait;

            Amazon.DabaseHelper.Update(_currentWishCode,
                new Amazon.DabaseHelper.LoadedHandler(
                    (a, b) => { LoadWishItems(_currentWishCode); this.Cursor = oldCur; }));
        }


        public MainWindow()
        {
            InitializeComponent();
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {

            try
            {
                if (Amazon.DabaseHelper.GetWishListTypesFromDB() == null)
                    Amazon.DabaseHelper.CreateDB();

                (Amazon.DabaseHelper.GetWishListTypesFromDB()
                    as List<Model.WishListType>).ForEach((f) =>
                            f.Codes.ForEach((item) =>
                    _listCodes.Add(new WishListCode { Id = item.Id, Description = item.Description })
                    ));

                GenerateMenuItemCodeWish(_listCodes);
            }
            catch (Exception)
            {


            }




        }

        private void GenerateMenuItemCodeWish(ObservableCollection<Model.WishListCode> _listCodes)
        {

            _listCodesMenu.Clear();

            foreach (var item in _listCodes)
            {
                _listCodesMenu.Add(new MenuItem
                {
                    IsCheckable = true,
                    Header = item.Description,
                    Command = WishChangeCmd,
                    CommandParameter = item.Id
                });
            }

            _listCodesMenu.Add(new Separator());

            //< Separator />
            //< MenuItem Header = "_Aggiorna" Click = "button_Update_Click" />
            _listCodesMenu.Add(new MenuItem
            {
                IsCheckable = false,
                Header = "_Aggiorna",
                Command = WishRefreshCmd
            });




            CollectionViewSource col = (CollectionViewSource)this.Resources["wishListCodes"];
            col.Source = _listCodesMenu;
        }


        private void LoadWishItems(string id)
        {
            try
            {
                CollectionViewSource col = (CollectionViewSource)this.Resources["wishList"];
                col.Source = null;

                Wish listWishNew = Amazon.DabaseHelper.GetFromDB(id);
                if (listWishNew == null)
                    WishRefreshCmd_Executed(this, null);


                col.Source = listWishNew.Items;
            }
            catch (Exception ex)
            {
            }
        }


        private void Button_Close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void MnuModificaListe_Click(object sender, RoutedEventArgs e)
        {
            EditWishList edt = new EditWishList();
            edt.Show(ref _listCodes);
            GenerateMenuItemCodeWish(_listCodes);
        }

    }
}
