using SprawdzRozklad.Views;
using System;
using System.Reflection;
using System.Windows;
using System.Windows.Controls;


namespace SprawdzRozklad
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            var version = Assembly.GetExecutingAssembly().GetName().Version;

            this.Title = $"Informica Migration Tool {version.Major}.{version.Minor}";

            MainContent.Content = new ViewA();
        }

        private readonly Dictionary<string, UserControl> views = new();
        private UserControl GetView(string key)
        {
            if(!views.TryGetValue(key, out var view))
            {
                view = key switch
                {
                    "A" => new ViewA(),
                    "B" => new ViewB(),
                    _ => throw new NotSupportedException()
                };
                views[key] = view;
            }
            return view;
        }
     
        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            var tab = (TabItem)((TabControl)sender).SelectedItem;
            MainContent.Content = GetView(tab.Tag.ToString());
        }
    }
}