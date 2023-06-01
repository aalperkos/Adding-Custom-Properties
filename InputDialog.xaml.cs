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

namespace CustomPa
{
    
    public partial class InputDialog : Window
    {
        // properties to get
        public string CategoryName { get; set; }
        public string PropertyName { get; set; }
        public string PropertyValue { get; set; }

        public InputDialog()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //get user-input data
            CategoryName = cateName.Text;
            PropertyName = propName.Text;
            PropertyValue = propValue.Text;
            // form close
            this.Close();
        }
    }
}