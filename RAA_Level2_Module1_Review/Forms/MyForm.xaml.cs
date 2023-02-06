using Autodesk.Revit.UI;
using Microsoft.Win32;
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


namespace RAA_Level2_Module1_Review
{
    /// <summary>
    /// Interaction logic for Window.xaml
    /// </summary>
    public partial class MyForm : Window
    {
        public MyForm()
        {
            InitializeComponent();
        }

        private void btnSelect_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog selectFile = new OpenFileDialog();
            selectFile.Multiselect = false;
            //selectFile.InitialDirectory = @"C:\";
            selectFile.RestoreDirectory = true;
            selectFile.Filter = "*.csv file (*.csv)|*.csv";
            
            if(selectFile.ShowDialog() == true)
            {
                tbxSelect.Text = selectFile.FileName;
            }
        }

        private void btnOK_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = true;
            this.Close();
        }

        private void btnCancel_Click(object sender, RoutedEventArgs e)
        {
            this.DialogResult = false;
            this.Close();
        }

        internal string GetFilePath()
        {
            if (tbxSelect.Text != "")
                return tbxSelect.Text;
            else
                TaskDialog.Show("Error", "Please select a file!");

            return null;
        }

        internal string GetUnits()
        {
            if(rbImperial.IsChecked == true)
            {
                return "imperial";
            }

            return "metric";
        }

        internal bool CreateFloorPlan()
        {
            if (cbFloorPlan.IsChecked == true)
                return true;

            return false;
        }

        internal bool CreateCeilingPlan()
        {
            if(cbCeilingPlan.IsChecked == true)
                return true;

            return false;
        }
    }
}
