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
using System.Windows.Shapes;

namespace ConvertSpecLevel
{
    /// <summary>
    /// Interaction logic for frmAdjustWindows.xaml
    /// </summary>
    public partial class frmAdjustWindows : Window
    {
        public frmAdjustWindows()
        {
            InitializeComponent();
        }

        public string GetSelectedSpecLevel()
        {
            if (rbCompleteHome.IsChecked == true)
                return rbCompleteHome.Content.ToString();
            else
                return rbCompleteHomePlus.Content.ToString();
        }

        private void chkAdjustWindowHeadHeights_Checked(object sender, RoutedEventArgs e)
        {
            chkAdjustWindowHeights.IsEnabled = true;
        }

        private void chkAdjustWindowHeadHeights_Unchecked(object sender, RoutedEventArgs e)
        {
            chkAdjustWindowHeights.IsEnabled = false;
            chkAdjustWindowHeights.IsChecked = false;
        }

        public bool IsAdjustWindowHeadHeightsChecked()
        {
            return chkAdjustWindowHeadHeights.IsChecked == true;
        }

        public bool IsAdjustWindowHeightsChecked()
        {
            return chkAdjustWindowHeights.IsChecked == true;
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

        private void btnHelp_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // launch the help site with user's default browser
                string helpUrl = "https://lifestyle-usa-design.atlassian.net/wiki/spaces/MFS/pages/472711169/Spec+Level+Conversion?atlOrigin=eyJpIjoiMmU4MzM3NzFmY2NlNDdiNjk1MjY2M2MyYzZkMjY2YWQiLCJwIjoiYyJ9";
                Process.Start(new ProcessStartInfo
                {
                    FileName = helpUrl,
                    UseShellExecute = true
                });

            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show("An error occurred while trying to display help: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
    }
}

