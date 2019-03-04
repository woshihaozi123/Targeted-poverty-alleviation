using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace WpfApplication1.Forms
{
    /// <summary>
    /// DataChange.xaml 的交互逻辑
    /// </summary>
    public partial class DataChange : Window
    {
        public DataChange()
        {
            InitializeComponent();
        }
      
        private void min_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }
        //窗体关闭
        private void close_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
           
        }

        private void Window_MouseMove(object sender, MouseEventArgs e)
        {
            if (e.LeftButton == MouseButtonState.Pressed)
            {
                this.DragMove();
                //Window.DragMove();
            }
        }

    }
}
