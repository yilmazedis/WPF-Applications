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
using System.IO;
using Microsoft.Win32;
using Microsoft.Office.Interop.Word;
using Window = System.Windows.Window;

namespace DocAnalysis
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }



        private void btnOpenFile_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Office.Interop.Word.Application wordApp = new Microsoft.Office.Interop.Word.Application();

            // Create OpenFileDialog 
            OpenFileDialog dlg = new OpenFileDialog();

            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".doc";
            dlg.Filter = "Word documents|*.doc;*.docx";

            // Get the selected file name and display in a TextBox 
            if (dlg.ShowDialog() == true)
            {
                // Open document 
                string filename = dlg.FileName;

                wordApp.Documents.Add(filename);

                var contents = wordApp.ActiveDocument.Content;

                //Console.WriteLine(contents.Text);

                String[] lines = contents.Text.Split(
                                new[] { "\r\n", "\r", "\n" },
                                StringSplitOptions.RemoveEmptyEntries
                            );

                string temp = "";
                foreach (string l in lines) {

                    //Console.WriteLine(l);

                    String[] sentences = l.Split(
                                new[] { "." },
                                StringSplitOptions.RemoveEmptyEntries
                            );

                   foreach(string sentence in sentences)
                    {

                        if (sentence.Split(new[] {" "}, StringSplitOptions.RemoveEmptyEntries).Length > 20)
                        {
                            Console.WriteLine(sentence);
                            txtEditor.Text += sentence;
                            txtEditor.Text += "\n\n";
                            
                        }
                    }
                }


            }
        }
    }
}
