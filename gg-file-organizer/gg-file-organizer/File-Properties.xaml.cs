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
using System.ComponentModel;
using MessageBox = System.Windows.Forms.MessageBox;
using MessageBoxImage = System.Windows.Forms.MessageBoxIcon;
using MessageBoxButton = System.Windows.Forms.MessageBoxButtons;
using MessageBoxResult = System.Windows.Forms.DialogResult;
using System.Runtime.InteropServices;

namespace gg_file_organizer
{
    public class Model2 : INotifyPropertyChanged
    {
        string _text;
        public string Text { get { return _text; } set { _text = value; OnPropertyChanged("Text"); } }

        public event PropertyChangedEventHandler PropertyChanged;

        protected virtual void OnPropertyChanged(string propertyName)
        {
            if (PropertyChanged != null)
            {
                PropertyChanged(
                    this,
                    new PropertyChangedEventArgs(propertyName)
                    );
            }
        }
    }
    /// <summary>
    /// Interaction logic for File_Properties.xaml
    /// </summary>
    public partial class File_Properties : Window
    {
        public File_Properties()
        {
            InitializeComponent();
            DataContext = new Model2();
            copy_key_handler();
        }
        private void FolderF1_MouseEnter(object sender, MouseEventArgs e)
        {
            string packUri = @"pack://application:,,,/Resources/Folder_onhover.png";
            F1.Source = new ImageSourceConverter().ConvertFromString(packUri) as ImageSource;
        }
        private void FolderF2_MouseEnter(object sender, MouseEventArgs e)
        {
            string packUri = @"pack://application:,,,/Resources/copy_onhover.png";
            F2.Source = new ImageSourceConverter().ConvertFromString(packUri) as ImageSource;
        }
        public string SourceFilePath = string.Empty;

        private void FolderF1_MouseLeave(object sender, MouseEventArgs e)
        {
            string packUri = @"pack://application:,,,/Resources/Folder.png";
            F1.Source = new ImageSourceConverter().ConvertFromString(packUri) as ImageSource;
        }
        private void FolderF2_MouseLeave(object sender, MouseEventArgs e)
        {
            string packUri = @"pack://application:,,,/Resources/copy.png";
            F2.Source = new ImageSourceConverter().ConvertFromString(packUri) as ImageSource;
        }

        private void Source_open_btn_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.OpenFileDialog Source_File_dialog = new System.Windows.Forms.OpenFileDialog();
            Source_File_dialog.Title = "Please choose the file to get the properties";
            MessageBoxResult result = Source_File_dialog.ShowDialog();
            if (result.ToString() == "OK")
            {
                Source_txt_box.Text = Source_File_dialog.FileName;
                SourceFilePath = Source_txt_box.Text;
            }
        }
        List<FileData> items = new List<FileData>();
        private void Verify_btn_Click(object sender, RoutedEventArgs e)
        {
            FileProperties.Items.Clear();
            List<string> arrHeaders = new List<string>();
            Shell32.Shell shell = new Shell32.Shell();
            var strFileName = Source_txt_box.Text;
            Shell32.Folder objFolder = shell.NameSpace(System.IO.Path.GetDirectoryName(strFileName));
            Shell32.FolderItem folderItem = objFolder.ParseName(System.IO.Path.GetFileName(strFileName));
            
            for (int i = 0; i < 320; i++)
            {
                string header = objFolder.GetDetailsOf(null, i);
                //if (String.IsNullOrEmpty(header))
                //break;
                arrHeaders.Add(header);
            }
            for (int i = 0; i < arrHeaders.Count; i++)
            {
                //Console.WriteLine("{0}\t{1}: {2}", i, arrHeaders[i], objFolder.GetDetailsOf(folderItem, i));
                if(!string.IsNullOrEmpty(objFolder.GetDetailsOf(folderItem, i)))
                    items.Add(new FileData() { ID = i, Name = arrHeaders[i], Value = objFolder.GetDetailsOf(folderItem,i) });
            }
            FileProperties.ItemsSource = items;
        }

        private void Copy_btn_Click(object sender, RoutedEventArgs e)
        {
            CopyFileProperties();
        }

        private void CopyFileProperties()
        {
            if (FileProperties.SelectedItems.Count != 0)
            {
                //where MyType is a custom datatype and the listview is bound to a 
                //List<MyType> called original_list_bound_to_the_listview
                List<FileData> selected = new List<FileData>();
                var sb = new StringBuilder();
                foreach (FileData s in FileProperties.SelectedItems)
                    selected.Add(s);
                foreach (FileData s in items)
                    if (selected.Contains(s))
                        sb.AppendLine("ID:" + s.ID + " Name:" + s.Name + " Value:" + s.Value);//or whatever format you want
                try
                {
                    System.Windows.Clipboard.SetData(DataFormats.Text, sb.ToString());
                }
                catch (COMException)
                {
                    MessageBox.Show("Sorry, unable to copy surveys to the clipboard. Try again.");
                }
            }
        }

        private void copy_key_handler()
        {
            ExecutedRoutedEventHandler handler = (sender_, arg_) => { CopyFileProperties(); };
            var command = new RoutedCommand("Copy", typeof(GridView));
            command.InputGestures.Add(new KeyGesture(Key.C, ModifierKeys.Control, "Copy"));
            FileProperties.CommandBindings.Add(new CommandBinding(command, handler));
            try
            { System.Windows.Clipboard.SetData(DataFormats.Text, ""); }
            catch (COMException)
            { }
        }
    }
    public class FileData
    {
        public int ID { get; set; }

        public string Name { get; set; }

        public string Value { get; set; }
    }
}
