using System;
using System.Collections.Generic;
using System.IO;
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
using Microsoft.Win32;


namespace WpfEmbeddedDocx
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

		private void ReadDocx(string path)
		{
			using (var stream = File.Open(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
			{
				var flowDocumentConverter = new DocxToFlowDocumentConverter(stream);
				flowDocumentConverter.Read();
				FlowDocumentScrollViewer.Document = flowDocumentConverter.Document;
				this.Title = Path.GetFileName(path);
			}
		}

		private void OnOpenFileClicked(object sender, RoutedEventArgs e)
		{
			var openFileDialog = new OpenFileDialog()
			{
				DefaultExt = ".docx",
				Filter = "Word documents (.docx)|*.docx"
			};

			if (openFileDialog.ShowDialog() == true)
				this.ReadDocx(openFileDialog.FileName);
		}
    }


}
