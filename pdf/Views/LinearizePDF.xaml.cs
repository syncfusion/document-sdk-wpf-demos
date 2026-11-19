#region Copyright Syncfusion Inc. 2001-2020.
// Copyright Syncfusion Inc. 2001-2020. All rights reserved.
// Use of this code is subject to the terms of our license.
// A copy of the current license can be obtained at any time by e-mailing
// licensing@syncfusion.com. Any infringement will be prosecuted under
// applicable laws. 
#endregion
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Microsoft.Win32;
using Syncfusion.Pdf;
using Syncfusion.Pdf.Parsing;
using Syncfusion.Windows.Shared;
using syncfusion.demoscommon.wpf;
using System.IO;

namespace syncfusion.pdfdemos.wpf
{
    /// <summary>
    /// Interaction logic for LinearizePDF.xaml
    /// </summary>
    public partial class LinearizePDF : DemoControl
    {
        # region Private Members
        private string m_fullPath;
        # endregion

        # region Constructor
        /// <summary>
        /// Window constructor
        /// </summary>
        public LinearizePDF()
        {
            InitializeComponent();

            m_fullPath = @"Assets\PDF\jQuery_Succinctly.pdf";
            textBox1.Text = "jQuery_Succinctly.pdf";
        }
        #endregion
        #region Dispose
        protected override void Dispose(bool disposing)
        {
            //Release all resources
            base.Dispose(disposing);
        }
        # endregion
        # region Events
        /// <summary>
        /// Linearizes PDF
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnLinearize_Click(object sender, RoutedEventArgs e)
        {
            if (this.textBox1.Text == "")
            {
                System.Windows.MessageBox.Show("Please select a PDF document");
                textBox1.Focus();
            }
            else
            {
                // FILE SIZE LIMIT (10 MB)
                const long MAX_FILE_SIZE = 10 * 1024 * 1024;
                FileInfo fileInfo = new FileInfo(m_fullPath);
                if (fileInfo.Length > MAX_FILE_SIZE)
                {
                    System.Windows.MessageBox.Show(string.Format("The selected PDF file exceeds the maximum allowed size of 10 MB. File size: {0:F2} MB.", fileInfo.Length / (1024.0 * 1024.0)),
                        "File too large", System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Warning);
                    return;
                }

                try
                {
                    // Load the existing PDF document
                    PdfLoadedDocument loadedDocument = new PdfLoadedDocument(m_fullPath);

                    // Linearize the PDF document for fast web viewing
                    loadedDocument.FileStructure.IsLinearized = true;

                    // Save the linearized PDF file with the input filename
                    string outputFileName = System.IO.Path.GetFileName(m_fullPath);
                    loadedDocument.Save(outputFileName);
                    loadedDocument.Close(true);

                    //Message box confirmation to view the created PDF document.
                    if (System.Windows.MessageBox.Show("PDF file linearized successfully! Do you want to view the PDF file?", "PDF File Linearized",
                        MessageBoxButton.YesNo, MessageBoxImage.Information) == MessageBoxResult.Yes)
                    {
                        //Launching the PDF file using the default Application.[Acrobat Reader]
                        System.Diagnostics.Process process = new System.Diagnostics.Process();
                        process.StartInfo = new System.Diagnostics.ProcessStartInfo(outputFileName) { UseShellExecute = true };
                        process.Start();
                    }
                }
                catch (Exception ex)
                {
                    System.Windows.MessageBox.Show("Error: " + ex.Message, "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        /// <summary>
        /// Gets the source document
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnBrowse_Click(object sender, RoutedEventArgs e)
        {
            Microsoft.Win32.OpenFileDialog file = new Microsoft.Win32.OpenFileDialog();
            file.Filter = "PDF Documents (*.pdf)|*.PDF";
            file.Title = "Choose PDF document";

            if (file.ShowDialog().Value)
            {
                textBox1.Text = file.SafeFileName;
                m_fullPath = file.FileName;
            }
        }

        # endregion
    }
}
