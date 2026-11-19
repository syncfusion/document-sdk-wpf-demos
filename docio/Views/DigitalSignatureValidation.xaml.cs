using Microsoft.Win32;
using syncfusion.demoscommon.wpf;
using Syncfusion.DocIO;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System;
using System.IO;
using System.Text;
using System.Windows;

namespace syncfusion.dociodemos.wpf
{
    public partial class DigitalSignatureValidation : DemoControl
    {
        private string documentPath = string.Empty;

        public DigitalSignatureValidation()
        {
            InitializeComponent();
        }

        protected override void Dispose(bool disposing)
        {
            base.Dispose(disposing);
        }

        private void btnBrowseDocument_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();

            dialog.Filter =
                "Word Documents (*.docx)|*.docx";

            if (dialog.ShowDialog() == true)
            {
                documentPath = dialog.FileName;
                txtDocument.Text =
                    Path.GetFileName(documentPath);
            }
        }

        private void btnValidate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                txtValidationResult.Clear();

                if (string.IsNullOrEmpty(documentPath))
                {
                    txtValidationResult.Text =
                        "Browse a DOCX document and then click the button to validate digital signature.";
                    return;
                }

                using (WordDocument document =
                    new WordDocument(documentPath, FormatType.Docx))
                {
                    StringBuilder builder =
                        new StringBuilder();

                    if (document.DigitalSignatures.Count > 0)
                    {
                        builder.AppendLine(
                            "The following digital signatures were found in " +
                            Path.GetFileName(documentPath));

                        int index = 1;

                        foreach (OfficeDigitalSignature signature
                            in document.DigitalSignatures)
                        {
                            builder.AppendLine();
                            builder.AppendLine(
                                string.Format("{0:D3}:", index));

                            if (signature.IsValid)
                                builder.AppendLine(
                                    "The signature is valid");
                            else
                                builder.AppendLine(
                                    "The signature is invalid");

                            builder.AppendLine(
                                "Signed at : " +
                                signature.SigningTime);

                            builder.AppendLine(
                                "Subject name : " +
                                signature.Certificate.Subject);

                            builder.AppendLine(
                                "Issuer name : " +
                                signature.Certificate.Issuer);

                            builder.AppendLine(
                                "Valid From : " +
                                signature.Certificate.ValidFrom);

                            builder.AppendLine(
                                "Valid To : " +
                                signature.Certificate.ValidTo);

                            builder.AppendLine();

                            index++;
                        }

                    }
                    else
                    {
                        builder.AppendLine(
                            "No digital signatures were found in " +
                            Path.GetFileName(documentPath));
                    }

                    txtValidationResult.Text =
                        builder.ToString();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    ex.Message,
                    "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }
    }
}
