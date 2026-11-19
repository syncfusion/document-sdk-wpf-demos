using Microsoft.Win32;
using syncfusion.dociodemos.wpf;
using Syncfusion.DocIO.DLS;
using Syncfusion.Office;
using System;
using System.IO;
using System.Windows;
using syncfusion.demoscommon.wpf;
using Syncfusion.DocIO;

namespace syncfusion.dociodemos.wpf
{
    public partial class DigitalSignature : DemoControl
    {
        private string documentPath = string.Empty;
        private string imagePath = string.Empty;
        private string certificatePath = string.Empty;

        public DigitalSignature()
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
                "Word Documents (*.docx)|*.docx|Word Documents (*.doc)|*.doc";

            if (dialog.ShowDialog() == true)
            {
                documentPath = dialog.FileName;
                txtDocument.Text = Path.GetFileName(documentPath);
            }
        }

        private void btnBrowseImage_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();

            dialog.Filter =
                "Image Files (*.png;*.jpg;*.jpeg)|*.png;*.jpg;*.jpeg";

            if (dialog.ShowDialog() == true)
            {
                imagePath = dialog.FileName;
                txtImage.Text = Path.GetFileName(imagePath);
            }
        }

        private void btnBrowseCertificate_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog dialog = new OpenFileDialog();

            dialog.Filter = "Certificate (*.pfx)|*.pfx";

            if (dialog.ShowDialog() == true)
            {
                certificatePath = dialog.FileName;
                txtCertificate.Text = Path.GetFileName(certificatePath);
            }
        }

        private void btnCreate_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(certificatePath))
                {
                    MessageBox.Show(
                        "Please choose a certificate.",
                        "Certificate Required",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                using (WordDocument document = new WordDocument())
                {
                    IWSection section = document.AddSection();

                    section.AddParagraph()
                           .AppendText("Digital Signature Sample");

                    OfficeDigitalSignatureCertificate certificate =
                        new OfficeDigitalSignatureCertificate(
                            certificatePath,
                            passwordBox.Password);

                    SignatureSettings settings = new SignatureSettings
                    {
                        SignTime = DateTime.Now,
                        Comments = txtComment.Text
                    };

                    if (radioXmlDsig.IsChecked == true)
                        settings.XmlDsigLevel = XmlDsigLevel.XmlDSig;
                    else
                        settings.XmlDsigLevel = XmlDsigLevel.XAdEsEpes;

                    if (radioVisible.IsChecked == true)
                    {
                        SignatureLineSettings signatureLine =
                            new SignatureLineSettings()
                            {
                                Signer = "John Doe",
                                SignerTitle = "Manager",
                                Email = "john.doe@company.com",
                                ShowDate = true
                            };

                        WPicture picture =
                            section.AddParagraph()
                                   .AppendSignatureLine(signatureLine, 192, 96)
                            as WPicture;

                        settings.SignatureLineId =
                            picture.SignatureLine.Id;

                        if (!string.IsNullOrEmpty(imagePath) &&
                            File.Exists(imagePath))
                        {
                            settings.SignatureLineImage =
                                File.ReadAllBytes(imagePath);
                        }
                    }

                    document.AddDigitalSignature(
                        certificate,
                        settings);
                    bool isValid = document.DigitalSignatures.IsValid;
                    string outputFile = "DigitalSignature.docx";

                    document.Save(outputFile, FormatType.Docx);

                    if (MessageBox.Show(
                        "Do you want to view the generated Word document?",
                        "Document Created",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Information)
                        == MessageBoxResult.Yes)
                    {
                        System.Diagnostics.Process process =
                            new System.Diagnostics.Process();

                        process.StartInfo =
                            new System.Diagnostics.ProcessStartInfo(outputFile)
                            {
                                UseShellExecute = true
                            };

                        process.Start();
                    }
                }
            }
            catch
            {

            }
        }
        private void btnSign_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(documentPath))
                {
                    MessageBox.Show(
                        "Please choose a Word document.",
                        "Document Required",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                if (string.IsNullOrEmpty(certificatePath))
                {
                    MessageBox.Show(
                        "Please choose a certificate.",
                        "Certificate Required",
                        MessageBoxButton.OK,
                        MessageBoxImage.Warning);
                    return;
                }

                using (WordDocument document =
                    new WordDocument(documentPath, FormatType.Automatic))
                {
                    OfficeDigitalSignatureCertificate certificate =
                        new OfficeDigitalSignatureCertificate(
                            certificatePath,
                            passwordBox.Password);

                    SignatureSettings settings =
                        new SignatureSettings
                        {
                            SignTime = DateTime.Now,
                            Comments = txtComment.Text
                        };

                    if (radioXmlDsig.IsChecked == true)
                        settings.XmlDsigLevel = XmlDsigLevel.XmlDSig;
                    else
                        settings.XmlDsigLevel = XmlDsigLevel.XAdEsEpes;

                    if(radioVisible.IsChecked == true)
                    {
                        WPicture picture = GetSignatureLine(document);

                        if (picture != null)
                        {
                            settings.SignatureLineId =
                                picture.SignatureLine.Id;

                            if (!string.IsNullOrEmpty(imagePath) &&
                                File.Exists(imagePath))
                            {
                                settings.SignatureLineImage =
                                    File.ReadAllBytes(imagePath);
                            }
                        }
                    }

                    document.AddDigitalSignature(
                        certificate,
                        settings);

                    string outputFile = "SignedWordDocument.docx";

                    document.Save(outputFile, FormatType.Docx);

                    if (MessageBox.Show(
                        "Do you want to view the signed Word document?",
                        "Document Created",
                        MessageBoxButton.YesNo,
                        MessageBoxImage.Information)
                        == MessageBoxResult.Yes)
                    {
                        System.Diagnostics.Process process =
                            new System.Diagnostics.Process();

                        process.StartInfo =
                            new System.Diagnostics.ProcessStartInfo(outputFile)
                            {
                                UseShellExecute = true
                            };

                        process.Start();
                    }
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

        private WPicture GetSignatureLine(WordDocument document)
        {
            foreach (WSection section in document.Sections)
            {
                foreach (WParagraph paragraph in section.Paragraphs)
                {
                    foreach (Entity entity in paragraph.ChildEntities)
                    {
                        if (entity is WPicture picture &&
                            picture.IsSignatureLine)
                        {
                            return picture;
                        }
                    }
                }
            }

            return null;
        }
    }
}
