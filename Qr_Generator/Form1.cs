using PdfSharp.Drawing;
using PdfSharp.Pdf;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using ZXing.Common;
using ZXing;
using OfficeOpenXml;

namespace Qr_Generator
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        private string outputFolder;
        private Bitmap labelBitmap;

        private void BtnConvert_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(txtBoxXLocation.Text))
                {
                    MessageBox.Show("Select the Excel Template, then press Generate QR button!", "Selection Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
                using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(txtBoxXLocation.Text)))
                {
                    ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[0];
                    int rowCount = 0;

                    for (int row = 4; row <= worksheet.Dimension.End.Row && !string.IsNullOrWhiteSpace(worksheet.Cells[row, 2].Text); row++)
                    {
                        rowCount++;
                    }

                    for (int row = 4; row <= worksheet.Dimension.End.Row; row++)
                    {
                        string partNumber = worksheet.Cells[row, 2].Text;
                        if (string.IsNullOrWhiteSpace(partNumber))
                            break;
                        string lotNumber = worksheet.Cells[row, 3].Text;
                        string serialNumber = worksheet.Cells[row, 4].Text;
                        string description = worksheet.Cells[row, 5].Text; 
                        
                        string webLink = worksheet.Cells[row, 6].Text;
                        string gtin = worksheet.Cells[row, 7].Text;
                        double.TryParse(gtin, out double gtinResult);
                        gtin = gtinResult.ToString("0");
                        string packQty = worksheet.Cells[row, 9].Text;
                        string dominoProduct = worksheet.Cells[row, 10].Text;
                        string batch = worksheet.Cells[row, 11].Text;

                        string countryOfOrgin = worksheet.Cells[row, 12].Text;
                        string QTY = worksheet.Cells[row, 13].Text;
                        string Manufacture = worksheet.Cells[row, 14].Text;
                        string date = worksheet.Cells[row, 15].Text;

                        string qrLink = GenerateQrLink(webLink, gtin, serialNumber, packQty, dominoProduct, batch);
                        GenerateLabel(partNumber, lotNumber, serialNumber, date, packQty, description, qrLink, countryOfOrgin, QTY, Manufacture);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (!string.IsNullOrWhiteSpace(txtBoxXLocation.Text))
                {
                    MessageBox.Show("Label saved as PDF successfully in " + outputFolder, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
        }
        private string GenerateQrLink(string webLink, string gtin, string serial, string packQty, string dominoProduct, string batch)
        {
            return $"https://{webLink}01{gtin}21{serial}30{packQty}240{dominoProduct}10{batch}";
        }

        private void BtnBrowse_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.Title = "Select the Excel File";
                openFileDialog.Filter = "Excel Files (*.xls;*.xlsx)|*.xls;*.xlsx";

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    txtBoxXLocation.Text = openFileDialog.FileName;
                }
                else
                {
                    MessageBox.Show("Load a valid file location!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
            }
        }

        private void GenerateLabel(string pn, string lotNo, string sn, string date, string qty, string des, string qrLink, string countryOfOrgin, string QTY, string Manufacture)
        {
            try
            {
                string filename = Path.Combine(Application.StartupPath, "Domino.png");

                if (!File.Exists(filename))
                {
                    MessageBox.Show("Image file not found: " + filename);
                    return;
                }

                // Define output folder
                outputFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Output");
                if (!Directory.Exists(outputFolder))
                    Directory.CreateDirectory(outputFolder);

                string pdfPath = Path.Combine(outputFolder, sn + ".pdf");
                string pngPath = Path.ChangeExtension(pdfPath, ".png");

                int width1 = 290;
                int height1 = 181;

                using (Bitmap bitmap1 = new Bitmap(width1, height1))
                {
                    //bitmap1.SetResolution(100, 100);
                    using (Graphics graphics = Graphics.FromImage(bitmap1))
                    {
                        graphics.Clear(Color.White);

                        // Load and resize logo
                        using (Image original = Image.FromFile(filename))
                        using (Bitmap logo = new Bitmap(original, new Size(80, 40))) // Adjusted logo size
                        {
                            graphics.DrawImage(logo, new PointF(5, 1)); // Adjusted position
                        }

                        using (Font font = new Font("Arial", 6f)) // Reduced font size
                            graphics.DrawString("Genuine Spare", font, Brushes.Black, new PointF(100f, 20f));

                        using (Font font = new Font("Arial", 8f, FontStyle.Regular))
                        {
                            graphics.DrawString("PN:", font, Brushes.Black, new PointF(10f, 35f));
                            graphics.DrawString(pn, font, Brushes.Black, new PointF(60f, 35f));
                        }

                        // Generate barcode for PN
                        using (Bitmap barcodePN = GenerateBarcode(pn, 150, 25)) // Adjusted size
                            graphics.DrawImage(barcodePN, 15, 48);

                        using (Font font = new Font("Arial", 6f))
                        {
                            StringFormat format = new StringFormat();
                            format.Alignment = StringAlignment.Near;
                            format.LineAlignment = StringAlignment.Near;
                            format.FormatFlags = StringFormatFlags.LineLimit; // Ensures text wraps within the rectangle

                            RectangleF layoutRectangle = new RectangleF(170f, 45f, 100f, 50f); // Adjust width and height

                            graphics.DrawString(des, font, Brushes.Black, layoutRectangle, format);
                        }
                            //graphics.DrawString(des, font, Brushes.Black, new PointF(170f, 45f));

                        using (Font font = new Font("Arial", 5f,FontStyle.Bold))
                        {
                            StringFormat format = new StringFormat();
                            format.Alignment = StringAlignment.Near;
                            format.LineAlignment = StringAlignment.Near;
                            format.FormatFlags = StringFormatFlags.LineLimit; // Ensures text wraps within the rectangle

                            RectangleF layoutRectangle = new RectangleF(130f, 100f, 110f, 60f); // Adjust width and height

                            graphics.DrawString(Manufacture, font, Brushes.Black, layoutRectangle, format);
                        }

                        using (Font font = new Font("Arial", 6f))
                            graphics.DrawString(date, font, Brushes.Black, new PointF(220f, 10f));

                        using (Font font = new Font("Arial", 7f, FontStyle.Regular))
                        {
                            graphics.DrawString("LOT", font, Brushes.Black, new PointF(12f, 75f));
                            graphics.DrawString("no:", font, Brushes.Black, new PointF(12f, 83f));
                            graphics.DrawString(lotNo, font, Brushes.Black, new PointF(70f, 80f));
                        }

                        // Generate barcode for LOT No
                        using (Bitmap barcodeLot = GenerateBarcode(lotNo, 95, 25)) // Adjusted size
                            graphics.DrawImage(barcodeLot, 25, 93);

                        using (Font font = new Font("Arial", 8f, FontStyle.Regular))
                        {
                            graphics.DrawString("SN:", font, Brushes.Black, new PointF(10f, 123f));
                            graphics.DrawString(sn, font, Brushes.Black, new PointF(60f, 123f));
                        }

                        // Generate barcode for SN
                        using (Bitmap barcodeSN = GenerateBarcode(sn, 80, 25)) // Adjusted size
                            graphics.DrawImage(barcodeSN, 25, 135);

                        using (Font font = new Font("Arial", 5f, FontStyle.Bold))
                            graphics.DrawString("Country of Origin: " + countryOfOrgin, font, Brushes.Black, new PointF(30f, 165f));

                        using (Font font = new Font("Arial", 7f, FontStyle.Regular))
                            graphics.DrawString("QTY: " + QTY, font, Brushes.Black, new PointF(150f, 162f));

                        // Generate QR Code
                        using (Bitmap qrCode = GenerateQRCode(qrLink, 60,60)) // Reduced QR size
                            graphics.DrawImage(qrCode, 220,125);

                        // Save label
                        SaveLabelAsPdf(bitmap1, pdfPath, width1, height1);
                        SaveLabelAsPng(bitmap1, pngPath);
                    }
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
        }
        private void SaveLabelAsPng(Bitmap bitmap, string pngPath)
        {
            try
            {
                bitmap.Save(pngPath, ImageFormat.Png);
            }
            catch (Exception ex)
            {
                MessageBox.Show("PNG Save Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // Generate Barcode (Code-128)
        private Bitmap GenerateBarcode(string data, int width, int height)
        {
            BarcodeWriter barcodeWriter = new BarcodeWriter
            {
                Format = BarcodeFormat.CODE_128,
                Options = new EncodingOptions
                {
                    Height = height,
                    Width = width,
                    PureBarcode = true
                },
                Renderer = new ZXing.Rendering.BitmapRenderer() // Ensure correct rendering
            };
            return barcodeWriter.Write(data);
        }

        // Generate QR Code
        private Bitmap GenerateQRCode(string data, int width, int height)
        {
            BarcodeWriter qrCodeWriter = new BarcodeWriter
            {
                Format = BarcodeFormat.QR_CODE,
                Options = new EncodingOptions
                {
                    Height = height,
                    Width = width,
                    Margin = 0
                },
                Renderer = new ZXing.Rendering.BitmapRenderer()
            };
            return qrCodeWriter.Write(data);
        }

        private void SaveLabelAsPdf(Bitmap bitmap, string pdfPath, int width, int height)
        {
            try
            {
                PdfDocument pdfDocument = new PdfDocument();
                PdfPage page = pdfDocument.AddPage();
                page.Width = XUnit.FromPoint(width);
                page.Height = XUnit.FromPoint(height);

                using (XGraphics xgraphics = XGraphics.FromPdfPage(page))
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    // Increase DPI for better text clarity
                    
                    bitmap.SetResolution(300, 300);
                    bitmap.Save(memoryStream, ImageFormat.Png);
                    memoryStream.Position = 0;

                    XImage image = XImage.FromStream(memoryStream);
                    xgraphics.DrawImage(image, 0, 0, width, height);
                }

                pdfDocument.Save(pdfPath);
                pdfDocument.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("PDF Save Error: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void DrawWrappedText(Graphics g, string text, Font font, RectangleF layoutRect)
        {
            using (StringFormat format = new StringFormat())
            {
                format.FormatFlags = StringFormatFlags.LineLimit;
                format.Trimming = StringTrimming.EllipsisWord;
                g.DrawString(text, font, Brushes.Black, layoutRect, format);
            }
        }

        private void BtnDownloadTemplate_Click(object sender, EventArgs e)
        {
            try
            {
                string templatePath = Path.Combine(Application.StartupPath, "Template", "Label_Template.xlsx");
                string outputFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Output");

                if (!Directory.Exists(outputFolder))
                    Directory.CreateDirectory(outputFolder);

                string destinationPath = Path.Combine(outputFolder, "Label_Template.xlsx");
                int fileIndex = 1;

                while (File.Exists(destinationPath))
                {
                    destinationPath = Path.Combine(outputFolder, $"Label_Template_{fileIndex}.xlsx");
                    fileIndex++;
                }

                if (File.Exists(templatePath))
                {
                    File.Copy(templatePath, destinationPath, true);
                    MessageBox.Show("Template has been downloaded to " + destinationPath, "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Template file not found!", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}
