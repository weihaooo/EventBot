using Microsoft.Bot.Builder.Dialogs;
using System;
using System.Threading.Tasks;
using Microsoft.Bot.Connector;
using Microsoft.WindowsAzure.Storage;
using System.Configuration;
using Microsoft.WindowsAzure.Storage.Blob;
using System.Net.Http;
using ZXing;
using System.Drawing;
using System.IO;
using System.Drawing.Imaging;

namespace Microsoft.Bot.Sample.ProactiveBot
{
    [Serializable]
    public class QRDialog : IDialog<object>
    {
        string code;

        public QRDialog(string message)
        {
            code = message;
        }

        public async Task StartAsync(IDialogContext context)
        {
            var storageAccount = CloudStorageAccount.Parse(ConfigurationManager.AppSettings["AzureWebJobsStorage"]);

            string destinationContainer = "qrcode";
            CloudBlobClient blobClient = storageAccount.CreateCloudBlobClient();
            var blobContainer = blobClient.GetContainerReference(destinationContainer);
            blobContainer.CreateIfNotExists();
            String newFileName = "";
            
            var barcodeWriter = new BarcodeWriter();
            barcodeWriter.Format = BarcodeFormat.QR_CODE;

            //Write qrcode into memorystream
            var ms = new MemoryStream();
            var imageAsBytes = barcodeWriter.Write(code);
            imageAsBytes.Save(ms, ImageFormat.Jpeg);

            try
            {
                newFileName = code + "_qrcode";

                // Set the permissions so the blobs are public. //ON HOLD
                BlobContainerPermissions permissions = new BlobContainerPermissions
                {
                    PublicAccess = BlobContainerPublicAccessType.Blob
                };
                blobContainer.SetPermissions(permissions);

                var newBlockBlob = blobContainer.GetBlockBlobReference(newFileName);

                try
                {
                    using (var memorystream1 = new MemoryStream(ms.ToArray()))
                    {
                        newBlockBlob.UploadFromStream(memorystream1);
                        newBlockBlob.Properties.ContentType = "image/jpeg";
                        newBlockBlob.SetProperties();
                    }
                    ms.Close();
                }
                catch (Exception ex)
                {
                    await context.PostAsync("Something is wrong here");
                }
            }
            catch (Exception e)
            {
                await context.PostAsync("Something went wrong with generating code. Please try again.");
                context.Done(this);
            }

            try
            {
                CloudBlockBlob blob = blobContainer.GetBlockBlobReference(newFileName);
                // create a barcode reader instance
                var barcodeReader = new BarcodeReader();
                var memoryStream = new MemoryStream();

                blob.DownloadToStream(memoryStream);

                // create an in memory bitmap
                var barcodeBitmap = (Bitmap)Bitmap.FromStream(memoryStream);

                // decode the barcode from the in memory bitmap
                var barcodeResult = barcodeReader.Decode(barcodeBitmap);

                var message = context.MakeMessage();
                memoryStream.Close();

                message.Text = "This is the qr code of " + barcodeResult.Text;

                message.Attachments.Add(new Attachment()
                {
                    ContentUrl = blob.Uri.ToString(),
                    ContentType = blob.Properties.ContentType,
                    Name = blob.Name
                });

                await context.PostAsync(message);
            }
            catch (Exception e)
            {
                await context.PostAsync("Something went wrong with generating code. Please try again.");
            }
            context.Done(this);
        }
    }
}