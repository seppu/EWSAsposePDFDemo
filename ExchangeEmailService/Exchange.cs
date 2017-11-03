using System;
using System.Collections.Generic;
using Microsoft.Exchange.WebServices.Data;
using System.IO;
using Aspose.Pdf;
using Aspose.Pdf.Facades;
using System.Drawing.Imaging;
using Aspose.Pdf.Text;
using MsgReader;
using System.Linq;

namespace ExchangeEmailService
{
    public class Exchange
    {
        ExchangeService service;
        string user;
        string domain = "DOMAINNAME"; //Please add the proper domain name.
        public Exchange(string username, string password,string email)
        {
            user = email;
            service = new ExchangeService(ExchangeVersion.Exchange2010_SP1); //Change the Exchange version to your relevant one.
            service.UseDefaultCredentials = false;
            service.Credentials = new WebCredentials(username, password, domain);
            service.AutodiscoverUrl(email, RedirectionUrlValidationCallback);
            //service.ImpersonatedUserId = new ImpersonatedUserId(ConnectingIdType.SmtpAddress, email);
            service.TraceEnabled = true;
            service.TraceFlags = TraceFlags.All;
        }
        private static bool RedirectionUrlValidationCallback(string redirectionUrl)
        {
            // The default for the validation callback is to reject the URL.
            bool result = false;

            Uri redirectionUri = new Uri(redirectionUrl);

            // Validate the contents of the redirection URL. In this simple validation
            // callback, the redirection URL is considered valid if it is using HTTPS
            // to encrypt the authentication credentials.
            if (redirectionUri.Scheme == "https")
            {
                result = true;
            }
            return result;
        }

        public void SendEmail(string toEmail,string subject,string message)
        {
            EmailMessage email = new EmailMessage(service);

            email.ToRecipients.Add(toEmail);

            email.Subject = subject;
            email.Body = new MessageBody(message);

            email.Send();
        }

        public void SaveUnreadEmail()
        {
            List<SearchFilter> searchFilterCollection = new List<SearchFilter>();
            searchFilterCollection.Add(new SearchFilter.IsEqualTo(EmailMessageSchema.IsRead, false));
            SearchFilter searchFilter = new SearchFilter.SearchFilterCollection(LogicalOperator.And, searchFilterCollection);

            ItemView view = new ItemView(100);
            view.PropertySet = new PropertySet(BasePropertySet.IdOnly, ItemSchema.Subject);
            view.Traversal = ItemTraversal.Shallow;

            FindItemsResults<Item> findResults;

            //Redemption.RDOSession rs;
            PropertySet props = new PropertySet(EmailMessageSchema.MimeContent);


            HtmlLoadOptions options = new HtmlLoadOptions();
            string dataDir = @"C:\Test\";

            do
            {
                findResults = service.FindItems(new FolderId(WellKnownFolderName.Inbox, new Mailbox(user)), searchFilter, view);
                foreach (var item in findResults)
                {
                    var email = EmailMessage.Bind(service, item.Id, props);

                    using (MemoryStream ms = new MemoryStream(email.MimeContent.Content))
                    {

                       var msgReader = new Reader();
                        List<MemoryStream> streams = msgReader.ExtractToStream(ms, true);


                        List<Document> documents = new List<Document>();
                        Document[] docs = new Document[3];
                        Document mailDoc = new Document(streams[0], options);

                        documents.Add(mailDoc);

                        int j = 1;
                        item.Load();
                        if (item.HasAttachments)
                        {
                            foreach (var i in item.Attachments)
                            {

                                // Load the file attachment into memory. This gives you access to the attachment content, which
                                // is a byte array that you can use to attach this file to another item. This results in a GetAttachment operation
                                // call to EWS.

                                string attchType = Path.GetExtension(i.Name);
                                switch (attchType)
                                {
                                    case ".txt":
                                    case ".rtf":
                                    case ".doc":
                                    case ".docx":
                                        documents.Add(new Document(ConvertDocToPdf(streams[j])));
                                        break;
                                    case ".pdf":
                                        documents.Add(new Document(streams[j]));
                                        break;
                                    case ".jpg":
                                    case ".tif":
                                    case ".tiff":
                                    case ".png":
                                    case ".wmf":
                                    case ".gif":
                                        documents.Add(new Document(EmbedImageToPdf(streams[j])));
                                        //documents.Add(new Document(ConvertImageToPdf(streams[j])));
                                        break;
                                    default:
                                        //documents.Add(CreateEmptyPageWithErrorMessage(i.Name));
                                        break;
                                }
                                j++;
                            }
                        }

                        Document footerDoc = new Document(streams[j], options);
                        documents.Add(footerDoc);

                        // Save the output as PDF document
                        string pdfFilename = dataDir + Guid.NewGuid().ToString() + ".pdf";
                        Document pdfDoc = new Document();
                        PdfFileEditor pdfEditor = new PdfFileEditor();
                        pdfEditor.Concatenate(documents.ToArray(), pdfDoc);
                        pdfDoc.Save(pdfFilename);
                    }
                }
                var itemIds = from item in findResults
                              select item.Id;
                service.MoveItems(itemIds, (Folder.Bind(service, new FolderId(WellKnownFolderName.MsgFolderRoot,user))
                    .FindFolders(new SearchFilter.IsEqualTo(FolderSchema.DisplayName, "Archive"), new FolderView(1))
                    .FirstOrDefault(x => x.DisplayName == "Archive")).Id);

                view.Offset = Convert.ToInt32(findResults.NextPageOffset);

            } while (findResults.MoreAvailable);


        }

        public Document CreateEmptyPageWithErrorMessage(string name)
        {
            Document emptyDoc = new Document();
            emptyDoc.Pages.Insert(1);
            Page pdfPage = (Page)emptyDoc.Pages[1];
            // Create text fragment
            TextFragment textFragment = new TextFragment("The attachment " + name + "is not a supported attachment type.");
            textFragment.Position = new Position(100, 600);

            // Set text properties
            textFragment.TextState.FontSize = 12;
            textFragment.TextState.Font = FontRepository.FindFont("TimesNewRoman");
            textFragment.TextState.BackgroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.Color.LightGray);
            textFragment.TextState.ForegroundColor = Aspose.Pdf.Color.FromRgb(System.Drawing.Color.Red);

            // Create TextBuilder object
            TextBuilder textBuilder = new TextBuilder(pdfPage);

            // Append the text fragment to the PDF page
            textBuilder.AppendText(textFragment);
            return emptyDoc;
        }
        public MemoryStream ConvertDocToPdf(MemoryStream inputStream)
        {
            Aspose.Words.Saving.PdfSaveOptions saveOptions = new Aspose.Words.Saving.PdfSaveOptions();
            saveOptions.SaveFormat = Aspose.Words.SaveFormat.Pdf;
            Aspose.Words.Document doc = new Aspose.Words.Document(inputStream);
            MemoryStream stream = new MemoryStream();
            doc.Save(stream, saveOptions);
            return stream;
        }

        protected virtual MemoryStream EmbedImageToPdf(MemoryStream inputStream)
        {
            int lowerLeftX = 50;
            int lowerLeftY = 50;
            int upperRightX = 500;
            int upperRightY = 800;
            Document emptyDoc = new Document();
            emptyDoc.Pages.Insert(1);
            Page pdfPage = (Page)emptyDoc.Pages[1];
            pdfPage.Resources.Images.Add(inputStream);
            pdfPage.Contents.Add(new Operator.GSave());
            Rectangle rectangle = new Aspose.Pdf.Rectangle(lowerLeftX, lowerLeftY, upperRightX, upperRightY);
            Matrix matrix = new Matrix(new double[] { rectangle.URX - rectangle.LLX, 0, 0, rectangle.URY - rectangle.LLY, rectangle.LLX, rectangle.LLY });
            pdfPage.Contents.Add(new Operator.ConcatenateMatrix(matrix));
            XImage ximage = pdfPage.Resources.Images[pdfPage.Resources.Images.Count];
            pdfPage.Contents.Add(new Operator.Do(ximage.Name));
            pdfPage.Contents.Add(new Operator.GRestore());
            MemoryStream stream = new MemoryStream();
            emptyDoc.Save(stream);
            //PdfFileMend mendor = new PdfFileMend(emptyDoc);
            //mendor.AddImage(inputStream, 1, 50, 50, 300, 300);
            //mendor.Close();
            //BackgroundArtifact background = new BackgroundArtifact();
            //background.BackgroundImage = inputStream;
            //pdfPage.Artifacts.Add(background);
            return stream;
        }
        public MemoryStream ConvertImageToPdf(MemoryStream inputStream)
        {
            Aspose.Words.Saving.PdfSaveOptions saveOptions = new Aspose.Words.Saving.PdfSaveOptions();
            saveOptions.SaveFormat = Aspose.Words.SaveFormat.Pdf;
            Aspose.Words.Document doc = new Aspose.Words.Document();
            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
            // Read the image from file, ensure it is disposed.
            using (System.Drawing.Image image = System.Drawing.Image.FromStream(inputStream))
            {
                // Find which dimension the frames in this image represent. For example
                // The frames of a BMP or TIFF are "page dimension" whereas frames of a GIF image are "time dimension".
                FrameDimension dimension = new FrameDimension(image.FrameDimensionsList[0]);
                // Get the number of frames in the image.
                int framesCount = image.GetFrameCount(dimension);
                // Loop through all frames.
                for (int frameIdx = 0; frameIdx < framesCount; frameIdx++)
                {
                    // Insert a section break before each new page, in case of a multi-frame TIFF.
                    if (frameIdx != 0)
                        builder.InsertBreak(Aspose.Words.BreakType.SectionBreakNewPage);

                    // Select active frame.
                    image.SelectActiveFrame(dimension, frameIdx);

                    // We want the size of the page to be the same as the size of the image.
                    // Convert pixels to points to size the page to the actual image size.
                    Aspose.Words.PageSetup ps = builder.PageSetup;
                    ps.PageWidth = Aspose.Words.ConvertUtil.PixelToPoint(image.Width, image.HorizontalResolution);
                    ps.PageHeight = Aspose.Words.ConvertUtil.PixelToPoint(image.Height, image.VerticalResolution);

                    // Insert the image into the document and position it at the top left corner of the page.
                    builder.InsertImage(
                        image,
                        Aspose.Words.Drawing.RelativeHorizontalPosition.Page,
                        0,
                        Aspose.Words.Drawing.RelativeVerticalPosition.Page,
                        0,
                        ps.PageWidth,
                        ps.PageHeight,
                        Aspose.Words.Drawing.WrapType.None);
                }
            }

            MemoryStream stream = new MemoryStream();
            doc.Save(stream, saveOptions);
            return stream;
        }

        public string GenerateFolderName(string context)
        {
            return context + "_" + DateTime.Now.ToString("yyyyMMddHHmmssfff") + "_" + Guid.NewGuid().ToString("N");
        }

        private string CleanPath(string emlfilename)
        {
            string invalidChars = new string(Path.GetInvalidFileNameChars());// + new string(Path.GetInvalidPathChars());
            foreach (char c in invalidChars)
            {
                if (c.Equals(@"\") || c.Equals("."))
                {
                    continue;
                }
                emlfilename = emlfilename.Replace(c.ToString(), ""); // or with "."
            }
            return emlfilename;
        }
        public string GetTemporaryFolder()
        {
            var tempDirectory = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString());
            Directory.CreateDirectory(tempDirectory);
            return tempDirectory;
        }
    }
}
