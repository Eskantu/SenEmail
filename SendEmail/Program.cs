using System;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using SevenZip;
namespace SendEmail
{
    class Program
    {
        static void Main(string[] args)
        {
            CreateTestMessage2("smtp.gmail.com", 587);
            Console.ReadLine();
        }

        public static void CreateTestMessage2(string host, int port)
        {
            try
            {
                const string to = "";
                const string from = "";
                var smtpClient = new SmtpClient("")
                {
                    Port = 0,
                    Credentials = new NetworkCredential(from, "your_password"),
                    EnableSsl = false,
                };
                string file = @"";
                string zipfle = @"";
                SevenZipCompressor.SetLibraryPath(@"C:\Users\maescalante\source\repos\SendEmail\SendEmail\dll\7z.dll");
                SevenZipCompressor sevenZipCompressor = new SevenZipCompressor();
                sevenZipCompressor.CompressionLevel = CompressionLevel.Ultra;
                sevenZipCompressor.CompressionMethod = CompressionMethod.Lzma;
                sevenZipCompressor.CompressFiles(zipfle, file);

                Attachment attachment = new Attachment(zipfle, MediaTypeNames.Application.Octet);
                ContentDisposition disposition = attachment.ContentDisposition;
                disposition.CreationDate = System.IO.File.GetCreationTime(zipfle);
                disposition.ModificationDate = System.IO.File.GetLastWriteTime(zipfle);
                disposition.ReadDate = System.IO.File.GetLastAccessTime(zipfle);
                var mailMessage = new MailMessage
                {
                    From = new MailAddress(from),
                    Subject = $"Email de prueba {DateTime.Now}",
                    Body = "Reporte",

                };
                mailMessage.Attachments.Add(attachment);
                mailMessage.To.Add(to);

                smtpClient.Send(mailMessage);
                Console.WriteLine("mensaje enviado correctamente");

            }
            catch (Exception ex)
            {
                Console.WriteLine($"Exception caught in CreateTestMessage2(): {0}",
                    ex.ToString());
            }
        }
    }
}
