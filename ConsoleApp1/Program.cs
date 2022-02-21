using SevenZip;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Mail;
using System.Net.Mime;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApp1
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
                const string to = "marioescalante3@gmail.com,eskantu@hotmail.com,ma.escalante@mavi.mx";
                const string from = "ma.escalante@mavi.mx";
                var smtpClient = new SmtpClient("mail.mavi.mx")
                {
                    Port = 26,
                    //Port = 587,
                    Credentials = new NetworkCredential(from, "43qw9xqqoj"),
                    EnableSsl = false,
                };
                string file = @"C:\Users\maescalante\Desktop\ReporteHistoricoSubNatural.csv";
                string zipfle = @"C:\Users\maescalante\Desktop\ReporteHistoricoSubNatural.7z";
                SevenZipCompressor.SetLibraryPath(@"C:\Users\maescalante\source\repos\SendEmail\SendEmail\dll\7za.dll");
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
                Console.WriteLine("Exception caught in CreateTestMessage2(): {0}",
                    ex.ToString());
            }
        }
    }
}
