
using System.Collections.Generic;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
//using System.Drawing;

namespace testFichierPdf
{
    class Program
    {
        private const float HAUTEUR_MAX_PAGE = 841;

        public static void AddImageOnEachPage(string cheminFichierFrom, string cheminFichierTo, string cheminImage)
        {
            // open the reader
            PdfReader reader = new PdfReader(cheminFichierFrom);
            Rectangle size = reader.GetPageSizeWithRotation(1);
            Document document = new Document(size);

            // open the writer
            FileStream fs = new FileStream(cheminFichierTo, FileMode.Create, FileAccess.Write);
            PdfWriter writer = PdfWriter.GetInstance(document, fs);
            document.Open();

            // the pdf content
            PdfContentByte cb = writer.DirectContent;

            if (!File.Exists(cheminImage))
                return;

            iTextSharp.text.Image image = iTextSharp.text.Image.GetInstance(cheminImage);
            image.SetAbsolutePosition(0 + 20 + 1, HAUTEUR_MAX_PAGE - 25 - 20 - image.ScaledHeight);

            for (int pageNumber = 1; pageNumber <= reader.NumberOfPages; pageNumber++)
            {
                // create the new page and add it to the pdf
                PdfImportedPage page = writer.GetImportedPage(reader, pageNumber);
                cb.AddTemplate(page, 0, 0);

                //ajout de l'image
                cb.AddImage(image);

                if (pageNumber < reader.NumberOfPages)
                    document.NewPage();
            }

            document.Close();
        }



        static void Main(string[] args)
        {
        }
    }
}
