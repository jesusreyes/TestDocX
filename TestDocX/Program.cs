using System;
using System.Drawing;
using System.Linq;
//using Xceed.Document.NET;
//using Xceed.Words.NET;
using Novacode;

namespace TestDocX
{
    class Program
    {
        static void Main(string[] args)
        {
            string folio = "SDAT-11-2022";
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string pathDocumento = path + "/edocs/SDAT-11_2022.docx";
            var documento = DocX.Load(pathDocumento);

            InsertarFolioALaDerecha(documento, folio);
        }

        static void InsertarFolioALaDerecha(DocX documento, string folio)
        {
            Console.WriteLine("Insertando folio en la esquina superior derecha.....");
            documento.DifferentFirstPage = false;
            documento.DifferentOddAndEvenPages = false;

            var p = documento.Paragraphs.First().InsertParagraphBeforeSelf(documento.InsertParagraph());
            p.Append(folio).Bold().FontSize(12).Font(new FontFamily(@"Arial"));
            p.Alignment = Alignment.right;

            documento.Save();

            Console.WriteLine("Inserción del folio completa.");
            Console.ReadLine();
        }
    }
}
