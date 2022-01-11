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

            documento.DifferentFirstPage = false;
            documento.DifferentOddAndEvenPages = false;

            //Formatting textFormat = new Formatting();
            //textFormat.Bold = true;
            //textFormat.Size = 12;
            //textFormat.FontFamily = new Font(@"Arial");

            //var p = documento.InsertParagraph(folio, false, textFormat);
            //p.Alignment = Alignment.right;

            var p = documento.Paragraphs.First().InsertParagraphBeforeSelf(documento.InsertParagraph());
            p.Append(folio).Bold().FontSize(12).Font(new FontFamily(@"Arial"));
            p.Alignment = Alignment.right;

            //var header = documento.Headers.odd;

            //var p = header.Paragraphs.Last() ?? header.InsertParagraph();

            //p.Append(folio).Bold().FontSize(12).Font(new FontFamily(@"Arial"));

            //p.Alignment = Alignment.right;

            //header.InsertParagraph();

            documento.Save();
        }
    }
}
