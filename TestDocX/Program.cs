using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
//using Novacode;
using Telerik.Windows.Documents.Extensibility;
using Telerik.Windows.Documents.Flow.Model;
using Telerik.Windows.Documents.Flow.Model.Styles;

namespace TestDocX
{
    class Program
    {
        static void Main(string[] args)
        {
            string folio = "SDAT-11-2022";
            string path = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string pathDocumento = path + "/edocs/SDAT-11_2022.docx";
            string pathPdf = path + "/edocs/SDAT-11_2022.pdf";
            var documento = Novacode.DocX.Load(pathDocumento);
            string firmaDigital = "EuFPlpB031UnloXm5A6BDSjFh+/YKBoYyv7MTn1zcwUawR1jyq1TDjQmHhytCnOyFejj9Yv0xgAa" +
                "Fe6AOw7xKmUrgDPfYGO3/lrjgsjTW9pGfBlMz7ZoxGaYilMBBCJYvyBy70JkV2lx8leR7KZAXFBq+66cETND2cJVOQwqOPd+dnnce" +
                "APaNSfvBCkELaFKQO56X2hQWPvWdDjyQZPrhTobcuE/GgkC3VVpOoj1EmPv4hRXvEN9B4yNSjQE4GBdphd6GG6Ptn9ldhpHfxIpvHaj" +
                "3lBmvMif37iKGF/2UuQG19dsRvS0Rui1Iz5jEGSP6HbaonOVCXIYmcFEq8w8gA==";
            string digestionArchivo = "X5/b+n89COpRm2PzMBK60ccxedg=";

            //InsertarFolioALaDerecha(documento, folio);
            IncluirFirmaADocumento(documento, firmaDigital, digestionArchivo);

            ConvertirDocxToPdf(pathDocumento, pathPdf);


            Console.Read();
        }

        static void ConvertirDocxToPdf(string docxPath, string pdfPath)
        {
            Console.WriteLine("Convirtiendo documento a PDF.");

            Telerik.Windows.Documents.Extensibility.JpegImageConverterBase jpegImageConverter = new Telerik.Documents.ImageUtils.JpegImageConverter();
            Telerik.Windows.Documents.Extensibility.FixedExtensibilityManager.JpegImageConverter = jpegImageConverter;

            var docxPRovider = new Telerik.Windows.Documents.Flow.FormatProviders.Docx.DocxFormatProvider();
            var pdfProvider = new Telerik.Windows.Documents.Flow.FormatProviders.Pdf.PdfFormatProvider();

            var docBytes = File.ReadAllBytes(docxPath);
            var document = docxPRovider.Import(docBytes);

            IEnumerable<Table> tables = document.EnumerateChildrenOfType<Table>();

            foreach (Table table in tables)
            {
                if (table.PreferredWidth.Type == TableWidthUnitType.Fixed)
                {
                    table.PreferredWidth = new TableWidthUnit(TableWidthUnitType.Auto, table.PreferredWidth.Value);
                }
            }


            var resultBytes = pdfProvider.Export(document);
            File.WriteAllBytes(pdfPath, resultBytes);

            Console.WriteLine("Documento PDF terminado.");
        }

        static void IncluirFirmaADocumento(Novacode.DocX documento, string firma, string digestion)
        {
            Console.WriteLine("Incluyendo Firma al documento ...");

            InsertarTablaFirma(documento, firma, digestion);
        }

        static void InsertarTablaFirma(Novacode.DocX documento, string firma, string digestion)
        {
            Console.WriteLine("Insertando tabla del Documento..");
            //Tabla Firma
            Novacode.Table tabla = documento.AddTable(2, 1);
            tabla.Design = Novacode.TableDesign.TableGrid;

            //Encabezado Firma
            Novacode.Cell celdaEncabezado = tabla.Rows[0].Cells[0];
            celdaEncabezado.FillColor = Color.Black;
            Novacode.Paragraph parrafoHeader = celdaEncabezado.Paragraphs[0];
            parrafoHeader.Alignment = Novacode.Alignment.center;
            parrafoHeader.Append("Firma Digital").FontSize(12).Font(new FontFamily(@"Arial"));

            //Firma
            Novacode.Cell celdaFirma = tabla.Rows[1].Cells[0];
            Novacode.Paragraph parrafoFirma = celdaFirma.Paragraphs[0];
            parrafoFirma.Alignment = Novacode.Alignment.center;
            parrafoFirma.Append(firma).FontSize(12).Font(new FontFamily(@"Arial"));

            //Se inserta tabla firma
            documento.InsertTable(tabla);

            //Tabla Digestión
            Novacode.Table tabla2 = documento.AddTable(2, 1);
            tabla2.Design = Novacode.TableDesign.TableGrid;

            //Encabezado Digestión
            Novacode.Cell celdaEncabezado2 = tabla2.Rows[0].Cells[0];
            celdaEncabezado2.FillColor = Color.Black;
            celdaEncabezado.Width = 2000;
            Novacode.Paragraph parrafoHeader2 = celdaEncabezado2.Paragraphs[0];
            parrafoHeader2.Alignment = Novacode.Alignment.center;
            parrafoHeader2.Append("Digestión Archivo").FontSize(12).Font(new FontFamily(@"Arial"));

            //Digestión
            Novacode.Cell celdaDigestion = tabla2.Rows[1].Cells[0];
            Novacode.Paragraph parrafoDigestion = celdaDigestion.Paragraphs[0];
            parrafoDigestion.Alignment = Novacode.Alignment.center;
            parrafoDigestion.Append(digestion).FontSize(12).Font(new FontFamily(@"Arial"));

            //Se inserta la tabla digestión
            documento.InsertTable(tabla2);

            //Se guarda documento
            documento.Save();

            Console.WriteLine("Tabla Insertada.");
        }

        static void InsertarFolioALaDerecha(Novacode.DocX documento, string folio)
        {
            Console.WriteLine("Insertando folio en la esquina superior derecha.....");
            documento.DifferentFirstPage = false;
            documento.DifferentOddAndEvenPages = false;

            var p = documento.Paragraphs.First().InsertParagraphBeforeSelf(documento.InsertParagraph());
            p.Append(folio).Bold().FontSize(12).Font(new FontFamily(@"Arial"));
            p.Alignment = Novacode.Alignment.right;

            documento.Save();

            Console.WriteLine("Inserción del folio completa.");
            Console.ReadLine();
        }
    }
}
