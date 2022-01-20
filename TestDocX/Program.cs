using System;
using System.Drawing;
using System.IO;
using System.Linq;
using Telerik.Windows.Documents.Flow.FormatProviders.Docx;
using Telerik.Windows.Documents.Flow.FormatProviders.Pdf;
using Telerik.Windows.Documents.Flow.Model;
using Telerik.Windows.Documents.Flow.Model.Editing;
using Telerik.Windows.Documents.Spreadsheet.Model;
using Telerik.Windows.Documents.Flow.Model.Styles;



namespace TestDocX
{
    class Program
    {
        static string pathDesktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        static void Main(string[] args)
        {
            string folio = "SDAT-11-2022";
            string pathDocumento = pathDesktop + "/edocs/SDAT-11_2022.docx";
            string pathPdf = pathDesktop + "/edocs/SDAT-11_2022.pdf";
            var documento = Novacode.DocX.Load(pathDocumento);
            string firmaDigital = "EuFPlpB031UnloXm5A6BDSjFh+/YKBoYyv7MTn1zcwUawR1jyq1TDjQmHhytCnOyFejj9Yv0xgAa" +
                "Fe6AOw7xKmUrgDPfYGO3/lrjgsjTW9pGfBlMz7ZoxGaYilMBBCJYvyBy70JkV2lx8leR7KZAXFBq+66cETND2cJVOQwqOPd+dnnce" +
                "APaNSfvBCkELaFKQO56X2hQWPvWdDjyQZPrhTobcuE/GgkC3VVpOoj1EmPv4hRXvEN9B4yNSjQE4GBdphd6GG6Ptn9ldhpHfxIpvHaj" +
                "3lBmvMif37iKGF/2UuQG19dsRvS0Rui1Iz5jEGSP6HbaonOVCXIYmcFEq8w8gA==";

            string digestionArchivo = "X5/b+n89COpRm2PzMBK60ccxedg=";

            //InsertarFolioALaDerecha(documento, folio);
            //IncluirFirmaADocumento(documento, firmaDigital, digestionArchivo);

            InsertarFirmaTelerik(pathDocumento, firmaDigital, digestionArchivo);

            string pathClone = pathDesktop + "/edocs/clone.docx";

            ConvertirDocxToPdf(pathClone, pathPdf);
            Console.Read();
        }

        static void ConvertirDocxToPdf(string docxPath, string pathPdf)
        {
            Console.WriteLine("Convirtiendo a PDF ...");

            RadFlowDocument document = new DocxFormatProvider().Import(File.ReadAllBytes(docxPath));
            var documentClone = document.Clone();

            File.WriteAllBytes(pathPdf, new PdfFormatProvider().Export(documentClone));
            Console.WriteLine("PDF Terminado...");
        }

        static void InsertarFirmaTelerik(string docxPath, string firma, string digestion)
        {
            Console.WriteLine("Insertando firma con Telerik ...");
            //Se carga el archivo .docx
            DocxFormatProvider docxPRovider = new DocxFormatProvider();
            var docBytes = File.ReadAllBytes(docxPath);
            var document = docxPRovider.Import(docBytes);
            var documentClone = document.Clone();

            //Se posisiona el editor en el ultimo elemento del documento
            RadFlowDocumentEditor editor = new RadFlowDocumentEditor(documentClone);
            Run run = documentClone.EnumerateChildrenOfType<Run>().Last();
            editor.MoveToInlineEnd(run);

            editor.InsertBreak(BreakType.LineBreak);
            editor.InsertBreak(BreakType.LineBreak);
            editor.InsertBreak(BreakType.LineBreak);

            var table = editor.InsertTable();
            table.Borders = new TableBorders(new Border(1, BorderStyle.Single, ThemableColor.FromArgb(255, 0, 0, 0)));
            ThemableColor cellBackground = new ThemableColor(System.Windows.Media.Colors.Black);

            var firstRow = table.Rows.AddTableRow();
            var firstCell = firstRow.Cells.AddTableCell();
            firstCell.Shading.BackgroundColor = cellBackground;
            var secondRow = table.Rows.AddTableRow();
            var secondCell = secondRow.Cells.AddTableCell();

            //inserting text in the cell
            firstCell.Blocks.AddParagraph().Inlines.AddRun("Firma Electrónica");
            
            secondCell.Blocks.AddParagraph().Inlines.AddRun(firma);
            table.Alignment = Alignment.Center;

            string pathClone = pathDesktop + "/edocs/clone.docx";
            using (Stream output = new FileStream(pathClone, FileMode.OpenOrCreate))
            {
                DocxFormatProvider provider = new DocxFormatProvider();
                provider.Export(documentClone, output);
            }

            Console.WriteLine("Terminando de insertar tabla con Telerik ...");
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
            tabla.Alignment = Novacode.Alignment.center;

            //Encabezado Firma
            Novacode.Cell celdaEncabezado = tabla.Rows[0].Cells[0];
            celdaEncabezado.FillColor = System.Drawing.Color.Black;
            Novacode.Paragraph parrafoHeader = celdaEncabezado.Paragraphs[0];
            parrafoHeader.Alignment = Novacode.Alignment.center;
            parrafoHeader.Append("Firma Digital").FontSize(12).Font(new FontFamily(@"Arial"));
            tabla.Rows[0].Cells[0].Width = 100;

            //Firma
            Novacode.Cell celdaFirma = tabla.Rows[1].Cells[0];
            Novacode.Paragraph parrafoFirma = celdaFirma.Paragraphs[0];
            parrafoFirma.Alignment = Novacode.Alignment.center;
            parrafoFirma.Append(firma).FontSize(12).Font(new FontFamily(@"Arial"));
            tabla.Rows[1].Cells[0].Width = 100;

            //Se inserta tabla firma
            documento.InsertTable(tabla);

            //Tabla Digestión
            Novacode.Table tabla2 = documento.AddTable(2, 1);
            tabla2.Design = Novacode.TableDesign.TableGrid;

            //Encabezado Digestión
            Novacode.Cell celdaEncabezado2 = tabla2.Rows[0].Cells[0];
            celdaEncabezado2.FillColor = System.Drawing.Color.Black;
            Novacode.Paragraph parrafoHeader2 = celdaEncabezado2.Paragraphs[0];
            parrafoHeader2.Alignment = Novacode.Alignment.center;
            parrafoHeader2.Append("Digestión Archivo").FontSize(12).Font(new FontFamily(@"Arial"));
            tabla2.Rows[0].Cells[0].Width = 100;

            //Digestión
            Novacode.Cell celdaDigestion = tabla2.Rows[1].Cells[0];
            Novacode.Paragraph parrafoDigestion = celdaDigestion.Paragraphs[0];
            parrafoDigestion.Alignment = Novacode.Alignment.center;
            parrafoDigestion.Append(digestion).FontSize(12).Font(new FontFamily(@"Arial"));
            tabla2.Rows[1].Cells[0].Width = 100;

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
