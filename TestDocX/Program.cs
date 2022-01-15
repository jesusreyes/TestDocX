using System;
using System.Drawing;
using System.Linq;
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
            string firmaDigital = "EuFPlpB031UnloXm5A6BDSjFh+/YKBoYyv7MTn1zcwUawR1jyq1TDjQmHhytCnOyFejj9Yv0xgAa" +
                "Fe6AOw7xKmUrgDPfYGO3/lrjgsjTW9pGfBlMz7ZoxGaYilMBBCJYvyBy70JkV2lx8leR7KZAXFBq+66cETND2cJVOQwqOPd+dnnce" +
                "APaNSfvBCkELaFKQO56X2hQWPvWdDjyQZPrhTobcuE/GgkC3VVpOoj1EmPv4hRXvEN9B4yNSjQE4GBdphd6GG6Ptn9ldhpHfxIpvHaj" +
                "3lBmvMif37iKGF/2UuQG19dsRvS0Rui1Iz5jEGSP6HbaonOVCXIYmcFEq8w8gA==";
            string digestionArchivo = "X5/b+n89COpRm2PzMBK60ccxedg=";

            //InsertarFolioALaDerecha(documento, folio);
            IncluirFirmaADocumento(documento, firmaDigital, digestionArchivo);

            
        }

        static void IncluirFirmaADocumento(DocX documento, string firma, string digestion)
        {
            Console.WriteLine("Incluyendo Firma al documento ...");

            InsertarTablaFirma(documento, firma, digestion);
            Console.Read();
        }

        static void InsertarTablaFirma(DocX documento, string firma, string digestion)
        {
            Console.WriteLine("Insertando tabla del Documento..");
            //Tabla Firma
            Table tabla = documento.AddTable(2, 1);
            tabla.Design = TableDesign.TableGrid;

            //Encabezado Firma
            Cell celdaEncabezado = tabla.Rows[0].Cells[0];
            celdaEncabezado.FillColor = Color.Black;
            Novacode.Paragraph parrafoHeader = celdaEncabezado.Paragraphs[0];
            parrafoHeader.Alignment = Alignment.center;
            parrafoHeader.Append("Firma Digital").FontSize(12).Font(new FontFamily(@"Arial"));

            //Firma
            Cell celdaFirma = tabla.Rows[1].Cells[0];
            Novacode.Paragraph parrafoFirma = celdaFirma.Paragraphs[0];
            parrafoFirma.Alignment = Alignment.center;
            parrafoFirma.Append(firma).FontSize(12).Font(new FontFamily(@"Arial"));

            //Se inserta tabla firma
            documento.InsertTable(tabla);
           
            //Tabla Digestión
            Table tabla2 = documento.AddTable(2, 1);
            tabla2.Design = TableDesign.TableGrid;

            //Encabezado Digestión
            Cell celdaEncabezado2 = tabla2.Rows[0].Cells[0];
            celdaEncabezado2.FillColor = Color.Black;
            celdaEncabezado.Width = 2000;
            Novacode.Paragraph parrafoHeader2 = celdaEncabezado2.Paragraphs[0];
            parrafoHeader2.Alignment = Alignment.center;
            parrafoHeader2.Append("Digestión Archivo").FontSize(12).Font(new FontFamily(@"Arial"));

            //Digestión
            Cell celdaDigestion = tabla2.Rows[1].Cells[0];
            Novacode.Paragraph parrafoDigestion = celdaDigestion.Paragraphs[0];
            parrafoDigestion.Alignment = Alignment.center;
            parrafoDigestion.Append(digestion).FontSize(12).Font(new FontFamily(@"Arial"));

            //Se inserta la tabla digestión
            documento.InsertTable(tabla2);

            //Se guarda documento
            documento.Save();

            Console.WriteLine("Tabla Insertada.");
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
