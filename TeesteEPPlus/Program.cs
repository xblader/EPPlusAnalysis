using FizzWare.NBuilder;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TeesteEPPlus
{
    class Program
    {
        static void Main(string[] args)
        {
            // criando/abrindo o arquivo:
            FileInfo caminhoNomeArquivo = new FileInfo(AppDomain.CurrentDomain.BaseDirectory + @"Excel\Template.xlsx");

            ExcelPackage arquivoExcel = new ExcelPackage(caminhoNomeArquivo);

            ExcelWorksheet sheet = arquivoExcel.Workbook.Worksheets["Planilha1"];

            //sheet.Cells["A13"].Value = "teste";

            var lista = Builder<Etapa>.CreateListOfSize(10).Build();
            int linha = 11;
            int numerocolunas = 7;
            //copia cabecalho...
            sheet.Cells[1, 1, 4, numerocolunas].Copy(sheet.Cells[7, 1, 10, numerocolunas]);

            var modelTable = sheet.Cells[linha, 1, linha + lista.Count, numerocolunas];

            //// Assign borders
            modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            modelTable.Style.WrapText = true;
            var cellFont = modelTable.Style.Font;
            cellFont.SetFromFont(new Font("Arial", 8));
            //modelTable.AutoFitColumns();

            foreach (var item in lista)
            {
                sheet.Cells[linha, 1].Value = item.Inicio.ToShortDateString();
                sheet.Cells[linha, 2].Value = item.Fim.ToShortDateString(); 
                sheet.Cells[linha, 3].Value = item.Localizacao;
                sheet.Cells[linha, 4].Value = item.Tipo;
                sheet.Cells[linha, 5].Value = item.Descricao;
                sheet.Cells[linha, 6].Value = item.HoraInicio;
                sheet.Cells[linha, 7].Value = item.HoraFim;
                linha++;
            }

            sheet.Cells[linha, 1, linha, numerocolunas].Merge = true;
            sheet.Cells[linha, 1].Value = $"Total de Registros {lista.Count}";

            //sheet.DeleteColumn(2);

            // salvando e fechando o arquivo: MUITO IMPORTANTE HEIN!!!
            arquivoExcel.Save();
            arquivoExcel.Dispose();
        }
    }
}
