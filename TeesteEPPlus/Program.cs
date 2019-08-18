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
            var lista = Builder<Etapa>
                .CreateListOfSize(10)
                .All()
                .TheFirst(5)
                   .With(x => x.Tipo = "Trabalho Eletrico")
                .TheLast(5)
                   .With(x => x.Tipo = "Trabalho a Frio")
                .Build();

            using (ExcelEtapaPTHandler handler = new ExcelEtapaPTHandler(AppDomain.CurrentDomain.BaseDirectory + @"Excel\Template.xlsx"))
            {
                int linha = 10;

                handler.SelecionarSheet("Planilha1");

                handler.CopiarCabecalho();

                handler.SetFormatacaoLinhas(linha, lista.Count);

                handler.InserirValores(linha, lista);

                handler.InsereTotal(linha + lista.Count, lista.Count);

                handler.SalvaExcel();
            }
        }
    }
}
