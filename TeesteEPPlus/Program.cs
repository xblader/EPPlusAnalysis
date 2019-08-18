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

            var grupos = lista.GroupBy(x => x.Tipo).Select(x => new GrupoModel { Key = x.Key, Lista = x.ToList() }).ToList();

            using (ExcelEtapaPTHandler handler = new ExcelEtapaPTHandler(AppDomain.CurrentDomain.BaseDirectory + @"Excel\Template.xlsx"))
            {
                int linha = 3;

                handler.SelecionarSheet("Planilha1");

                foreach (var grupo in grupos)
                {
                    linha = handler.CopiarCabecalho(linha, grupo.Key);

                    handler.SetFormatacaoLinhas(linha, grupo.Lista.Count);

                    handler.InserirValores(linha, grupo.Lista);

                    handler.InsereTotal(linha + grupo.Lista.Count, grupo.Lista.Count);

                    linha += grupo.Lista.Count + 3;
                }

                handler.RemoveColuna();

                handler.SalvaExcel();
            }
        }
    }
}
