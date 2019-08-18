﻿using OfficeOpenXml;
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
    public class ExcelEtapaPTHandler : IDisposable
    {
        private ExcelPackage arquivoExcel;
        private ExcelWorksheet sheetselecionado;
        private int linha = 11;
        private int numerocolunas = 7;
        public ExcelEtapaPTHandler(string path)
        {
            FileInfo caminhoNomeArquivo = new FileInfo(path);

            arquivoExcel = new ExcelPackage(caminhoNomeArquivo);
        }

        public void SelecionarSheet(string sheetname)
        {
            this.sheetselecionado = arquivoExcel.Workbook.Worksheets[sheetname];
        }

        public void Dispose()
        {
            arquivoExcel.Dispose();
        }

        internal void CopiarCabecalho()
        {
            sheetselecionado.Cells[1, 1, 4, numerocolunas].Copy(sheetselecionado.Cells[7, 1, 10, numerocolunas]);
        }

        internal void InsereTotal(int linha, int count)
        {
            sheetselecionado.Cells[linha, 1, linha, numerocolunas].Merge = true;
            sheetselecionado.Cells[linha, 1].Value = $"Total de Registros {count}";
        }

        internal void SalvaExcel()
        {
            arquivoExcel.Save();
        }

        internal void InserirValores(int linha, IList<Etapa> lista)
        {
            foreach (var item in lista)
            {
                sheetselecionado.Cells[linha, 1].Value = item.Inicio.ToShortDateString();
                sheetselecionado.Cells[linha, 2].Value = item.Fim.ToShortDateString();
                sheetselecionado.Cells[linha, 3].Value = item.Localizacao;
                sheetselecionado.Cells[linha, 4].Value = item.Tipo;
                sheetselecionado.Cells[linha, 5].Value = item.Descricao;
                sheetselecionado.Cells[linha, 6].Value = item.HoraInicio;
                sheetselecionado.Cells[linha, 7].Value = item.HoraFim;
                linha++;
            }
        }

        internal void SetFormatacaoLinhas(int linha, int totallinhas)
        {
            var modelTable = sheetselecionado.Cells[linha, 1, linha + totallinhas, numerocolunas];
            modelTable.Style.Border.Top.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            modelTable.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            modelTable.Style.WrapText = true;
            var cellFont = modelTable.Style.Font;
            cellFont.SetFromFont(new Font("Arial", 8));
        }
    }
}
