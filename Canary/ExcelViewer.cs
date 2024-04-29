using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Canary
{
    public partial class ExcelViewer : Form
    {
        public ExcelViewer(string path, int endColumn)
        {
            InitializeComponent();

            reoGridControl.Load(path);

            var sheet = reoGridControl.CurrentWorksheet;

            sheet.CreateColumnFilter("A", ToAlphabet(endColumn));
        }

        private string ToAlphabet(int index)
        {
            string alphabet = string.Empty;
            if (index < 1) return alphabet;

            while (index > 0)
            {
                // A-Zの変換を0-25にするため1を引く
                index--;
                // ASCIIではAは10進数で65
                alphabet = Convert.ToChar(index % 26 + 65) + alphabet;
                index = index / 26;
            }

            return alphabet;
        }
    }
}
