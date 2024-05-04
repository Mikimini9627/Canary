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
    /// <summary>
    /// ビューア画面クラス
    /// </summary>
    public partial class ExcelViewer : Form
    {
        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="path"></param>
        /// <param name="endColumn"></param>
        public ExcelViewer(string path, int endColumn)
        {
            InitializeComponent();

            // ファイルを読み込む
            reoGridControl.Load(path);

            // シートを取得する
            var sheet = reoGridControl.CurrentWorksheet;

            // フィルターをかける
            sheet.CreateColumnFilter("A", ToAlphabet(endColumn));
        }

        /// <summary>
        /// 数値から列名へ変換する
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
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
