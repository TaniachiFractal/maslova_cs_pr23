using System;
using System.IO;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace maslova_cs_pr23
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        string[] names = { "Alex", "Bonnie", "Clara", "Drake", "Egor" };
        int[] numbers = { 144, 3232, 335, 4224, 5453 };

        private void btWord_Click(object sender, EventArgs e)
        {
            var WordApp = new Word.Application();
            Word.Document document = WordApp.Documents.Add();

            Word.Paragraph userParagraph = document.Paragraphs.Add();
            Word.Range userRange = userParagraph.Range;
            userRange.Text = "Справочник телефонов";
            userRange.InsertParagraphAfter();

            Word.Paragraph tableParagraph = document.Paragraphs.Add();
            Word.Range tableRange = tableParagraph.Range;
            Word.Table numbersTable = document.Tables.Add(tableRange, names.Length, 2);
            numbersTable.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleDoubleWavy;
            numbersTable.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDashDotDot;

            Word.Range cellRange;
            cellRange = numbersTable.Cell(1, 1).Range;
            cellRange.Text = "ФИО";
            cellRange = numbersTable.Cell(1, 2).Range;
            cellRange.Text = "Телефон";

            numbersTable.Rows[1].Range.Bold = 1;
            numbersTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphThaiJustify;

            for (int i = 0; i < names.Length; i++)
            {
                cellRange = numbersTable.Cell(i + 2, 1).Range;
                cellRange.Text = names[i];
                cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                cellRange = numbersTable.Cell(i+2,2).Range;
                cellRange.Text = numbers[i].ToString();
                cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphRight;
            }

            WordApp.Visible = true;
            document.SaveAs2(Directory.GetCurrentDirectory() + @"\test.docx");
        }
    }
}
