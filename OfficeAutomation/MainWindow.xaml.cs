using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using Microsoft.Win32;
//using Microsoft.Office.Interop.Word;

namespace OfficeAutomation
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Build_Word(object sender, RoutedEventArgs e)
        {
            CreateDocument();
        }


        object missing = System.Reflection.Missing.Value;

        private Microsoft.Office.Interop.Word.Document CreateWordDocument(Microsoft.Office.Interop.Word.Application app)
        {
            return app.Documents.Add(ref missing, ref missing, ref missing, ref missing);
        }

        private void InsertHeaderFooter(Microsoft.Office.Interop.Word.Document doc)
        {
            //Add header into the document
            foreach (Microsoft.Office.Interop.Word.Section section in doc.Sections)
            {
                //Get the header range and add the header details.
                Microsoft.Office.Interop.Word.Range headerRange = section.Headers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                headerRange.Fields.Add(headerRange, Microsoft.Office.Interop.Word.WdFieldType.wdFieldPage);
                headerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;
                headerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdBlue;
                headerRange.Font.Size = 10;
                headerRange.Font.NameBi = "B Nazanin";
                headerRange.Text = "اتصال گیردار جوشی به کمک ورق های روسری و زیر سری";
            }
            //Add the footers into the document
            foreach (Microsoft.Office.Interop.Word.Section wordSection in doc.Sections)
            {
                //Get the footer range and add the footer details.
                Microsoft.Office.Interop.Word.Range footerRange = wordSection.Footers[Microsoft.Office.Interop.Word.WdHeaderFooterIndex.wdHeaderFooterPrimary].Range;
                footerRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdDarkRed;
                footerRange.Font.Size = 10;
                footerRange.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                footerRange.Text = "تمامی حقوق محفوظ است";
            }
        }

        private void InsertHeading(Microsoft.Office.Interop.Word.Document doc,string text)
        {
            Microsoft.Office.Interop.Word.Paragraph para = doc.Content.Paragraphs.Add(missing);
            para.Range.Text = text;
            para.Range.Font.SizeBi = 16;
            para.Range.Font.NameBi = "B Titr";
            para.Range.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorWhite;
            para.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;
            para.Range.Shading.BackgroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorLightBlue;
            para.Range.InsertParagraphAfter();
            InsertText(doc, "");
        }

        private void InsertText(Microsoft.Office.Interop.Word.Document doc, string text)
        {
            Microsoft.Office.Interop.Word.Paragraph para2 = doc.Content.Paragraphs.Add(missing);
            para2.Range.Text = text;
            para2.Range.Font.SizeBi = 14;
            para2.Range.Font.NameBi = "B Nazanin";
            para2.Range.Font.Color = Microsoft.Office.Interop.Word.WdColor.wdColorBlack;
            para2.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphRight;
            para2.Range.Shading.BackgroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorWhite;
            para2.Range.InsertParagraphAfter();
        }

        private void SaveWordDocument(Microsoft.Office.Interop.Word.Document doc)
        {
            //Save the document
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.AddExtension = true;
            sfd.DefaultExt = ".docx";
            sfd.Title = "Save Example Word File";
            if (sfd.ShowDialog() == true)
            {
                object filename = sfd.FileName;
                doc.SaveAs2(ref filename);
                doc.Close(ref missing, ref missing, ref missing);
                doc = null;
                MessageBox.Show("Document created successfully !");
            }
        }

        private void AddEquation(Microsoft.Office.Interop.Word.Document doc)
        {
            /*doc.OMaths.Add(missing);
            Microsoft.Office.Interop.Word.OMathFunction wdFunction = winword.Selection.OMaths[1].Functions.Add(winword.Selection.Range,
                   Microsoft.Office.Interop.Word.WdOMathFunctionType.wdOMathFunctionNary);
            Microsoft.Office.Interop.Word.OMathNary wdNary = wdFunction.Nary;
            wdNary.Char = 8721;
            wdNary.Grow = false;
            wdNary.SubSupLim = false;
            wdNary.HideSub = false;
            wdNary.HideSup = false;
            //Following code will setup value in Nary Function
            Microsoft.Office.Interop.Word.Selection wdSelection = winword.Selection;
            object unit = Microsoft.Office.Interop.Word.WdUnits.wdCharacter;
            object lu = Microsoft.Office.Interop.Word.WdUnits.wdLine;
            object count = 1;
            object dcount = 2;
            object tcount = 3;
            wdSelection.MoveLeft(ref unit, ref count);
            wdSelection.TypeText("11");
            wdSelection.MoveLeft(ref unit, ref tcount);
            wdSelection.TypeText("12");
            wdSelection.MoveDown(ref lu, ref count);
            wdSelection.TypeText("13");
            wdNary.Application.Visible = true;*/
        }

        public void AddTable(Microsoft.Office.Interop.Word.Document doc)
        {
            //Create a 5X5 table and insert some dummy record
            /*Microsoft.Office.Interop.Word.Table firstTable = document.Tables.Add(para1.Range, 5, 5, ref missing, ref missing);

            firstTable.Borders.Enable = 1;
            foreach (Microsoft.Office.Interop.Word.Row row in firstTable.Rows)
            {
                foreach (Microsoft.Office.Interop.Word.Cell cell in row.Cells)
                {
                    //Header row
                    if (cell.RowIndex == 1)
                    {
                        cell.Range.Text = "Column " + cell.ColumnIndex.ToString();
                        cell.Range.Font.Bold = 1;
                        //other format properties goes here
                        cell.Range.Font.Name = "verdana";
                        cell.Range.Font.Size = 10;
                        //cell.Range.Font.ColorIndex = WdColorIndex.wdGray25;                         

                        cell.Shading.BackgroundPatternColor = Microsoft.Office.Interop.Word.WdColor.wdColorGray25;
                        //Center alignment for the Header cells
                        cell.VerticalAlignment = Microsoft.Office.Interop.Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        cell.Range.ParagraphFormat.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    }
                    //Data row
                    else
                    {
                        cell.Range.Text = (cell.RowIndex - 2 + cell.ColumnIndex).ToString();
                    }
                }
            }*/
        }


        //Create document method
        private void CreateDocument()
        {
            try
            {
                //Open Word
                Microsoft.Office.Interop.Word.Application winword = new Microsoft.Office.Interop.Word.Application();
                winword.ShowAnimation = false;
                winword.Visible = false;

                Microsoft.Office.Interop.Word.Document document = CreateWordDocument(winword);

                InsertHeaderFooter(document);

                InsertHeading(document, "محاسبه مقاومت خمشی و برشی مورد نیاز");
                InsertText(document, "لورم ایپسوم متن ساختگی با تولید سادگی نامفهوم از صنعت چاپ و با استفاده از طراحان گرافیک است. چاپگرها و متون بلکه روزنامه و مجله در ستون و سطرآنچنان که لازم است و برای شرایط فعلی تکنولوژی مورد نیاز و کاربردهای متنوع با هدف بهبود ابزارهای کاربردی می باشد. کتابهای زیادی در شصت و سه درصد گذشته، حال و آینده شناخت فراوان جامعه و متخصصان را می طلبد تا با نرم افزارها شناخت بیشتری را برای طراحان رایانه ای علی الخصوص طراحان خلاقی و فرهنگ پیشرو در زبان فارسی ایجاد کرد. در این صورت می توان امید داشت که تمام و دشواری موجود در ارائه راهکارها و شرایط سخت تایپ به پایان رسد وزمان مورد نیاز شامل حروفچینی دستاوردهای اصلی و جوابگوی سوالات پیوسته اهل دنیای موجود طراحی اساسا مورد استفاده قرار گیرد.");
                InsertHeading(document, "نیرو ها در وجه ستون سمت راست");

                SaveWordDocument(document);

                //Close Word
                winword.Quit(ref missing, ref missing, ref missing);
                winword = null;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

    }
}
