using Microsoft.Office.Interop.Word;
using System;

namespace WordDocWork
{
    class Program
    {

        /// <summary>
        /// номера раздела, ==0 - нет разделов
        /// </summary>
        static uint _sectionNumber = 0;
        /// <summary>
        /// номера рисунков, ==0 - нет картинок
        /// </summary>
        static uint _pictureNumber = 0;
        /// <summary>
        /// номера таблиц, ==0 - нет таблиц
        /// </summary>
        static uint _tableNumber = 0;

        static void Main(string[] args)
        {

            string sourcePath = @"w:\stud_work\шаблон.rtf";//путь до исходного шаблона
            string distPath = @"w:\stud_work\result.rtf";//путь до выходного файла
            string csvPath = @"w:\stud_work\data.csv";//путь до csv файла для создания таблицы

            //список закладок
            string[] templateStringList =
                {
                "[*имя раздела*]",///0
                "[*имя рисунка*]",///1
                "[*ссылка на следующий рисунок*]",///2
                "[*ссылка на предыдущий рисунок*]",///3
                "[*ссылка на таблицу*]",///4
                "[*таблица первая*]"///5
                };
            var application = new Application();
            application.Visible = true;

            var document = application.Documents.Open(sourcePath);

            Paragraph prevParagraph = null;

            Object missing = System.Type.Missing;

            foreach (Paragraph paragraph in document.Paragraphs)
            {
                for (int i = 0; i < templateStringList.Length; i++)
                {
                    if (paragraph.Range.Text.Contains(templateStringList[i]))
                    {
                        switch (i)
                        {
                            case 0:
                                {
                                    paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                    paragraph.Range.Font.Name = "Times New Roman";
                                    paragraph.Range.Font.Size = 15;
                                    paragraph.Format.SpaceAfter = 12;
                                    paragraph.Range.Font.Bold = 1;
                                    paragraph.Range.HighlightColorIndex = 0;

                                    _sectionNumber++;
                                    string replaceString = _sectionNumber.ToString();

                                    paragraph.Range.Find.Execute(templateStringList[i],
                                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                        0, ref missing, replaceString, 2, ref missing, ref missing,
                                        ref missing, ref missing);
                                }
                                break;
                            case 1:
                                {
                                    paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                    paragraph.Range.Font.Name = "Times New Roman";
                                    paragraph.Range.Font.Size = 12;
                                    paragraph.Format.SpaceAfter = 12;
                                    paragraph.Range.HighlightColorIndex = 0;

                                    if (prevParagraph != null)
                                    {
                                        prevParagraph.Format.SpaceBefore = 12;
                                        prevParagraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                    }

                                    _pictureNumber++;
                                    string replaceString = "Рисунок " + _sectionNumber.ToString() + "." + _pictureNumber.ToString() + " -";

                                    paragraph.Range.Find.Execute(templateStringList[i],
                                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                        0, ref missing, replaceString, 2, ref missing, ref missing,
                                        ref missing, ref missing);
                                }
                                break;
                            case 2:
                                {
                                    paragraph.Range.HighlightColorIndex = 0;

                                    string replaceString = _sectionNumber.ToString() + "." + (_pictureNumber + 1).ToString();

                                    paragraph.Range.Find.Execute(templateStringList[i],
                                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                        0, ref missing, replaceString, 2, ref missing, ref missing,
                                        ref missing, ref missing);
                                }
                                break;
                            case 3:
                                {
                                    paragraph.Range.HighlightColorIndex = 0;

                                    string replaceString = _sectionNumber.ToString() + "." + _pictureNumber.ToString();

                                    paragraph.Range.Find.Execute(templateStringList[i],
                                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                        0, ref missing, replaceString, 2, ref missing, ref missing,
                                        ref missing, ref missing);
                                }
                                break;
                            case 4:
                                {

                                    paragraph.Range.HighlightColorIndex = 0;

                                    _tableNumber++;
                                    string replaceString = _sectionNumber.ToString() + "." + _tableNumber.ToString();

                                    paragraph.Range.Find.Execute(templateStringList[i],
                                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                                        0, ref missing, replaceString, 2, ref missing, ref missing,
                                        ref missing, ref missing);
                                }
                                break;
                            case 5:
                                {
                                    paragraph.Range.InsertParagraphBefore();
                                    paragraph.Range.InsertBefore("TABLA");
                                    paragraph.Alignment = WdParagraphAlignment.wdAlignParagraphCenter;
                                    paragraph.Range.Font.Name = "Times New Roman";
                                    paragraph.Range.Font.Size = 15;
                                    paragraph.Format.SpaceAfter = 12;
                                    paragraph.Range.HighlightColorIndex = 0;

                                    application.Selection.Find.Execute(templateStringList[i]);
                                    var range = application.Selection.Range;
                                    //range.HighlightColorIndex = 0;
                                                                        
                                    string[] listRows = System.IO.File.ReadAllText(csvPath).Split("\r\n".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                                    string[] listTitle = listRows[0].Split(";,".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);

                                    var wordTable = document.Tables.Add(range, listRows.Length, listTitle.Length);

                                    wordTable.Range.Font.Name = "Times New Roman";
                                    wordTable.Range.Font.Size = 11;
                                    wordTable.Range.Columns.DistributeWidth();
                                    wordTable.Range.Borders[WdBorderType.wdBorderLeft].LineStyle = WdLineStyle.wdLineStyleDouble;
                                    wordTable.Range.Borders[WdBorderType.wdBorderRight].LineStyle = WdLineStyle.wdLineStyleTriple;
                                    wordTable.Range.Borders[WdBorderType.wdBorderTop].LineStyle = WdLineStyle.wdLineStyleEmboss3D;
                                    wordTable.Range.Borders[WdBorderType.wdBorderBottom].LineStyle = WdLineStyle.wdLineStyleDoubleWavy;
                                    wordTable.Range.Borders.InsideLineStyle = WdLineStyle.wdLineStyleDashDot;
                                                                        
                                    for (var k = 0; k < listTitle.Length; k++)
                                    {
                                        wordTable.Cell(1, k + 1).Range.Text = listTitle[k].ToString();
                                    }

                                    for (var j = 1; j < listRows.Length; j++)
                                    {
                                        string[] listValues = listRows[j].Split(";,".ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
                                        for (var k = 0; k < listValues.Length; k++)
                                        {
                                            wordTable.Cell(j + 1, k + 1).Range.Text = listValues[k].ToString();
                                        }
                                    }                                                                        
                                }
                                break;
                        }
                    }
                }
                prevParagraph = paragraph;
            }

            document.SaveAs2(distPath);
            System.Console.In.Read();
            // application.Quit();
        }
    }
}
