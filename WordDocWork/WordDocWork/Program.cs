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
