using Microsoft.Office.Interop.Word;

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

            string sourcePath = @"f:\stud_work\шаблон.rtf";//путь до исходного шаблона
            string distPath = @"f:\stud_work\result.rtf";//путь до выходного файла
            string csvPath = @"f:\stud_work\data.csv";//путь до csv файла для создания таблицы

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

        }
    }
}
