using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelObj = Microsoft.Office.Interop.Excel;
using System.Reflection;
using Newtonsoft.Json.Linq;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;

namespace PARSEEXCEL
{
    public partial class mainForm : Form
    {
        /// <summary>
        /// Массив времени пар
        /// </summary>
        static string[] timeOfLessons = new string[] {
            "8:30-10:00",
            "10:15-11:45",
            "12:00-13:30",
            "14:10-15:40",
            "15:55-17:25",
            "17:40-19:10"
        };

        /// <summary>
        /// Массив групп
        /// </summary>
        static JArray Groups = new JArray();

        /// <summary>
        /// Экземпляр приложения Excel
        /// </summary>
        static ExcelObj.Application app;

        /// <summary>
        /// Текущий лист в таблице
        /// </summary>
        static ExcelObj.Workbook workbook;

        /// <summary>
        /// Переменная хранитель ячеек со значениями в таблице
        /// </summary>
        static ExcelObj.Range ShtRange;

        /// <summary>
        /// Стартовый индекс для парса(индекс строки с которой начинаются предметы)
        /// </summary>
        static int startIndex = 0;

        public mainForm()
        {
            InitializeComponent();
        }

        private void openTable()
        {
            OpenFileDialog _dialog = new OpenFileDialog();

            if (_dialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    workbook = app.Workbooks.Open(_dialog.FileName, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                    Missing.Value);
                }
                catch (Exception message)
                {
                    throw new Exception("Не удалось открыть таблицу\n" + message.Message);
                }
            }
            else
            {
                throw new Exception("Не удалось открыть таблицу\nЗакрыто окно выбора файла!");
            }
        }

        private void GetGroupsList()
        {
            try
            {
                // первый проход по таблице и выявление индексов столбцов, по которым распологаются предметы у
                // конкретной группы
                for (int WSnum = 1; WSnum <= workbook.Sheets.Count; WSnum++)
                {
                    ShtRange = ((ExcelObj.Worksheet)workbook.Sheets.get_Item(WSnum)).UsedRange;
                    bool isStoped = false;
                    for (int Rnum = 3; Rnum <= ShtRange.Rows.Count; Rnum++)
                    {
                        if (!isStoped)
                        {
                            for (int Cnum = 1; Cnum <= ShtRange.Columns.Count; Cnum++)
                            {
                                if ((ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2 != null)
                                {
                                    // по регулярке вытаскиваем здесь группы и добавляем в json 
                                    if (Regex.Matches((ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2,
                                        @"[А-я]+\-[0-9]+").Count != 0)
                                    {
                                        Data.AddToGroupList(
                                            (ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2, Cnum, WSnum);
                                        Groups.Add(getTemplateForGroup(
                                            (ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2));
                                        startIndex = Rnum;
                                    }
                                    if ((ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range)
                                        .Value2.ToLower().Contains("понедельник"))
                                    {
                                        isStoped = true;
                                        break;
                                    }
                                }
                            }
                        }
                        else
                        {
                            break;
                        }
                    }
                }
                //Data.SaveGroupList();
            }
            catch (Exception message)
            {
                throw new Exception("Не удалось получить список групп\n" + message.Message);
            }
        }

        private void GetShapesList()
        {
            try
            {
                for (int WSnum = 1; WSnum <= workbook.Sheets.Count; WSnum++)
                {
                    Data.ShapesRanges.Add(new List<Data.Shapes>());

                    foreach (ExcelObj.Shape _sh in ((ExcelObj.Worksheet)workbook.Sheets.get_Item(WSnum)).Shapes)
                    {
                        if (_sh.Height > 1 && _sh.Width > 1)
                        {
                            Data.ShapesRanges[WSnum - 1].Add(new Data.Shapes(_sh.TextFrame.Characters(0, 100).Text,
                                _sh.TopLeftCell.Column,
                                _sh.BottomRightCell.Column,
                                _sh.TopLeftCell.Row,
                                _sh.BottomRightCell.Row));
                        }
                    }
                }
                //Data.SaveShapesList();
            }
            catch (Exception message)
            {
                throw new Exception("Не удалось получить список фигур\n" + message.Message);
            }
        }

        private void GetScheduleForGroups()
        {
            for (int WSnum = 1; WSnum <= workbook.Sheets.Count; WSnum++)
            {

                ShtRange = ((ExcelObj.Worksheet)workbook.Sheets.get_Item(WSnum)).UsedRange;
                int curentDayOfWeek = 1;
                int currentNumberOfLesson = 1;

                // проход по строкам
                for (int Rnum = startIndex + 1; Rnum <= ShtRange.Rows.Count; Rnum++)
                {
                    if ((ShtRange.Cells[Rnum, 1] as ExcelObj.Range).Value2 != null)
                    {
                        // заносим в переменную - хранитель дня недели - текущий день
                        if (Regex.Matches((ShtRange.Cells[Rnum, 1] as ExcelObj.Range).Value2, @"[А-я]+").Count != 0)
                        {
                            string value = (ShtRange.Cells[Rnum, 1] as ExcelObj.Range).Value2;
                            switch (value.ToLower())
                            {
                                case "понедельник":
                                    curentDayOfWeek = 1;
                                    break;
                                case "вторник":
                                    curentDayOfWeek = 2;
                                    break;
                                case "среда":
                                    curentDayOfWeek = 3;
                                    break;
                                case "четверг":
                                    curentDayOfWeek = 4;
                                    break;
                                case "пятница":
                                    curentDayOfWeek = 5;
                                    break;
                                case "суббота":
                                    curentDayOfWeek = 6;
                                    break;
                            }
                            if (value.ToLower().Contains("Д/недели".ToLower()))
                            {
                                continue;
                            }
                        }
                    }

                    // проход по столбцам
                    for (int Cnum = 2; Cnum <= ShtRange.Columns.Count; Cnum++)
                    {
                        mainTextBox.Text = (ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Address;

                        if ((ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2 != null)
                        {
                            var cellValue = ((ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2).ToString();

                            // заносим в переменную - хранитель номера пары - текущую пару
                            foreach (Match m in Regex.Matches(cellValue, @"[0-9]"))
                            {
                                if (Regex.Matches(cellValue, @"[0-9]").Count == 1)
                                {
                                    if (int.Parse(m.ToString()) != 0 && int.Parse(m.ToString()) < 7)
                                    {
                                        currentNumberOfLesson = int.Parse(m.ToString());
                                    }
                                }
                                else
                                {
                                    break;
                                }
                            }

                            // проверка на соответствие текущего столбца - столбцу группы
                            for (int k = 0; k < Data.GroupsIndexes.Count; k++)
                            {
                                if (Cnum == Data.GroupsIndexes[k].Index
                                && Data.GroupsIndexes[k].NumberOfSheet == WSnum)
                                {
                                    // создание нового экземпляра объекта с полями предмета
                                    JObject lesson = new JObject();
                                    lesson.Add("time", timeOfLessons[currentNumberOfLesson - 1]);
                                    lesson.Add("name", (ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2);
                                    lesson.Add("teacher", null);
                                    lesson.Add("audience", null);

                                    int currentCnt = 0;
                                    for (int q = 0; q < Groups.Count; q++)
                                    {
                                        if (Groups[q]["name"].ToString() == Data.GroupsIndexes[k].Name)
                                        {
                                            currentCnt = q;
                                            break;
                                        }
                                    }

                                    #region чекаем все картинки в таблице

                                    // если ячейка объедененная
                                    if ((ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).MergeCells)
                                    {
                                        bool isContinue = false;
                                        for (int cnt = 0; cnt < Data.ShapesRanges[WSnum - 1].Count; cnt++)
                                        {
                                            isContinue = false;

                                            int startRowPoint = (ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).MergeArea.Row;
                                            int mergeRowCount = (ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).MergeArea.Rows.Count;

                                            int startColumnPoint = (ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).MergeArea.Column;
                                            int mergeColumnsCount = (ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).MergeArea.Columns.Count;

                                            for (int i = Data.ShapesRanges[WSnum - 1][cnt].TopBorder; i <= Data.ShapesRanges[WSnum - 1][cnt].BottomBorder && !isContinue; i++)
                                            {
                                                for (int j = Data.ShapesRanges[WSnum - 1][cnt].LeftBorder; j <= Data.ShapesRanges[WSnum - 1][cnt].RightBorder && !isContinue; j++)
                                                {
                                                    for (int _c = startColumnPoint; _c < startColumnPoint + mergeColumnsCount && !isContinue; _c++)
                                                    {
                                                        if ((ShtRange.Cells[startRowPoint, _c] as ExcelObj.Range).Column == j)
                                                        {
                                                            if ((ShtRange.Cells[startRowPoint, _c] as ExcelObj.Range).Row == i)
                                                            {
                                                                bool isAudienceHaveChars = false;
                                                                // исключение
                                                                if (Regex.Matches(Data.ShapesRanges[WSnum - 1][cnt].Text, @"к/кл").Count > 0)
                                                                {
                                                                    lesson["audience"] = Data.ShapesRanges[WSnum - 1][cnt].Text;
                                                                    isAudienceHaveChars = true;

                                                                }
                                                                // если в картинке цифры - заносим в кабинет
                                                                if (Regex.Matches(Data.ShapesRanges[WSnum - 1][cnt].Text, @"[0-9]+").Count > 0)
                                                                {

                                                                    lesson["audience"] = Data.ShapesRanges[WSnum - 1][cnt].Text;

                                                                }
                                                                // если в картинке текст - значит препод
                                                                if (Regex.Matches(Data.ShapesRanges[WSnum - 1][cnt].Text, @"[А-я]+").Count > 0 && !isAudienceHaveChars)
                                                                {
                                                                    if (Regex.Matches(Data.ShapesRanges[WSnum - 1][cnt].Text, @"[0-9]+").Count == 0)
                                                                    {
                                                                        lesson["teacher"] = Data.ShapesRanges[WSnum - 1][cnt].Text;
                                                                    }
                                                                }
                                                                isContinue = true;
                                                            }
                                                            if ((ShtRange.Cells[startRowPoint + 1, _c] as ExcelObj.Range).Row == i)
                                                            {
                                                                bool isAudienceHaveChars = false;
                                                                // исключение
                                                                if (Regex.Matches(Data.ShapesRanges[WSnum - 1][cnt].Text, @"к/кл").Count > 0)
                                                                {
                                                                    if (lesson["audience"].ToString() == "")
                                                                    {
                                                                        lesson["audience"] = Data.ShapesRanges[WSnum - 1][cnt].Text;
                                                                        isAudienceHaveChars = true;
                                                                    }
                                                                }
                                                                // если в картинке цифры - заносим в кабинет
                                                                if (Regex.Matches(Data.ShapesRanges[WSnum - 1][cnt].Text, @"[0-9]+").Count > 0)
                                                                {
                                                                    if (lesson["audience"].ToString() == "")
                                                                    {
                                                                        lesson["audience"] = Data.ShapesRanges[WSnum - 1][cnt].Text;
                                                                    }
                                                                }
                                                                // если в картинке текст - значит препод
                                                                if (Regex.Matches(Data.ShapesRanges[WSnum - 1][cnt].Text, @"[А-я]+").Count > 0 && !isAudienceHaveChars)
                                                                {
                                                                    if (Regex.Matches(Data.ShapesRanges[WSnum - 1][cnt].Text, @"[0-9]+").Count == 0)
                                                                    {
                                                                        if (lesson["teacher"].ToString() == "")
                                                                        {
                                                                            lesson["teacher"] = Data.ShapesRanges[WSnum - 1][cnt].Text;
                                                                        }
                                                                    }
                                                                }
                                                                isContinue = true;
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    // если ячейка не объединенная
                                    else
                                    {
                                        bool isContinue = false;
                                        for (int cnt = 0; cnt < Data.ShapesRanges[WSnum - 1].Count; cnt++)
                                        {
                                            isContinue = false;

                                            if (Cnum >= Data.ShapesRanges[WSnum - 1][cnt].LeftBorder && Cnum <= Data.ShapesRanges[WSnum - 1][cnt].RightBorder && !isContinue)
                                            {
                                                if ((Rnum >= Data.ShapesRanges[WSnum - 1][cnt].TopBorder || Rnum + 1 >= Data.ShapesRanges[WSnum - 1][cnt].TopBorder) && Rnum <= Data.ShapesRanges[WSnum - 1][cnt].BottomBorder && !isContinue)
                                                {
                                                    for (int i = Data.ShapesRanges[WSnum - 1][cnt].TopBorder; i <= Data.ShapesRanges[WSnum - 1][cnt].BottomBorder && !isContinue; i++)
                                                    {
                                                        for (int j = Data.ShapesRanges[WSnum - 1][cnt].LeftBorder; j <= Data.ShapesRanges[WSnum - 1][cnt].RightBorder && !isContinue; j++)
                                                        {
                                                            if ((ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Column == j)
                                                            {
                                                                if ((ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).Row == i)
                                                                {
                                                                    bool isAudienceHaveChars = false;
                                                                    // исключение
                                                                    if (Regex.Matches(Data.ShapesRanges[WSnum - 1][cnt].Text, @"к/кл").Count > 0)
                                                                    {
                                                                        lesson["audience"] = Data.ShapesRanges[WSnum - 1][cnt].Text;
                                                                        isAudienceHaveChars = true;

                                                                    }
                                                                    // если в картинке цифры - заносим в кабинет
                                                                    if (Regex.Matches(Data.ShapesRanges[WSnum - 1][cnt].Text, @"[0-9]+").Count > 0)
                                                                    {
                                                                        lesson["audience"] = Data.ShapesRanges[WSnum - 1][cnt].Text;

                                                                    }
                                                                    // если в картинке текст - значит препод
                                                                    if (Regex.Matches(Data.ShapesRanges[WSnum - 1][cnt].Text, @"[А-я]+").Count > 0 && !isAudienceHaveChars)
                                                                    {
                                                                        if (Regex.Matches(Data.ShapesRanges[WSnum - 1][cnt].Text, @"[0-9]+").Count == 0)
                                                                        {

                                                                            lesson["teacher"] = Data.ShapesRanges[WSnum - 1][cnt].Text;
                                                                        }
                                                                    }
                                                                    isContinue = true;
                                                                }
                                                            }
                                                            if ((ShtRange.Cells[Rnum + 1, Cnum] as ExcelObj.Range).Row == i && !isContinue)
                                                            {
                                                                bool isAudienceHaveChars = false;
                                                                // исключение
                                                                if (Regex.Matches(Data.ShapesRanges[WSnum - 1][cnt].Text, @"к/кл").Count > 0)
                                                                {
                                                                    if (lesson["audience"].ToString() == "")
                                                                    {
                                                                        lesson["audience"] = Data.ShapesRanges[WSnum - 1][cnt].Text;
                                                                        isAudienceHaveChars = true;
                                                                    }
                                                                }
                                                                // если в картинке цифры - заносим в кабинет
                                                                if (Regex.Matches(Data.ShapesRanges[WSnum - 1][cnt].Text, @"[0-9]+").Count > 0)
                                                                {
                                                                    if (lesson["audience"].ToString() == "")
                                                                    {
                                                                        lesson["audience"] = Data.ShapesRanges[WSnum - 1][cnt].Text;
                                                                    }
                                                                }
                                                                // если в картинке текст - значит препод
                                                                if (Regex.Matches(Data.ShapesRanges[WSnum - 1][cnt].Text, @"[А-я]+").Count > 0 && !isAudienceHaveChars)
                                                                {
                                                                    if (Regex.Matches(Data.ShapesRanges[WSnum - 1][cnt].Text, @"[0-9]+").Count == 0)
                                                                    {
                                                                        if (lesson["teacher"].ToString() == "")
                                                                        {
                                                                            lesson["teacher"] = Data.ShapesRanges[WSnum - 1][cnt].Text;
                                                                        }
                                                                    }
                                                                }
                                                                isContinue = true;
                                                            }

                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    #endregion

                                        (Groups[currentCnt]["scheduleOnWeek"]
                                        [curentDayOfWeek - 1][currentNumberOfLesson - 1] as JArray).Add(lesson);

                                    #region если ячейка объединена с другими - добавим во все объедененные, значения этой ячейки
                                    if ((ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).MergeCells)
                                    {
                                        int startColumnPoint = (ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).MergeArea.Column;
                                        int mergeColumnsCount = (ShtRange.Cells[Rnum, Cnum] as ExcelObj.Range).MergeArea.Columns.Count;

                                        for (int j = startColumnPoint + 1; j < startColumnPoint + mergeColumnsCount; j++) // изменение 17.04.2015 22:12
                                        {
                                            // записываем в группу которая в диапозоне объединения - это значение
                                            for (int k2 = 0; k2 < Data.GroupsIndexes.Count; k2++)
                                            {
                                                if (j == Data.GroupsIndexes[k2].Index
                                                && Data.GroupsIndexes[k2].NumberOfSheet == WSnum)
                                                {
                                                    for (int q = 0; q < Groups.Count; q++)
                                                    {
                                                        // записываем предмет, если его нет; записываем второй предмет, 
                                                        // если уже есть первый и они разные
                                                        if (Groups[q]["name"].ToString() == Data.GroupsIndexes[k2].Name)
                                                        {
                                                            if ((Groups[q]["scheduleOnWeek"]
                                                                [curentDayOfWeek - 1][currentNumberOfLesson - 1] as JArray).Count == 0)
                                                            {
                                                                (Groups[q]["scheduleOnWeek"]
                                                                    [curentDayOfWeek - 1][currentNumberOfLesson - 1] as JArray).Add(lesson);
                                                            }
                                                            else
                                                            {
                                                                // идем по массиву предметов - если там есть такая же ячейка - не записываем новую
                                                                bool isHave = false;
                                                                for (int asd = 0; asd < (Groups[q]["scheduleOnWeek"]
                                                                [curentDayOfWeek - 1][currentNumberOfLesson - 1] as JArray).Count; asd++)
                                                                {
                                                                    if ((Groups[q]["scheduleOnWeek"]
                                                                   [curentDayOfWeek - 1][currentNumberOfLesson - 1] as JArray)[asd] == lesson)
                                                                    {
                                                                        isHave = true;
                                                                        break;
                                                                    }
                                                                }
                                                                if (!isHave)
                                                                {
                                                                    (Groups[q]["scheduleOnWeek"]
                                                                    [curentDayOfWeek - 1][currentNumberOfLesson - 1] as JArray).Add(lesson);
                                                                }
                                                            }
                                                            break;
                                                        }
                                                    }
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                    #endregion

                                    break;
                                }
                            }

                        }
                    }
                }
            }
            app.Quit();

            JObject data = new JObject();
            data.Add("data", Groups);
            try
            {
                File.WriteAllText(@"D:\schedule.json", data.ToString());
            }
            catch (Exception a)
            {
                mainTextBox.Text += "\n" + a.Message;
            }
            
        }

        JObject getTemplateForGroup(string nameOfGroup)
        {
            JObject temp = new JObject();
            JArray weekArray = new JArray();
            for (int i = 0; i < 6; i++)
            {
                JArray lessonsArray = new JArray();
                for (int j = 0; j < 6; j++)
                {
                    JArray lessonsOnDay = new JArray(); // каждый день - массив
                    lessonsArray.Add(lessonsOnDay);
                }
                weekArray.Add(lessonsArray);
            }
            temp.Add("name", nameOfGroup);
            temp.Add("scheduleOnWeek", weekArray);
            return temp;
        }

        private void GetScheduleFromTable()
        {
            app = new ExcelObj.Application();

            startButton.Enabled = false;
            mainComboBox.Enabled = false;
            try
            {
                openTable();
                mainTextBox.Text += "Таблица открыта\n";
                GetGroupsList();
                mainTextBox.Text += "Список групп составлен и сохранен\n";
                GetShapesList();
                mainTextBox.Text += "Список картинок составлен и сохранен\n";
                string start = "Время начала парса" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + "\n";
                GetScheduleForGroups();
                string finish = "Время оканчания парса" + DateTime.Now.Hour + ":" + DateTime.Now.Minute + "\n";
                mainTextBox.Text += "РАСПИСАНИЕ СОСТАВЛЕНО!\n";
                mainTextBox.Text += start + finish;
            }
            catch (Exception message)
            {
                mainTextBox.Text += message.Message + "\n";
            }
            finally
            {
                mainComboBox.Enabled = true;
                app.Quit();
            }
        }

        private void GetScheduleFromFile()
        {
            try
            {
                Groups = JArray.Parse(JObject.Parse(File.ReadAllText(@"D:\schedule.json"))["data"].ToString());
                mainProgressBar.Maximum = Groups.Count;
                for (int i = 0; i < Groups.Count; i++)
                {
                    mainComboBox.Items.Add(Groups[i]["name"]);
                    mainProgressBar.Value++;
                }
                startButton.Enabled = false;
            }
            catch (Exception message)
            {
                mainTextBox.Text = message.Message;
            }
        }

        private void mainComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            mainTextBox.Text = null;

            mainProgressBar.Maximum = 6 * 6;

            for (int i = 0; i < Groups.Count; i++)
            {
                if (mainComboBox.SelectedItem == Groups[i]["name"])
                {
                    JArray groupWeek = JArray.Parse(Groups[i]["scheduleOnWeek"].ToString());
                    for (int j = 0; j < groupWeek.Count; j++)
                    {
                        switch (j)
                        {
                            case 0:
                                mainTextBox.Text += "Понедельник:\n";
                                break;
                            case 1:
                                mainTextBox.Text += "Вторник:\n";
                                break;
                            case 2:
                                mainTextBox.Text += "Среда:\n";
                                break;
                            case 3:
                                mainTextBox.Text += "Четверг:\n";
                                break;
                            case 4:
                                mainTextBox.Text += "Пятница:\n";
                                break;
                            case 5:
                                mainTextBox.Text += "Суббота:\n";
                                break;
                        }
                        mainTextBox.Text += "\n";

                        JArray lessonsInDay = JArray.Parse(groupWeek[j].ToString());
                        for (int k = 0; k < lessonsInDay.Count; k++)
                        {
                            if (lessonsInDay[k].HasValues)
                            {
                                JArray lessonsInOneDay = JArray.Parse(lessonsInDay[k].ToString());

                                mainTextBox.Text += lessonsInOneDay[0]["time"].ToString() + "\n\n";

                                for (int n = 0; n < lessonsInOneDay.Count; n++)
                                {
                                    if (lessonsInOneDay.Count == 2)
                                    {
                                        if (n == 0)
                                        {
                                            mainTextBox.Text += "По четным неделям:\n";
                                        }
                                        else
                                        {
                                            mainTextBox.Text += "По нечетным:\n";
                                        }
                                    }
                                    mainTextBox.Text += lessonsInOneDay[n]["name"].ToString() + "\n";
                                    if (lessonsInOneDay[n]["teacher"] != null)
                                    {
                                        mainTextBox.Text += "Преподователь: " + lessonsInOneDay[n]["teacher"].ToString() + "\n";
                                    }
                                    if (lessonsInOneDay[n]["audience"] != null)
                                    {
                                        mainTextBox.Text += "Аудитория: " + lessonsInOneDay[n]["audience"].ToString() + "\n";
                                    }
                                    mainTextBox.Text += "\n";
                                }
                            }
                            try
                            {
                                mainProgressBar.Value++;
                            }
                            catch (Exception)
                            {

                            }
                        }
                    }
                    break;
                }
            }
        }

        private void startButton_Click(object sender, EventArgs e)
        {
            GetScheduleFromFile();
        }
    }
}