using Google.GData.Client;
using Google.GData.Spreadsheets;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;
namespace SendEmailGoogleApps2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.WindowState = FormWindowState.Maximized;
        }

        SpreadsheetsService service = new SpreadsheetsService("EmailSpreadsheet");
        string worksheetQuery = "https://spreadsheets.google.com/feeds/worksheets/15fuEC4H5sfe2KQO1hV28Ng0W48eCbmMBwxNIlyE3KW4/private/full";

        private void Form1_Load(object sender, EventArgs e)
        {
            
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            loadWorksheets();
        }

        private void loadWorksheets()
        {
            service.setUserCredentials("lawyerhelperapp@gmail.com", "9219190202");    

            WorksheetQuery query = new WorksheetQuery(worksheetQuery);

            listView1.Items.Clear();

            AtomEntryCollection entries = service.Query(query).Entries;
            for (int i = 0; i < entries.Count; i++)
            {
                AtomLink cellsLink = entries[i].Links.FindService(GDataSpreadsheetsNameTable.CellRel, AtomLink.ATOM_TYPE);
                AtomLink listLink = entries[i].Links.FindService(GDataSpreadsheetsNameTable.ListRel, AtomLink.ATOM_TYPE);
                ListViewItem item = new ListViewItem(new string[3]{ entries[i].Title.Text, cellsLink.HRef.Content, listLink.HRef.Content});
                listView1.Items.Add(item);
            }

            listView1.Items[0].Selected = true;
            listView1.Select();
        }

        private string[] aimArray()
        {
            string[] array = new string[114];
            array[0] = "Получение исполнительного листа";
            array[1] = "Возбуждение исполнительного производства";
            array[2] = "Установление должнику срока для добровольного исполнения требований, содержащихся в исполнительном документе";
            array[3] = "Вызов должника";
            array[4] = "Получение объяснений должника";
            array[5] = "Выход на адрес должника";
            array[6] = "Совершение запроса в инспекции Федеральной налоговой службы";
            array[7] = "Получение ответа на запрос в инспекции Федеральной налоговой службы";
            array[8] = "Совершение запросов в кредитные организации";
            array[9] = "Сбербанк";
            array[10] = "ВТБ";
            array[11] = "Банк «Санкт-Петербург";
            array[12] = "другой банк (1)";
            array[13] = "другой банк (2)";
            array[14] = "Получение ответов на запросы из кредитных организаций";
            array[15] = "Сбербанк";
            array[16] = "ВТБ";
            array[17] = "Банк «Санкт-Петербург";
            array[18] = "другой банк (1)";
            array[19] = "другой банк (2)";
            array[20] = "Наложение ареста на счета должника в кредитных организациях";
            array[21] = "Сбербанк";
            array[22] = "ВТБ";
            array[23] = "Банк «Санкт-Петербург";
            array[24] = "другой банк (1)";
            array[25] = "другой банк (2)";
            array[26] = "Совершение запроса в Пенсионный фонд Российской Федерации";
            array[27] = "Получение ответов на запросы из Пенсионного фонда Российской Федерации";
            array[28] = "Совершение запроса в Управление Федеральной службы госрегистрации";
            array[29] = "Получение ответа на запрос из Управления Федеральной службы госрегистрации";
            array[30] = "Совершение запросов в иные, чем указаны в пункте 14, управления Федеральной службы госрегистрации";
            array[31] = "Получение ответов на запросы в иные, чем указаны в пункте 14, управления Федеральной службы госрегистрации";
            array[32] = "Совершение запросов в Проектно-инвентаризационные бюро";
            array[33] = "Получение ответов на запросы из Проектно-инвентаризационных бюро";
            array[34] = "Совершение запроса в Государственную инспецию по безопасности дорожного движения";
            array[35] = "Получение ответа на запрос из Государственной инспеции по безопасности дорожного движения";
            array[36] = "Совершение запроса в Государственную инспецию по маломерным судам";
            array[37] = "Получение ответа на запрос из Государственной инспеции по маломерным судам";
            array[38] = "Совершение запроса в Министерство внутренних дел";
            array[39] = "Получение ответа на запрос из Министерства внутренних дел";
            array[40] = "Совершение запросов в туристические фирмы";
            array[41] = "Получение ответов на запросы в туристические фирмы";
            array[42] = "Совершение запросов операторам сотовой связи";
            array[43] = "Получение ответов на запросы операторам сотовой связи";
            array[44] = "Совершение запросов в единые рассчетно-кассовые центры";
            array[45] = "Получение ответов на запросы в единые рассчетно-кассовые центры";
            array[46] = "Совершение запроса о маршрутах передвижения должника (постановка в систему «Розыск-Магистраль»)";
            array[47] = "Получение ответов на запрос о маршрутах передвижения должника из системы «Розыск-Магистраль»";
            array[48] = "Постановка должника в систему «Поток-Д»";
            array[49] = "Контроль процесса постановки должника в систему «Поток-Д»";
            array[50] = "Вынесение постановления о розыске имущества должника";
            array[51] = "Контроль процесса розыска имущества должника";
            array[52] = "Наложение ареста на имущество должника";
            array[53] = "Конфискация имущества должника";
            array[54] = "Совершение запроса в отдел Записи актов гражданского состояния о наличии у должника брака";
            array[55] = "Вызов супруги должника";
            array[56] = "Получение объяснений супруги должника";
            array[57] = "Выход в адрес супруги должника";
            array[58] = "Совершение запроса в инспекции Федеральной налоговой службы на супругу должника";
            array[59] = "Получение ответа на запрос в инспекции Федеральной налоговой службы на супругу должника";
            array[60] = "Совершение запросов в кредитные организации на супругу должника";
            array[61] = "Сбербанк";
            array[62] = "ВТБ";
            array[63] = "Банк «Санкт-Петербург";
            array[64] = "другой банк (1)";
            array[65] = "другой банк (2)";
            array[66] = "Получение ответов на запросы из кредитных организаций на супругу должника";
            array[67] = "Сбербанк";
            array[68] = "ВТБ";
            array[69] = "Банк «Санкт-Петербург";
            array[70] = "другой банк (1)";
            array[71] = "другой банк (2)";
            array[72] = "Наложение ареста на счета супруги должника в кредитных организациях в части совместной собственности супругов";
            array[73] = "Сбербанк";
            array[74] = "ВТБ";
            array[75] = "Банк «Санкт-Петербург";
            array[76] = "другой банк (1)";
            array[77] = "другой банк (2)";
            array[78] = "Совершение запроса в Пенсионный фонд Российской Федерации на супругу должника";
            array[79] = "Получение ответов на запросы из Пенсионного фонда Российской Федерации на супругу должника";
            array[80] = "Совершение запроса в Управление Федеральной службы госрегистрации на супругу должника";
            array[81] = "Получение ответа на запрос из Управления Федеральной службы госрегистрации на супругу должника";
            array[82] = "Совершение запросов в иные, чем указаны в пункте 51, управления Федеральной службы госрегистрации на супругу должника";
            array[83] = "Получение ответов на запросы в иные, чем указаны в пункте 51, управления Федеральной службы госрегистрации";
            array[84] = "Совершение запросов в Проектно-инвентаризационные бюро на супругу должника";
            array[85] = "Получение ответов на запросы из Проектно-инвентаризационных бюро на супругу должника";
            array[86] = "Совершение запроса в Государственную инспецию по безопасности дорожного движения на супругу должника";
            array[87] = "Получение ответа на запрос из Государственной инспеции по безопасности дорожного движения на супругу должника";
            array[88] = "Совершение запроса в Государственную инспецию по маломерным судам на супругу должника";
            array[89] = "Получение ответа на запрос из Государственной инспеции по маломерным судам на супругу должника";
            array[90] = "Совершение запроса в Министерство внутренних дел на супругу должника";
            array[91] = "Получение ответа на запрос из Министерства внутренних дел на супругу должника";
            array[92] = "Совершение запросов в туристические фирмы на супругу должника";
            array[93] = "Получение ответов на запросы в туристические фирмы на супругу должника";
            array[94] = "Совершение запросов операторам сотовой связи на супругу должника";
            array[95] = "Получение ответов на запросы операторам сотовой связи на супругу должника";
            array[96] = "Совершение запросов в единые рассчетно-кассовые центры на супругу должника";
            array[97] = "Получение ответов на запросы в единые рассчетно-кассовые центры на супругу должника";
            array[98] = "Совершение запроса о маршрутах передвижения супруги должника посредством постановки информации в систему «Розыск-Магистраль»";
            array[99] = "Получение ответов на запрос о маршрутах передвижения супруги должника из системы «Розыск-Магистраль»";
            array[100] = "Постановка супруги должника в систему «Поток-Д»";
            array[101] = "Контроль процесса постановки супруги должника в систему «Поток-Д»";
            array[102] = "Вынесение постановления о розыске имущества супруги должника";
            array[103] = "Контроль процесса розыска имущества супруги должника";
            array[104] = "Наложение ареста на имущество супруги должника";
            array[105] = "Конфискация имущества супруги должника";
            array[106] = "Совершение запрета на выезд должника за пределы Российской Федерации";
            array[107] = "Совершение требования об исполнении судебного акта";
            array[108] = "Повторное совершение требования об исполнении судебного акта";
            array[109] = "Предупреждение должника о привлечении к административной ответственности";
            array[110] = "Привлечение должника к административной ответственности";
            array[111] = "Предупреждение должника о привлечении к уголовной ответственности";
            array[112] = "Повторное предупреждение должника о привлечении к уголовной ответственности";
            array[113] = "Возбуждение уголовного дела в отношении должника";
            return array;
        }

        private double[] indexArray()
        {
            double[] array = new double[114];
            for (int i = 0; i < 9; ++i)
            {
                array[i] = i + 1;
            }
            array[9] = 9.1;
            array[10] = 9.2;
            array[11] = 9.3;
            array[12] = 9.4;
            array[13] = 9.5;
            array[14] = 10;
            array[15] = 10.1;
            array[16] = 10.2;
            array[17] = 10.3;
            array[18] = 10.4;
            array[19] = 10.5;
            array[20] = 11;
            array[21] = 11.1;
            array[22] = 11.2;
            array[23] = 11.3;
            array[24] = 11.4;
            array[25] = 11.5;
            for (int i = 26; i < 61; ++i)
            {
                array[i] = i - 14;
            }
            array[61] = 46.1;
            array[62] = 46.2;
            array[63] = 46.3;
            array[64] = 46.4;
            array[65] = 46.5;
            array[66] = 47;
            array[67] = 47.1;
            array[68] = 47.2;
            array[69] = 47.3;
            array[70] = 47.4;
            array[71] = 47.5;
            array[72] = 48;
            array[73] = 48.1;
            array[74] = 48.2;
            array[75] = 48.3;
            array[76] = 48.4;
            array[77] = 48.5;
            for (int i = 78; i < 114; ++i)
            {
                array[i] = i - 29;
            }
            return array;
        }

        private void loadSells(string query)
        {
            DataTable datatable = new DataTable();
            CellQuery cellQuery = new CellQuery(query);
            cellQuery.MinimumRow = 1;
            cellQuery.MinimumColumn = 1;
            cellQuery.MaximumColumn = 7;

            DataTable dt = new DataTable();
            dt.Columns.Add("Наименование");
            dt.Columns.Add("Значение");
            dt.Rows.Add("Исполнительное производство №");
            dt.Rows.Add("Дата начала исполнительного производства::");
            dt.Rows.Add("Служба судебных приставов:");
            dt.Rows.Add("Судебный пристав-исполнитель:");
            dt.Rows.Add("Дознаватель:");

                CellFeed cellFeed = service.Query(cellQuery);
                datatable.Columns.Add("№");
                datatable.Columns.Add("Наименование действия");
                datatable.Columns.Add("Дата(План)");
                datatable.Columns.Add("Срок с начала");
                datatable.Columns.Add("Дата(Факт)");
                datatable.Columns.Add("Просрочка");
                datatable.Columns.Add("Примечания");
                
                for (int i = 0; i < 114; i++)
                {
                    datatable.Rows.Add(indexArray()[i], aimArray()[i], "", "", "", "", "");
                }

                DataRow workRow;
                for (int i = 0; i <= Convert.ToInt32(cellFeed.RowCount.Value); i++)
                {
                    workRow = datatable.NewRow();
                    workRow[0] = "";
                    workRow[1] = "";
                    workRow[2] = "";
                    workRow[3] = "";
                    workRow[4] = "";
                    workRow[5] = "";
                    workRow[6] = "";
                    datatable.Rows.Add(workRow);
                }

                foreach (CellEntry cell in cellFeed.Entries)
                {
                    if (Convert.ToInt32(cell.Row) >= 7)
                        datatable.Rows[Convert.ToInt32(cell.Row) - 7][Convert.ToInt32(cell.Column) - 1] = cell.Value;
                    else
                        dt.Rows[Convert.ToInt32(cell.Row) - 1][Convert.ToInt32(cell.Column) - 1] = cell.Value;
                }

                dataGridView2.DataSource = dt;
                dataGridView1.DataSource = datatable;
                dataGridView1.Columns[0].Width = 30;
                dataGridView1.Columns[1].Width = 700;
                dataGridView1.Columns[2].Width = 70;
                dataGridView1.Columns[3].Width = 55;
                dataGridView1.Columns[4].Width = 70;
                dataGridView1.Columns[5].Width = 65;
                dataGridView1.Columns[6].Width = 300;


                if (cellFeed.Entries.Count == 0)
                {
                    makeNewSheet();
                }
        }

        private void addWorksheet(string name)
        {
            WorksheetEntry worksheet = new WorksheetEntry();
            worksheet.Title.Text = name;
            worksheet.Cols = 7;
            worksheet.Rows = 250;
            WorksheetQuery query = new WorksheetQuery(worksheetQuery);
            service.Insert(service.Query(query), worksheet);
        }

        private void makeNewSheet()
        {
            CellEntry temp = new CellEntry();
            label2.Text = "Пожалуйста, подождите";
            CellEntry.CellElement cellE = new CellEntry.CellElement();
            cellE.Row = 1;
            cellE.Column = 1;
            cellE.InputValue = "Исполнительное производство";
            temp.Cell = cellE;
            service.Insert(new Uri(listView1.SelectedItems[0].SubItems[1].Text), temp);

            cellE.Row = 2;
            cellE.Column = 1;
            cellE.InputValue = "Дата начала исполнительного производства::";
            temp.Cell = cellE;
            service.Insert(new Uri(listView1.SelectedItems[0].SubItems[1].Text), temp);

            cellE.Row = 3;
            cellE.Column = 1;
            cellE.InputValue = "Служба судебных приставов:";
            temp.Cell = cellE;
            service.Insert(new Uri(listView1.SelectedItems[0].SubItems[1].Text), temp);

            cellE.Row = 4;
            cellE.Column = 1;
            cellE.InputValue = "Судебный пристав-исполнитель:";
            temp.Cell = cellE;
            service.Insert(new Uri(listView1.SelectedItems[0].SubItems[1].Text), temp);

            cellE.Row = 5;
            cellE.Column = 1;
            cellE.InputValue = "Дознаватель:";
            temp.Cell = cellE;
            service.Insert(new Uri(listView1.SelectedItems[0].SubItems[1].Text), temp);

            for (int i = 0; i < 114; ++i)
            {
                cellE.Row = Convert.ToUInt32(i + 7);
                cellE.Column = Convert.ToUInt32(1);
                cellE.InputValue = indexArray()[i].ToString();
                temp.Cell = cellE;
                service.Insert(new Uri(listView1.SelectedItems[0].SubItems[1].Text), temp);

                CellEntry temp1 = new CellEntry();
                cellE.Row = Convert.ToUInt32(i + 7);
                cellE.Column = Convert.ToUInt32(2);
                cellE.InputValue = aimArray()[i].ToString();
                temp1.Cell = cellE;
                service.Insert(new Uri(listView1.SelectedItems[0].SubItems[1].Text), temp1);
            }
            label2.Text = "";
        }

        private void delWorksheet(string name)
        {
            WorksheetQuery query = new WorksheetQuery(worksheetQuery);
            WorksheetFeed wsFeed = service.Query(query);
            foreach (WorksheetEntry wsheet in wsFeed.Entries)
            {
                if (wsheet.Title.Text == name)
                {
                    wsheet.Delete();
                }
            }
        }

        private void updateCell(int row, int col, string value, string query)
        {
            CellEntry entry = new CellEntry();
            CellEntry.CellElement cellE = new CellEntry.CellElement();
            cellE.InputValue = value;
            cellE.Row = Convert.ToUInt32(row);
            cellE.Column = Convert.ToUInt32(col);
            entry.Cell = cellE;
            service.Insert(new Uri(query), entry);
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                loadSells(listView1.SelectedItems[0].SubItems[1].Text);                
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            loadWorksheets();
            CellEntry.CellElement cellE = new CellEntry.CellElement();
            //WorksheetQuery query = new WorksheetQuery(worksheetQuery);
            CellQuery cellQuery = new CellQuery(worksheetQuery);
            CellFeed cellFeed = service.Query(cellQuery);
                try
                {
                    for (int i = 0; i < 114; i++)
                    {
                        CellEntry temp = new CellEntry();

                            foreach (CellEntry cell in cellFeed.Entries)
                            {
                                if ((cell.Column == 1) && (cell.Row == i + 1))
                                {
                                    temp = cell;
                                    break;
                                }
                            }
                        if (temp.Value == null)
                        {
                            cellE.Row = Convert.ToUInt32(i + 1);
                            cellE.Column = Convert.ToUInt32(1);
                            cellE.InputValue = indexArray()[i].ToString();
                            temp.Cell = cellE;
                            service.Insert(new Uri(worksheetQuery), temp);
                        }
                        else
                            break;
                    }
                    for (int i = 0; i < 114; i++)
                    {
                        CellEntry temp = new CellEntry();
                        foreach (CellEntry cell in cellFeed.Entries)
                        {
                            if ((cell.Column == 2) && (cell.Row == i + 1))
                            {
                                temp = cell;
                                break;
                            }
                        }
                        if (temp.Value == null)
                        {
                            cellE.Row = Convert.ToUInt32(i + 1);
                            cellE.Column = Convert.ToUInt32(2);
                            cellE.InputValue = aimArray()[i].ToString();
                            temp.Cell = cellE;
                            service.Insert(new Uri(listView1.SelectedItems[0].SubItems[1].Text), temp);
                        }
                        else
                            break;
                    }
                }
                catch (System.NullReferenceException)
                {
                    for (int i = 0; i < 114; i++)
                    {
                        CellEntry temp = new CellEntry();
                        cellE.Row = Convert.ToUInt32(i + 1);
                        cellE.Column = Convert.ToUInt32(1);
                        cellE.InputValue = indexArray()[i].ToString();
                        temp.Cell = cellE;
                        service.Insert(new Uri(listView1.SelectedItems[0].SubItems[1].Text), temp);

                        CellEntry temp1 = new CellEntry();
                        cellE.Row = Convert.ToUInt32(i + 1);
                        cellE.Column = Convert.ToUInt32(2);
                        cellE.InputValue = aimArray()[i].ToString();
                        temp1.Cell = cellE;
                        service.Insert(new Uri(listView1.SelectedItems[0].SubItems[1].Text), temp1);
                    }
                }
            MessageBox.Show("Синхронизация закончена");
        }        

        private void button2_Click(object sender, EventArgs e)
        {
            string str;
            if (InputBox.InputBox.Input("Добавить", "Название:", out str))
            {
                addWorksheet(str);
                loadWorksheets();
            }
        }        

        private void button3_Click(object sender, EventArgs e)
        {
            if (listView1.SelectedItems.Count > 0)
            {
                string name = listView1.SelectedItems[0].SubItems[0].Text;
                DialogResult dialogResult = MessageBox.Show("Удалить документ '" + name + "'", "Удаление", MessageBoxButtons.YesNo);
                if (dialogResult == DialogResult.Yes)
                {
                    delWorksheet(name);
                    loadWorksheets();
                }                
            }
        }

        private void dataGridView1_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {            
            int row = dataGridView1.CurrentCell.RowIndex + 7;
            int column = dataGridView1.CurrentCell.ColumnIndex + 1;
            string value = dataGridView1.CurrentCell.Value.ToString();
            updateCell(row, column, value, listView1.SelectedItems[0].SubItems[1].Text);
            if ((column == 5) && (value != ""))
            {
                CellQuery cellQuery = new CellQuery(listView1.SelectedItems[0].SubItems[1].Text);
                CellFeed cellFeed = service.Query(cellQuery);
                CellEntry temp = new CellEntry();
                int row1 = row;
                int column1 = 6;
                foreach (CellEntry cell in cellFeed.Entries)
                {
                    if ((cell.Column == 3) && (cell.Row == row))
                    {
                        temp = cell;
                        break;
                    }
                }
                DateTime value1 = DateTime.Parse(dataGridView1.CurrentCell.Value.ToString());
                DateTime temp1 = DateTime.Parse(temp.Value.ToString());
                TimeSpan nDay = value1 - temp1;
                int days = nDay.Days;
                string value2 = days.ToString();
                dataGridView1.Rows[row1 - 7].Cells[column1 - 1].Value = value2;
                updateCell(row1, column1, value2, listView1.SelectedItems[0].SubItems[1].Text);
            }

            if ((column == 4) && (value != ""))
            {
                CellQuery cellQuery = new CellQuery(listView1.SelectedItems[0].SubItems[1].Text);
                CellFeed cellFeed = service.Query(cellQuery);
                CellEntry temp = new CellEntry();
                foreach (CellEntry cell in cellFeed.Entries)
                {
                    if ((cell.Column == 2) && (cell.Row == 2))
                    {
                        temp = cell;
                        break;
                    }
                }
                TimeSpan ts = TimeSpan.FromDays(double.Parse(value));
                DateTime tempDt = DateTime.Parse(temp.Value);
                dataGridView1.Rows[row - 7].Cells[2].Value = (tempDt + ts).Date;
                updateCell(row, 3, (tempDt + ts).Date.ToString(), listView1.SelectedItems[0].SubItems[1].Text);
            }
        }

        private void dataGridView2_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            int row = dataGridView2.CurrentCell.RowIndex + 1;
            int column = dataGridView2.CurrentCell.ColumnIndex + 1;
            string value = dataGridView2.CurrentCell.Value.ToString();
            CellEntry entry = new CellEntry();
            CellEntry.CellElement cellE = new CellEntry.CellElement();
            cellE.InputValue = value;
            cellE.Row = Convert.ToUInt32(row);
            cellE.Column = Convert.ToUInt32(column);
            entry.Cell = cellE;
            service.Insert(new Uri(listView1.SelectedItems[0].SubItems[1].Text), entry);

        }
        //это типо изменять лист: добавлять столбцы, строки, менять имя. мб решим, что нахер не надо. он его не написал полностью. 
        //надо разбирать
        //private void updateWorksheet(string name)
        //{
        //    WorksheetQuery query = new WorksheetQuery(worksheetQuery);
        //    WorksheetFeed wsFeed = service.Query(query);
        //    foreach (WorksheetEntry wsheet in wsFeed.Entries)
        //    {
        //        if (wsheet.Title.Text == name)
        //        {
        //            worksheet.Title.Text = "новое название";
        //            worksheet.Cols = 100;
        //            worksheet.Rows = 15000;
        //            wsheet.Update();
        //        }
        //    }
        //}
    }
}