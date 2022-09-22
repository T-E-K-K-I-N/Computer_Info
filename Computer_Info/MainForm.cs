using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Management;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;
using System.Diagnostics;

namespace Computer_Info
{
    public partial class Computer_Info : Form
    {




        /// <summary>
        /// Путь к файлу
        /// </summary>
        string _path = @"..\Computer_Information_Report.xlsx";

        public Computer_Info()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Метод отвечающий за получение информации по ключу
        /// </summary>
        private void GetHardwareInfo (string key, ListView list)
        {
            list.Items.Clear();

            ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM " + key);

            try
            {

                foreach (var obj in searcher.Get())
                {
                    ListViewGroup listViewGroup;

                    try
                    {
                        listViewGroup = list.Groups.Add(obj["Name"].ToString(), obj["Name"].ToString());
                    }
                    catch (Exception ex)
                    {
                        listViewGroup = list.Groups.Add(obj.ToString(), obj.ToString());
                    }

                    if (obj.Properties.Count == 0)
                    {
                        MessageBox.Show("Не удалось получить информацию","Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }

                    foreach (PropertyData data in obj.Properties)
                    {
                        ListViewItem item = new ListViewItem(listViewGroup);

                        if(list.Items.Count % 2 !=0)
                        {
                            item.BackColor = Color.White;
                        }
                        else
                        {
                            item.BackColor = Color.WhiteSmoke;
                        }

                        item.Text = data.Name;

                        if(data.Value != null && !string.IsNullOrEmpty(data.Value.ToString()))
                        {
                            switch (data.Value.GetType().ToString())
                            {
                                case "System.String[]":

                                    string[] stringData = data.Value as string[];
                                    string resStr1 = string.Empty;
                                    foreach (string str in stringData)
                                    {
                                        resStr1 += $"{str} ";
                                    }
                                    item.SubItems.Add(resStr1);

                                    break;

                                case "System.UInt16[]":

                                    ushort[] ushortData = data.Value as ushort[];
                                    string resStr2 = string.Empty;
                                    foreach (ushort ush in ushortData)
                                    {
                                        resStr2 += $"{Convert.ToString(ush)} ";
                                    }
                                    item.SubItems.Add(resStr2);

                                    break;
                                case "System.UInt32[]":

                                    uint[] uintData = data.Value as uint[];
                                    string resStr3 = string.Empty;
                                    foreach (uint value in uintData)
                                    {
                                        resStr3 += $"{Convert.ToString(value)} ";
                                    }
                                    item.SubItems.Add(resStr3);

                                    break;

                                default:

                                    item.SubItems.Add(data.Value.ToString());

                                    break;
                            }

                            list.Items.Add(item);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void toolStripComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            string key = string.Empty;

            switch (toolStripComboBox1.SelectedItem.ToString())
            {
                case "Процессор":
                    key = "Win32_Processor";
                    break;
                case "Видеокарта":
                    key = "Win32_VideoController";
                    break;
                case "Чипсет":
                    key = "Win32_IDEController";
                    break;
                case "Батарея":
                    key = "Win32_Battery";
                    break;
                case "Биос":
                    key = "Win32_BIOS";
                    break;
                case "Оперативная память":
                    key = "Win32_PhysicalMemory";
                    break;
                case "Кэш":
                    key = "Win32_CacheMemory";
                    break;
                case "USB":
                    key = "Win32_USBController";
                    break;
                case "Диск":
                    key = "Win32_DiskDrive";
                    break;
                case "Логические диски":
                    key = "Win32_LogicalDisk";
                    break;
                case "Монитор":
                    key = "Win32_DesktopMonitor";
                    break;
                case "Клавиатура":
                    key = "Win32_Keyboard";
                    break;
                case "Мышь":
                    key = "Win32_PointingDevice";
                    break;
                case "Сеть":
                    key = "Win32_NetworkAdapter";
                    break;
                case "Пользователи":
                    key = "Win32_Account";
                    break;
                default:
                    key = "Win32_Processor";
                    break;
            }

            GetHardwareInfo(key,listView1);
        }


        private void Computer_Info_Load(object sender, EventArgs e)
        {
            toolStripComboBox1.SelectedIndex = 0;
        }

        #region Сохранение файла

        private void сохранитьToolStripButton_Click(object sender, EventArgs e)
        {
            SaveInFile();
        }

        private void сохранитьВExcelФайлToolStripMenuItem_Click(object sender, EventArgs e)
        {
            SaveInFile();
        }

        private void SaveInFile()
        {
            Dictionary<string, string> dictionary = new Dictionary<string, string>();

            dictionary.Add("Процессор", "Win32_Processor");
            dictionary.Add("Видеокарта", "Win32_VideoController");
            dictionary.Add("Чипсет", "Win32_IDEController");
            dictionary.Add("Батарея", "Win32_Battery");
            dictionary.Add("Биос", "Win32_BIOS");
            dictionary.Add("Оперативная память", "Win32_PhysicalMemory");
            dictionary.Add("Кэш", "Win32_CacheMemory");
            dictionary.Add("USB", "Win32_USBController");
            dictionary.Add("Диск", "Win32_DiskDrive");
            dictionary.Add("Логические диски", "Win32_LogicalDisk");
            dictionary.Add("Монитор", "Win32_DesktopMonitor");
            dictionary.Add("Клавиатура", "Win32_Keyboard");
            dictionary.Add("Мышь", "Win32_PointingDevice");
            dictionary.Add("Сеть", "Win32_NetworkAdapter");
            dictionary.Add("Пользователи", "Win32_Account");

            try
            {
                ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                var package = new ExcelPackage();

                var reportExcel = WritingInfoInExcelFile(dictionary, package);
                File.WriteAllBytes(_path, reportExcel);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            MessageBox.Show("Отчет был успешно сохранен", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// Запись в Excel
        /// </summary>
        private byte[] WritingInfoInExcelFile(Dictionary<string, string> dictionary, ExcelPackage excelPackage)
        {
            using(excelPackage)
            {

                foreach (var item in dictionary)
                {
                    var sheet = excelPackage.Workbook.Worksheets.Add(item.Key);

                    int row = 1;

                    sheet.Cells[row, 1].Value = "Название";
                    sheet.Cells[row, 2].Value = "Значение";

                    row = 2;

                    ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT * FROM " + item.Value);
                    foreach (var obj in searcher.Get())
                    {
                        if (obj.Properties.Count == 0)
                        {
                            MessageBox.Show("Не удалось получить информацию", "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            break;
                        }

                        foreach (PropertyData data in obj.Properties)
                        {

                            sheet.Cells[row, 1].Value = data.Name;
                            sheet.Cells[row, 2].Value = data.Value;
                            row++;

                        }
                    }
                    
                    #region Работа со стилями

                    sheet.Column(1).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
                    sheet.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;


                    int lastRow = 1;

                    for (int i = 1; i < sheet.Rows.EndRow; i++)
                    {
                        if (sheet.Cells[i,1].Value == null)
                        {
                            lastRow = i;
                            break;
                        }
                    }


                    for (int i = 1; i < lastRow; i++)
                    {
                        sheet.Cells[i, 1].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                        sheet.Cells[i, 2].Style.Border.BorderAround(ExcelBorderStyle.Thin);
                    }

                    sheet.Column(1).Width = 40;
                    sheet.Column(1).Style.Font.Size = 10;
                    sheet.Column(1).Style.Font.Color.SetColor(Color.Black);
                    sheet.Row(1).Style.Font.Size = 20;

                    sheet.Column(2).Width = 50;
                    sheet.Column(2).Style.Font.Size = 10;
                    sheet.Column(2).Style.Font.Color.SetColor(Color.DarkBlue);

                    sheet.Row(1).Style.Font.Size = 13;
                    sheet.Cells["A1:B1"].Style.Border.BorderAround(ExcelBorderStyle.Medium);

                    sheet.Cells[$"A1:A{lastRow -1}"].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                    sheet.Cells[$"B1:B{lastRow -1}"].Style.Border.BorderAround(ExcelBorderStyle.Medium);

                    sheet.Cells["A1:B1"].Style.Fill.PatternType = ExcelFillStyle.Solid;
                    sheet.Cells["A1:B1"].Style.Fill.BackgroundColor.SetColor(Color.Coral);

                    #endregion

                    sheet.Protection.IsProtected = false;
                }

                return excelPackage.GetAsByteArray();
            }
        }

        #endregion

        #region Открытие файла

        private void открытьПоследнийОтчетToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFile();
        }

        private void открытьToolStripButton_Click(object sender, EventArgs e)
        {
            OpenFile();
        }

        private void OpenFile ()
        {
            try
            {
                if (!File.Exists(_path))
                {
                    ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;

                    using (var package = new ExcelPackage())
                    {
                        var sheet = package.Workbook.Worksheets.Add("List 1");
                        File.WriteAllBytes(_path, package.GetAsByteArray());
                        MessageBox.Show("Программа не нашла нужный файл и создала новый пустой документ.", "Информация", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }

                }
                else
                {
                    Process.Start(new ProcessStartInfo { FileName = _path });
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Ошибка", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        #endregion

    }
}
