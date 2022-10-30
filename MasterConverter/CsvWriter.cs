using System;
using System.IO;
using System.Text;
using System.Reflection;
using System.Globalization;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.VisualBasic;
using Microsoft.Office.Interop.Excel;
using System.Text.RegularExpressions;
using System.Linq;

namespace MasterConverter
{
    class CsvWriter
    {
        private readonly Regex PurseFile = new Regex(".*Pursemst_0000_00_00_.*\\.bin$");

        private readonly Regex CyberneFile = new Regex(".*CYBERNE_\\d{8}.*\\.xl[mst][xm]?$");

        private readonly Regex CompanyNameFile = new Regex(".*事業者コードマスタ_\\d{8}.*\\.xl[mst][xm]?$");

        private readonly Regex SheetNameRegex = new Regex("全量_\\d{8}$");

        private string newFile = string.Empty;

        public void ConvertFileToCsv(string file)
        {
            try
            {
                this.newFile = GetNewFileName(file);
                LogWriter.WriteLog(file, newFile, "変換開始");
                List<ICsvConvertible> records = GetRecords(file);
                if (records.Count != 0) WriteCsvFile(file, records);
            }
            catch (FileNotTargetException)
            {
                var sb = new StringBuilder();
                string fileName = Path.GetFileName(file);

                sb.AppendLine($"{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}");
                sb.AppendLine($"ファイル名：{fileName}");
                sb.AppendLine($"内容：処理対象ファイル異常");

                MessageBox.Show($"処理対象のファイルではありません。{Environment.NewLine}「Pursemst_0000_00_00_」を含み「.bin」で終わる、「CYBERNE_(8桁年月日)」を含むExcelファイル、「事業者コードマスタ_(8桁年月日)」を含むExcelファイルのみ処理対象です。", Assembly.GetExecutingAssembly().GetName().Name, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                LogWriter.WriteErrorLog(sb.ToString());
            }
            catch (IndexOutOfRangeException ex)
            {
                var sb = new StringBuilder();
                string fileName = Path.GetFileName(file);

                sb.AppendLine($"{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}");
                sb.AppendLine($"ファイル名：{fileName}");
                sb.AppendLine($"内容：フォーマット異常");
                sb.AppendLine($"{ex.Message}");
                sb.AppendLine($"{ex.StackTrace}");

                MessageBox.Show($"ファイルフォーマット異常。{Environment.NewLine}フォーマットが想定と違います。フォーマットが変わっていないか確認してください。", Assembly.GetExecutingAssembly().GetName().Name, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                LogWriter.WriteErrorLog(sb.ToString());
            }
            catch (Exception ex)
            {
                var sb = new StringBuilder();
                string fileName = Path.GetFileName(file);

                sb.AppendLine($"{DateTime.Now.ToString("yyyy/MM/dd HH:mm:ss")}");
                sb.AppendLine($"ファイル名：{fileName}");
                sb.AppendLine($"内容：例外");
                sb.AppendLine($"{ex.Message}");
                sb.AppendLine($"{ex.StackTrace}");

                MessageBox.Show("エラー発生、処理を中止します。", Assembly.GetExecutingAssembly().GetName().Name, MessageBoxButtons.OK, MessageBoxIcon.Error);
                LogWriter.WriteErrorLog(sb.ToString());
            }
            finally
            {
                GC.Collect();
                GC.WaitForFullGCComplete();
            }
        }

        private void WriteCsvFile(string file, List<ICsvConvertible> records)
        {
            var sb = new StringBuilder();
            var isHeaderWritten = false;

            foreach (var record in records)
            {
                if (!isHeaderWritten)
                {
                    sb.AppendLine(record.header);
                    isHeaderWritten = true;
                }

                PropertyInfo[] properties = record.GetType().GetProperties().Where(pi => pi.GetCustomAttributes(typeof(Ignore), true).Length == 0).ToArray(); ;
                foreach (PropertyInfo p in properties)
                {
                    string value = p.GetValue(record).ToString();
                    if (HasAtt(p, typeof(Trim))) value = value.Trim();
                    if (HasAtt(p, typeof(DoubleQuotes))) value = $"\"{value}\"";
                    if (HasAtt(p, typeof(EndComma))) value += ",";
                    sb.Append(value);
                }
                sb.AppendLine();
            }

            using (var writer = new StreamWriter(newFile, false, Encoding.GetEncoding(932)))
            {
                writer.Write(sb.ToString());
            }
            LogWriter.WriteLog(file, newFile, "変換終了");
            MessageBox.Show($"変換完了。{Environment.NewLine}変換前：{Path.GetFileName(file)}{Environment.NewLine}変換後：{Path.GetFileName(newFile)}", Assembly.GetExecutingAssembly().GetName().Name, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private string GetNewFileName(string file)
        {
            string newfileName = string.Empty;
            string fileName = Path.GetFileName(file);
            string fileDirectory = Path.GetDirectoryName(file);

            if (PurseFile.IsMatch(fileName)) newfileName = "PURSE";
            else if (CyberneFile.IsMatch(fileName)) newfileName = "CYBERNE";
            else if (CompanyNameFile.IsMatch(fileName)) newfileName = "COMPANYNAME";

            return $"{fileDirectory}\\{newfileName}{DateTime.Now.ToString("yyyyMMddHHmmss")}.csv";
        }

        private bool HasAtt(PropertyInfo p, Type att)
        {
            return p.GetCustomAttribute(att, true) != null ? true : false;
        }

        private List<ICsvConvertible> GetRecords(string file)
        {
            bool defaultPassword = false;
            var records = new List<ICsvConvertible>();

            string fileName = Path.GetFileName(file);

            if (PurseFile.IsMatch(fileName))
            {
                try
                {
                    byte[] bytes = File.ReadAllBytes(file);

                    int noOfCompanies = bytes[20]; // 登録事業者数（ループ回数）
                    int address = 36; // ループ開始アドレス

                    for (int i = 0; i < noOfCompanies; i++)
                    {
                        var record = new PurseCsv();

                        // 識別コード
                        record.IdentityFlag = int.Parse(bytes[address].ToString("X"), NumberStyles.HexNumber).ToString();
                        address++;

                        // 事業者コード
                        record.CompanyCode = GetValue(bytes, ref address, 2, true, false);

                        // パース上限金額-大人
                        record.MaxAdultCharge = GetValue(bytes, ref address, 3, false, true);

                        // パース上限金額 - 小児
                        record.MaxChildCharge = GetValue(bytes, ref address, 3, false, true);

                        // パース上限金額 - 大割
                        record.MaxAdultDiscount = GetValue(bytes, ref address, 3, false, true);

                        // パース上限金額 - 小割
                        record.MaxChildDiscount = GetValue(bytes, ref address, 3, false, true);

                        // 予備
                        record.Yobi = "";
                        address++;

                        records.Add(record);
                    }
                }
                catch(Exception)
                {
                    throw;
                }
            }
            else if (CyberneFile.IsMatch(fileName) || CompanyNameFile.IsMatch(fileName))
            {
                Microsoft.Office.Interop.Excel.Application excelApplication = new Microsoft.Office.Interop.Excel.Application();
                Workbooks excelWorkbooks = null;
                Workbook excelWorkbook = null;
                Worksheet excelWorkSheet = null;

                try
                {
                    excelWorkbooks = excelApplication.Workbooks;
                    excelWorkbook = excelWorkbooks.Open(file, ReadOnly: true, Password: "");

                    var sheetNames = new List<string>();
                    foreach (Worksheet sheet in excelWorkbook.Sheets)
                    {
                        sheetNames.Add(sheet.Name);
                    }

                    string sheetName = GetNewestSheet(sheetNames);

                    if (SheetNameRegex.IsMatch(sheetName))
                    {
                        excelWorkSheet = excelWorkbook.Sheets[sheetName];
                        if (CyberneFile.IsMatch(fileName)) AddRecords(excelWorkSheet, records);
                        else AddRecordsCompany(excelWorkSheet, records);
                    }
                    else
                    {
                        LogWriter.WriteLog(file, newFile, "シート無");
                        MessageBox.Show($"対象のシートがありません。{Environment.NewLine}シート名：{sheetName}。{Environment.NewLine}「全量_yyyymmdd」シートがありません。", Assembly.GetExecutingAssembly().GetName().Name, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }

                    if (excelWorkbook != null) excelWorkbook.Close(false);
                    if (excelWorkbooks != null) excelWorkbooks.Close();
                    if (excelApplication != null) excelApplication.Quit();
                }
                catch (Exception)
                {
                    while (true)
                    {
                        try
                        {
                            if (excelWorkbook != null) excelWorkbook.Close(false);
                            if (excelWorkbooks != null) excelWorkbooks.Close();

                            if (excelWorkSheet != null) Marshal.ReleaseComObject(excelWorkSheet);
                            if (excelWorkbook != null) Marshal.ReleaseComObject(excelWorkbook);
                            if (excelWorkbooks != null) Marshal.ReleaseComObject(excelWorkbooks);

                            string input;
                            if (defaultPassword)
                            {
                                input = Interaction.InputBox($"パスワードが違います。{Environment.NewLine}パスワードを入力してください。", "マスタコンバータ", "", 0, 0);
                            }
                            else
                            {
                                defaultPassword = true;
                                if (CyberneFile.IsMatch(fileName)) input = Properties.Settings.Default["CybernePassword"].ToString();
                                else input = Properties.Settings.Default["CompanyPassword"].ToString();
                            }
                            if (input.Length != 0)
                            {
                                LogWriter.WriteLog(file, newFile, $"パスワード指定 - {input}");

                                excelWorkbooks = excelApplication.Workbooks;
                                excelWorkbook = excelWorkbooks.Open(file, ReadOnly: true, Password: input);

                                var sheetNames = new List<string>();
                                foreach (Worksheet sheet in excelWorkbook.Sheets)
                                {
                                    sheetNames.Add(sheet.Name);
                                }
                                string sheetName = GetNewestSheet(sheetNames);

                                if (SheetNameRegex.IsMatch(sheetName))
                                {
                                    excelWorkSheet = excelWorkbook.Sheets[sheetName];
                                    if (CyberneFile.IsMatch(fileName)) AddRecords(excelWorkSheet, records);
                                    else AddRecordsCompany(excelWorkSheet, records);
                                }
                                else
                                {
                                    LogWriter.WriteLog(file, newFile, "シート無");
                                    MessageBox.Show($"対象のシートがありません。{Environment.NewLine}シート名：{sheetName}。{Environment.NewLine}「全量_yyyymmdd」シートがありません。", Assembly.GetExecutingAssembly().GetName().Name, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                                }

                                if (excelWorkbook != null) excelWorkbook.Close(false);
                                if (excelWorkbooks != null) excelWorkbooks.Close();
                                if (excelApplication != null) excelApplication.Quit();

                                break;
                            }
                            else
                            {
                                LogWriter.WriteLog(file, newFile, "変換キャンセル");
                                break;
                            }
                        }
                        catch (Exception)
                        {
                            continue;
                        }
                    }
                }
                finally
                {
                    if (excelWorkSheet != null) Marshal.ReleaseComObject(excelWorkSheet);
                    if (excelWorkbook != null) Marshal.ReleaseComObject(excelWorkbook);
                    if (excelWorkbooks != null) Marshal.ReleaseComObject(excelWorkbooks);
                    if (excelApplication != null) Marshal.ReleaseComObject(excelApplication);

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                }
            }
            else
            {
                throw new FileNotTargetException();
            }

            return records;
        }

        private string GetNewestSheet(List<string> sheetNames)
        {
            sheetNames.RemoveAll(item => !item.Contains("全量_"));

            List<int> dates = new List<int>();
            foreach (var name in sheetNames)
            {
                dates.Add(int.Parse(name.Replace("全量_", "")));
            }
            dates.Sort((a, b) => b.CompareTo(a));

            return dates.Count != 0 ? "全量_" + dates[0] : "全量_yyyymmddシート無し";
        }

        private void AddRecordsCompany(Worksheet excelSheet, List<ICsvConvertible> records)
        {
            var col = excelSheet.UsedRange.Columns["G:G", Type.Missing];
            foreach (var item in col.Cells)
            {
                if (item.Value == null) continue;

                var record = new CompanyNameCsv();
                var fullText = (string)item.Value;
                string[] text = fullText.Split(',');

                record.IdentityFlag = "0";
                record.CompanyCode = text[0];
                record.CompanyName = text[2];

                records.Add(record);
            }
        }

        private void AddRecords(Worksheet excelSheet, List<ICsvConvertible> records)
        {
            var col = excelSheet.UsedRange.Columns["J:J", Type.Missing];
            foreach (var item in col.Cells)
            {
                if (item.Value == null) continue;

                var record = new CyberneCsv();
                var fullText = (string)item.Value;
                string[] text = fullText.Split(',');

                record.AreaIdentityCode = text[0];
                record.CyberneCode = text[1];
                record.ExpiryDate = text[2];
                record.Name8Letters = text[3];
                record.Name4Letters = text[4];
                record.Name10Letters = text[5];
                record.CompanyCode = text[6];
                record.CompanyName = text[7];
                
                records.Add(record);
            }
        }

        private string GetValue(byte[] bytes, ref int address, int size, bool isBigEndian, bool isDecimal)
        {
            string result = string.Empty;

            for (int i = 0; i < size; i++)
            {
                if (isBigEndian)
                {
                    result += bytes[address].ToString("X2");
                }
                else
                {
                    result = bytes[address].ToString("X2") + result;
                }
                address++;
            }

            if (isDecimal) result = int.Parse(result, NumberStyles.HexNumber).ToString();

            return result;
        }
    }
}