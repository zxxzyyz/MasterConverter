using System;
using System.Windows.Forms;

namespace MasterConverter
{
    static class Program
    {
        /// <summary>
        /// アプリケーションのメイン エントリ ポイントです。
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            LogWriter.DeleteLog();
            if (args.Length != 0)
            {
                var csvWriter = new CsvWriter();
                for (int i = 0; i < args.Length; i++)
                {
                    csvWriter.ConvertFileToCsv(args[i]);
                }
            }
            else
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new MasterConverter());
            }
        }
    }
}
