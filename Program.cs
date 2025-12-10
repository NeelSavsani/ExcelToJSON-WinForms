using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;                 // REQUIRED
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SampleProject1
{
    internal static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            // REQUIRED for ExcelDataReader to read XLS and XLSX files
            Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
