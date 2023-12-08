using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace APP_超晉線圈特性檢測機
{
    static class Program
    {

        /// <summary>
        /// 應用程式的主要進入點。
        /// </summary>
        [STAThread]
        static void Main()
        {
            bool createdNew;
            string mutexName = "APP_日發檢測機"; // 唯一的互斥體名稱
            Mutex mutex = new Mutex(true, mutexName, out createdNew);
            if (createdNew)
            {
                Application.EnableVisualStyles();
                Application.SetCompatibleTextRenderingDefault(false);
                Application.Run(new Form1());
            }

            else
            {
                MessageBox.Show("應用程序已經運行。");
            }
        }

    }
}
