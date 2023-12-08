using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using MyUI;
using Basic;

namespace APP_超晉線圈特性檢測機
{
    public partial class Form1 : Form
    {
        MySerialPort GOM_804_SerialPort = new MySerialPort();
        PLC_Device GOM_804通訊已連線指示 = new PLC_Device("S10100");
        PLC_Device GOM_804測試Ready = new PLC_Device("S10101");
        PLC_Device GOM_804測試完成 = new PLC_Device("S10102");
        PLC_Device GOM_804檢測觸發 = new PLC_Device("S6025");
        PLC_Device PLC_NumBox_GOM804檢測歐姆值 = new PLC_Device("D3000");
        MyTimer MyTimerGOM_wait_recieve = new MyTimer();
        MyTimer MyTimerGOM_TimeOut = new MyTimer();
        double Ohm_Value;
        DialogResult GOM_ComportInit_result;
        bool GOM_FLAG_ERR = false;

        void GOM_804_ComPort通訊()
        {
            this.GOM_804_Recieve();
            this.GOM_804_檢測開始();

        }

        public void GOM_804_Init(string PortName, int BaudRate)
        {

            GOM_804_SerialPort.Init(PortName, BaudRate, 8, System.IO.Ports.Parity.None, System.IO.Ports.StopBits.One);
            if (GOM_804_SerialPort.SerialPortOpen())
            {
                GOM_804通訊已連線指示.Bool = true;
            }
            //else
            //{
            //    cnt = 2;
            //    GOM_ComportInit_result = MessageBox.Show("微歐姆計通訊初始化失敗", "連線異常", MessageBoxButtons.OK, MessageBoxIcon.Error);

            //}


        }

        private void plC_Button208_btnClick(object sender, EventArgs e)
        {
            GOM_804_Init("COM4", 115200);
        }
        private void plC_Button_GOM_804_TRI_btnClick(object sender, EventArgs e)
        {

        }
        private void GOM_804_檢測開始()
        {
            int cnt = 0;
            if(cnt == 0)
            {
                GOM_804測試Ready.Bool = true;
                if (GOM_804檢測觸發.Bool)
                {
                    GOM_804測試Ready.Bool = false;
                    GOM_804測試完成.Bool = false;
                    cnt++;
                }
                
            }

            if (cnt == 1)
            {
                List<byte> Trigger_list_value = new List<byte>();
                byte[] value = new byte[] { 0x52, 0x45, 0x41, 0x44, 0x3F, 0x0D };//READ?歐姆計量測

                foreach (byte temp in value)
                {
                    Trigger_list_value.Add(temp);
                }
                try
                {
                    GOM_804_SerialPort.WriteByte(Trigger_list_value.ToArray());
                }
                catch
                {
                    if (!GOM_FLAG_ERR)
                    {
                        GOM_FLAG_ERR = true;
                        GOM_ComportInit_result = MessageBox.Show("微歐姆計通訊失敗", "連線異常", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    }
                }
                MyTimerGOM_wait_recieve.TickStop();
                MyTimerGOM_wait_recieve.StartTickTime(200);
                cnt = 0;
            }

        }
        public void GOM_804_Recieve()
        {
            int retry = 0;
            byte[] recieve_bytes = GOM_804_SerialPort.ReadByte();

            while (true)
            {
                if (retry >= 3)
                {
                    break;
                }
                if (MyTimerGOM_TimeOut.IsTimeOut())
                {
                    retry++;
                }
                if (recieve_bytes != null)
                {
                    if (recieve_bytes.Length >= 11)
                    {
                        Invoke(new EventHandler(delegate
                        {
                            textBox_GOM804_正負符號.Text = char.ConvertFromUtf32(recieve_bytes[0]);
                        }));
                        Ohm_Value = (recieve_bytes[1] - 48) * 10000 + (recieve_bytes[3] - 48) * 1000 +
                            (recieve_bytes[4] - 48) * 100 + (recieve_bytes[5] - 48) * 10 + (recieve_bytes[6] - 48) * 1;

                        Ohm_Value /= 10000;
                        Ohm_Value *= 尾數E(recieve_bytes[9], recieve_bytes[8]);
                        PLC_NumBox_GOM804檢測歐姆值.Value = (int)(Math.Round(Ohm_Value, 4) * 1000);

                        GOM_804_SerialPort.ClearReadByte();
                        GOM_804測試完成.Bool = true;
                        //break;
                    }

                }
            }


        }
        
        private int 尾數E(byte E,byte symbol)
        {
            int E_result = 1;

            if(symbol == 43)
            {
                if (E == 48) E_result *= 1;
                if (E == 49) E_result *= 10;
                if (E == 50) E_result *= 100;
                if (E == 51) E_result *= 1000;
                if (E == 52) E_result *= 10000;
                if (E == 53) E_result *= 100000;
                if (E == 54) E_result *= 1000000;
                if (E == 55) E_result *= 10000000;
                if (E == 56) E_result *= 100000000;
                if (E == 57) E_result *= 1000000000;
            }


            if (symbol == 45)
            {
                if (E == 48) E_result /= 1;
                if (E == 49) E_result /= 10;
                if (E == 50) E_result /= 100;
                if (E == 51) E_result /= 1000;
                if (E == 52) E_result /= 10000;
                if (E == 53) E_result /= 100000;
                if (E == 54) E_result /= 1000000;
                if (E == 55) E_result /= 10000000;
                if (E == 56) E_result /= 100000000;
                if (E == 57) E_result /= 1000000000;
            }




            return E_result;
        }


    }



}
