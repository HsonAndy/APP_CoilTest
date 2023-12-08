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
        void IWT_5000A_ComPort通訊()
        {
            IWT_5000A_Recieve();
            IWT_5000A_檢測開始();

        }
        MySerialPort IWT_5000A_SerialPort = new MySerialPort();
        PLC_Device IWT_5000A通訊已連線指示 = new PLC_Device("S10110");
        PLC_Device IWT_5000A測試Ready = new PLC_Device("S10111");
        PLC_Device IWT_5000A測試完成 = new PLC_Device("S10112");
        PLC_Device IWT_5000A採樣完成 = new PLC_Device("S10113");
        PLC_Device IWT_5000A_檢測按下 = new PLC_Device("S6051");
        PLC_Device IWT_5000A_檢測開始執行 = new PLC_Device("S6060");
        PLC_Device IWT_5000A匝間採樣觸發 = new PLC_Device("S6250");
        PLC_Device IWT_5000A匝間採樣觸發oFF = new PLC_Device("S6251");
        PLC_Device IWT_5000A檢測觸發 = new PLC_Device("S6075");
        PLC_Device flag_IWT_5000A匝間採樣觸發過 = new PLC_Device("S6076");
       // bool flag_IWT_5000A匝間採樣觸發過 = false;


        PLC_Device PLC_NumBox_IWT5000A檢測匝間面積比 = new PLC_Device("D3020");
        PLC_Device PLC_NumBox_IWT5000A檢測匝間電暈數 = new PLC_Device("D3025");
        MyTimer MyTimerIWT_wait_recieve = new MyTimer();
        MyTimer MyTimerIWT_TimeOut = new MyTimer();
        double 面積比結果_Value;
        double 電暈數結果_Value;
        DialogResult result;
        DialogResult IWT_ComportInit_result;
        bool IWT_FLAG_ERR = false;

        public void IWT_5000A_Init(string PortName, int BaudRate)
        {
            IWT_5000A採樣完成.Bool = false;
            flag_IWT_5000A匝間採樣觸發過.Bool = false;
            IWT_5000A_SerialPort.Init(PortName, BaudRate, 8, System.IO.Ports.Parity.None, System.IO.Ports.StopBits.One);
            if (IWT_5000A_SerialPort.SerialPortOpen())
            {
                IWT_5000A通訊已連線指示.Bool = true;
            }

        }
        private void plC_Button_匝間採樣_btnClick(object sender, EventArgs e)
        {
            //IWT_5000A_SerialPort.ClearReadByte();
            //List<byte> Trigger_list_value = new List<byte>();
            //byte[] value = new byte[] { 0x3A, 0x53, 0x4D, 0x0D, 0x0A };//:SM匝間採樣

            //foreach (byte temp in value)
            //{
            //    Trigger_list_value.Add(temp);
            //}
            //try
            //{
            //    IWT_5000A_SerialPort.WriteByte(Trigger_list_value.ToArray());
            //}
            //catch
            //{
            //    if (!IWT_FLAG_ERR)
            //    {
            //        IWT_FLAG_ERR = true;
            //        IWT_ComportInit_result = MessageBox.Show("匝間儀器通訊失敗", "連線異常", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    }
            //}
            //MyTimerIWT_wait_recieve.TickStop();
            //MyTimerIWT_wait_recieve.StartTickTime(200);
            //MyTimerIWT_TimeOut.TickStop();
            //MyTimerIWT_TimeOut.StartTickTime(200);

        }
        private void IWT_5000A_檢測開始()
        {
            int cnt = 0;

            if (cnt == 0)
            {
                if (IWT_5000A採樣完成.Bool)
                {
                    IWT_5000A測試Ready.Bool = true;
                    flag_IWT_5000A匝間採樣觸發過.Bool = true;
                    if (IWT_5000A檢測觸發.Bool)
                    {
                        IWT_5000A測試Ready.Bool = false;
                        IWT_5000A測試完成.Bool = false;
                        cnt++;
                    }
                }
                else if (!IWT_5000A採樣完成.Bool && IWT_5000A檢測觸發.Bool)
                {
                    //IWT_5000A測試Ready.Bool = false;
                    IWT_5000A測試完成.Bool = false;
                    this.IWT_5000A_檢測按下.Bool = false;
                    this.IWT_5000A_檢測開始執行.Bool = false;
                    result = MessageBox.Show("請先執行匝間採樣", "尚未採樣", MessageBoxButtons.OK, MessageBoxIcon.Information, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                   // if (result == DialogResult.OK) cnt = 0;
                }

            }
            if (cnt == 1)
            {
                IWT_5000A_SerialPort.ClearReadByte();
                List<byte> Trigger_list_value = new List<byte>();
                byte[] value = new byte[] { 0x3A, 0x43, 0x53, 0x0D, 0x0A };//:CS匝間測量開始

                foreach (byte temp in value)
                {
                    Trigger_list_value.Add(temp);
                }
                try
                {
                    IWT_5000A_SerialPort.WriteByte(Trigger_list_value.ToArray());
                }

                catch
                {
                    if (!IWT_FLAG_ERR)
                    {
                        IWT_FLAG_ERR = true;
                        IWT_ComportInit_result = MessageBox.Show("匝間儀器通訊失敗", "連線異常", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    }
                }
                MyTimerGOM_wait_recieve.TickStop();
                MyTimerGOM_wait_recieve.StartTickTime(200);
                cnt = 0;
            }
            if (IWT_5000A匝間採樣觸發.Bool)
            {
                
                IWT_5000A採樣完成.Bool = false;
                if (!flag_IWT_5000A匝間採樣觸發過.Bool)
                {
                    IWT_5000A_SerialPort.ClearReadByte();
                    List<byte> Trigger_list_value = new List<byte>();
                    byte[] value = new byte[] { 0x3A, 0x53, 0x4D, 0x0D, 0x0A };//:SM匝間採樣

                    foreach (byte temp in value)
                    {
                        Trigger_list_value.Add(temp);
                    }
                    try
                    {
                        IWT_5000A_SerialPort.WriteByte(Trigger_list_value.ToArray());
                        IWT_5000A匝間採樣觸發oFF.Bool = true;
                        
                    }
                    catch
                    {
                        if (!IWT_FLAG_ERR)
                        {
                            IWT_FLAG_ERR = true;
                            IWT_ComportInit_result = MessageBox.Show("匝間儀器通訊失敗", "連線異常", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);

                        }
                    }
                    MyTimerIWT_wait_recieve.TickStop();
                    MyTimerIWT_wait_recieve.StartTickTime(200);
                    MyTimerIWT_TimeOut.TickStop();
                    MyTimerIWT_TimeOut.StartTickTime(200);
                }
            
            }


        }
        private void IWT_5000A_檢測收值()
        {
            IWT_5000A_SerialPort.ClearReadByte();
            List<byte> Trigger_list_value = new List<byte>();
            byte[] value = new byte[] { 0x3A, 0x41, 0x52, 0x0D, 0x0A };//:AR收取量測數值

            foreach (byte temp in value)
            {
                Trigger_list_value.Add(temp);
            }
            IWT_5000A_SerialPort.WriteByte(Trigger_list_value.ToArray());


        }


        public void IWT_5000A_Recieve()
        {
            int retry = 0;
            string recieve_string = IWT_5000A_SerialPort.ReadString();
            byte[] recieve_bytes = IWT_5000A_SerialPort.ReadByte();

            while (true)
            {
                if (retry >= 3)
                {
                    break;
                }
                if (MyTimerIWT_TimeOut.IsTimeOut())
                {
                    retry++;
                }

                if (recieve_string != null)
                {
                    if (recieve_string.Length >= 5)
                    {
                        if (recieve_string == "BUSY\r\nSample Success!\r\n")
                        {
                            IWT_5000A採樣完成.Bool = true;
                            flag_IWT_5000A匝間採樣觸發過.Bool = true;
                            IWT_5000A_SerialPort.ClearReadByte();
                            break;
                        }
                        if(recieve_string == "BUSY\r\nPASS\r\n" || recieve_string == "BUSY\r\nFAIL\r\n")
                        {
                            IWT_5000A_檢測收值();
                            IWT_5000A_SerialPort.ClearReadByte();
                            break;
                        }

                        
                    }
                }
                if (recieve_bytes != null)
                {
                    if (recieve_bytes.Length >= 35)
                    {
                        if (recieve_bytes[0] == 'A' && recieve_bytes[1] == 'r' && recieve_bytes[2] == 'e' && recieve_bytes[3] == 'a')
                        {
                            Invoke(new EventHandler(delegate
                            {
                                textBox_IWT5000A正負符號.Text = char.ConvertFromUtf32(recieve_bytes[6]);
                            }));
                            //if (recieve_bytes[7] - 48 < 0) recieve_bytes[7] = 48;
                            //if (recieve_bytes[8] - 48 < 0) recieve_bytes[8] = 48;
                            //if (recieve_bytes[10] - 48 < 0) recieve_bytes[10] = 48;
                            if (recieve_bytes[18] - 48 < 0) recieve_bytes[18] = 48;
                            if (recieve_bytes[19] - 48 < 0) recieve_bytes[19] = 48;
                            if (recieve_bytes[21] - 48 < 0) recieve_bytes[21] = 48;
                            if (recieve_bytes[30] - 48 < 0) recieve_bytes[30] = 48;
                            if (recieve_bytes[31] - 48 < 0) recieve_bytes[31] = 48;
                            if (recieve_bytes[32] - 48 < 0) recieve_bytes[32] = 48;

                            this.面積比結果_Value = (recieve_bytes[18] - 48) * 100 + (recieve_bytes[19] - 48) * 10 + (recieve_bytes[21] - 48) * 1;
                            this.PLC_NumBox_IWT5000A檢測匝間面積比.Value = (int)this.面積比結果_Value;

                            this.電暈數結果_Value = (recieve_bytes[30] - 48) * 100 + (recieve_bytes[31] - 48) * 10 + (recieve_bytes[32] - 48) * 1;
                            this.PLC_NumBox_IWT5000A檢測匝間電暈數.Value = (int)this.電暈數結果_Value;
                            IWT_5000A測試完成.Bool = true;
                            IWT_5000A_SerialPort.ClearReadByte();
                            break;

                        }
                    }
                }


            }


        }

    }
}
