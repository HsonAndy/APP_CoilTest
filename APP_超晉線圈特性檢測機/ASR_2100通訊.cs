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
        void ASR_2100_ComPort通訊()
        {
            this.ASR_2100_檢測開始_OUTP1();
            this.ASR_2100_Recieve();

        }
        MySerialPort ASR_2100_SerialPort = new MySerialPort();
        PLC_Device ASR_2100通訊已連線指示 = new PLC_Device("S10130");
        PLC_Device ASR_2100測試Ready = new PLC_Device("S10131");
        PLC_Device ASR_2100測試完成 = new PLC_Device("S10132");
        PLC_Device ASR_2100檢測觸發 = new PLC_Device("S6175");
        PLC_Device PLC_NumBox_ASR_2100電功率Vrms量測值 = new PLC_Device("D3060");
        PLC_Device PLC_NumBox_ASR_2100電功率Irms量測值 = new PLC_Device("D3062");
        PLC_Device PLC_NumBox_ASR_2100電功率P量測值 = new PLC_Device("D3064");
        MyTimer MyTimerASR_wait_recieve = new MyTimer();
        MyTimer MyTimerASR_TimeOut = new MyTimer();
        MyTimer MyTimerASR_wait_test = new MyTimer();
        DialogResult ASR_ComportInit_result;
        double ASR_Vrms_Value;
        double ASR_Irms_Value;
        double ASR_P_Value;
        bool ASR_FLAG_ERR = false;
        public void ASR_2100_Init(string PortName, int BaudRate)
        {

            ASR_2100_SerialPort.Init(PortName, BaudRate, 8, System.IO.Ports.Parity.None, System.IO.Ports.StopBits.One);
            if (ASR_2100_SerialPort.SerialPortOpen())
            {
                ASR_2100通訊已連線指示.Bool = true;
            }

        }
        private void ASR_2100_檢測開始_OUTP1()
        {
            int cnt = 0;
            if (cnt == 0)
            {
                this.ASR_2100測試Ready.Bool = true;
                if (ASR_2100檢測觸發.Bool)
                {
                    this.ASR_2100測試Ready.Bool = false;
                    this.ASR_2100測試完成.Bool = false;
                    cnt++;
                }

            }

            if (cnt == 1)
            {
                this.ASR_2100_SerialPort.ClearReadByte();
                List<byte> Trigger_list_value = new List<byte>();
                byte[] value = new byte[] { 0x3A, 0x4F, 0x55, 0x54, 0x50, 0x20, 0x31, 0x0D, 0x0A };//:OUTP 1測量OUT開啟

                foreach (byte temp in value)
                {
                    Trigger_list_value.Add(temp);
                }
                try
                {
                    this.ASR_2100_SerialPort.WriteByte(Trigger_list_value.ToArray());
                }
                catch
                {
                    if (!ASR_FLAG_ERR)
                    {
                        ASR_FLAG_ERR = true;
                        this.ASR_ComportInit_result = MessageBox.Show("耐壓絕緣儀器通訊失敗", "連線異常", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    }
                }

                ASR_2100_檢測開始_OUTPask();
                cnt = 0;
            }

                    

            


        }

         private void ASR_2100_檢測開始_OUTPask()
        {

            this.ASR_2100_SerialPort.ClearReadByte();
            List<byte> Trigger_list_value = new List<byte>();
            byte[] value = new byte[] { 0x3A, 0x4F, 0x55, 0x54, 0x50, 0x3F, 0x0D, 0x0A };//:OUTP?查詢out狀態

            foreach (byte temp in value)
            {
                Trigger_list_value.Add(temp);
            }
            try
            {
                this.ASR_2100_SerialPort.WriteByte(Trigger_list_value.ToArray());
            }
            catch
            {
                if (!ASR_FLAG_ERR)
                {
                    ASR_FLAG_ERR = true;
                    this.ASR_ComportInit_result = MessageBox.Show("耐壓絕緣儀器通訊失敗", "連線異常", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                }
            }
            MyTimerASR_wait_test.TickStop();
            MyTimerASR_wait_test.StartTickTime(500);

        }
        private void ASR_2100_檢測開始_OUTP0()
        {

            this.ASR_2100_SerialPort.ClearReadByte();
            List<byte> Trigger_list_value = new List<byte>();
            byte[] value = new byte[] { 0x3A, 0x4F, 0x55, 0x54, 0x50, 0x20, 0x30, 0x0D, 0x0A };//:OUTP 0測量OUT關閉

            foreach (byte temp in value)
            {
                Trigger_list_value.Add(temp);
            }
            try
            {
                this.ASR_2100_SerialPort.WriteByte(Trigger_list_value.ToArray());
            }
            catch
            {
                if (!ASR_FLAG_ERR)
                {
                    ASR_FLAG_ERR = true;
                    this.ASR_ComportInit_result = MessageBox.Show("耐壓絕緣儀器通訊失敗", "連線異常", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                }
            }


        }
        private void ASR_2100_檢測開始_READ()
        {

            this.ASR_2100_SerialPort.ClearReadByte();
            List<byte> Trigger_list_value = new List<byte>();
            byte[] value = new byte[] { 0x3A, 0x52, 0x45, 0x41, 0x44, 0x3F, 0x0D, 0x0A };//:READ?測量結果

            foreach (byte temp in value)
            {
                Trigger_list_value.Add(temp);
            }
            try
            {
                this.ASR_2100_SerialPort.WriteByte(Trigger_list_value.ToArray());
            }
            catch
            {
                if (!ASR_FLAG_ERR)
                {
                    ASR_FLAG_ERR = true;
                    this.ASR_ComportInit_result = MessageBox.Show("耐壓絕緣儀器通訊失敗", "連線異常", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                }
            }


        }
        private void ASR_2100_Recieve()
        {
            int retry = 0;
            string recieve_string = ASR_2100_SerialPort.ReadString();
            byte[] recieve_bytes = ASR_2100_SerialPort.ReadByte();

            while(true)
            {
                if(retry == 3)
                {
                    break;
                }
                if(MyTimerASR_TimeOut.IsTimeOut())
                {
                    retry++;
                }
                if (recieve_string != null)
                {
                    if(recieve_string.Length >= 3)
                    {
                        if(recieve_string == "+1\n")
                        {
                            if(MyTimerASR_wait_test.IsTimeOut())
                            {
                                ASR_2100_檢測開始_READ();
                                break;
                            }

                        }

                    }
                }
                if(recieve_bytes != null)
                {
                    if(recieve_string.Length >= 135 && recieve_string.Length < 142)//電壓個位數
                    {

                        Invoke(new EventHandler(delegate
                        {
                            textBox_電功率Vrms符號.Text = char.ConvertFromUtf32(recieve_bytes[0]);
                            textBox_電功率Irms符號.Text = char.ConvertFromUtf32(recieve_bytes[36]);
                            textBox_電功率P符號.Text = char.ConvertFromUtf32(recieve_bytes[76]);
                        }));

                        this.ASR_Vrms_Value = (recieve_bytes[1] - 48) * 100000 + (recieve_bytes[2] - 48) * 10000 +
                                    (recieve_bytes[4] - 48) * 1000 + (recieve_bytes[5] - 48) * 100 + (recieve_bytes[6] - 48) * 10 + (recieve_bytes[7] - 48) * 1;

                        this.ASR_Vrms_Value /= 1000;
                        this.PLC_NumBox_ASR_2100電功率Vrms量測值.Value = (int)(Math.Round(ASR_Vrms_Value, 4) * 1000);


                        this.ASR_Irms_Value = (recieve_bytes[37] - 48) * 10000 + (recieve_bytes[39] - 48) * 1000 +
                                    (recieve_bytes[40] - 48) * 100 + (recieve_bytes[41] - 48) * 10 + (recieve_bytes[42] - 48) * 1;
                        this.ASR_Irms_Value /= 1000;
                        this.PLC_NumBox_ASR_2100電功率Irms量測值.Value = (int)(Math.Round(ASR_Irms_Value, 4) * 1000);


                        this.ASR_P_Value = (recieve_bytes[77] - 48) * 100000 + (recieve_bytes[78] - 48) * 10000 +
                                    (recieve_bytes[80] - 48) * 1000 + (recieve_bytes[81] - 48) * 100 + (recieve_bytes[82] - 48) * 10 + (recieve_bytes[83] - 48) * 1;
                        this.ASR_P_Value /= 1000;
                        this.PLC_NumBox_ASR_2100電功率P量測值.Value = (int)(Math.Round(ASR_P_Value, 4) * 1000);

                        this.ASR_2100_檢測開始_OUTP0();
                        this.ASR_2100測試完成.Bool = true;

                    }
                    if (recieve_string.Length >= 139 && recieve_string.Length < 143)//電壓十位數
                    {

                        Invoke(new EventHandler(delegate
                        {
                            textBox_電功率Vrms符號.Text = char.ConvertFromUtf32(recieve_bytes[0]);

                        }));

                        ASR_2100_檢測開始_OUTP0();
                        this.ASR_2100測試完成.Bool = true;

                    }
                    if (recieve_string.Length >= 143)//電壓百位數
                    {

                        Invoke(new EventHandler(delegate
                        {
                            textBox_電功率Vrms符號.Text = char.ConvertFromUtf32(recieve_bytes[0]);

                        }));

                        ASR_2100_檢測開始_OUTP0();
                        this.ASR_2100測試完成.Bool = true;

                    }
                }



            }



        }

    }
}
