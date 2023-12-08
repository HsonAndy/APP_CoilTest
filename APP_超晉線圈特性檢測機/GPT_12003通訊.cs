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
        MySerialPort GPT_12003_SerialPort = new MySerialPort();
        PLC_Device GPT_12003通訊已連線指示 = new PLC_Device("S10120");
        PLC_Device 線圈90度模式 = new PLC_Device("S5000");
        PLC_Device GPT_12003測試Ready = new PLC_Device("S10121");
        PLC_Device GPT_12003測試完成 = new PLC_Device("S10122");
        PLC_Device GPT_12003檢測觸發 = new PLC_Device("S6125");
        PLC_Device GPT_12003檢測結果 = new PLC_Device("S5020");
        MyTimer MyTimerGPT_wait_recieve = new MyTimer();
        MyTimer MyTimerGPT_TimeOut = new MyTimer();
        DialogResult GPT_ComportInit_result;
        double GPT_Value;
        bool ACW檢測結果;
        bool DCW檢測結果;
        bool IR檢測結果;
        bool CONT檢測結果;
        bool GPT_FLAG_ERR = false;

        void GPT_12003_ComPort通訊()
        {
            this.GPT_12003_檢測開始_TESTON();
            this.GPT_12003_Recieve();          
        }
        public void GPT_12003_Init(string PortName, int BaudRate)
        {
            this.GPT_12003通訊已連線指示.Bool = false;
            GPT_12003_SerialPort.Init(PortName, BaudRate, 8, System.IO.Ports.Parity.None, System.IO.Ports.StopBits.One);
            if (GPT_12003_SerialPort.SerialPortOpen())
            {
                this.GPT_12003通訊已連線指示.Bool = true;
                
            }
            this.GPT_12003_檢測開始_ReturnON();
        }

        private void GPT_12003_檢測開始_ReturnON()
        {
            GPT_12003_SerialPort.ClearReadByte();
            List<byte> Trigger_list_value = new List<byte>();
            byte[] value = new byte[] { 0x54, 0x45, 0x53, 0x54, 0x3A, 0x52, 0x45, 0x54, 0x55, 0x52, 0x4E, 0x20, 0x4F, 0x4E, 0x0D };//TEST:RETURN ON 耐壓絕緣測量開始

            foreach (byte temp in value)
            {
                Trigger_list_value.Add(temp);
            }
            try
            {
                GPT_12003_SerialPort.WriteByte(Trigger_list_value.ToArray());
            }
            catch
            {
                if(!GPT_FLAG_ERR)
                {
                    GPT_FLAG_ERR = true;
                    GPT_ComportInit_result = MessageBox.Show("耐壓絕緣儀器通訊失敗", "連線異常", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    
                }
                
            }

        }
        private void GPT_12003_檢測開始_TESTON()
        {
            int cnt = 0;
            if (cnt == 0)
            {
                GPT_12003測試Ready.Bool = true;
                if (GPT_12003檢測觸發.Bool)
                {
                    GPT_12003測試Ready.Bool = false;
                    GPT_12003測試完成.Bool = false;
                    cnt++;
                }

            }

            if (cnt == 1)
            {
                GPT_12003_SerialPort.ClearReadByte();
                List<byte> Trigger_list_value = new List<byte>();
                byte[] value = new byte[] { 0x46, 0x55, 0x4E, 0x43, 0x74, 0x69, 0x6F, 0x6E, 0x3A, 0x54, 0x45, 0x53, 0x54, 0x20, 0x4F, 0x4E, 0x0D };//FUNCtion:TEST ON 耐壓絕緣測量開始

                foreach (byte temp in value)
                {
                    Trigger_list_value.Add(temp);
                }
                try
                {
                    GPT_12003_SerialPort.WriteByte(Trigger_list_value.ToArray());
                }
                catch
                {
                    if (!GPT_FLAG_ERR)
                    {
                        GPT_FLAG_ERR = true;
                        GPT_ComportInit_result = MessageBox.Show("耐壓絕緣儀器通訊失敗", "連線異常", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    }
                }

                cnt = 0;
            }

        }

        private void GPT_12003_檢測開始_TESTOFF()
        {

            GPT_12003_SerialPort.ClearReadByte();
            List<byte> Trigger_list_value = new List<byte>();
            byte[] value = new byte[] { 0x46, 0x55, 0x4E, 0x43, 0x74, 0x69, 0x6F, 0x6E, 0x3A, 0x54, 0x45, 0x53, 0x54, 0x20, 0x4F, 0x46, 0x46, 0x0D };//FUNCtion:TEST OFF 耐壓絕緣測量關閉警報

            foreach (byte temp in value)
            {
                Trigger_list_value.Add(temp);
            }
            try
            {
                GPT_12003_SerialPort.WriteByte(Trigger_list_value.ToArray());
            }

            catch
            {
                if (!GPT_FLAG_ERR)
                {
                    GPT_FLAG_ERR = true;
                    GPT_ComportInit_result = MessageBox.Show("耐壓絕緣儀器通訊失敗", "連線異常", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                }
            }
        }

        private void GPT_12003_檢測開始_MEAS1()
        {

            GPT_12003_SerialPort.ClearReadByte();
            List<byte> Trigger_list_value = new List<byte>();
            byte[] value = new byte[] { 0x4D, 0x45, 0x41, 0x53, 0x31, 0x3F, 0x0D };//MEAS1? 回傳CH1量測結果

            foreach (byte temp in value)
            {
                Trigger_list_value.Add(temp);
            }
            GPT_12003_SerialPort.WriteByte(Trigger_list_value.ToArray());
        }
        private void GPT_12003_檢測開始_MEAS2()
        {

            GPT_12003_SerialPort.ClearReadByte();
            List<byte> Trigger_list_value = new List<byte>();
            byte[] value = new byte[] { 0x4D, 0x45, 0x41, 0x53, 0x32, 0x3F, 0x0D };//MEAS2? 回傳CH2量測結果

            foreach (byte temp in value)
            {
                Trigger_list_value.Add(temp);
            }
            GPT_12003_SerialPort.WriteByte(Trigger_list_value.ToArray());
        }
        private void GPT_12003_檢測開始_MEAS3()
        {

            GPT_12003_SerialPort.ClearReadByte();
            List<byte> Trigger_list_value = new List<byte>();
            byte[] value = new byte[] { 0x4D, 0x45, 0x41, 0x53, 0x33, 0x3F, 0x0D };//MEAS3? 回傳CH3量測結果

            foreach (byte temp in value)
            {
                Trigger_list_value.Add(temp);
            }
            GPT_12003_SerialPort.WriteByte(Trigger_list_value.ToArray());
        }
        private void GPT_12003_檢測開始_MEAS4()
        {

            GPT_12003_SerialPort.ClearReadByte();
            List<byte> Trigger_list_value = new List<byte>();
            byte[] value = new byte[] { 0x4D, 0x45, 0x41, 0x53, 0x34, 0x3F, 0x0D };//MEAS4? 回傳CH4量測結果

            foreach (byte temp in value)
            {
                Trigger_list_value.Add(temp);
            }
            GPT_12003_SerialPort.WriteByte(Trigger_list_value.ToArray());
        }
        private void GPT_12003_Recieve()
        {
            int retry = 0;
            string recieve_string = GPT_12003_SerialPort.ReadString();
            byte[] recieve_bytes = GPT_12003_SerialPort.ReadByte();

            while(true)
            {

                if (retry >= 3)
                {
                    break;
                }
                if (MyTimerGPT_TimeOut.IsTimeOut())
                {
                    retry++;
                }
                if (recieve_string != null)
                {
                    if (recieve_string.Length >= 3)
                    {
                        if (recieve_string == "OK\r\n")
                        {
                            this.GPT_12003_檢測開始_TESTOFF();
                            this.GPT_12003_檢測開始_MEAS1();
                            break;
                        }
                    }
                    if (recieve_string.Length >= 38)
                    {
                        if (recieve_bytes[0] == 'A' && recieve_bytes[1] == 'C' && recieve_bytes[2] == 'W')
                        {
                            if(recieve_bytes[4] == 'P' && recieve_bytes[5] == 'A' && recieve_bytes[6] == 'S' && recieve_bytes[7] == 'S')
                            {
                                Invoke(new EventHandler(delegate
                                {
                                    textBox_ACW量測值.ForeColor = Color.Lime;
                                    textBox_ACW量測值.Text = char.ConvertFromUtf32(recieve_bytes[20]) + char.ConvertFromUtf32(recieve_bytes[21]) + char.ConvertFromUtf32(recieve_bytes[22])
                                    + char.ConvertFromUtf32(recieve_bytes[23]) + char.ConvertFromUtf32(recieve_bytes[24]) + char.ConvertFromUtf32(recieve_bytes[25]);
                                }));
                                this.ACW檢測結果 = true;
                            }
                            else
                            {
                                Invoke(new EventHandler(delegate
                                {
                                    textBox_ACW量測值.ForeColor = Color.Red;
                                    textBox_ACW量測值.Text = char.ConvertFromUtf32(recieve_bytes[20]) + char.ConvertFromUtf32(recieve_bytes[21]) + char.ConvertFromUtf32(recieve_bytes[22])
                                    + char.ConvertFromUtf32(recieve_bytes[23]) + char.ConvertFromUtf32(recieve_bytes[24]) + char.ConvertFromUtf32(recieve_bytes[25]);
                                }));
                                this.ACW檢測結果 = false;
                            }

                            this.GPT_12003_SerialPort.ClearReadByte();
                            this.GPT_12003_檢測開始_MEAS2();
                            break;                            
                        }
                        if (recieve_bytes[0] == 'I' && recieve_bytes[1] == 'R')
                        {

                            if (recieve_bytes[4] == 'P' && recieve_bytes[5] == 'A' && recieve_bytes[6] == 'S' && recieve_bytes[7] == 'S')
                            {
                                Invoke(new EventHandler(delegate
                                {
                                    textBox_IR絕緣量測值.ForeColor = Color.Lime;
                                    textBox_IR絕緣量測值.Text = char.ConvertFromUtf32(recieve_bytes[18]) + char.ConvertFromUtf32(recieve_bytes[19]) +
                                    char.ConvertFromUtf32(recieve_bytes[20]) + char.ConvertFromUtf32(recieve_bytes[21]) + char.ConvertFromUtf32(recieve_bytes[22])
                                    + char.ConvertFromUtf32(recieve_bytes[23]) + char.ConvertFromUtf32(recieve_bytes[24]) + char.ConvertFromUtf32(recieve_bytes[25]) + char.ConvertFromUtf32(recieve_bytes[26]);
                                }));
                                this.IR檢測結果 = true;
                            }
                            else
                            {
                                Invoke(new EventHandler(delegate
                                {
                                    textBox_IR絕緣量測值.ForeColor = Color.Red;
                                    textBox_IR絕緣量測值.Text = char.ConvertFromUtf32(recieve_bytes[20]) + char.ConvertFromUtf32(recieve_bytes[21]) + char.ConvertFromUtf32(recieve_bytes[22])
                                    + char.ConvertFromUtf32(recieve_bytes[23]) + char.ConvertFromUtf32(recieve_bytes[24]) + char.ConvertFromUtf32(recieve_bytes[25]) + char.ConvertFromUtf32(recieve_bytes[26]);
                                }));
                                this.IR檢測結果 = false;
                            }

                            if (ACW檢測結果 && IR檢測結果) this.GPT_12003檢測結果.Bool = true;
                            else this.GPT_12003檢測結果.Bool = false;
                            this.GPT_12003_SerialPort.ClearReadByte();
                            this.GPT_12003測試完成.Bool = true;
                            //this.GPT_12003_檢測開始_MEAS3();
                            break;
                        }
                        if (recieve_bytes[0] == 'D' && recieve_bytes[1] == 'C' && recieve_bytes[2] == 'W')
                        {

                            if (recieve_bytes[4] == 'P' && recieve_bytes[5] == 'A' && recieve_bytes[6] == 'S' && recieve_bytes[7] == 'S')
                            {
                                Invoke(new EventHandler(delegate
                                {
                                    textBox_DCW量測值.ForeColor = Color.Lime;
                                    textBox_DCW量測值.Text = char.ConvertFromUtf32(recieve_bytes[20]) + char.ConvertFromUtf32(recieve_bytes[21]) + char.ConvertFromUtf32(recieve_bytes[22])
                                    + char.ConvertFromUtf32(recieve_bytes[23]) + char.ConvertFromUtf32(recieve_bytes[24]) + char.ConvertFromUtf32(recieve_bytes[25]);
                                }));
                                this.DCW檢測結果 = true;
                            }
                            else
                            {
                                Invoke(new EventHandler(delegate
                                {
                                    textBox_DCW量測值.ForeColor = Color.Red;
                                    textBox_DCW量測值.Text = char.ConvertFromUtf32(recieve_bytes[20]) + char.ConvertFromUtf32(recieve_bytes[21]) + char.ConvertFromUtf32(recieve_bytes[22])
                                    + char.ConvertFromUtf32(recieve_bytes[23]) + char.ConvertFromUtf32(recieve_bytes[24]) + char.ConvertFromUtf32(recieve_bytes[25]);
                                }));
                                this.DCW檢測結果 = false;
                            }
                            this.GPT_12003_SerialPort.ClearReadByte();
                            if(線圈90度模式.Bool)
                            {
                                this.GPT_12003_檢測開始_MEAS4();
                                break;
                            }
                            else
                            {
                                this.GPT_12003測試完成.Bool = true;
                                break;
                            }
                            
                        }
                        if (線圈90度模式.Bool)
                        {
                            if (recieve_bytes[0] == 'C' && recieve_bytes[1] == 'O' && recieve_bytes[2] == 'N')
                            {
                                if (recieve_bytes[4] == 'P' && recieve_bytes[5] == 'A' && recieve_bytes[6] == 'S' && recieve_bytes[7] == 'S')
                                {
                                    Invoke(new EventHandler(delegate
                                    {
                                        textBox_接地導通量測值.ForeColor = Color.Lime;
                                        textBox_接地導通量測值.Text =
                                        char.ConvertFromUtf32(recieve_bytes[18]) + char.ConvertFromUtf32(recieve_bytes[19]) + char.ConvertFromUtf32(recieve_bytes[20]) + char.ConvertFromUtf32(recieve_bytes[21])
                                        + char.ConvertFromUtf32(recieve_bytes[22]) + char.ConvertFromUtf32(recieve_bytes[23]) + char.ConvertFromUtf32(recieve_bytes[24]) + char.ConvertFromUtf32(recieve_bytes[25])
                                        + char.ConvertFromUtf32(recieve_bytes[26]);

                                    }));
                                    this.CONT檢測結果 = true;
                                }
                                else
                                {
                                    Invoke(new EventHandler(delegate
                                    {
                                        textBox_接地導通量測值.ForeColor = Color.Red;
                                        textBox_接地導通量測值.Text =
                                        char.ConvertFromUtf32(recieve_bytes[18]) + char.ConvertFromUtf32(recieve_bytes[19]) + char.ConvertFromUtf32(recieve_bytes[20]) + char.ConvertFromUtf32(recieve_bytes[21])
                                        + char.ConvertFromUtf32(recieve_bytes[22]) + char.ConvertFromUtf32(recieve_bytes[23]) + char.ConvertFromUtf32(recieve_bytes[24]) + char.ConvertFromUtf32(recieve_bytes[25])
                                        + char.ConvertFromUtf32(recieve_bytes[26]);

                                    }));
                                    this.CONT檢測結果 = false;

                                }
                                //if (ACW檢測結果 && DCW檢測結果 && IR檢測結果 && CONT檢測結果) this.GPT_12003檢測結果.Bool = true;


                                this.GPT_12003測試完成.Bool = true;
                                this.GPT_12003_SerialPort.ClearReadByte();
                                break;
                            }
                        }
                    }

                }


            }


        }


    }
}
