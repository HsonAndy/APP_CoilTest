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
        public Form1()
        {
            InitializeComponent();
        }

        MyThread MyThread_儀器通訊;

        private void Form1_Load(object sender, EventArgs e)
        {
            plC_UI_Init1.Run(this.FindForm(), lowerMachine_Panel1);
            this.timer_mythread.Enabled = true;
            MyMessageBox.form = this.FindForm();

        }

        private void timer_mythread_Tick(object sender, EventArgs e)
        {
            if(this.plC_UI_Init1.Init_Finish)
            {
                MyThread_儀器通訊 = new MyThread();
                this.MyThread_儀器通訊.Add_Method(GPT_12003_ComPort通訊);
                this.MyThread_儀器通訊.Add_Method(GOM_804_ComPort通訊);
                this.MyThread_儀器通訊.Add_Method(ASR_2100_ComPort通訊);
                this.MyThread_儀器通訊.Add_Method(IWT_5000A_ComPort通訊);
                this.MyThread_儀器通訊.Add_Method(sub_PLC_Device_ROW輸出);
                this.MyThread_儀器通訊.SetSleepTime(2);
                this.MyThread_儀器通訊.AutoRun(true);
                this.MyThread_儀器通訊.AutoStop(true);
                this.MyThread_儀器通訊.Trigger();

                this.GOM_804_Init("COM4", 115200);
                this.IWT_5000A_Init("COM6", 57600);
                this.GPT_12003_Init("COM5", 115200);
                this.ASR_2100_Init("COM7", 115200);
                this.Program_表單下載_Init();
                this.timer_mythread.Enabled = false;
                this.WindowState = FormWindowState.Maximized;
            }

        }
        #region 視窗關閉輸出程式
        List<PLC_Device> PLC_輸出 = new List<PLC_Device>();
        List<PLC_Device> PLC_輔助輸出 = new List<PLC_Device>();
        List<PLC_Device> PLC_按鈕 = new List<PLC_Device>();
        List<PLC_Device> PLC_按鈕按下 = new List<PLC_Device>();
        #region Y0~Y77
        PLC_Device PLC_輸出_Y00 = new PLC_Device("Y00");
        PLC_Device PLC_輸出_Y01 = new PLC_Device("Y01");
        PLC_Device PLC_輸出_Y02 = new PLC_Device("Y02");
        PLC_Device PLC_輸出_Y03 = new PLC_Device("Y03");
        PLC_Device PLC_輸出_Y04 = new PLC_Device("Y04");
        PLC_Device PLC_輸出_Y05 = new PLC_Device("Y05");
        PLC_Device PLC_輸出_Y06 = new PLC_Device("Y06");
        PLC_Device PLC_輸出_Y07 = new PLC_Device("Y07");

        PLC_Device PLC_輸出_Y10 = new PLC_Device("Y10");
        PLC_Device PLC_輸出_Y11 = new PLC_Device("Y11");
        PLC_Device PLC_輸出_Y12 = new PLC_Device("Y12");
        PLC_Device PLC_輸出_Y13 = new PLC_Device("Y13");
        PLC_Device PLC_輸出_Y14 = new PLC_Device("Y14");
        PLC_Device PLC_輸出_Y15 = new PLC_Device("Y15");
        PLC_Device PLC_輸出_Y16 = new PLC_Device("Y16");
        PLC_Device PLC_輸出_Y17 = new PLC_Device("Y17");

        PLC_Device PLC_輸出_Y20 = new PLC_Device("Y20");
        PLC_Device PLC_輸出_Y21 = new PLC_Device("Y21");
        PLC_Device PLC_輸出_Y22 = new PLC_Device("Y22");
        PLC_Device PLC_輸出_Y23 = new PLC_Device("Y23");
        PLC_Device PLC_輸出_Y24 = new PLC_Device("Y24");
        PLC_Device PLC_輸出_Y25 = new PLC_Device("Y25");
        PLC_Device PLC_輸出_Y26 = new PLC_Device("Y26");
        PLC_Device PLC_輸出_Y27 = new PLC_Device("Y27");

        PLC_Device PLC_輸出_Y30 = new PLC_Device("Y30");
        PLC_Device PLC_輸出_Y31 = new PLC_Device("Y31");
        PLC_Device PLC_輸出_Y32 = new PLC_Device("Y32");
        PLC_Device PLC_輸出_Y33 = new PLC_Device("Y33");
        PLC_Device PLC_輸出_Y34 = new PLC_Device("Y34");
        PLC_Device PLC_輸出_Y35 = new PLC_Device("Y35");
        PLC_Device PLC_輸出_Y36 = new PLC_Device("Y36");
        PLC_Device PLC_輸出_Y37 = new PLC_Device("Y37");

        PLC_Device PLC_輸出_Y40 = new PLC_Device("Y40");
        PLC_Device PLC_輸出_Y41 = new PLC_Device("Y41");
        PLC_Device PLC_輸出_Y42 = new PLC_Device("Y42");
        PLC_Device PLC_輸出_Y43 = new PLC_Device("Y43");
        PLC_Device PLC_輸出_Y44 = new PLC_Device("Y44");
        PLC_Device PLC_輸出_Y45 = new PLC_Device("Y45");
        PLC_Device PLC_輸出_Y46 = new PLC_Device("Y46");
        PLC_Device PLC_輸出_Y47 = new PLC_Device("Y47");

        PLC_Device PLC_輸出_Y50 = new PLC_Device("Y50");
        PLC_Device PLC_輸出_Y51 = new PLC_Device("Y51");
        PLC_Device PLC_輸出_Y52 = new PLC_Device("Y52");
        PLC_Device PLC_輸出_Y53 = new PLC_Device("Y53");
        PLC_Device PLC_輸出_Y54 = new PLC_Device("Y54");
        PLC_Device PLC_輸出_Y55 = new PLC_Device("Y55");
        PLC_Device PLC_輸出_Y56 = new PLC_Device("Y56");
        PLC_Device PLC_輸出_Y57 = new PLC_Device("Y57");

        PLC_Device PLC_輸出_Y60 = new PLC_Device("Y60");
        PLC_Device PLC_輸出_Y61 = new PLC_Device("Y61");
        PLC_Device PLC_輸出_Y62 = new PLC_Device("Y62");
        PLC_Device PLC_輸出_Y63 = new PLC_Device("Y63");
        PLC_Device PLC_輸出_Y64 = new PLC_Device("Y64");
        PLC_Device PLC_輸出_Y65 = new PLC_Device("Y65");
        PLC_Device PLC_輸出_Y66 = new PLC_Device("Y66");
        PLC_Device PLC_輸出_Y67 = new PLC_Device("Y67");

        PLC_Device PLC_輸出_Y70 = new PLC_Device("Y70");
        PLC_Device PLC_輸出_Y71 = new PLC_Device("Y71");
        PLC_Device PLC_輸出_Y72 = new PLC_Device("Y72");
        PLC_Device PLC_輸出_Y73 = new PLC_Device("Y73");
        PLC_Device PLC_輸出_Y74 = new PLC_Device("Y74");
        PLC_Device PLC_輸出_Y75 = new PLC_Device("Y75");
        PLC_Device PLC_輸出_Y76 = new PLC_Device("Y76");
        PLC_Device PLC_輸出_Y77 = new PLC_Device("Y77");
        #endregion
        #region S0~S77
        PLC_Device PLC_按鈕_S00 = new PLC_Device("S00");
        PLC_Device PLC_按鈕_S01 = new PLC_Device("S01");
        PLC_Device PLC_按鈕_S02 = new PLC_Device("S02");
        PLC_Device PLC_按鈕_S03 = new PLC_Device("S03");
        PLC_Device PLC_按鈕_S04 = new PLC_Device("S04");
        PLC_Device PLC_按鈕_S05 = new PLC_Device("S05");
        PLC_Device PLC_按鈕_S06 = new PLC_Device("S06");
        PLC_Device PLC_按鈕_S07 = new PLC_Device("S07");

        PLC_Device PLC_按鈕_S10 = new PLC_Device("S10");
        PLC_Device PLC_按鈕_S11 = new PLC_Device("S11");
        PLC_Device PLC_按鈕_S12 = new PLC_Device("S12");
        PLC_Device PLC_按鈕_S13 = new PLC_Device("S13");
        PLC_Device PLC_按鈕_S14 = new PLC_Device("S14");
        PLC_Device PLC_按鈕_S15 = new PLC_Device("S15");
        PLC_Device PLC_按鈕_S16 = new PLC_Device("S16");
        PLC_Device PLC_按鈕_S17 = new PLC_Device("S17");

        PLC_Device PLC_按鈕_S20 = new PLC_Device("S20");
        PLC_Device PLC_按鈕_S21 = new PLC_Device("S21");
        PLC_Device PLC_按鈕_S22 = new PLC_Device("S22");
        PLC_Device PLC_按鈕_S23 = new PLC_Device("S23");
        PLC_Device PLC_按鈕_S24 = new PLC_Device("S24");
        PLC_Device PLC_按鈕_S25 = new PLC_Device("S25");
        PLC_Device PLC_按鈕_S26 = new PLC_Device("S26");
        PLC_Device PLC_按鈕_S27 = new PLC_Device("S27");

        PLC_Device PLC_按鈕_S30 = new PLC_Device("S30");
        PLC_Device PLC_按鈕_S31 = new PLC_Device("S31");
        PLC_Device PLC_按鈕_S32 = new PLC_Device("S32");
        PLC_Device PLC_按鈕_S33 = new PLC_Device("S33");
        PLC_Device PLC_按鈕_S34 = new PLC_Device("S34");
        PLC_Device PLC_按鈕_S35 = new PLC_Device("S35");
        PLC_Device PLC_按鈕_S36 = new PLC_Device("S36");
        PLC_Device PLC_按鈕_S37 = new PLC_Device("S37");

        PLC_Device PLC_按鈕_S40 = new PLC_Device("S40");
        PLC_Device PLC_按鈕_S41 = new PLC_Device("S41");
        PLC_Device PLC_按鈕_S42 = new PLC_Device("S42");
        PLC_Device PLC_按鈕_S43 = new PLC_Device("S43");
        PLC_Device PLC_按鈕_S44 = new PLC_Device("S44");
        PLC_Device PLC_按鈕_S45 = new PLC_Device("S45");
        PLC_Device PLC_按鈕_S46 = new PLC_Device("S46");
        PLC_Device PLC_按鈕_S47 = new PLC_Device("S47");

        PLC_Device PLC_按鈕_S50 = new PLC_Device("S50");
        PLC_Device PLC_按鈕_S51 = new PLC_Device("S51");
        PLC_Device PLC_按鈕_S52 = new PLC_Device("S52");
        PLC_Device PLC_按鈕_S53 = new PLC_Device("S53");
        PLC_Device PLC_按鈕_S54 = new PLC_Device("S54");
        PLC_Device PLC_按鈕_S55 = new PLC_Device("S55");
        PLC_Device PLC_按鈕_S56 = new PLC_Device("S56");
        PLC_Device PLC_按鈕_S57 = new PLC_Device("S57");

        PLC_Device PLC_按鈕_S60 = new PLC_Device("S60");
        PLC_Device PLC_按鈕_S61 = new PLC_Device("S61");
        PLC_Device PLC_按鈕_S62 = new PLC_Device("S62");
        PLC_Device PLC_按鈕_S63 = new PLC_Device("S63");
        PLC_Device PLC_按鈕_S64 = new PLC_Device("S64");
        PLC_Device PLC_按鈕_S65 = new PLC_Device("S65");
        PLC_Device PLC_按鈕_S66 = new PLC_Device("S66");
        PLC_Device PLC_按鈕_S67 = new PLC_Device("S67");

        PLC_Device PLC_按鈕_S70 = new PLC_Device("S70");
        PLC_Device PLC_按鈕_S71 = new PLC_Device("S71");
        PLC_Device PLC_按鈕_S72 = new PLC_Device("S72");
        PLC_Device PLC_按鈕_S73 = new PLC_Device("S73");
        PLC_Device PLC_按鈕_S74 = new PLC_Device("S74");
        PLC_Device PLC_按鈕_S75 = new PLC_Device("S75");
        PLC_Device PLC_按鈕_S76 = new PLC_Device("S76");
        PLC_Device PLC_按鈕_S77 = new PLC_Device("S77");
        #endregion
        #region M0~M77
        PLC_Device PLC_輔助輸出_M00 = new PLC_Device("M00");
        PLC_Device PLC_輔助輸出_M01 = new PLC_Device("M01");
        PLC_Device PLC_輔助輸出_M02 = new PLC_Device("M02");
        PLC_Device PLC_輔助輸出_M03 = new PLC_Device("M03");
        PLC_Device PLC_輔助輸出_M04 = new PLC_Device("M04");
        PLC_Device PLC_輔助輸出_M05 = new PLC_Device("M05");
        PLC_Device PLC_輔助輸出_M06 = new PLC_Device("M06");
        PLC_Device PLC_輔助輸出_M07 = new PLC_Device("M07");

        PLC_Device PLC_輔助輸出_M10 = new PLC_Device("M10");
        PLC_Device PLC_輔助輸出_M11 = new PLC_Device("M11");
        PLC_Device PLC_輔助輸出_M12 = new PLC_Device("M12");
        PLC_Device PLC_輔助輸出_M13 = new PLC_Device("M13");
        PLC_Device PLC_輔助輸出_M14 = new PLC_Device("M14");
        PLC_Device PLC_輔助輸出_M15 = new PLC_Device("M15");
        PLC_Device PLC_輔助輸出_M16 = new PLC_Device("M16");
        PLC_Device PLC_輔助輸出_M17 = new PLC_Device("M17");

        PLC_Device PLC_輔助輸出_M20 = new PLC_Device("M20");
        PLC_Device PLC_輔助輸出_M21 = new PLC_Device("M21");
        PLC_Device PLC_輔助輸出_M22 = new PLC_Device("M22");
        PLC_Device PLC_輔助輸出_M23 = new PLC_Device("M23");
        PLC_Device PLC_輔助輸出_M24 = new PLC_Device("M24");
        PLC_Device PLC_輔助輸出_M25 = new PLC_Device("M25");
        PLC_Device PLC_輔助輸出_M26 = new PLC_Device("M26");
        PLC_Device PLC_輔助輸出_M27 = new PLC_Device("M27");

        PLC_Device PLC_輔助輸出_M30 = new PLC_Device("M30");
        PLC_Device PLC_輔助輸出_M31 = new PLC_Device("M31");
        PLC_Device PLC_輔助輸出_M32 = new PLC_Device("M32");
        PLC_Device PLC_輔助輸出_M33 = new PLC_Device("M33");
        PLC_Device PLC_輔助輸出_M34 = new PLC_Device("M34");
        PLC_Device PLC_輔助輸出_M35 = new PLC_Device("M35");
        PLC_Device PLC_輔助輸出_M36 = new PLC_Device("M36");
        PLC_Device PLC_輔助輸出_M37 = new PLC_Device("M37");

        PLC_Device PLC_輔助輸出_M40 = new PLC_Device("M40");
        PLC_Device PLC_輔助輸出_M41 = new PLC_Device("M41");
        PLC_Device PLC_輔助輸出_M42 = new PLC_Device("M42");
        PLC_Device PLC_輔助輸出_M43 = new PLC_Device("M43");
        PLC_Device PLC_輔助輸出_M44 = new PLC_Device("M44");
        PLC_Device PLC_輔助輸出_M45 = new PLC_Device("M45");
        PLC_Device PLC_輔助輸出_M46 = new PLC_Device("M46");
        PLC_Device PLC_輔助輸出_M47 = new PLC_Device("M47");

        PLC_Device PLC_輔助輸出_M50 = new PLC_Device("M50");
        PLC_Device PLC_輔助輸出_M51 = new PLC_Device("M51");
        PLC_Device PLC_輔助輸出_M52 = new PLC_Device("M52");
        PLC_Device PLC_輔助輸出_M53 = new PLC_Device("M53");
        PLC_Device PLC_輔助輸出_M54 = new PLC_Device("M54");
        PLC_Device PLC_輔助輸出_M55 = new PLC_Device("M55");
        PLC_Device PLC_輔助輸出_M56 = new PLC_Device("M56");
        PLC_Device PLC_輔助輸出_M57 = new PLC_Device("M57");

        PLC_Device PLC_輔助輸出_M60 = new PLC_Device("M60");
        PLC_Device PLC_輔助輸出_M61 = new PLC_Device("M61");
        PLC_Device PLC_輔助輸出_M62 = new PLC_Device("M62");
        PLC_Device PLC_輔助輸出_M63 = new PLC_Device("M63");
        PLC_Device PLC_輔助輸出_M64 = new PLC_Device("M64");
        PLC_Device PLC_輔助輸出_M65 = new PLC_Device("M65");
        PLC_Device PLC_輔助輸出_M66 = new PLC_Device("M66");
        PLC_Device PLC_輔助輸出_M67 = new PLC_Device("M67");

        PLC_Device PLC_輔助輸出_M70 = new PLC_Device("M70");
        PLC_Device PLC_輔助輸出_M71 = new PLC_Device("M71");
        PLC_Device PLC_輔助輸出_M72 = new PLC_Device("M72");
        PLC_Device PLC_輔助輸出_M73 = new PLC_Device("M73");
        PLC_Device PLC_輔助輸出_M74 = new PLC_Device("M74");
        PLC_Device PLC_輔助輸出_M75 = new PLC_Device("M75");
        PLC_Device PLC_輔助輸出_M76 = new PLC_Device("M76");
        PLC_Device PLC_輔助輸出_M77 = new PLC_Device("M77");
        #endregion
        #region S1000~S1315
        PLC_Device PLC_按鈕按下_S1000 = new PLC_Device("S1000");
        PLC_Device PLC_按鈕按下_S1005 = new PLC_Device("S1005");
        PLC_Device PLC_按鈕按下_S1010 = new PLC_Device("S1010");
        PLC_Device PLC_按鈕按下_S1015 = new PLC_Device("S1015");
        PLC_Device PLC_按鈕按下_S1020 = new PLC_Device("S1020");
        PLC_Device PLC_按鈕按下_S1025 = new PLC_Device("S1025");
        PLC_Device PLC_按鈕按下_S1030 = new PLC_Device("S1030");
        PLC_Device PLC_按鈕按下_S1035 = new PLC_Device("S1035");

        PLC_Device PLC_按鈕按下_S1040 = new PLC_Device("S1040");
        PLC_Device PLC_按鈕按下_S1045 = new PLC_Device("S1045");
        PLC_Device PLC_按鈕按下_S1050 = new PLC_Device("S1050");
        PLC_Device PLC_按鈕按下_S1055 = new PLC_Device("S1055");
        PLC_Device PLC_按鈕按下_S1060 = new PLC_Device("S1060");
        PLC_Device PLC_按鈕按下_S1065 = new PLC_Device("S1065");
        PLC_Device PLC_按鈕按下_S1070 = new PLC_Device("S1070");
        PLC_Device PLC_按鈕按下_S1075 = new PLC_Device("S1075");

        PLC_Device PLC_按鈕按下_S1080 = new PLC_Device("S1080");
        PLC_Device PLC_按鈕按下_S1085 = new PLC_Device("S1085");
        PLC_Device PLC_按鈕按下_S1090 = new PLC_Device("S1090");
        PLC_Device PLC_按鈕按下_S1095 = new PLC_Device("S1095");
        PLC_Device PLC_按鈕按下_S1100 = new PLC_Device("S1100");
        PLC_Device PLC_按鈕按下_S1105 = new PLC_Device("S1105");
        PLC_Device PLC_按鈕按下_S1110 = new PLC_Device("S1110");
        PLC_Device PLC_按鈕按下_S1115 = new PLC_Device("S1115");

        PLC_Device PLC_按鈕按下_S1120 = new PLC_Device("S1120");
        PLC_Device PLC_按鈕按下_S1125 = new PLC_Device("S1125");
        PLC_Device PLC_按鈕按下_S1130 = new PLC_Device("S1130");
        PLC_Device PLC_按鈕按下_S1135 = new PLC_Device("S1135");
        PLC_Device PLC_按鈕按下_S1140 = new PLC_Device("S1140");
        PLC_Device PLC_按鈕按下_S1145 = new PLC_Device("S1145");
        PLC_Device PLC_按鈕按下_S1150 = new PLC_Device("S1150");
        PLC_Device PLC_按鈕按下_S1155 = new PLC_Device("S1155");

        PLC_Device PLC_按鈕按下_S1160 = new PLC_Device("S1160");
        PLC_Device PLC_按鈕按下_S1165 = new PLC_Device("S1165");
        PLC_Device PLC_按鈕按下_S1170 = new PLC_Device("S1170");
        PLC_Device PLC_按鈕按下_S1175 = new PLC_Device("S1175");
        PLC_Device PLC_按鈕按下_S1180 = new PLC_Device("S1180");
        PLC_Device PLC_按鈕按下_S1185 = new PLC_Device("S1185");
        PLC_Device PLC_按鈕按下_S1190 = new PLC_Device("S1190");
        PLC_Device PLC_按鈕按下_S1195 = new PLC_Device("S1195");

        PLC_Device PLC_按鈕按下_S1200 = new PLC_Device("S1200");
        PLC_Device PLC_按鈕按下_S1205 = new PLC_Device("S1205");
        PLC_Device PLC_按鈕按下_S1210 = new PLC_Device("S1210");
        PLC_Device PLC_按鈕按下_S1215 = new PLC_Device("S1215");
        PLC_Device PLC_按鈕按下_S1220 = new PLC_Device("S1220");
        PLC_Device PLC_按鈕按下_S1225 = new PLC_Device("S1225");
        PLC_Device PLC_按鈕按下_S1230 = new PLC_Device("S1230");
        PLC_Device PLC_按鈕按下_S1235 = new PLC_Device("S1235");

        PLC_Device PLC_按鈕按下_S1240 = new PLC_Device("S1240");
        PLC_Device PLC_按鈕按下_S1245 = new PLC_Device("S1245");
        PLC_Device PLC_按鈕按下_S1250 = new PLC_Device("S1250");
        PLC_Device PLC_按鈕按下_S1255 = new PLC_Device("S1255");
        PLC_Device PLC_按鈕按下_S1260 = new PLC_Device("S1260");
        PLC_Device PLC_按鈕按下_S1265 = new PLC_Device("S1265");
        PLC_Device PLC_按鈕按下_S1270 = new PLC_Device("S1270");
        PLC_Device PLC_按鈕按下_S1275 = new PLC_Device("S1275");

        PLC_Device PLC_按鈕按下_S1280 = new PLC_Device("S1280");
        PLC_Device PLC_按鈕按下_S1285 = new PLC_Device("S1285");
        PLC_Device PLC_按鈕按下_S1290 = new PLC_Device("S1290");
        PLC_Device PLC_按鈕按下_S1295 = new PLC_Device("S1295");
        PLC_Device PLC_按鈕按下_S1300 = new PLC_Device("S1300");
        PLC_Device PLC_按鈕按下_S1305 = new PLC_Device("S1305");
        PLC_Device PLC_按鈕按下_S1310 = new PLC_Device("S1310");
        PLC_Device PLC_按鈕按下_S1315 = new PLC_Device("S1315");
        #endregion
        PLC_Device PLC_初始化BIT_S8000 = new PLC_Device("S8000");
        bool flag_OutputClose = false;

        private void Form1_FormClosing_1(object sender, FormClosingEventArgs e)
        {
            #region Y0~Y77
            PLC_輸出.Add(PLC_輸出_Y00);
            PLC_輸出.Add(PLC_輸出_Y01);
            PLC_輸出.Add(PLC_輸出_Y02);
            PLC_輸出.Add(PLC_輸出_Y03);
            PLC_輸出.Add(PLC_輸出_Y04);
            PLC_輸出.Add(PLC_輸出_Y05);
            PLC_輸出.Add(PLC_輸出_Y06);
            PLC_輸出.Add(PLC_輸出_Y07);

            PLC_輸出.Add(PLC_輸出_Y10);
            PLC_輸出.Add(PLC_輸出_Y11);
            PLC_輸出.Add(PLC_輸出_Y12);
            PLC_輸出.Add(PLC_輸出_Y13);
            PLC_輸出.Add(PLC_輸出_Y14);
            PLC_輸出.Add(PLC_輸出_Y15);
            PLC_輸出.Add(PLC_輸出_Y16);
            PLC_輸出.Add(PLC_輸出_Y17);

            PLC_輸出.Add(PLC_輸出_Y20);
            PLC_輸出.Add(PLC_輸出_Y21);
            PLC_輸出.Add(PLC_輸出_Y22);
            PLC_輸出.Add(PLC_輸出_Y23);
            PLC_輸出.Add(PLC_輸出_Y24);
            PLC_輸出.Add(PLC_輸出_Y25);
            PLC_輸出.Add(PLC_輸出_Y26);
            PLC_輸出.Add(PLC_輸出_Y27);

            PLC_輸出.Add(PLC_輸出_Y30);
            PLC_輸出.Add(PLC_輸出_Y31);
            PLC_輸出.Add(PLC_輸出_Y32);
            PLC_輸出.Add(PLC_輸出_Y33);
            PLC_輸出.Add(PLC_輸出_Y34);
            PLC_輸出.Add(PLC_輸出_Y35);
            PLC_輸出.Add(PLC_輸出_Y36);
            PLC_輸出.Add(PLC_輸出_Y37);

            PLC_輸出.Add(PLC_輸出_Y40);
            PLC_輸出.Add(PLC_輸出_Y41);
            PLC_輸出.Add(PLC_輸出_Y42);
            PLC_輸出.Add(PLC_輸出_Y43);
            PLC_輸出.Add(PLC_輸出_Y44);
            PLC_輸出.Add(PLC_輸出_Y45);
            PLC_輸出.Add(PLC_輸出_Y46);
            PLC_輸出.Add(PLC_輸出_Y47);

            PLC_輸出.Add(PLC_輸出_Y50);
            PLC_輸出.Add(PLC_輸出_Y51);
            PLC_輸出.Add(PLC_輸出_Y52);
            PLC_輸出.Add(PLC_輸出_Y53);
            PLC_輸出.Add(PLC_輸出_Y54);
            PLC_輸出.Add(PLC_輸出_Y55);
            PLC_輸出.Add(PLC_輸出_Y56);
            PLC_輸出.Add(PLC_輸出_Y57);

            PLC_輸出.Add(PLC_輸出_Y60);
            PLC_輸出.Add(PLC_輸出_Y61);
            PLC_輸出.Add(PLC_輸出_Y62);
            PLC_輸出.Add(PLC_輸出_Y63);
            PLC_輸出.Add(PLC_輸出_Y64);
            PLC_輸出.Add(PLC_輸出_Y65);
            PLC_輸出.Add(PLC_輸出_Y66);
            PLC_輸出.Add(PLC_輸出_Y67);

            PLC_輸出.Add(PLC_輸出_Y70);
            PLC_輸出.Add(PLC_輸出_Y71);
            PLC_輸出.Add(PLC_輸出_Y72);
            PLC_輸出.Add(PLC_輸出_Y73);
            PLC_輸出.Add(PLC_輸出_Y74);
            PLC_輸出.Add(PLC_輸出_Y75);
            PLC_輸出.Add(PLC_輸出_Y76);
            PLC_輸出.Add(PLC_輸出_Y77);
            #endregion
            #region M0~M77
            PLC_輔助輸出.Add(PLC_輔助輸出_M00);
            PLC_輔助輸出.Add(PLC_輔助輸出_M01);
            PLC_輔助輸出.Add(PLC_輔助輸出_M02);
            PLC_輔助輸出.Add(PLC_輔助輸出_M03);
            PLC_輔助輸出.Add(PLC_輔助輸出_M04);
            PLC_輔助輸出.Add(PLC_輔助輸出_M05);
            PLC_輔助輸出.Add(PLC_輔助輸出_M06);
            PLC_輔助輸出.Add(PLC_輔助輸出_M07);

            PLC_輔助輸出.Add(PLC_輔助輸出_M10);
            PLC_輔助輸出.Add(PLC_輔助輸出_M11);
            PLC_輔助輸出.Add(PLC_輔助輸出_M12);
            PLC_輔助輸出.Add(PLC_輔助輸出_M13);
            PLC_輔助輸出.Add(PLC_輔助輸出_M14);
            PLC_輔助輸出.Add(PLC_輔助輸出_M15);
            PLC_輔助輸出.Add(PLC_輔助輸出_M16);
            PLC_輔助輸出.Add(PLC_輔助輸出_M17);

            PLC_輔助輸出.Add(PLC_輔助輸出_M20);
            PLC_輔助輸出.Add(PLC_輔助輸出_M21);
            PLC_輔助輸出.Add(PLC_輔助輸出_M22);
            PLC_輔助輸出.Add(PLC_輔助輸出_M23);
            PLC_輔助輸出.Add(PLC_輔助輸出_M24);
            PLC_輔助輸出.Add(PLC_輔助輸出_M25);
            PLC_輔助輸出.Add(PLC_輔助輸出_M26);
            PLC_輔助輸出.Add(PLC_輔助輸出_M27);

            PLC_輔助輸出.Add(PLC_輔助輸出_M30);
            PLC_輔助輸出.Add(PLC_輔助輸出_M31);
            PLC_輔助輸出.Add(PLC_輔助輸出_M32);
            PLC_輔助輸出.Add(PLC_輔助輸出_M33);
            PLC_輔助輸出.Add(PLC_輔助輸出_M34);
            PLC_輔助輸出.Add(PLC_輔助輸出_M35);
            PLC_輔助輸出.Add(PLC_輔助輸出_M36);
            PLC_輔助輸出.Add(PLC_輔助輸出_M37);

            PLC_輔助輸出.Add(PLC_輔助輸出_M40);
            PLC_輔助輸出.Add(PLC_輔助輸出_M41);
            PLC_輔助輸出.Add(PLC_輔助輸出_M42);
            PLC_輔助輸出.Add(PLC_輔助輸出_M43);
            PLC_輔助輸出.Add(PLC_輔助輸出_M44);
            PLC_輔助輸出.Add(PLC_輔助輸出_M45);
            PLC_輔助輸出.Add(PLC_輔助輸出_M46);
            PLC_輔助輸出.Add(PLC_輔助輸出_M47);

            PLC_輔助輸出.Add(PLC_輔助輸出_M50);
            PLC_輔助輸出.Add(PLC_輔助輸出_M51);
            PLC_輔助輸出.Add(PLC_輔助輸出_M52);
            PLC_輔助輸出.Add(PLC_輔助輸出_M53);
            PLC_輔助輸出.Add(PLC_輔助輸出_M54);
            PLC_輔助輸出.Add(PLC_輔助輸出_M55);
            PLC_輔助輸出.Add(PLC_輔助輸出_M56);
            PLC_輔助輸出.Add(PLC_輔助輸出_M57);

            PLC_輔助輸出.Add(PLC_輔助輸出_M60);
            PLC_輔助輸出.Add(PLC_輔助輸出_M61);
            PLC_輔助輸出.Add(PLC_輔助輸出_M62);
            PLC_輔助輸出.Add(PLC_輔助輸出_M63);
            PLC_輔助輸出.Add(PLC_輔助輸出_M64);
            PLC_輔助輸出.Add(PLC_輔助輸出_M65);
            PLC_輔助輸出.Add(PLC_輔助輸出_M66);
            PLC_輔助輸出.Add(PLC_輔助輸出_M67);

            PLC_輔助輸出.Add(PLC_輔助輸出_M70);
            PLC_輔助輸出.Add(PLC_輔助輸出_M71);
            PLC_輔助輸出.Add(PLC_輔助輸出_M72);
            PLC_輔助輸出.Add(PLC_輔助輸出_M73);
            PLC_輔助輸出.Add(PLC_輔助輸出_M74);
            PLC_輔助輸出.Add(PLC_輔助輸出_M75);
            PLC_輔助輸出.Add(PLC_輔助輸出_M76);
            PLC_輔助輸出.Add(PLC_輔助輸出_M77);
            #endregion
            #region S0~S77
            PLC_按鈕.Add(PLC_按鈕_S00);
            PLC_按鈕.Add(PLC_按鈕_S01);
            PLC_按鈕.Add(PLC_按鈕_S02);
            PLC_按鈕.Add(PLC_按鈕_S03);
            PLC_按鈕.Add(PLC_按鈕_S04);
            PLC_按鈕.Add(PLC_按鈕_S05);
            PLC_按鈕.Add(PLC_按鈕_S06);
            PLC_按鈕.Add(PLC_按鈕_S07);

            PLC_按鈕.Add(PLC_按鈕_S10);
            PLC_按鈕.Add(PLC_按鈕_S11);
            PLC_按鈕.Add(PLC_按鈕_S12);
            PLC_按鈕.Add(PLC_按鈕_S13);
            PLC_按鈕.Add(PLC_按鈕_S14);
            PLC_按鈕.Add(PLC_按鈕_S15);
            PLC_按鈕.Add(PLC_按鈕_S16);
            PLC_按鈕.Add(PLC_按鈕_S17);

            PLC_按鈕.Add(PLC_按鈕_S20);
            PLC_按鈕.Add(PLC_按鈕_S21);
            PLC_按鈕.Add(PLC_按鈕_S22);
            PLC_按鈕.Add(PLC_按鈕_S23);
            PLC_按鈕.Add(PLC_按鈕_S24);
            PLC_按鈕.Add(PLC_按鈕_S25);
            PLC_按鈕.Add(PLC_按鈕_S26);
            PLC_按鈕.Add(PLC_按鈕_S27);

            PLC_按鈕.Add(PLC_按鈕_S30);
            PLC_按鈕.Add(PLC_按鈕_S31);
            PLC_按鈕.Add(PLC_按鈕_S32);
            PLC_按鈕.Add(PLC_按鈕_S33);
            PLC_按鈕.Add(PLC_按鈕_S34);
            PLC_按鈕.Add(PLC_按鈕_S35);
            PLC_按鈕.Add(PLC_按鈕_S36);
            PLC_按鈕.Add(PLC_按鈕_S37);

            PLC_按鈕.Add(PLC_按鈕_S40);
            PLC_按鈕.Add(PLC_按鈕_S41);
            PLC_按鈕.Add(PLC_按鈕_S42);
            PLC_按鈕.Add(PLC_按鈕_S43);
            PLC_按鈕.Add(PLC_按鈕_S44);
            PLC_按鈕.Add(PLC_按鈕_S45);
            PLC_按鈕.Add(PLC_按鈕_S46);
            PLC_按鈕.Add(PLC_按鈕_S47);

            PLC_按鈕.Add(PLC_按鈕_S50);
            PLC_按鈕.Add(PLC_按鈕_S51);
            PLC_按鈕.Add(PLC_按鈕_S52);
            PLC_按鈕.Add(PLC_按鈕_S53);
            PLC_按鈕.Add(PLC_按鈕_S54);
            PLC_按鈕.Add(PLC_按鈕_S55);
            PLC_按鈕.Add(PLC_按鈕_S56);
            PLC_按鈕.Add(PLC_按鈕_S57);

            PLC_按鈕.Add(PLC_按鈕_S60);
            PLC_按鈕.Add(PLC_按鈕_S61);
            PLC_按鈕.Add(PLC_按鈕_S62);
            PLC_按鈕.Add(PLC_按鈕_S63);
            PLC_按鈕.Add(PLC_按鈕_S64);
            PLC_按鈕.Add(PLC_按鈕_S65);
            PLC_按鈕.Add(PLC_按鈕_S66);
            PLC_按鈕.Add(PLC_按鈕_S67);

            PLC_按鈕.Add(PLC_按鈕_S70);
            PLC_按鈕.Add(PLC_按鈕_S71);
            PLC_按鈕.Add(PLC_按鈕_S72);
            PLC_按鈕.Add(PLC_按鈕_S73);
            PLC_按鈕.Add(PLC_按鈕_S74);
            PLC_按鈕.Add(PLC_按鈕_S75);
            PLC_按鈕.Add(PLC_按鈕_S76);
            PLC_按鈕.Add(PLC_按鈕_S77);
            #endregion
            #region S1000~S1315
            PLC_按鈕按下.Add(PLC_按鈕按下_S1000);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1005);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1010);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1015);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1020);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1025);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1030);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1035);

            PLC_按鈕按下.Add(PLC_按鈕按下_S1040);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1045);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1050);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1055);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1060);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1065);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1070);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1075);

            PLC_按鈕按下.Add(PLC_按鈕按下_S1080);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1085);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1090);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1095);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1100);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1105);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1110);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1115);

            PLC_按鈕按下.Add(PLC_按鈕按下_S1120);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1125);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1130);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1135);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1140);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1145);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1150);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1155);

            PLC_按鈕按下.Add(PLC_按鈕按下_S1160);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1165);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1170);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1175);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1180);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1185);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1190);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1195);

            PLC_按鈕按下.Add(PLC_按鈕按下_S1200);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1205);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1210);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1215);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1220);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1225);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1230);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1235);

            PLC_按鈕按下.Add(PLC_按鈕按下_S1240);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1245);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1250);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1255);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1260);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1265);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1270);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1275);

            PLC_按鈕按下.Add(PLC_按鈕按下_S1280);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1285);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1290);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1295);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1300);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1305);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1310);
            PLC_按鈕按下.Add(PLC_按鈕按下_S1315);


            #endregion
            while (true)
            {
                for (int i = 0; i < 64; i++)
                {
                    PLC_輸出[i].Bool = false;
                    PLC_按鈕[i].Bool = false;
                    PLC_按鈕按下[i].Bool = false;
                    PLC_初始化BIT_S8000.Bool = false;
                    if (PLC_輸出[i].Bool || PLC_按鈕[i].Bool || PLC_按鈕按下[i].Bool || PLC_輔助輸出[i].Bool || PLC_初始化BIT_S8000.Bool) flag_OutputClose = true;
                    else flag_OutputClose = false;
                }
                if (!flag_OutputClose)
                {
                    break;
                }

            }
        }
        #endregion
        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void tabPage3_Click(object sender, EventArgs e)
        {

        }


    }
}
