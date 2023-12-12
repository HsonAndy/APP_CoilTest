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
using MyOffice;
using SQLUI;
namespace APP_超晉線圈特性檢測機
{
    public partial class Form1 : Form
    {
        enum enum_test_result
        {
            GUID,
            面積比,
            電暈數量,
            電阻值,
            漏電流,
            絕緣抵抗,
            Vrms,
            Irms,
            電功率,
            加入時間,
        }
        enum enum_test_result_匯出
        {
            GUID,
            面積比,
            電暈數量,
            電阻值,
            漏電流,
            絕緣抵抗,
            Vrms,
            Irms,
            電功率,
            加入時間,
        }
        public class ExcelResultClass
        {
            public class HeaderClass
            {
                private string _規格 = "";
                private string _批號 = "";
                private string _工令 = "";
                private string _日期 = "";
                private string _測試數 = "";
                private string _良品數 = "";
                private string _不良品數 = "";

                private string _面積比 = "";
                private string _電暈數量 = "";
                private string _電阻值 = "";
                private string _漏電流 = "";
                private string _絕緣抵抗 = "";
                private string _Vrms = "";
                private string _Irms = "";
                private string _電功率 = "";

                public string 規格 { get => _規格; set => _規格 = value; }
                public string 批號 { get => _批號; set => _批號 = value; }
                public string 工令 { get => _工令; set => _工令 = value; }
                public string 日期 { get => _日期; set => _日期 = value; }
                public string 測試數 { get => _測試數; set => _測試數 = value; }
                public string 良品數 { get => _良品數; set => _良品數 = value; }
                public string 不良品數 { get => _不良品數; set => _不良品數 = value; }
                public string 面積比 { get => _面積比; set => _面積比 = value; }
                public string 電暈數量 { get => _電暈數量; set => _電暈數量 = value; }
                public string 電阻值 { get => _電阻值; set => _電阻值 = value; }
                public string 漏電流 { get => _漏電流; set => _漏電流 = value; }
                public string 絕緣抵抗 { get => _絕緣抵抗; set => _絕緣抵抗 = value; }
                public string Vrms { get => _Vrms; set => _Vrms = value; }
                public string Irms { get => _Irms; set => _Irms = value; }
                public string 電功率 { get => _電功率; set => _電功率 = value; }
            }
            public class Row
            {
                private string _面積比 = "";
                private string _電暈數量 = "";
                private string _電阻值 = "";
                private string _漏電流 = "";
                private string _絕緣抵抗 = "";
                private string _Vrms = "";
                private string _Irms = "";
                private string _電功率 = "";
                private string _判定 = "";
                private string _匝間判定 = "";
                private string _歐姆判定 = "";
                private string _耐壓判定 = "";
                private string _電功率判定 = "";

                public string 面積比 { get => _面積比; set => _面積比 = value; }
                public string 電暈數量 { get => _電暈數量; set => _電暈數量 = value; }
                public string 電阻值 { get => _電阻值; set => _電阻值 = value; }
                public string 漏電流 { get => _漏電流; set => _漏電流 = value; }
                public string 絕緣抵抗 { get => _絕緣抵抗; set => _絕緣抵抗 = value; }
                public string Vrms { get => _Vrms; set => _Vrms = value; }
                public string Irms { get => _Irms; set => _Irms = value; }
                public string 電功率 { get => _電功率; set => _電功率 = value; }
                public string 判定 { get => _判定; set => _判定 = value; }
                public string 匝間判定 { get => _匝間判定; set => _匝間判定 = value; }
                public string 歐姆判定 { get => _歐姆判定; set => _歐姆判定 = value; }
                public string 耐壓判定 { get => _耐壓判定; set => _耐壓判定 = value; }
                public string 電功率判定 { get => _電功率判定; set => _電功率判定 = value; }

            }
            private HeaderClass header = new HeaderClass();
            private List<Row> rows = new List<Row>();
            public HeaderClass Header { get => header; set => header = value; }
            public List<Row> Rows { get => rows; set => rows = value; }

            public ExcelResultClass()
            {
        
            }
            public void AddRow(Row row)
            {
                this.Rows.Add(row);
            }
        }
        public string SheetFileName = @".\excel_header.txt";
        private void Program_表單下載_Init()
        {
            Table table = new Table("");
            table.DBName = "coil_mechine";
            table.TableName = "test_result";
            table.Server = "127.0.0.1";
            table.Username = "user";
            table.Password = "66437068";
            table.Port = "3306";
            table.AddColumnList("GUID", Table.StringType.VARCHAR, 50, Table.IndexType.PRIMARY);
            table.AddColumnList("面積比", Table.StringType.VARCHAR, 50, Table.IndexType.None);
            table.AddColumnList("電暈數量", Table.StringType.VARCHAR, 50, Table.IndexType.None);
            table.AddColumnList("電阻值", Table.StringType.VARCHAR, 50, Table.IndexType.None);
            table.AddColumnList("漏電流", Table.StringType.VARCHAR, 50, Table.IndexType.None);
            table.AddColumnList("絕緣抵抗", Table.StringType.VARCHAR, 50, Table.IndexType.None);
            table.AddColumnList("Vrms", Table.StringType.VARCHAR, 50, Table.IndexType.None);
            table.AddColumnList("Irms", Table.StringType.VARCHAR, 50, Table.IndexType.None);
            table.AddColumnList("電功率", Table.StringType.VARCHAR, 50, Table.IndexType.None);
            table.AddColumnList("加入時間", Table.DateType.DATETIME, Table.IndexType.INDEX);

            this.sqL_DataGridView_線圈測試結果.DataBaseName = table.DBName;
            this.sqL_DataGridView_線圈測試結果.TableName = table.TableName;
            this.sqL_DataGridView_線圈測試結果.Server = table.Server;
            this.sqL_DataGridView_線圈測試結果.UserName = table.Username;
            this.sqL_DataGridView_線圈測試結果.Password = table.Password;
            this.sqL_DataGridView_線圈測試結果.Port = table.Port.StringToUInt32();
            this.sqL_DataGridView_線圈測試結果.SSLMode = MySql.Data.MySqlClient.MySqlSslMode.None;
            this.sqL_DataGridView_線圈測試結果.Init(table);
            this.sqL_DataGridView_線圈測試結果.Set_ColumnVisible(false, new enum_test_result().GetEnumNames());
            this.sqL_DataGridView_線圈測試結果.Set_ColumnWidth(100, DataGridViewContentAlignment.MiddleCenter, enum_test_result.GUID);
            this.sqL_DataGridView_線圈測試結果.Set_ColumnWidth(80, DataGridViewContentAlignment.MiddleCenter, enum_test_result.面積比);
            this.sqL_DataGridView_線圈測試結果.Set_ColumnWidth(80, DataGridViewContentAlignment.MiddleCenter, enum_test_result.電暈數量);
            this.sqL_DataGridView_線圈測試結果.Set_ColumnWidth(80, DataGridViewContentAlignment.MiddleCenter, enum_test_result.電阻值);
            this.sqL_DataGridView_線圈測試結果.Set_ColumnWidth(80, DataGridViewContentAlignment.MiddleCenter, enum_test_result.漏電流);
            this.sqL_DataGridView_線圈測試結果.Set_ColumnWidth(80, DataGridViewContentAlignment.MiddleCenter, enum_test_result.絕緣抵抗);
            this.sqL_DataGridView_線圈測試結果.Set_ColumnWidth(80, DataGridViewContentAlignment.MiddleCenter, enum_test_result.Vrms);
            this.sqL_DataGridView_線圈測試結果.Set_ColumnWidth(80, DataGridViewContentAlignment.MiddleCenter, enum_test_result.Irms);
            this.sqL_DataGridView_線圈測試結果.Set_ColumnWidth(80, DataGridViewContentAlignment.MiddleCenter, enum_test_result.電功率);
            this.sqL_DataGridView_線圈測試結果.Set_ColumnWidth(150, DataGridViewContentAlignment.MiddleLeft, enum_test_result.加入時間);
            if (this.sqL_DataGridView_線圈測試結果.SQL_IsTableCreat() == false)
            {
                this.sqL_DataGridView_線圈測試結果.SQL_CreateTable();
            }
            else
            {
                this.sqL_DataGridView_線圈測試結果.SQL_CheckAllColumnName(true);
            }

            this.plC_Button_表單下載.btnClick += PlC_Button_表單下載_btnClick;

            this.plC_RJ_Button_新增資料.MouseDownEvent += PlC_RJ_Button_新增資料_MouseDownEvent;
            this.plC_RJ_Button_更新資料.MouseDownEvent += PlC_RJ_Button_更新資料_MouseDownEvent;
            this.plC_RJ_Button_刪除資料.MouseDownEvent += PlC_RJ_Button_刪除資料_MouseDownEvent;
            this.plC_RJ_Button_顯示全部.MouseDownEvent += PlC_RJ_Button_顯示全部_MouseDownEvent;
            this.plC_RJ_Button_匯出.MouseDownEvent += PlC_RJ_Button_匯出_MouseDownEvent;

        }

        private void PlC_RJ_Button_匯出_MouseDownEvent(MouseEventArgs mevent)
        {
            this.Invoke(new Action(delegate 
            {
                if (this.saveFileDialog_SaveExcel.ShowDialog() == DialogResult.OK)
                {
                    List<object[]> list_value = this.sqL_DataGridView_線圈測試結果.GetAllRows();
                    DataTable dataTable = list_value.ToDataTable(new enum_test_result());
                    dataTable = dataTable.ReorderTable(enum_test_result.面積比.GetEnumName(), enum_test_result.電暈數量.GetEnumName(), enum_test_result.電阻值.GetEnumName(), enum_test_result.漏電流.GetEnumName()
                        , enum_test_result.絕緣抵抗.GetEnumName(), enum_test_result.Vrms.GetEnumName(), enum_test_result.Irms.GetEnumName(), enum_test_result.電功率.GetEnumName(), enum_test_result.加入時間.GetEnumName());
                    dataTable.SaveFile(this.saveFileDialog_SaveExcel.FileName);
                    MyMessageBox.ShowDialog("匯出完成!!");
                }
            }));
        }

        private void PlC_RJ_Button_顯示全部_MouseDownEvent(MouseEventArgs mevent)
        {
            DateTime dt_st = rJ_DatePicker_起始時間.Value;
            dt_st = new DateTime(dt_st.Year, dt_st.Month, dt_st.Day, 00, 00, 00);
            DateTime dt_end = rJ_DatePicker_結束時間.Value;
            dt_end = new DateTime(dt_end.Year, dt_end.Month, dt_end.Day, 23, 59, 59);

            List<object[]> list_value = this.sqL_DataGridView_線圈測試結果.SQL_GetRowsByBetween((int)enum_test_result.加入時間, dt_st, dt_end, true);
        }
        private void PlC_RJ_Button_刪除資料_MouseDownEvent(MouseEventArgs mevent)
        {
            List<object[]> list_value = this.sqL_DataGridView_線圈測試結果.Get_All_Select_RowsValues();

            if (list_value.Count == 0)
            {
                MyMessageBox.ShowDialog("未選取資料!!");
                return;
            }
 
            this.sqL_DataGridView_線圈測試結果.SQL_DeleteExtra(list_value, false);
            this.sqL_DataGridView_線圈測試結果.DeleteExtra(list_value, true);
        }
        private void PlC_RJ_Button_更新資料_MouseDownEvent(MouseEventArgs mevent)
        {
            List<object[]> list_value = this.sqL_DataGridView_線圈測試結果.Get_All_Select_RowsValues();

            if (list_value.Count == 0)
            {
                MyMessageBox.ShowDialog("未選取資料!!");
                return;
            }
            object[] value = list_value[0];
            // value[(int)enum_test_result.規格] = textBox_規格.Text;
            value[(int)enum_test_result.面積比] = (PLC_NumBox_IWT5000A檢測匝間面積比.Value / 10D).ToString("0.0");
            value[(int)enum_test_result.電暈數量] = PLC_NumBox_IWT5000A檢測匝間電暈數.Value.ToString();
            value[(int)enum_test_result.電阻值] = (PLC_NumBox_GOM804檢測歐姆值.Value / 1000D).ToString("0.000");
            value[(int)enum_test_result.漏電流] = textBox_ACW量測值.Text;
            value[(int)enum_test_result.絕緣抵抗] = textBox_IR絕緣量測值.Text;
            value[(int)enum_test_result.Vrms] = (PLC_NumBox_ASR_2100電功率Vrms量測值.Value / 10000D).ToString("0.0000");
            value[(int)enum_test_result.Irms] = (PLC_NumBox_ASR_2100電功率Irms量測值.Value / 10000D).ToString("0.0000");
            value[(int)enum_test_result.電功率] = (PLC_NumBox_ASR_2100電功率P量測值.Value / 10000D).ToString("0.0000");
            value[(int)enum_test_result.加入時間] = DateTime.Now.ToDateTimeString();
            this.sqL_DataGridView_線圈測試結果.SQL_ReplaceExtra(value, false);
            this.sqL_DataGridView_線圈測試結果.ReplaceExtra(value, true);
        }
        private void PlC_RJ_Button_新增資料_MouseDownEvent(MouseEventArgs mevent)
        {
            object[] value = new object[new enum_test_result().GetLength()];
            value[(int)enum_test_result.GUID] = Guid.NewGuid().ToString();
            value[(int)enum_test_result.面積比] = (PLC_NumBox_IWT5000A檢測匝間面積比.Value / 10D).ToString("0.0");
            value[(int)enum_test_result.電暈數量] = PLC_NumBox_IWT5000A檢測匝間電暈數.Value.ToString();
            value[(int)enum_test_result.電阻值] = (PLC_NumBox_GOM804檢測歐姆值.Value / 1000D).ToString("0.000");
            value[(int)enum_test_result.漏電流] = textBox_ACW量測值.Text;
            value[(int)enum_test_result.絕緣抵抗] = textBox_IR絕緣量測值.Text;
            value[(int)enum_test_result.Vrms] = (PLC_NumBox_ASR_2100電功率Vrms量測值.Value / 10000D).ToString("0.0000");
            value[(int)enum_test_result.Irms] = (PLC_NumBox_ASR_2100電功率Irms量測值.Value / 10000D).ToString("0.0000");
            value[(int)enum_test_result.電功率] = (PLC_NumBox_ASR_2100電功率P量測值.Value / 10000D).ToString("0.0000");
            value[(int)enum_test_result.加入時間] = DateTime.Now.ToDateTimeString();

            this.sqL_DataGridView_線圈測試結果.SQL_AddRow(value, false);
            this.sqL_DataGridView_線圈測試結果.AddRow(value, true);
        }

        private void PlC_RJ_Button1_MouseClickEvent(MouseEventArgs mevent)
        {

        }
        List<ExcelResultClass.Row> Row = new List<ExcelResultClass.Row>();
        PLC_Device PLC_Device_輸出一次ROW = new PLC_Device("M4000");
        PLC_Device PLC_Device_表單重置 = new PLC_Device("M4002");
        PLC_Device PLC_Device_輸出結果 = new PLC_Device("M4005");
        PLC_Device PLC_Device_輸出ROW數量 = new PLC_Device("D4000");
        PLC_Device PLC_Device_測試數量 = new PLC_Device("D1040");
        PLC_Device PLC_Device_測試數量_OK = new PLC_Device("D1030");
        PLC_Device PLC_Device_測試數量_NG = new PLC_Device("D1035");

        PLC_Device PLC_Device_電功率測試_OK = new PLC_Device("S5010");
        PLC_Device PLC_Device_耐壓測試_OK = new PLC_Device("S5011");
        PLC_Device PLC_Device_匝間測試_OK = new PLC_Device("S5012");
        PLC_Device PLC_Device_微歐姆測試_OK = new PLC_Device("S5013");
        private void PlC_Button_表單下載_btnClick(object sender, EventArgs e)
        {
            this.Invoke(new Action(delegate 
            {
                if (this.saveFileDialog_SaveExcel.ShowDialog() != DialogResult.OK) return;
            }));

            string SavefileName = this.saveFileDialog_SaveExcel.FileName;
            string loadText = Basic.MyFileStream.LoadFileAllText(SheetFileName, "utf-8");
            SheetClass sheetClass = loadText.JsonDeserializet<SheetClass>();

            #region 範例
            //excelResultClass.Header.規格 = "測試規格";
            //excelResultClass.Header.批號 = "測試批號";
            //excelResultClass.Header.工令 = "測試工令";
            //excelResultClass.Header.日期 = "2023/07/15";
            //excelResultClass.Header.測試數 = "測試測試數";
            //excelResultClass.Header.良品數 = "測試良品數";
            //excelResultClass.Header.不良品數 = "測試不良品數";
            //excelResultClass.Header.面積比 = "測試面積比";
            //excelResultClass.Header.電暈數量 = "測試電暈數量";
            //excelResultClass.Header.電阻值 = "測試電阻值";
            //excelResultClass.Header.漏電流 = "測試漏電流";
            //excelResultClass.Header.絕緣抵抗 = "測試絕緣抵抗";
            //excelResultClass.Header.Vrms = "測試Vrms";
            //excelResultClass.Header.Irms = "測試Irms";
            //excelResultClass.Header.電功率 = "測試電功率";
            #endregion
            ExcelResultClass excelResultClass = new ExcelResultClass();

            excelResultClass.Header.規格 = plC_WordBox_測試規格.Text;
            excelResultClass.Header.批號 = plC_WordBox_測試批號.Text;
            excelResultClass.Header.工令 = plC_WordBox_測試工令.Text;
            excelResultClass.Header.日期 = plC_WordBox_測試日期.Text;
            excelResultClass.Header.測試數 = PLC_Device_測試數量.Value.ToString();
            excelResultClass.Header.良品數 = PLC_Device_測試數量_OK.Value.ToString();
            excelResultClass.Header.不良品數 = PLC_Device_測試數量_NG.Value.ToString();
            excelResultClass.Header.面積比 = "測試面積比";
            excelResultClass.Header.電暈數量 = "測試電暈數量";
            excelResultClass.Header.電阻值 = "測試電阻值";
            excelResultClass.Header.漏電流 = "測試漏電流";
            excelResultClass.Header.絕緣抵抗 = "測試絕緣抵抗";
            excelResultClass.Header.Vrms = "測試Vrms";
            excelResultClass.Header.Irms = "測試Irms";
            excelResultClass.Header.電功率 = "測試電功率";


            for (int i = 0; i < PLC_Device_輸出ROW數量.Value; i++)
            {
                excelResultClass.AddRow(Row[i]);
            }

            #region 範例
            //ExcelResultClass.Row row1 = new ExcelResultClass.Row();
            //row1.面積比 = PLC_NumBox_IWT5000A檢測匝間面積比.Value.ToString();
            //row1.電暈數量 = "1";
            //row1.電阻值 = "20000";
            //row1.漏電流 = "0.155";
            //row1.絕緣抵抗 = "71822";
            //row1.Vrms = "無";
            //row1.Irms = "無";
            //row1.電功率 = "無";
            //row1.判定 = "PASS";
            //excelResultClass.AddRow(row1);

            //ExcelResultClass.Row row2 = new ExcelResultClass.Row();
            //row2.面積比 = "14985.53906";
            //row2.電暈數量 = "1";
            //row2.電阻值 = "20000";
            //row2.漏電流 = "0.155";
            //row2.絕緣抵抗 = "71822";
            //row2.Vrms = "無";
            //row2.Irms = "無";
            //row2.電功率 = "無";
            //row2.判定 = "FAIL";
            //excelResultClass.AddRow(row2);


            //for (int i = 0; i < 6; i++)
            //{
            //    row_[i].面積比 = PLC_NumBox_IWT5000A檢測匝間面積比.Value.ToString();
            //    row_[i].電暈數量 = PLC_NumBox_IWT5000A檢測匝間電暈數.Value.ToString();
            //    row_[i].電阻值 = PLC_NumBox_GOM804檢測歐姆值.Value.ToString();
            //    row_[i].漏電流 = textBox_ACW量測值.Text;
            //    row_[i].絕緣抵抗 = textBox_IR絕緣量測值.Text;
            //    row_[i].Vrms = PLC_NumBox_ASR_2100電功率Vrms量測值.Value.ToString();
            //    row_[i].Irms = PLC_NumBox_ASR_2100電功率Irms量測值.Value.ToString();
            //    row_[i].電功率 = PLC_NumBox_ASR_2100電功率P量測值.Value.ToString();
            //    row_[i].判定 = "FAIL";
            //    excelResultClass.AddRow(row_[i]);
            //}

            #endregion

            this.GetExcelResultSheet(excelResultClass, ref sheetClass);
            sheetClass.NPOI_SaveFile(SavefileName);
            MyMessageBox.ShowDialog("存檔完成!");
            PLC_Device_表單重置.Bool = true;
        }

        private void sub_PLC_Device_ROW輸出()
        {
            this.PLC_Device_測試數量.Value = this.PLC_Device_測試數量_OK.Value + this.PLC_Device_測試數量_NG.Value;
            if(PLC_Device_表單重置.Bool)
            {
                PLC_Device_輸出ROW數量.Value = 0;
                Row.Clear();
            }

            if (PLC_Device_輸出一次ROW.Bool)
            {
                PLC_Device_輸出ROW數量.Value = PLC_Device_輸出ROW數量.Value + 1;
                for (int i = 0; i < PLC_Device_輸出ROW數量.Value; i++)
                {
                    this.Row.Add(new ExcelResultClass.Row());
                }
                for (int i = PLC_Device_輸出ROW數量.Value - 1; i < PLC_Device_輸出ROW數量.Value; i++)
                {
                    Row[i].面積比 = enum_test_result.面積比.ToString("0.0");
                    Row[i].電暈數量 = enum_test_result.電暈數量.ToString();
                    Row[i].電阻值 = enum_test_result.電阻值.ToString("0.000");
                    Row[i].漏電流 = enum_test_result.漏電流.ToString();
                    Row[i].絕緣抵抗 = enum_test_result.絕緣抵抗.ToString();
                    Row[i].Vrms = enum_test_result.Vrms.ToString("0.0000");
                    Row[i].Irms = enum_test_result.Irms.ToString("0.0000");
                    Row[i].電功率 = enum_test_result.電功率.ToString("0.0000");
                    if(PLC_Device_電功率測試_OK.Bool)
                    {
                        Row[i].電功率判定 = "PASS";
                    }
                    else Row[i].電功率判定 = "FAIL";
                    if (PLC_Device_耐壓測試_OK.Bool)
                    {
                        Row[i].耐壓判定 = "PASS";
                    }
                    else Row[i].耐壓判定 = "FAIL";
                    if (PLC_Device_匝間測試_OK.Bool)
                    {
                        Row[i].匝間判定 = "PASS";
                    }
                    else Row[i].匝間判定 = "FAIL";
                    if (PLC_Device_微歐姆測試_OK.Bool)
                    {
                        Row[i].歐姆判定 = "PASS";
                    }
                    else Row[i].歐姆判定 = "FAIL";
                    if (PLC_Device_輸出結果.Bool)
                    {
                        Row[i].判定 = "PASS";
                    }
                    else Row[i].判定 = "FAIL";
                }
                PLC_Device_輸出一次ROW.Bool = false;
                
            }

        }

        private void GetExcelResultSheet(ExcelResultClass excelResultClass, ref SheetClass sheetClass)
        {
            sheetClass.ReplaceCell(2, 1, $"{excelResultClass.Header.規格}");
            sheetClass.ReplaceCell(2, 2, $"{excelResultClass.Header.批號}");
            sheetClass.ReplaceCell(2, 3, $"{excelResultClass.Header.工令}");
            sheetClass.ReplaceCell(2, 4, $"{excelResultClass.Header.日期}");

            sheetClass.ReplaceCell(2, 6, $"{excelResultClass.Header.測試數}");
            sheetClass.ReplaceCell(2, 7, $"{excelResultClass.Header.良品數}");
            sheetClass.ReplaceCell(2, 8, $"{excelResultClass.Header.不良品數}");

            sheetClass.ReplaceCell(5, 1, $"{excelResultClass.Header.面積比}");
            sheetClass.ReplaceCell(5, 2, $"{excelResultClass.Header.電暈數量}");
            sheetClass.ReplaceCell(5, 3, $"{excelResultClass.Header.電阻值}");
            sheetClass.ReplaceCell(5, 4, $"{excelResultClass.Header.漏電流}");
            sheetClass.ReplaceCell(5, 5, $"{excelResultClass.Header.絕緣抵抗}");
            sheetClass.ReplaceCell(5, 6, $"{excelResultClass.Header.Vrms}");
            sheetClass.ReplaceCell(5, 7, $"{excelResultClass.Header.Irms}");
            sheetClass.ReplaceCell(5, 8, $"{excelResultClass.Header.電功率}");

            for (int i = 0; i < excelResultClass.Rows.Count; i++)
            {
                ExcelResultClass.Row row = excelResultClass.Rows[i];

                sheetClass.AddNewCell_Webapi(7 + i, 0, $"{i + 1}", "微軟正黑體", 12, false, NPOI_Color.BLACK, 430, NPOI.SS.UserModel.HorizontalAlignment.Left, NPOI.SS.UserModel.VerticalAlignment.Bottom, NPOI.SS.UserModel.BorderStyle.Thin);
                if (row.匝間判定 == "PASS") sheetClass.AddNewCell_Webapi(7 + i, 1, $"{row.面積比}", "微軟正黑體", 12, false, NPOI_Color.BLACK, 430, NPOI.SS.UserModel.HorizontalAlignment.Left, NPOI.SS.UserModel.VerticalAlignment.Bottom, NPOI.SS.UserModel.BorderStyle.Thin);
                if (row.匝間判定 == "PASS") sheetClass.AddNewCell_Webapi(7 + i, 2, $"{row.電暈數量}", "微軟正黑體", 12, false, NPOI_Color.BLACK, 430, NPOI.SS.UserModel.HorizontalAlignment.Left, NPOI.SS.UserModel.VerticalAlignment.Bottom, NPOI.SS.UserModel.BorderStyle.Thin);
                if (row.歐姆判定 == "PASS") sheetClass.AddNewCell_Webapi(7 + i, 3, $"{row.電阻值}", "微軟正黑體", 12, false, NPOI_Color.BLACK, 430, NPOI.SS.UserModel.HorizontalAlignment.Left, NPOI.SS.UserModel.VerticalAlignment.Bottom, NPOI.SS.UserModel.BorderStyle.Thin);
                if (row.耐壓判定 == "PASS") sheetClass.AddNewCell_Webapi(7 + i, 4, $"{row.漏電流}", "微軟正黑體", 12, false, NPOI_Color.BLACK, 430, NPOI.SS.UserModel.HorizontalAlignment.Left, NPOI.SS.UserModel.VerticalAlignment.Bottom, NPOI.SS.UserModel.BorderStyle.Thin);
                if (row.耐壓判定 == "PASS") sheetClass.AddNewCell_Webapi(7 + i, 5, $"{row.絕緣抵抗}", "微軟正黑體", 12, false, NPOI_Color.BLACK, 430, NPOI.SS.UserModel.HorizontalAlignment.Left, NPOI.SS.UserModel.VerticalAlignment.Bottom, NPOI.SS.UserModel.BorderStyle.Thin);
                if (row.電功率判定 == "PASS") sheetClass.AddNewCell_Webapi(7 + i, 6, $"{row.Vrms}", "微軟正黑體", 12, false, NPOI_Color.BLACK, 430, NPOI.SS.UserModel.HorizontalAlignment.Left, NPOI.SS.UserModel.VerticalAlignment.Bottom, NPOI.SS.UserModel.BorderStyle.Thin);
                if (row.電功率判定 == "PASS") sheetClass.AddNewCell_Webapi(7 + i, 7, $"{row.Irms}", "微軟正黑體", 12, false, NPOI_Color.BLACK, 430, NPOI.SS.UserModel.HorizontalAlignment.Left, NPOI.SS.UserModel.VerticalAlignment.Bottom, NPOI.SS.UserModel.BorderStyle.Thin);
                if (row.電功率判定 == "PASS") sheetClass.AddNewCell_Webapi(7 + i, 8, $"{row.電功率}", "微軟正黑體", 12, false, NPOI_Color.BLACK, 430, NPOI.SS.UserModel.HorizontalAlignment.Left, NPOI.SS.UserModel.VerticalAlignment.Bottom, NPOI.SS.UserModel.BorderStyle.Thin);

                if (row.匝間判定 == "FAIL") sheetClass.AddNewCell_Webapi(7 + i, 1, $"{row.面積比}", "微軟正黑體", 12, false, NPOI_Color.RED, 430, NPOI.SS.UserModel.HorizontalAlignment.Left, NPOI.SS.UserModel.VerticalAlignment.Bottom, NPOI.SS.UserModel.BorderStyle.Thin);
                if (row.匝間判定 == "FAIL") sheetClass.AddNewCell_Webapi(7 + i, 2, $"{row.電暈數量}", "微軟正黑體", 12, false, NPOI_Color.RED, 430, NPOI.SS.UserModel.HorizontalAlignment.Left, NPOI.SS.UserModel.VerticalAlignment.Bottom, NPOI.SS.UserModel.BorderStyle.Thin);
                if (row.歐姆判定 == "FAIL") sheetClass.AddNewCell_Webapi(7 + i, 3, $"{row.電阻值}", "微軟正黑體", 12, false, NPOI_Color.RED, 430, NPOI.SS.UserModel.HorizontalAlignment.Left, NPOI.SS.UserModel.VerticalAlignment.Bottom, NPOI.SS.UserModel.BorderStyle.Thin);
                if (row.耐壓判定 == "FAIL") sheetClass.AddNewCell_Webapi(7 + i, 4, $"{row.漏電流}", "微軟正黑體", 12, false, NPOI_Color.RED, 430, NPOI.SS.UserModel.HorizontalAlignment.Left, NPOI.SS.UserModel.VerticalAlignment.Bottom, NPOI.SS.UserModel.BorderStyle.Thin);
                if (row.耐壓判定 == "FAIL") sheetClass.AddNewCell_Webapi(7 + i, 5, $"{row.絕緣抵抗}", "微軟正黑體", 12, false, NPOI_Color.RED, 430, NPOI.SS.UserModel.HorizontalAlignment.Left, NPOI.SS.UserModel.VerticalAlignment.Bottom, NPOI.SS.UserModel.BorderStyle.Thin);
                if (row.電功率判定 == "FAIL") sheetClass.AddNewCell_Webapi(7 + i, 6, $"{row.Vrms}", "微軟正黑體", 12, false, NPOI_Color.RED, 430, NPOI.SS.UserModel.HorizontalAlignment.Left, NPOI.SS.UserModel.VerticalAlignment.Bottom, NPOI.SS.UserModel.BorderStyle.Thin);
                if (row.電功率判定 == "FAIL") sheetClass.AddNewCell_Webapi(7 + i, 7, $"{row.Irms}", "微軟正黑體", 12, false, NPOI_Color.RED, 430, NPOI.SS.UserModel.HorizontalAlignment.Left, NPOI.SS.UserModel.VerticalAlignment.Bottom, NPOI.SS.UserModel.BorderStyle.Thin);
                if (row.電功率判定 == "FAIL") sheetClass.AddNewCell_Webapi(7 + i, 8, $"{row.電功率}", "微軟正黑體", 12, false, NPOI_Color.RED, 430, NPOI.SS.UserModel.HorizontalAlignment.Left, NPOI.SS.UserModel.VerticalAlignment.Bottom, NPOI.SS.UserModel.BorderStyle.Thin);
                if (row.判定 == "PASS") sheetClass.AddNewCell_Webapi(7 + i, 9, $"{row.判定}", "微軟正黑體", 12, false, NPOI_Color.BLACK, 430, NPOI.SS.UserModel.HorizontalAlignment.Left, NPOI.SS.UserModel.VerticalAlignment.Bottom, NPOI.SS.UserModel.BorderStyle.Thin);
                if (row.判定 == "FAIL") sheetClass.AddNewCell_Webapi(7 + i, 9, $"{row.判定}", "微軟正黑體", 12, false, NPOI_Color.RED, 430, NPOI.SS.UserModel.HorizontalAlignment.Left, NPOI.SS.UserModel.VerticalAlignment.Bottom, NPOI.SS.UserModel.BorderStyle.Thin);
            }

        }

    }
}
