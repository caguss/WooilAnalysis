using DevExpress.XtraCharts;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WooilAnalysis
{
    public partial class Form1 : Form
    {
        int trycnt = 0;

        System.Data.DataTable dt;
        System.Data.DataTable SWdt;
        System.Data.DataTable Tdt;
        System.Data.DataTable DTdt;
        bool _isDrawing = false;
        public Form1()
        {
            InitializeComponent();
            ColumnDictionary test = new ColumnDictionary();
            dtpStart_SW.Value = DateTime.Today.AddDays(-1).AddHours(7);
            dtpEnd_SW.Value = DateTime.Today.AddHours(06).AddMinutes(59).AddSeconds(59);
            dtpStart_T.Value = DateTime.Today.AddDays(-1).AddHours(7);
            dtpEnd_T.Value = DateTime.Today.AddHours(06).AddMinutes(59).AddSeconds(59);
            ChartSetting();
            cb_code_DT.SelectedIndex = 0;
        }
        #region Button Method
        /// <summary>
        /// Steam&Water 조회 버튼
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSearch_SW_Click(object sender, EventArgs e)
        {
            if ((dtpEnd_SW.Value - dtpStart_SW.Value).Seconds < 0)
            {
                MessageBox.Show("날짜 확인 부탁드립니다.");
                return;
            }
            if ((dtpEnd_SW.Value - dtpStart_SW.Value).Hours > 24)
            {
                MessageBox.Show("24시간 이상 조회 불가");
                return;
            }
            lblStatus_SW.Text = "Loading...";

            Thread t1 = new Thread(new ThreadStart(Search_SW));
            t1.Start();

        }

        private async void btn_Excel_SW_Click(object sender, EventArgs e)
        {
            if (SWdt == null)
            {
                MessageBox.Show("검색을 먼저 실행해 주세요.");
                return;
            }
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "경로 설정";
            saveFileDialog.DefaultExt = "csv";
            saveFileDialog.Filter = "csv 파일|*.csv|xlsx 파일|*.xlsx|xls 파일|*.xls";
            //saveFileDialog.DefaultExt = "csv";
            //saveFileDialog.Filter = "csv 파일|*.csv";


            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                lblStatus_SW.Text = "Loading...";
                dt = SWdt;
                CSafeSetBool(btn_Excel_SW, false);
                CSafeSetBool(btnSearch_SW, false);
                await Saving(saveFileDialog.FileName);
            }
        }

        /// <summary>
        /// TENTOR 조회 버튼
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnSearch_T_Click(object sender, EventArgs e)
        {
            if ((dtpEnd_T.Value - dtpStart_T.Value).Seconds < 0)
            {
                MessageBox.Show("날짜 확인 부탁드립니다.");
                return;
            }
            //if ((dtpEnd_T.Value - dtpStart_T.Value).Hours > 24)
            //{
            //    MessageBox.Show("24시간 이상 조회 불가");
            //    return;
            //}
            lblStatus_T.Text = "Loading...";
            Thread t1 = new Thread(new ThreadStart(Search_T));
            t1.Start();

        }


        private async void btn_Excel_T_Click(object sender, EventArgs e)
        {
            if (Tdt == null)
            {
                MessageBox.Show("검색을 먼저 실행해 주세요.");
                return;
            }
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "경로 설정";
            saveFileDialog.DefaultExt = "xlsx";
            saveFileDialog.Filter = "xlsx 파일|*.xlsx|xls 파일|*.xls";
            //saveFileDialog.DefaultExt = "csv";
            //saveFileDialog.Filter = "csv 파일|*.csv";


            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                lblStatus_T.Text = "Loading...";
                dt = Tdt;
                CSafeSetBool(btn_Excel_T, false);
                CSafeSetBool(btnSearch_T, false);
                await Task.Run(() => Saving_T(saveFileDialog.FileName));
            }
        }



        #endregion



        #region System Method

        delegate void CrossThreadSafetySetBool(Control ctl, bool gubun);
        private void CSafeSetBool(Control ctl, bool gubun)
        {
            if (ctl.InvokeRequired)
                ctl.Invoke(new CrossThreadSafetySetBool(CSafeSetBool), ctl, gubun);
            else
            {
                ctl.Enabled = gubun;
            }
        }

        delegate void CrossThreadSafetySetString(Control ctl, string gubun);
        private void CSafeSetString(Control ctl, string gubun)
        {
            if (ctl.InvokeRequired)
                ctl.Invoke(new CrossThreadSafetySetString(CSafeSetString), ctl, gubun);
            else
            {
                ctl.Text = gubun;
            }
        }
        public System.Data.DataTable GetTable_T(System.Data.SqlClient.SqlDataReader reader)
        {
            System.Data.DataTable table = reader.GetSchemaTable();
            System.Data.DataTable dt = new System.Data.DataTable();

            System.Data.DataColumn dc;
            System.Data.DataRow row;
            System.Collections.ArrayList aList = new System.Collections.ArrayList();

            for (int i = 0; i < table.Rows.Count; i++)
            {
                dc = new System.Data.DataColumn();

                if (!dt.Columns.Contains(table.Rows[i]["ColumnName"].ToString()))
                {
                    dc.ColumnName = table.Rows[i]["ColumnName"].ToString();
                    dc.Unique = Convert.ToBoolean(table.Rows[i]["IsUnique"]);
                    dc.AllowDBNull = Convert.ToBoolean(table.Rows[i]["AllowDBNull"]);
                    dc.ReadOnly = Convert.ToBoolean(table.Rows[i]["IsReadOnly"]);
                    aList.Add(dc.ColumnName);
                    dt.Columns.Add(dc);
                }
            }

            while (reader.Read())
            {
                row = dt.NewRow();
                for (int i = 0; i < aList.Count; i++)
                {
                    row[((System.String)aList[i])] = reader[table.Rows[i]["ColumnName"].ToString()];
                }
                dt.Rows.Add(row);
            }
            return dt;
        }

        public System.Data.DataTable GetTable_SW(System.Data.SqlClient.SqlDataReader reader)
        {
            System.Data.DataTable table = reader.GetSchemaTable();
            System.Data.DataTable dt = new System.Data.DataTable();

            System.Data.DataColumn dc;
            System.Data.DataRow row;
            System.Collections.ArrayList aList = new System.Collections.ArrayList();

            for (int i = 0; i < table.Rows.Count; i++)
            {
                dc = new System.Data.DataColumn();

                if (!dt.Columns.Contains(ColumnDictionary.GetColumnName(table.Rows[i]["ColumnName"].ToString())))
                {
                    dc.ColumnName = ColumnDictionary.GetColumnName(table.Rows[i]["ColumnName"].ToString());
                    dc.Unique = Convert.ToBoolean(table.Rows[i]["IsUnique"]);
                    dc.AllowDBNull = Convert.ToBoolean(table.Rows[i]["AllowDBNull"]);
                    dc.ReadOnly = Convert.ToBoolean(table.Rows[i]["IsReadOnly"]);
                    aList.Add(dc.ColumnName);
                    dt.Columns.Add(dc);
                }
            }

            while (reader.Read())
            {
                row = dt.NewRow();
                for (int i = 0; i < aList.Count; i++)
                {
                    row[((System.String)aList[i])] = reader[table.Rows[i]["ColumnName"].ToString()];
                }
                dt.Rows.Add(row);
            }
            return dt;
        }
        #endregion

        #region Chart Method
        private void ChartSetting()
        {
            // 마우스를 올리면 해당 위치에 값이 보임
            chartControl1.CrosshairEnabled = DevExpress.Utils.DefaultBoolean.True;
            chartControl1.CrosshairOptions.ShowGroupHeaders = true; // X축 시간값 표시
            chartControl2.CrosshairEnabled = DevExpress.Utils.DefaultBoolean.True;
            chartControl2.CrosshairOptions.ShowGroupHeaders = true; // X축 시간값 표시
            chartControl3.CrosshairEnabled = DevExpress.Utils.DefaultBoolean.True;
            chartControl3.CrosshairOptions.ShowGroupHeaders = true; // X축 시간값 표시
            chartControl4.CrosshairEnabled = DevExpress.Utils.DefaultBoolean.True;
            chartControl4.CrosshairOptions.ShowGroupHeaders = true; // X축 시간값 표시
            chartControl5.CrosshairEnabled = DevExpress.Utils.DefaultBoolean.True;
            chartControl5.CrosshairOptions.ShowGroupHeaders = true; // X축 시간값 표시
            // 범례 체크박스로 변경
            chartControl1.Legend.MarkerMode = LegendMarkerMode.CheckBox;
            chartControl1.Legend.MaxHorizontalPercentage = 100;
            chartControl2.Legend.MarkerMode = LegendMarkerMode.CheckBox;
            chartControl2.Legend.MaxHorizontalPercentage = 100;
            chartControl3.Legend.MarkerMode = LegendMarkerMode.CheckBox;
            chartControl3.Legend.MaxHorizontalPercentage = 100;
            chartControl4.Legend.MarkerMode = LegendMarkerMode.CheckBox;
            chartControl4.Legend.MaxHorizontalPercentage = 100;
            chartControl5.Legend.MarkerMode = LegendMarkerMode.CheckBox;
            chartControl5.Legend.MaxHorizontalPercentage = 100;

        }

        /// <summary>
        /// 그래프 그리기
        /// </summary>
        private void DrawingGraph()
        {
            if (!_isDrawing)
            {
                _isDrawing = true;
                // 초기화
                chartControl1.Series.Clear();
                chartControl2.Series.Clear();
                chartControl3.Series.Clear();
                chartControl4.Series.Clear();
                chartControl5.Series.Clear();
                string _pSenName;
                DevExpress.XtraCharts.Series _Series_Param;
                PointSeriesView myView1;

                _pSenName = "설정온도";
                _Series_Param = new DevExpress.XtraCharts.Series(_pSenName, ViewType.Point);
                myView1 = (PointSeriesView)_Series_Param.View;
                myView1.PointMarkerOptions.Size = 1;
                GraphSeriesAdd(chartControl1, _Series_Param);

                _Series_Param = new DevExpress.XtraCharts.Series(_pSenName, ViewType.Point);
                myView1 = (PointSeriesView)_Series_Param.View;
                myView1.PointMarkerOptions.Size = 1;
                GraphSeriesAdd(chartControl2, _Series_Param);
                _Series_Param = new DevExpress.XtraCharts.Series(_pSenName, ViewType.Point);
                myView1 = (PointSeriesView)_Series_Param.View;
                myView1.PointMarkerOptions.Size = 1;
                GraphSeriesAdd(chartControl3, _Series_Param);
                _Series_Param = new DevExpress.XtraCharts.Series(_pSenName, ViewType.Point);
                myView1 = (PointSeriesView)_Series_Param.View;
                myView1.PointMarkerOptions.Size = 1;
                GraphSeriesAdd(chartControl4, _Series_Param);
                _Series_Param = new DevExpress.XtraCharts.Series(_pSenName, ViewType.Point);
                myView1 = (PointSeriesView)_Series_Param.View;
                myView1.PointMarkerOptions.Size = 1;
                GraphSeriesAdd(chartControl5, _Series_Param);

                _pSenName = "현재온도";
                _Series_Param = new DevExpress.XtraCharts.Series(_pSenName, ViewType.Point);
                myView1 = (PointSeriesView)_Series_Param.View;
                myView1.PointMarkerOptions.Size = 1;
                GraphSeriesAdd(chartControl1, _Series_Param);
                _Series_Param = new DevExpress.XtraCharts.Series(_pSenName, ViewType.Point);
                myView1 = (PointSeriesView)_Series_Param.View;
                myView1.PointMarkerOptions.Size = 1;
                GraphSeriesAdd(chartControl2, _Series_Param);
                _Series_Param = new DevExpress.XtraCharts.Series(_pSenName, ViewType.Point);
                myView1 = (PointSeriesView)_Series_Param.View;
                myView1.PointMarkerOptions.Size = 1;
                GraphSeriesAdd(chartControl3, _Series_Param);
                _Series_Param = new DevExpress.XtraCharts.Series(_pSenName, ViewType.Point);
                myView1 = (PointSeriesView)_Series_Param.View;
                myView1.PointMarkerOptions.Size = 1;
                GraphSeriesAdd(chartControl4, _Series_Param);
                _Series_Param = new DevExpress.XtraCharts.Series(_pSenName, ViewType.Point);
                myView1 = (PointSeriesView)_Series_Param.View;
                myView1.PointMarkerOptions.Size = 1;
                GraphSeriesAdd(chartControl5, _Series_Param);

                _pSenName = "포온도";
                _Series_Param = new DevExpress.XtraCharts.Series(_pSenName, ViewType.Point);
                myView1 = (PointSeriesView)_Series_Param.View;
                myView1.PointMarkerOptions.Size = 1;
                GraphSeriesAdd(chartControl1, _Series_Param);
                _Series_Param = new DevExpress.XtraCharts.Series(_pSenName, ViewType.Point);
                myView1 = (PointSeriesView)_Series_Param.View;
                myView1.PointMarkerOptions.Size = 1;
                GraphSeriesAdd(chartControl2, _Series_Param);
                _Series_Param = new DevExpress.XtraCharts.Series(_pSenName, ViewType.Point);
                myView1 = (PointSeriesView)_Series_Param.View;
                myView1.PointMarkerOptions.Size = 1;
                GraphSeriesAdd(chartControl3, _Series_Param);
                _Series_Param = new DevExpress.XtraCharts.Series(_pSenName, ViewType.Point);
                myView1 = (PointSeriesView)_Series_Param.View;
                myView1.PointMarkerOptions.Size = 1;
                GraphSeriesAdd(chartControl4, _Series_Param);
                _Series_Param = new DevExpress.XtraCharts.Series(_pSenName, ViewType.Point);
                myView1 = (PointSeriesView)_Series_Param.View;
                myView1.PointMarkerOptions.Size = 1;
                GraphSeriesAdd(chartControl5, _Series_Param);

                for (int i = 0; i < SWdt.Rows.Count; i++)
                {

                    DateTime rowtime;
                    DateTime.TryParseExact(SWdt.Rows[i]["생산일자"].ToString() + SWdt.Rows[i]["생산시간"].ToString(), "yyyyMMddHHmmss", null, DateTimeStyles.None, out rowtime);

                    // 포인트 추가부분
                    decimal data;
                    SeriesPoint sp;

                    data = Convert.ToDecimal(SWdt.Rows[i]["챔바3설정온도1"].ToString());
                    sp = new SeriesPoint(rowtime, data);
                    chartControl1.Series["설정온도"].Points.Add(sp);

                    data = Convert.ToDecimal(SWdt.Rows[i]["챔바3현재온도1"].ToString());
                    sp = new SeriesPoint(rowtime, data);
                    chartControl1.Series["현재온도"].Points.Add(sp);

                    data = Convert.ToDecimal(SWdt.Rows[i]["포온도3현재온도"].ToString());
                    sp = new SeriesPoint(rowtime, data);
                    chartControl1.Series["포온도"].Points.Add(sp);

                    data = Convert.ToDecimal(SWdt.Rows[i]["챔바4설정온도1"].ToString());
                    sp = new SeriesPoint(rowtime, data);
                    chartControl2.Series["설정온도"].Points.Add(sp);

                    data = Convert.ToDecimal(SWdt.Rows[i]["챔바4현재온도1"].ToString());
                    sp = new SeriesPoint(rowtime, data);
                    chartControl2.Series["현재온도"].Points.Add(sp);

                    data = Convert.ToDecimal(SWdt.Rows[i]["포온도4현재온도"].ToString());
                    sp = new SeriesPoint(rowtime, data);
                    chartControl2.Series["포온도"].Points.Add(sp);

                    data = Convert.ToDecimal(SWdt.Rows[i]["챔바5설정온도1"].ToString());
                    sp = new SeriesPoint(rowtime, data);
                    chartControl3.Series["설정온도"].Points.Add(sp);

                    data = Convert.ToDecimal(SWdt.Rows[i]["챔바5현재온도1"].ToString());
                    sp = new SeriesPoint(rowtime, data);
                    chartControl3.Series["현재온도"].Points.Add(sp);

                    data = Convert.ToDecimal(SWdt.Rows[i]["포온도5현재온도"].ToString());
                    sp = new SeriesPoint(rowtime, data);
                    chartControl3.Series["포온도"].Points.Add(sp);


                    data = Convert.ToDecimal(SWdt.Rows[i]["챔바6설정온도1"].ToString());
                    sp = new SeriesPoint(rowtime, data);
                    chartControl4.Series["설정온도"].Points.Add(sp);

                    data = Convert.ToDecimal(SWdt.Rows[i]["챔바6현재온도1"].ToString());
                    sp = new SeriesPoint(rowtime, data);
                    chartControl4.Series["현재온도"].Points.Add(sp);

                    data = Convert.ToDecimal(SWdt.Rows[i]["포온도6현재온도"].ToString());
                    sp = new SeriesPoint(rowtime, data);
                    chartControl4.Series["포온도"].Points.Add(sp);



                    data = Convert.ToDecimal(SWdt.Rows[i]["챔바7설정온도1"].ToString());
                    sp = new SeriesPoint(rowtime, data);
                    chartControl5.Series["설정온도"].Points.Add(sp);

                    data = Convert.ToDecimal(SWdt.Rows[i]["챔바7현재온도1"].ToString());
                    sp = new SeriesPoint(rowtime, data);
                    chartControl5.Series["현재온도"].Points.Add(sp);

                    data = Convert.ToDecimal(SWdt.Rows[i]["포온도7현재온도"].ToString());
                    sp = new SeriesPoint(rowtime, data);
                    chartControl5.Series["포온도"].Points.Add(sp);

                }
                ChartReSizing(chartControl1);
                ChartReSizing(chartControl2);
                ChartReSizing(chartControl3);
                ChartReSizing(chartControl4);
                ChartReSizing(chartControl5);
                _isDrawing = false;
            }
        }
        /// <summary>
        /// 시리즈 추가 후 속성 세팅
        /// </summary>
        /// <param name="series"></param>


        private void XScaleSetting(XYDiagram diagram)
        {

            diagram.AxisX.WholeRange.SideMarginsValue = 0; // 왼쪽에 공백 0으로 설정해서 없에기
            diagram.AxisX.Label.TextPattern = "{A:dd'일' HH:mm}"; // 시간,분,초만 표시
            diagram.AxisX.DateTimeScaleOptions.ScaleMode = ScaleMode.Manual;
            diagram.AxisX.DateTimeScaleOptions.MeasureUnit = DateTimeMeasureUnit.Second; // 데이터들을 MeasureUnit단위로 묶어서 평균을 내 한개의 점으로 찍게함
            diagram.AxisX.DateTimeScaleOptions.GridAlignment = DateTimeGridAlignment.Hour; // X축에 날짜를 시간단위로 보여줌
            diagram.AxisX.DateTimeScaleOptions.GridSpacing = 1; // 1 DateTimeMeasureUnit.Minute(분) 단위로 X축 표시


        }
        /// <summary>
        /// 시리즈 추가 후 속성 세팅
        /// </summary>
        /// <param name="series"></param>
        private void GraphSeriesAdd(ChartControl _chartMstView, DevExpress.XtraCharts.Series _Series_Param2)
        {
            // 범례 체크박스는 기본적으로 true
            _Series_Param2.CheckedInLegend = true;
            // 없을경우 추가
            _chartMstView.Series.Add(_Series_Param2);

            // X,Y축 설정용 변수. 시리즈를 추가 한 뒤 변수를 조회하여야 올바른 값이 반환되는듯 함.
            XYDiagram diagram = (XYDiagram)_chartMstView.Diagram;
            // X축 스크롤 및 줌 설정
            diagram.EnableAxisXScrolling = true;
            diagram.EnableAxisXZooming = true;
            // Y축 스크롤 및 줌 설정
            diagram.EnableAxisYScrolling = false;
            diagram.EnableAxisYZooming = false;
            XScaleSetting(diagram);
            //diagram.AxisX.WholeRange.SideMarginsValue = 0; // 왼쪽에 공백 0으로 설정해서 없에기
            //diagram.AxisX.Label.TextPattern = "{A:dd'일' HH:mm}"; // 시간,분,초만 표시
            //diagram.AxisX.DateTimeScaleOptions.ScaleMode = ScaleMode.Manual;
            //diagram.AxisX.DateTimeScaleOptions.MeasureUnit = DateTimeMeasureUnit.Hour; // 데이터들을 MeasureUnit단위로 묶어서 평균을 내 한개의 점으로 찍게함
            //diagram.AxisX.DateTimeScaleOptions.GridAlignment = DateTimeGridAlignment.Hour; // X축에 날짜를 시간단위로 보여줌
            //diagram.AxisX.DateTimeScaleOptions.GridSpacing = 1; // 1 DateTimeMeasureUnit.Minute(분) 단위로 X축 표시

            // 1번째 ROW는 기본 Y축을 이용한다.
            // Y축설정
            //diagram.AxisY.WholeRange.MaxValue = GetMaxValue(ScaleName);
            //diagram.AxisY.WholeRange.MinValue = GetMinValue(ScaleName);
            diagram.AxisY.WholeRange.Auto = true; // Y축 자동설정
            diagram.AxisY.WholeRange.AlwaysShowZeroLevel = false;
            diagram.AxisY.NumericScaleOptions.GridSpacing = 1;
            diagram.AxisY.NumericScaleOptions.GridOffset = 0;
            diagram.AxisY.NumericScaleOptions.MeasureUnit = NumericMeasureUnit.Ones;
            diagram.AxisY.NumericScaleOptions.AutoGrid = true;
            diagram.AxisY.GridLines.Visible = false;
            diagram.AxisY.Label.TextPattern = "{V:N0}"; // 시리즈 Y축 값 표시형식 정수형태로 변경

        }
        private void ChartReSizing(ChartControl _chartMstView)
        {
            int commonPadding = 6; // 컨트롤의 기본 패딩 3 + 3 
            // 넓이는 폼길이 - 좌우측 padding(테이블레이아웃의 0번, 마지막번 열 넓이 더한 값)
            int widthScale = (int)tableLayoutPanel3.ColumnStyles[0].Width + (int)tableLayoutPanel3.ColumnStyles[tableLayoutPanel3.ColumnCount - 1].Width + commonPadding + 50;
            _chartMstView.Width = this.Width - widthScale;

            // 높이는 폼높이 - 0행 높이 + {생길 리소스 줄 수 *체크박스높이 + 3}
            //( (리소스글자수 + 10(체크박스높이) + 기본패딩(체크박스와 리소스글자 사이 padding)) / 넓이)  => 한 줄에 들어갈 리소스 수
            // 총 리소스 개수/한 줄에 들어갈 리소스 수  = 생길 리소스 줄 수 
            // 
            int ZeroRowHeight = (int)tableLayoutPanel3.RowStyles[0].Height;
            int ResourceCnt = _chartMstView.Series.Count(); // 총 리소스 개수
            int SumResourceNameLength = 0; // 리소스 글자 수
            int CheckboxSize = _chartMstView.Legend.MarkerSize.Height + 7; // 체크박스 높이
            int CntOfRow = 0; // 생길 리소스 줄 수
            foreach (DevExpress.XtraCharts.Series param in _chartMstView.Series)
            {
                SumResourceNameLength += param.Name.Length;
            }
            SumResourceNameLength = SumResourceNameLength * (int)_chartMstView.Legend.Font.Size;
            int ResourceSize = SumResourceNameLength / ResourceCnt; ; // 리소스 1개 평균 넓이
            if (ResourceCnt > 0)
            {
                //CntOfRow = (int)Math.Ceiling((((SumResourceNameLength *8)+ CheckboxSize + commonPadding)/(double)(_chartMstView.Width)));
                CntOfRow = ResourceCnt / (_chartMstView.Width / ResourceSize); // 1줄에 들어갈 리소스 개수

                _chartMstView.Height = this.Height - ZeroRowHeight + (CntOfRow * CheckboxSize);
            }
            else
            {
                _chartMstView.Height = this.Height - ZeroRowHeight - 50;
            }
        }


        #endregion


        #region Data Method
        private void Search_T()
        {
            SqlConnection conn = new SqlConnection("Server=db2.coever.co.kr,1897; Database=CoFAS_WOOIL; uid=wooriuser; pwd=wooriuser1!");

            try
            {

                CSafeSetBool(btnSearch_T, false);
                CSafeSetBool(btn_Excel_T, false);
                using (conn)
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand("sp_SearchResult_T", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 30;

                    // Input param
                    SqlParameter startdate = new SqlParameter("@v_startDate", SqlDbType.VarChar, 14);
                    SqlParameter enddate = new SqlParameter("@v_endDate", SqlDbType.VarChar, 14);
                    startdate.Value = dtpStart_T.Value.ToString("yyyyMMddHHmmss");
                    enddate.Value = dtpEnd_T.Value.ToString("yyyyMMddHHmmss");

                    cmd.Parameters.Add(startdate);
                    cmd.Parameters.Add(enddate);

                    // 데이타는 서버에서 가져오도록 실행
                    SqlDataReader rdr = cmd.ExecuteReader();

                    Tdt = new System.Data.DataTable();
                    Tdt = GetTable_T(rdr);
                }

                decimal jiggersteam;
                decimal rapidsteam;
                decimal trainersteam;
                decimal jiggerwater;
                decimal rapidwater;
                decimal trainerwater;

                System.Data.DataTable dt1;
                for (int i = 0; i < Tdt.Rows.Count; i++)
                {
                    dt1 = new System.Data.DataTable();
                    switch (Tdt.Rows[i]["place_name"].ToString())
                    {
                        case "직가":
                            dt1 = Tdt.AsEnumerable().Where(Row => Row.Field<string>("place_cd") == "cd_001").OrderBy(Row => Row.Field<string>("ndate")).CopyToDataTable();

                            jiggersteam = Convert.ToDecimal(dt1.Rows[1]["steam"].ToString() == " " ? "0" : dt1.Rows[1]["steam"].ToString()) - Convert.ToDecimal(dt1.Rows[0]["steam"].ToString() == " " ? "0" : dt1.Rows[0]["steam"].ToString());
                            jiggerwater = Convert.ToDecimal(dt1.Rows[1]["water"].ToString() == " " ? "0" : dt1.Rows[1]["water"].ToString()) - Convert.ToDecimal(dt1.Rows[0]["water"].ToString() == " " ? "0" : dt1.Rows[0]["water"].ToString());
                            CSafeSetString(lblJiggerSteam, jiggersteam.ToString());
                            CSafeSetString(lblJiggerWater, jiggerwater.ToString());
                            CSafeSetString(lblCurrentJiggerSteam, dt1.Rows[1]["steam"].ToString() == " " ? "0" : dt1.Rows[1]["steam"].ToString());
                            CSafeSetString(lblCurrentJiggerWater, dt1.Rows[1]["water"].ToString() == " " ? "0" : dt1.Rows[1]["water"].ToString());
                            break;
                        case "래피드":
                            dt1 = Tdt.AsEnumerable().Where(Row => Row.Field<string>("place_cd") == "cd_003").OrderBy(Row => Row.Field<string>("ndate")).CopyToDataTable();

                            rapidsteam = Convert.ToDecimal(dt1.Rows[1]["steam"].ToString() == " " ? "0" : dt1.Rows[1]["steam"].ToString()) - Convert.ToDecimal(dt1.Rows[0]["steam"].ToString() == " " ? "0" : dt1.Rows[0]["steam"].ToString());
                            rapidwater = Convert.ToDecimal(dt1.Rows[1]["water"].ToString() == " " ? "0" : dt1.Rows[1]["water"].ToString()) - Convert.ToDecimal(dt1.Rows[0]["water"].ToString() == " " ? "0" : dt1.Rows[0]["water"].ToString());
                            CSafeSetString(lblRapidSteam, rapidsteam.ToString());
                            CSafeSetString(lblRapidWater, rapidwater.ToString());
                            CSafeSetString(lblCurrentRapidSteam, dt1.Rows[1]["steam"].ToString() == " " ? "0" : dt1.Rows[1]["steam"].ToString());
                            CSafeSetString(lblCurrentRapidWater, dt1.Rows[1]["water"].ToString() == " " ? "0" : dt1.Rows[1]["water"].ToString());
                            break;
                        case "정련기":

                            dt1 = Tdt.AsEnumerable().Where(Row => Row.Field<string>("place_cd") == "cd_002").OrderBy(Row => Row.Field<string>("ndate")).CopyToDataTable();

                            trainersteam = Convert.ToDecimal(dt1.Rows[1]["steam"].ToString() == " " ? "0" : dt1.Rows[1]["steam"].ToString()) - Convert.ToDecimal(dt1.Rows[0]["steam"].ToString() == " " ? "0" : dt1.Rows[0]["steam"].ToString());
                            trainerwater = Convert.ToDecimal(dt1.Rows[1]["water"].ToString() == " " ? "0" : dt1.Rows[1]["water"].ToString()) - Convert.ToDecimal(dt1.Rows[0]["water"].ToString() == " " ? "0" : dt1.Rows[0]["water"].ToString());
                            CSafeSetString(lblTrainerSteam, trainersteam.ToString());
                            CSafeSetString(lblTrainerWater, trainerwater.ToString());
                            CSafeSetString(lblCurrentTrainerSteam, dt1.Rows[1]["steam"].ToString() == " " ? "0" : dt1.Rows[1]["steam"].ToString());
                            CSafeSetString(lblCurrentTrainerWater, dt1.Rows[1]["water"].ToString() == " " ? "0" : dt1.Rows[1]["water"].ToString());
                            break;
                    }

                }

                trycnt = 0;
                CSafeSetString(lblStatus_T, "Finish!");
            }
            catch (Exception ex)
            {
                conn.Close();
                trycnt++;
                if (trycnt < 5)
                {
                    CSafeSetString(lblStatus_T, "Retrying...");
                    Search_T();
                }
                else
                {
                    MessageBox.Show(ex.ToString());
                    CSafeSetString(lblStatus_T, "Error!");
                }

            }
            finally
            {
                CSafeSetBool(btn_Excel_T, true);
                CSafeSetBool(btnSearch_T, true);

            }
        }
        private void Search_DT()
        {
            SqlConnection conn = new SqlConnection("Server=db2.coever.co.kr,1897; Database=CoFAS_WOOIL; uid=wooriuser; pwd=wooriuser1!");

            try
            {

                CSafeSetBool(btnSearch_T, false);
                CSafeSetBool(btn_Excel_T, false);
                using (conn)
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand("sp_SearchResult_DT", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 30;

                    // Input param
                    SqlParameter date = new SqlParameter("@v_Date", SqlDbType.VarChar, 8);
                    SqlParameter code = new SqlParameter("@v_Code", SqlDbType.VarChar, 14);
                    date.Value = dtp_date_DT.Value.ToString("yyyyMMdd");

                    string cb_code_vl = "";
                    cb_code_DT.Invoke((MethodInvoker)delegate ()
                    {
                        cb_code_vl = cb_code_DT.Text;
                    });

                    switch (cb_code_vl)
                    {
                        case "직거":
                            code.Value = "cd_001";
                            break;
                        case "래피드":
                            code.Value = "cd_002";
                            break;
                        case "정련기":
                            code.Value = "cd_003";
                            break;
                    }

                    cmd.Parameters.Add(date);
                    cmd.Parameters.Add(code);

                    // 데이타는 서버에서 가져오도록 실행
                    SqlDataReader rdr = cmd.ExecuteReader();
                    DTdt = GetTable_T(rdr);
                        dgv_main_DT.DataSource = DTdt;
                    dgv_main_DT.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.DisplayedCells;
                }

                trycnt = 0;
                lblStatus_DT.Text = "Finish!";
            }
            catch (Exception ex)
            {
                conn.Close();
                trycnt++;
                if (trycnt < 5)
                {
                lblStatus_DT.Text = "Retrying...";
                    Search_DT();
                }
                else
                {
                    MessageBox.Show(ex.ToString());
                lblStatus_DT.Text = "Error!";
                }

            }
            finally
            {
                CSafeSetBool(btn_Excel_T, true);
                CSafeSetBool(btnSearch_T, true);

            }
        }
        private void Search_SW()
        {
            SqlConnection conn = new SqlConnection("Server=db2.coever.co.kr,1897; Database=CoFAS_WOOIL; uid=wooriuser; pwd=wooriuser1!");

            try
            {


                CSafeSetBool(btnSearch_SW, false);
                CSafeSetBool(btn_Excel_SW, false);

                using (conn)
                {
                    conn.Open();
                    SqlCommand cmd = new SqlCommand("sp_SearchResult_SW", conn);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 30;

                    // Input param
                    SqlParameter startdate = new SqlParameter("@v_startDate", SqlDbType.VarChar, 14);
                    SqlParameter enddate = new SqlParameter("@v_endDate", SqlDbType.VarChar, 14);
                    startdate.Value = dtpStart_SW.Value.ToString("yyyyMMddHHmmss");
                    enddate.Value = dtpEnd_SW.Value.ToString("yyyyMMddHHmmss");

                    cmd.Parameters.Add(startdate);
                    cmd.Parameters.Add(enddate);



                    // 데이타는 서버에서 가져오도록 실행
                    SqlDataReader rdr = cmd.ExecuteReader();
                    SWdt = new System.Data.DataTable();
                    SWdt = GetTable_SW(rdr);
                    SWdt = SWdt.AsEnumerable().OrderBy(Row => Row.Field<string>("생산시간")).OrderBy(Row => Row.Field<string>("생산일자")).CopyToDataTable();
                    dgvRow.Invoke(new System.Action(delegate ()
                    {
                        dgvRow.DataSource = SWdt;
                    }));


                }
                tableLayoutPanel3.Invoke(new System.Action(delegate ()
                {
                    DrawingGraph();
                }));
                trycnt = 0;
                CSafeSetString(lblStatus_SW, "Finish!");

            }
            catch (Exception ex)
            {
                conn.Close();
                trycnt++;
                if (trycnt < 5)
                {
                    CSafeSetString(lblStatus_SW, "Retrying...");
                    Search_SW();
                }
                else
                {
                    MessageBox.Show(ex.ToString());
                    CSafeSetString(lblStatus_SW, "Error!");
                }

            }
            finally
            {
                CSafeSetBool(btnSearch_SW, true);
                CSafeSetBool(btn_Excel_SW, true);
            }
        }
        private bool XlsxSave_OneFile(string filename)
        {

            try
            {
                #region 통합 파일
                My_DataTable_Extensions.ExportToExcel(dt, filename);

                return true;
                #endregion
            }
            catch (Exception ex)
            {
                //wb = new XLWorkbook();

                CSafeSetString(lblStatus_SW, "Error!");
                MessageBox.Show(ex.Message);
                return false;

            }
            finally
            {
                CSafeSetString(lblStatus_SW, "Finish!");

            }
        }

        private void Saving_T(string fileName)
        {
            try
            {
                Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application();
                _Worksheet Worksheet;
                // load excel, and create a new workbook
                Excel.Workbooks.Add();
                int ColumnsCount;
                if (dt == null || (ColumnsCount = Tdt.Columns.Count) == 0)
                    throw new Exception("ExportToExcel: Null or empty input table!\n");
                // single worksheet
                Worksheet = (_Worksheet)Excel.ActiveSheet;

                object[] Header = new object[ColumnsCount];



                Microsoft.Office.Interop.Excel.Range HeaderRange = Worksheet.get_Range((Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[2, 3]), (Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[3, ColumnsCount - 1]));
                HeaderRange.Value = Header;
                HeaderRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                HeaderRange.Font.Bold = true;

                //조회 기간 입력
                Worksheet.Cells[1, 7] = "조회기간";
                Worksheet.Cells[1, 8] = dtpStart_T.Value.ToString();
                Worksheet.Cells[1, 9] = "~";
                Worksheet.Cells[1, 10] = dtpEnd_T.Value.ToString();
                //헤더양식 작성
                Worksheet.Cells[4, 2] = "직거";
                Worksheet.Cells[5, 2] = "래피드";
                Worksheet.Cells[6, 2] = "정련기";
                Worksheet.Cells[2, 3] = "수량카운터";
                Worksheet.Cells[2, 5] = "스팀량카운터";
                Worksheet.Cells[3, 3] = "사용량";
                Worksheet.Cells[3, 4] = "누적사용량";
                Worksheet.Cells[3, 5] = "사용량";
                Worksheet.Cells[3, 6] = "누적사용량";


                //value insert
                Worksheet.Cells[4, 3] = lblJiggerWater.Text;
                Worksheet.Cells[5, 3] = lblRapidWater.Text;
                Worksheet.Cells[6, 3] = lblTrainerWater.Text;
                Worksheet.Cells[4, 4] = lblCurrentJiggerWater.Text;
                Worksheet.Cells[5, 4] = lblCurrentRapidWater.Text;
                Worksheet.Cells[6, 4] = lblCurrentTrainerWater.Text;
                Worksheet.Cells[4, 5] = lblJiggerSteam.Text;
                Worksheet.Cells[5, 5] = lblRapidSteam.Text;
                Worksheet.Cells[6, 5] = lblTrainerSteam.Text;
                Worksheet.Cells[4, 6] = lblCurrentJiggerSteam.Text;
                Worksheet.Cells[5, 6] = lblCurrentRapidSteam.Text;
                Worksheet.Cells[6, 6] = lblCurrentTrainerSteam.Text;
                Worksheet.Columns.AutoFit();

                // check fielpath
                if (fileName != null && fileName != "")
                {

                    Worksheet.SaveAs(fileName);
                    Excel.Quit();

                }
                else    // no filepath is given
                {
                    Excel.Visible = true;
                }
                CSafeSetString(lblStatus_T, "Finish!");

            }
            catch (Exception ex)
            {
                CSafeSetString(lblStatus_T, "Error!");
                MessageBox.Show(ex.Message);
            }
            finally
            {
                CSafeSetBool(btn_Excel_T, true);
                CSafeSetBool(btnSearch_T, true);
                MessageBox.Show("저장이 완료되었습니다.");

            }


        }

        private async Task Saving(string filename)
        {
            try
            {
                await Task.Run(() =>
                {

                    //저장 방식에 따른 save method 분기
                    switch (filename.Split('.')[filename.Split('.').Count() - 1])
                    {
                        case "xlsx":
                        case "xls":
                            XlsxSave_OneFile(filename);
                            break;
                        case "csv":
                            CsvSave_EachFile(filename);
                            break;
                    }
                });
                CSafeSetString(lblStatus_SW, "Finish!");

            }
            catch (Exception ex)
            {
                CSafeSetString(lblStatus_SW, "Error!");
                MessageBox.Show(ex.Message);
            }
            finally
            {
                CSafeSetBool(btn_Excel_SW, true);
                CSafeSetBool(btnSearch_SW, true);
                MessageBox.Show("저장이 완료되었습니다.");

            }

        }

        private void CsvSave_EachFile(string filename)
        {
            string filepath = Path.GetDirectoryName(filename);
            if (File.Exists(filepath) == false)
            {
                DirectoryInfo di = new DirectoryInfo(filepath);
                di.Create();
            }


            Microsoft.Office.Interop.Excel.Application xls = new Microsoft.Office.Interop.Excel.ApplicationClass();
            xls.Visible = false;
            FileStream fs = new FileStream(filepath + "\\" + Path.GetFileNameWithoutExtension(filename) + Path.GetExtension(filename), FileMode.Create, FileAccess.Write);
            StreamWriter sw = new StreamWriter(fs, Encoding.UTF8);
            try
            {
                //컬럼 이름들을 ","로 나누고 저장.
                string line = string.Join(",", dt.Columns.Cast<object>());
                sw.WriteLine(line);

                //row들을 ","로 나누고 저장.
                foreach (DataRow item in dt.Rows)
                {
                    line = string.Join(",", item.ItemArray.Cast<object>());
                    sw.WriteLine(line);
                }
                sw.Close();
                fs.Close();
                //csv 파일 생성 끝
                xls.Quit();
            }
            catch (Exception ex)
            {
                sw.Close();
                fs.Close();
                xls.Quit();
                CSafeSetString(lblStatus_SW, "Error!");
                MessageBox.Show(ex.Message);
            }
            finally
            {
                CSafeSetString(lblStatus_SW, "Finish!");

            }
        }

        #endregion

        private void btn_Search_DT_Click(object sender, EventArgs e)
        {
            lblStatus_DT.Text = "Loading...";
            Search_DT();
        }

        private void btn_Excel_DT_Click(object sender, EventArgs e)
        {
            if (DTdt == null)
            {
                MessageBox.Show("검색을 먼저 실행해 주세요.");
                return;
            }
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Title = "경로 설정";
            saveFileDialog.DefaultExt = "xlsx";
            saveFileDialog.Filter = "xlsx 파일|*.xlsx|xls 파일|*.xls";


            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                lblStatus_DT.Text = "Loading...";
                dt = DTdt;
                CSafeSetBool(btn_Excel_DT, false);
                CSafeSetBool(btn_Search_DT, false);
                try
                {
                    My_DataTable_Extensions.ExportToExcel(dt, saveFileDialog.FileName);
                    MessageBox.Show("저장이 완료되었습니다.");
                    lblStatus_DT.Text = "Finish!";

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message);
                    lblStatus_DT.Text = "Error!";
                }
            }



        }
    }

    public static class My_DataTable_Extensions
    {
        /// <summary>
        /// Export DataTable to Excel file
        /// </summary>
        /// <param name="ds">Source DataTable</param>
        /// <param name="ExcelFilePath">Path to result file name</param>
        public static void ExportToExcel(this System.Data.DataTable dt, string ExcelFilePath = null)
        {
            Microsoft.Office.Interop.Excel.Application Excel = new Microsoft.Office.Interop.Excel.Application();
            _Worksheet Worksheet;
            int ColumnsCount;
            if (dt == null || (ColumnsCount = dt.Columns.Count) == 0)
                throw new Exception("ExportToExcel: Null or empty input table!\n");

            // load excel, and create a new workbook
            Excel.Workbooks.Add();

            // single worksheet
            Worksheet = (_Worksheet)Excel.ActiveSheet;

            object[] Header = new object[ColumnsCount];

            // column headings               
            for (int j = 0; j < ColumnsCount; j++)
                Header[j] = dt.Columns[j].ColumnName;

            Microsoft.Office.Interop.Excel.Range HeaderRange = Worksheet.get_Range((Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[1, 1]), (Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[1, ColumnsCount]));
            HeaderRange.Value = Header;
            HeaderRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
            HeaderRange.Font.Bold = true;

            // DataCells
            int RowsCount = dt.Rows.Count;
            object[,] Cells = new object[RowsCount, ColumnsCount];

            for (int k = 0; k < RowsCount; k++)
                for (int j = 0; j < ColumnsCount; j++)
                    Cells[k, j] = dt.Rows[k][j];

            Worksheet.get_Range((Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[2, 1]), (Microsoft.Office.Interop.Excel.Range)(Worksheet.Cells[RowsCount + 1, ColumnsCount])).Value = Cells;

            Worksheet.Columns.AutoFit();
            // check fielpath
            if (ExcelFilePath != null && ExcelFilePath != "")
            {

                Worksheet.SaveAs(ExcelFilePath);
                Excel.Quit();

            }
            else    // no filepath is given
            {
                Excel.Visible = true;
            }
        }
    }
}

