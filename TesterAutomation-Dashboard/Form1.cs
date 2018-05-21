using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Ivi.Visa;
using Ivi.Visa.FormattedIO;
using System.Windows.Forms.DataVisualization.Charting;

namespace TesterAutomation_Dashboard
{
    public partial class Form1 : Form
    {
        //Global Var
        string Func = "CpGpRp";
        bool GPIBstatus = false;
        bool PauseIsTrue = false;
        bool ratemode = false;
        string[] Cp_Unit = { "F", "mF", "μF", "nF", "pF" };
        string[] Gp_Unit = { "S", "mS", "μS", "nS", "pS" };
        string[] Rp_Unit = { "Ω", "kΩ", "MΩ" };
        int Cp_Unit_order = 0;
        int Gp_Unit_order = 0;
        int Rp_Unit_order = 0;
        int rate = 200;
        Int32 n = 0;
        System.Threading.ThreadStart MonitorThreadStart;
        System.Threading.Thread MonitorThread;
        DataTable dtdata;
        List<double> listdatax;
        List<double> listdatay;
        DataRow newrow;

        public enum DashboardState
        {
            noconnection,
            running,
            pause
        }
        DashboardState state = DashboardState.noconnection;

        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            labelInt_Value.Text = Inter().TrimEnd(',');
            labelMea1_Unit.Text = Cp_Unit[Cp_Unit_order];
            labelMea2_Unit.Text = Gp_Unit[Gp_Unit_order];
            labelMea3_Unit.Text = Rp_Unit[Rp_Unit_order];
            Control.CheckForIllegalCrossThreadCalls = false;
        }


        // GPIB instruments on the GPIB0 interface
        // Change this variable to the address of your instrument
        string VISA_ADDRESS = "GPIB0::25::INSTR";

        // Create a connection (session) to the instrument
        IMessageBasedSession session;

        // Create a formatted I/O object which will help us format the data we want to send/receive to/from the instrument
        MessageBasedFormattedIO formattedIO;

        private bool openGPIB()
        {
            try
            {
                session = GlobalResourceManager.Open(VISA_ADDRESS, AccessModes.None, 5) as IMessageBasedSession;
                formattedIO = new MessageBasedFormattedIO(session);
                MessageBox.Show("The instrument has been successfully connected on GPIB0::25", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return true;
            }
            catch (NativeVisaException visaException)
            {
                //Shell.WriteLine("Error is:\r\n{0}\r\nPress any key to exit...", visaException);
                MessageBox.Show("Please check GPIB conncetion. GPIB port of the instrument should be set as GPIB0::25", "Error: Connection Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        private string Inter()
        {
            string result = "";
            if (rbInterShort.Checked == true)
            {
                result = "SHOR,";
            }
            else if (rbInterMedium.Checked == true)
            {
                result = "MED,";
            }
            else if (rbInterLong.Checked == true)
            {
                result = "LONG,";
            }
            return result;
        }

        private string[] sendCommand(string comm)
        {
            string Comm;
            string idnResponse;
            string[] normalreturn = { DateTime.Now.ToString(), ":success" };
            try
            {
                switch (comm)
                {
                    case "APER":
                        Comm = comm + " " + Inter() + txtAveRate.Text;
                        formattedIO.WriteLine(Comm);
                        formattedIO.WriteLine("APER?");
                        idnResponse = formattedIO.ReadLine().Replace("\n", "");
                        return normalreturn;
                    case "Volt":
                        Comm = comm + " " + txtOscVoltage.Text + "mV";
                        formattedIO.WriteLine(Comm);
                        formattedIO.WriteLine("VOLT?");
                        idnResponse = formattedIO.ReadLine().Replace("\n", "");
                        return normalreturn;
                    case "Freq":
                        Comm = comm + " " + txtFrequency.Text + "Hz";
                        formattedIO.WriteLine(Comm);
                        formattedIO.WriteLine("Freq?");
                        idnResponse = formattedIO.ReadLine().Replace("\n", "");
                        return normalreturn;
                    case "Func:IMP CPG":
                        Comm = comm;
                        formattedIO.WriteLine(Comm);
                        formattedIO.WriteLine("Func:IMP?");
                        idnResponse = formattedIO.ReadLine().Replace("\n", "");
                        return normalreturn;
                    case "Func:IMP CPRP":
                        Comm = comm;
                        formattedIO.WriteLine(Comm);
                        formattedIO.WriteLine("Func:IMP?");
                        idnResponse = formattedIO.ReadLine().Replace("\n", "");
                        return normalreturn;
                    case "Func:IMP CSRS":
                        Comm = comm;
                        formattedIO.WriteLine(Comm);
                        formattedIO.WriteLine("Func:IMP?");
                        idnResponse = formattedIO.ReadLine().Replace("\n", "");
                        return normalreturn;
                    case "BIAS:STAT 1":
                        Comm = comm;
                        formattedIO.WriteLine(Comm);
                        formattedIO.WriteLine("BIAS:STAT?");
                        idnResponse = formattedIO.ReadLine().Replace("\n", "");
                        return normalreturn;
                    case "BIAS:STAT 0":
                        Comm = comm;
                        formattedIO.WriteLine(Comm);
                        formattedIO.WriteLine("BIAS:STAT?");
                        idnResponse = formattedIO.ReadLine().Replace("\n", "");
                        return normalreturn;
                    case "Fetch?":
                        string[] idnResponseFetch;
                        Comm = comm;
                        formattedIO.WriteLine(Comm);
                        idnResponse = formattedIO.ReadLine().Replace("\n", "");
                        idnResponseFetch = idnResponse.Split(new string[] { "," }, StringSplitOptions.None);
                        return idnResponseFetch;

                    //Special Fetch
                    case "Fetch?[Special]":
                        string[] idnResponseFetchSpecial;
                        Comm = "Fetch?";
                        System.Threading.Thread.Sleep(1000);
                        formattedIO.WriteLine(Comm);
                        while (1 == 1)
                        {
                            try
                            {
                                System.Threading.Thread.Sleep(1000);
                                idnResponse = formattedIO.ReadLine().Replace("\n", "");
                                break;
                            }
                            catch
                            {

                            }
                        }
                        idnResponseFetchSpecial = idnResponse.Split(new string[] { "," }, StringSplitOptions.None);
                        return idnResponseFetchSpecial;

                    case "Fetch?[Special2]":
                        string[] idnResponseFetchSpecial2;
                        Comm = "Fetch?";
                        System.Threading.Thread.Sleep(100);
                        formattedIO.WriteLine(Comm);
                        System.Threading.Thread.Sleep(100);
                        idnResponse = formattedIO.ReadLine().Replace("\n", "");
                        idnResponseFetchSpecial2 = idnResponse.Split(new string[] { "," }, StringSplitOptions.None);
                        return idnResponseFetchSpecial2;
                }
            }
            catch (NativeVisaException visaException)
            {
                return null;
            }
            return null;
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void labelFunc_Click(object sender, EventArgs e)
        {

        }

        private void buttonConn_Click(object sender, EventArgs e)
        {
            if(openGPIB())
            {

                pictureConn.Image = global::TesterAutomation_Dashboard.Properties.Resources.green;
                labelConn_Value.Text = "Success";
                pictureState.Image = global::TesterAutomation_Dashboard.Properties.Resources.green;
                labelState_Value.Text = "Running";
                GPIBstatus = true;
                state = DashboardState.running;
                initializemeasure();
            }
            else
            {
                pictureConn.Image = global::TesterAutomation_Dashboard.Properties.Resources.red;
                labelConn_Value.Text = "Fail";
                GPIBstatus = false;
                state = DashboardState.noconnection;
            }
        }

        private void pictureConn_Click(object sender, EventArgs e)
        {

        }

        private void rbInterShort_CheckedChanged(object sender, EventArgs e)
        {
            labelInt_Value.Text = Inter().TrimEnd(',');
        }

        private void rbInterMedium_CheckedChanged(object sender, EventArgs e)
        {
            labelInt_Value.Text = Inter().TrimEnd(',');
        }

        private void rbInterLong_CheckedChanged(object sender, EventArgs e)
        {
            labelInt_Value.Text = Inter().TrimEnd(',');
        }

        private void labelVersion_Click(object sender, EventArgs e)
        {
            MessageBox.Show("LRC Meter Dashboard \nDeveloped for Agilent 4284A Precision LCR Meter\nVersion 1.1\nBuilt on 05/21/2018", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        private void initializemeasure()
        {
            string Comm;
            Comm = "Freq" + " " + "1000" + "Hz";
            formattedIO.WriteLine(Comm);
            labelFreq_Value.Text ="1KHz";
            sendCommand("BIAS:STAT 1");
            Comm = "BIAS:VOLT " + "0" + "V";
            formattedIO.WriteLine(Comm);
            labelBias_Value.Text = "0";
            Comm = "Volt" + " " + "25" + "mV";
            formattedIO.WriteLine(Comm);
            labelLevel_Value.Text = "25mV";
            if (rbParallel.Checked == true)
            {
                sendCommand("Func:IMP CPG");
            }
            else if (rbSeries.Checked == true)
            {
                sendCommand("Func:IMP CSRS");
            }
            sendCommand("APER");

            buttonPause.Enabled = true;
            state = DashboardState.running;
            MonitorThreadStart = new System.Threading.ThreadStart(CGRmonitor);
            MonitorThread = new System.Threading.Thread(MonitorThreadStart);
            MonitorThread.IsBackground = true;
            MonitorThread.Start();
        }
        private void buttonSubmit_Click(object sender, EventArgs e)
        {
            string Comm;
            if(!GPIBstatus)
            {
                MessageBox.Show("Please Connect Instrument.", "No GPIB Connection", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            Comm = "Freq" + " " + txtFrequency.Text + "Hz";
            formattedIO.WriteLine(Comm);
            labelFreq_Value.Text = txtFrequency.Text + "Hz";
            sendCommand("BIAS:STAT 1");
            Comm = "BIAS:VOLT " + txtMeasVoltage.Text + "V";
            formattedIO.WriteLine(Comm);
            Comm = "Volt" + " " + txtOscVoltage.Text + "mV";
            formattedIO.WriteLine(Comm);
            if (rbParallel.Checked == true)
            {
                sendCommand("Func:IMP CPG");
            }
            else if (rbSeries.Checked == true)
            {
                sendCommand("Func:IMP CSRS");
            }
            sendCommand("APER");
        }

        public int MeaLabelModify(int MeaLabelNumber,double value)
        {
            if(PauseIsTrue==true)
            {
                return 0;
            }
            else
            {
                if(MeaLabelNumber==1)
                {
                    labelMea1_Value.Text = Convert.ToString(Cp_Unit_order==0?value:value*System.Math.Pow(10,3*Cp_Unit_order));
                }
                else if(MeaLabelNumber == 2)
                {
                    labelMea2_Value.Text = Convert.ToString(Gp_Unit_order == 0 ? value : value * System.Math.Pow(10, 3 * Gp_Unit_order));
                }
                else if(MeaLabelNumber == 3)
                {
                    labelMea3_Value.Text = Convert.ToString(Rp_Unit_order == 0 ? value : value * System.Math.Pow(10, -3 * Rp_Unit_order));
                }
                return 1;
            }
        }

        private double CpMonitor()
        {
            string[] FetchResult;
            FetchResult = sendCommand("Fetch?");
            return Convert.ToDouble(FetchResult[0]);
        }
        private double GpMonitor()
        {
            string[] FetchResult;
            FetchResult = sendCommand("Fetch?");
            return Convert.ToDouble(FetchResult[1]);
        }
        private double RpMonitor()
        {
            string[] FetchResult;
            FetchResult = sendCommand("Fetch?");
            return Convert.ToDouble(FetchResult[1]);
        }
        private void CGRmonitor()
        {
            while(true)
            {
                sendCommand("Func:IMP CPG");
                System.Threading.Thread.Sleep(rate);
                double Cp = CpMonitor();
                double Gp = GpMonitor();
                if (!ratemode)
                {
                    sendCommand("Func:IMP CPRP");
                    System.Threading.Thread.Sleep(250);
                    double Rp = RpMonitor();
                    MeaLabelModify(3, Rp);
                }
                MeaLabelModify(1, Cp);
                MeaLabelModify(2, Gp);
                if(ratemode)
                {
                    this.Invoke(new Action(delegate {
                        newrow = dtdata.NewRow();
                        newrow[0] = rate * n / 1000.0;
                        newrow[1] = Cp;
                        dtdata.Rows.InsertAt(newrow, 0);
                    }));
                    listdatax.Add(rate * n / 1000.0);
                    listdatay.Add(Cp);
                    n++;
                }
            }
        }

        private void cpGpRpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Func = "CpGpRp";
        }

        private void buttonFS1K_Click(object sender, EventArgs e)
        {
            string Comm = "Freq" + " 1000"  + "Hz";
            formattedIO.WriteLine(Comm);
            labelFreq_Value.Text = "1KHz";
        }
        private void buttonFS10K_Click(object sender, EventArgs e)
        {
            string Comm = "Freq" + " 10000" + "Hz";
            formattedIO.WriteLine(Comm);
            labelFreq_Value.Text = "10KHz";
        }
        private void buttonFS100K_Click(object sender, EventArgs e)
        {
            string Comm = "Freq" + " 100000" + "Hz";
            formattedIO.WriteLine(Comm);
            labelFreq_Value.Text = "100KHz";
        }
        private void buttonFS1M_Click(object sender, EventArgs e)
        {
            string Comm = "Freq" + " 1000000" + "Hz";
            formattedIO.WriteLine(Comm);
            labelFreq_Value.Text = "1MHz";
        }

        private void buttonUSCp_Click(object sender, EventArgs e)
        {
            Cp_Unit_order = Cp_Unit_order < 4 ? Cp_Unit_order+1 : 0;
            labelMea1_Unit.Text = Cp_Unit[Cp_Unit_order];
        }

        private void buttonUSGp_Click(object sender, EventArgs e)
        {
            Gp_Unit_order = Gp_Unit_order < 4 ? Gp_Unit_order + 1 : 0;
            labelMea2_Unit.Text = Gp_Unit[Gp_Unit_order];
        }

        private void buttonUSRp_Click(object sender, EventArgs e)
        {
            Rp_Unit_order = Rp_Unit_order < 2 ? Rp_Unit_order + 1 : 0;
            labelMea3_Unit.Text = Rp_Unit[Rp_Unit_order];
        }

        private void buttonUSDefault_Click(object sender, EventArgs e)
        {
            Cp_Unit_order = 0;
            Gp_Unit_order = 0;
            Rp_Unit_order = 0;
            labelMea1_Unit.Text = Cp_Unit[Cp_Unit_order];
            labelMea2_Unit.Text = Gp_Unit[Gp_Unit_order];
            labelMea3_Unit.Text = Rp_Unit[Rp_Unit_order];
        }

        private void buttonTest1_Click(object sender, EventArgs e)
        {
            buttonPause.Enabled = true;
            state =  DashboardState.running;
            MonitorThreadStart = new System.Threading.ThreadStart(CGRmonitor);
            MonitorThread = new System.Threading.Thread(MonitorThreadStart);
            MonitorThread.IsBackground = true;
            MonitorThread.Start();
        }

        private void buttonPause_Click(object sender, EventArgs e)
        {
            pictureState.Image = global::TesterAutomation_Dashboard.Properties.Resources.red;
            labelState_Value.Text = "Not Running";
            MonitorThread.Suspend();
            buttonPause.Enabled = false;
            buttonResume.Enabled = true;
        }

        private void buttonResume_Click(object sender, EventArgs e)
        {
            pictureState.Image = global::TesterAutomation_Dashboard.Properties.Resources.green;
            labelState_Value.Text = "Running";
            MonitorThread.Resume();
            buttonPause.Enabled = true;
            buttonResume.Enabled = false;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ratemode = true;
            if (Convert.ToInt32(textRate.Text) < 200)
            {
                MessageBox.Show("The minimum rate is 200ms.", "Error: Invalid Parameter", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            rate = (Convert.ToInt32(textRate.Text));
            dtdata = new DataTable();
            listdatax= new List<double>();
            listdatay= new List<double>();
            dtdata.Columns.Add("T(s)", typeof(float));
            dtdata.Columns.Add("C(F)", typeof(float));
            dgvdata.DataSource = dtdata;
            timerplot.Interval = 1000;
            timerplot.Enabled = true;
        }
        private void plot(Chart chart, string tag, List<double> listx, List<double> listy)
        {
            chart.Series.Clear();
            Series series = new Series(tag);
            series.ChartType = SeriesChartType.Spline;
            series.Points.DataBindXY(listx, listy);
            chart.Series.Add(series);
        }
        private void timerplot_Tick(object sender, EventArgs e)
        {
            plot(chartdata, "C-t", listdatax, listdatay);
        }

        private void buttonStop_Click(object sender, EventArgs e)
        {
            ratemode = false;
            timerplot.Enabled = false;
        }

        private void buttonSave_Click(object sender, EventArgs e)
        {
            DataGridViewToCSV(dgvdata);
        }

        private void buttonClear_Click(object sender, EventArgs e)
        {
            if(ratemode)
            {
                ratemode = false;
                timerplot.Enabled = false;
            }
            dtdata = new DataTable();
            dgvdata.DataSource = dtdata;
            n = 0;
        }
    }
}
