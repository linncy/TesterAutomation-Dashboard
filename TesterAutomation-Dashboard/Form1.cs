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

namespace TesterAutomation_Dashboard
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            labelInt_Value.Text = Inter().TrimEnd(',');

        }
        //Global Var
        string Func = "CpRp";
        bool GPIBstatus = false;

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
                session = GlobalResourceManager.Open(VISA_ADDRESS, AccessModes.None, 50000) as IMessageBasedSession;
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

        private void cpRpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Func = "CpRp";
            labelFunc_Value.Text = "Cp-Rp";
            labelMea2.Text = "Rp:";
            labelMea2_Unit.Text = "Ω";
            labelS1.Text = "Vm:";
            labelS2.Text = "Im:";
        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void labelFunc_Click(object sender, EventArgs e)
        {

        }

        private void cpDpToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Func = "CpDp";
            labelFunc_Value.Text = "Cp-Dp";
            labelMea2.Text = "Dp:";
            labelMea2_Unit.Text = "S";
        }

        private void buttonConn_Click(object sender, EventArgs e)
        {
            if(openGPIB())
            {
                pictureConn.Image = global::TesterAutomation_Dashboard.Properties.Resources.green;
                labelConn_Value.Text = "Success";
                GPIBstatus = true;
            }
            else
            {
                pictureConn.Image = global::TesterAutomation_Dashboard.Properties.Resources.red;
                labelConn_Value.Text = "Fail";
                GPIBstatus = false;
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
            MessageBox.Show("LRC Meter Dashboard \nDeveloped for Agilent 4284A Precision LCR Meter\nVersion Origin0.1\nBuilt on 11/17/2017", "Information", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void buttonSubmit_Click(object sender, EventArgs e)
        {
            string Comm;
            if(!GPIBstatus)
            {
                MessageBox.Show("Please Connect Instrument.", "No GPIB Connection", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            sendCommand("BIAS:STAT 1");
            Comm = "BIAS:VOLT " + txtMeasVoltage.Text + "V";
            formattedIO.WriteLine(Comm);
            Comm = "Volt" + " " + txtOscVoltage.Text + "mV";
            formattedIO.WriteLine(Comm);
        }
    }
}
