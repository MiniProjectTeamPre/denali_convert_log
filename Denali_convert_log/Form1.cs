using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Web.Script.Serialization;
using System.Text;
using System.Drawing;
using System.Threading;

namespace Denali_convert_log {
    public partial class Form1 : Form {
        public Form1() {
            InitializeComponent();
        }
        private string path_file_data_txt = "Denali_data_txt";
        private string path_file_data_csv = "Denali_data_csv";

        private void Form1_Load(object sender, EventArgs e) {
            if (!Directory.Exists(path_file_data_txt)) Directory.CreateDirectory(path_file_data_txt);
            if (!Directory.Exists(path_file_data_csv)) Directory.CreateDirectory(path_file_data_csv);
            Application.Idle += Application_Idle;
        }

        TestResult testResult;
        private string DataSummary = "";
        private string[] result_split_;
        private void Application_Idle(object sender, EventArgs e) {
            if (!flag_running) return;
            List<string> file_data = new List<string>();
            try {
                string[] zxc = Directory.GetFiles(path_file_data_txt);
                file_data = zxc.ToList<string>();
            } catch { }
            string s = "";
            for (int i = 0; i < 5; i++) {
                try { s = file_data[i].Replace(path_file_data_txt + "\\", ""); } catch { break; }
                string file_name = s.Replace(".txt", ".csv");
                if (!File.Exists(path_file_data_csv + "\\" + file_name)) {
                    StreamWriter swOut_ = new StreamWriter(path_file_data_csv + "\\" + file_name, true);
                    string ghf = "";
                    foreach (string zzxc in head_all) {
                        ghf += zzxc + ",";
                    }
                    swOut_.WriteLine(ghf);
                    swOut_.Close();
                }
                StreamWriter swOut = new StreamWriter(path_file_data_csv + "\\" + file_name, true);
                string[] result = File.ReadAllLines(file_data[i]);
                for (int j = 0; j < result.Count(); j++) {
                    DataSummary = "";
                    if (result[j].Contains("{\"Date\":\"")) {
                        string[] mnb = result[j].Split(',');
                        if(comboBox1.Text == FG_64T236_LF || comboBox1.Text == FG_63T336_LF) {
                            string[] replaceHead = mnb[0].Split('\t');
                            result[j] = result[j].Replace(replaceHead[0] + "\t", "");
                        } else {
                            result[j] = result[j].Replace(mnb[0] + ",", "");
                        }
                        
                        testResult = new TestResult();
                        testResult = string2json(result[j]);

                        if (comboBox1.Text == FG_64T236_LF || comboBox1.Text == FG_63T336_LF) {
                            string[] replaceHead = mnb[0].Split('\t');
                            DataSummary += replaceHead[0] + ",";
                        } else {
                            DataSummary += mnb[0] + ",";
                        }

                        if (comboBox1.Text == FG_63T305_LF || comboBox1.Text == FG_62T245_LF || comboBox1.Text == FG_63T334_LF) {
                            add_data_FG_63T305_LF();
                        }
                        if (comboBox1.Text == FG_63T306_LF || comboBox1.Text == FG_63T335_LF) {
                            add_data_FG_63T306_LF();
                        }
                        if (comboBox1.Text == FG_63T307_LF || comboBox1.Text == FG_63T336_LF) {
                            add_data_FG_63T307_LF();
                        }
                        if(comboBox1.Text == FG_64T236_LF || comboBox1.Text == FG_63T336_LF) {
                            add_data_FG_64T236_LF();
                        }

                        DataSummary += testResult.Failure + ",";
                        DataSummary += testResult.Result + ",";
                        string[] date_time_split = testResult.Date.Split('/');
                        string DateTime = date_time_split[2] + "." + date_time_split[1] + "." + date_time_split[0] + " ";
                        //string DateTime = Convert.ToDateTime(testResult.Date).ToString("yyyy.MM.dd") + " ";
                        DataSummary += DateTime + testResult.Time + ",";
                        DataSummary += comboBox2.Text + ",";
                        if (comboBox1.Text == FG_63T307_LF || comboBox1.Text == FG_63T336_LF) DataSummary += "'" + testResult.SN + ",";
                        DataSummary += testResult.LoginID + ",";
                        int seconds_end = Convert.ToInt32(TimeSpan.Parse(testResult.Time).TotalSeconds);
                        int seconds_ing = Convert.ToInt32(testResult.TestTime);
                        int seconds_old = seconds_end - seconds_ing;
                        DataSummary += DateTime + convert2time(seconds_old) + ",";
                        DataSummary += DateTime + testResult.Time + ",";
                        DataSummary += "'" + convert2time(seconds_ing) + ",";

                        if (comboBox1.Text == FG_63T305_LF || comboBox1.Text == FG_62T245_LF || comboBox1.Text == FG_63T334_LF) {
                            add_data2_FG_63T305_LF();
                        }
                        if (comboBox1.Text == FG_63T306_LF || comboBox1.Text == FG_63T335_LF) {
                            add_data2_FG_63T306_LF();
                        }
                        if (comboBox1.Text == FG_63T307_LF || comboBox1.Text == FG_63T336_LF) {
                            add_data2_FG_63T307_LF();
                        }
                        if (comboBox1.Text == FG_64T236_LF || comboBox1.Text == FG_63T336_LF) {
                            add_data2_FG_64T236_LF();
                        }

                    } else {
                        //DataSummary += result[j];
                        result_split_ = result[j].Split(',');
                        if (result_split_.Count() < 3) continue;
                        if (result_split_[0].Count() != 10) continue;
                        if (comboBox1.Text == FG_63T305_LF || comboBox1.Text == FG_62T245_LF || comboBox1.Text == FG_63T334_LF) {
                            add_data3_FG_63T305_LF();
                        }
                        if (comboBox1.Text == FG_63T306_LF || comboBox1.Text == FG_63T335_LF) {
                            add_data3_FG_63T306_LF();
                        }
                        if (comboBox1.Text == FG_63T307_LF || comboBox1.Text == FG_63T336_LF) {
                            add_data3_FG_63T307_LF();
                        }
                    }
                    while (true) {
                        try { swOut.WriteLine(DataSummary); break; } catch { MessageBox.Show("_กรุณาปิด log file csv ก่อน"); }
                    }
                }
                swOut.Close();
                File.Delete(file_data[i]);
            }
            Thread.Sleep(2000);
        }
        private TestResult string2json(string input) {
            TestResult result = new TestResult();
            string[] split_ResultString = input.Replace("\"ResultString\":[", "฿").Split('฿');
            List<string> values = new List<string>();
            List<string> keys = new List<string>();
            string pattern = @"\""(?<key>[^\""]+)\""\:\""?(?<value>[^\"",}]+)\""?\,?";
            foreach (Match m in Regex.Matches(split_ResultString[0] + "}", pattern)) {
                if (m.Success) {
                    values.Add(m.Groups["value"].Value);
                    //keys.Add(m.Groups["key"].Value);
                }
            }
            result.Date = values[0];
            result.Time = values[1];
            result.LoginID = values[2];
            result.VersionSW = values[3];
            result.VersionFW = values[4];
            result.VersionSpec = values[5];
            result.TestTime = values[6];
            result.LoadIn = values[7];
            result.Mode = values[8];
            result.Result = values[9];
            result.SN = values[10];
            try { result.Failure = values[11]; } catch { result.Failure = ""; }

            List<ResultStepDetail> resultString = new List<ResultStepDetail>();
            split_ResultString[1] = split_ResultString[1].Replace(",]}", "");
            string[] step = split_ResultString[1].Replace("},{", "฿").Split('฿');
            step[0] = step[0] + "}";
            step[step.Count() - 1] = "{" + step[step.Count() - 1];
            for (int h = 1; h < step.Count() - 1; h++) {
                step[h] = "{" + step[h] + "}";
            }
            for (int h = 0; h < step.Count(); h++) {
                values.Clear();
                foreach (Match m in Regex.Matches(step[h], pattern)) {
                    if (m.Success) {
                        values.Add(m.Groups["value"].Value);
                    }
                }
                while (values.Count < 5) { values.Add(""); }
                resultString.Add(new ResultStepDetail() { Step = values[0], Description = values[1], Tolerance = values[2], Measured = values[3], Result = values[4] });
            }
            result.ResultString = resultString;

            return result;
        }
        private void add_data_FG_63T305_LF() {
            DataSummary += getResult("2.2") + ",";
            DataSummary += getResult("2.4") + ",";
            DataSummary += getResult("2.5") + ",";
            DataSummary += "'" + getResult("3.15") + ",";
            DataSummary += "'" + getResult("3.14") + ",";
            DataSummary += "'" + getResult("3.13") + ",";
            DataSummary += getResult("4.5") + ",";
            DataSummary += getResult("4.4") + ",";
            DataSummary += getResult("2.12") + ",";
            DataSummary += getResult("2.10.") + ",";
        }
        private void add_data_FG_63T306_LF() {
            DataSummary += getResult("2.2") + ",";
            DataSummary += getResult("2.4") + ",";
            DataSummary += getResult("2.5") + ",";
            DataSummary += "'" + getResult("3.15") + ",";
            DataSummary += "'" + getResult("3.14") + ",";
            DataSummary += "'" + getResult("3.13") + ",";
            DataSummary += getResult("4.5") + ",";
            DataSummary += getResult("4.4") + ",";
            DataSummary += getResult("2.12") + ",";
            DataSummary += getResult("2.10.") + ",";
        }
        private void add_data_FG_63T307_LF() {
            DataSummary += getResult("2.2") + ",";
            DataSummary += getResult("2.4") + ",";
            DataSummary += getResult("2.5") + ",";
            DataSummary += "'" + getResult("3.15") + ",";
            DataSummary += "'" + getResult("3.14") + ",";
            DataSummary += "'" + getResult("3.13") + ",";
            DataSummary += getResult("4.5") + ",";
            DataSummary += getResult("4.4") + ",";
            DataSummary += getResult("2.12") + ",";
            DataSummary += getResult("2.10.") + ",";
        }
        private void add_data_FG_64T236_LF() {
            DataSummary += getResult("2.4") + ",";
            DataSummary += getResult("2.3") + ",";
            DataSummary += "-" + ","; //measure current running [tp102 +] [tp103 -] [mA]
            DataSummary += "'" + getResult("3.14") + ",";
            DataSummary += "'" + getResult("3.13") + ",";
            DataSummary += "'" + getResult("3.12") + ",";
            DataSummary += "-" + ",";//measure crystal 32 [PPM]
            DataSummary += "-" + ",";//measure crystal 32 [kHz]
            DataSummary += getResult("2.14") + ",";
            DataSummary += getResult("2.12") + ",";
        }
        private void add_data2_FG_63T305_LF() {
            DataSummary += "'" + getResult("3.12") + ",";
            DataSummary += getResult("2.6") + ",";
            DataSummary += getResult("3.3") + ",";
            DataSummary += getResult("3.4") + ",";
            DataSummary += getResult("2.8") + ",";
            DataSummary += getResult("3.5") + ",";
            DataSummary += getResult("3.6") + ",";
            DataSummary += getResult("3.8") + ",";
            DataSummary += getResult("3.9") + ",";
            DataSummary += getResult("3.10") + ",";
            DataSummary += getResult("3.7.1") + ",";
            DataSummary += getResult("3.7.2") + ",";
            DataSummary += getResult("3.7.3") + ",";
            DataSummary += getResult("3.7.4") + ",";
            DataSummary += getResult("3.7.5") + ",";
            DataSummary += getResult("3.7.6") + ",";
            DataSummary += "'" + getResult("3.16") + ",";
        }
        private void add_data2_FG_63T306_LF() {
            DataSummary += getResult("3.17") + ",";
            DataSummary += "'" + getResult("3.18") + ",";
            DataSummary += getResult("3.19") + ",";
            DataSummary += "'" + getResult("3.12") + ",";
            DataSummary += getResult("2.6") + ",";
            DataSummary += getResult("3.3") + ",";
            DataSummary += getResult("3.4") + ",";
            DataSummary += getResult("2.8") + ",";
            DataSummary += getResult("3.5") + ",";
            DataSummary += getResult("3.6") + ",";
            DataSummary += getResult("3.8") + ",";
            DataSummary += getResult("3.9") + ",";
            DataSummary += getResult("3.10") + ",";
            DataSummary += getResult("3.7.1") + ",";
            DataSummary += getResult("3.7.2") + ",";
            DataSummary += getResult("3.7.3") + ",";
            DataSummary += getResult("3.7.4") + ",";
            DataSummary += getResult("3.7.5") + ",";
            DataSummary += getResult("3.7.6") + ",";
            DataSummary += "'" + getResult("3.16") + ",";
        }
        private void add_data2_FG_63T307_LF() {
            DataSummary += getResult("3.17") + ",";
            DataSummary += "'" + getResult("3.18") + ",";
            DataSummary += getResult("3.19") + ",";
            DataSummary += "'" + getResult("3.12") + ",";
            DataSummary += getResult("2.6") + ",";
            DataSummary += getResult("3.3") + ",";
            DataSummary += getResult("3.4") + ",";
            DataSummary += getResult("2.8") + ",";
            DataSummary += getResult("3.5") + ",";
            DataSummary += getResult("3.6") + ",";
            DataSummary += getResult("3.8") + ",";
            DataSummary += getResult("3.9") + ",";
            DataSummary += getResult("3.10") + ",";
            DataSummary += getResult("3.7.1") + ",";
            DataSummary += getResult("3.7.2") + ",";
            DataSummary += getResult("3.7.3") + ",";
            DataSummary += getResult("3.7.4") + ",";
            DataSummary += getResult("3.7.5") + ",";
            DataSummary += getResult("3.7.6") + ",";
            DataSummary += "'" + getResult("3.16") + ",";
        }
        private void add_data2_FG_64T236_LF() {
            DataSummary += "'" + getResult("3.11") + ",";
            DataSummary += "-" + ",";
            DataSummary += getResult("3.3") + ",";
            DataSummary += getResult("3.4") + ",";
            DataSummary += getResult("2.6") + ",";
            DataSummary += getResult("3.5") + ",";
            DataSummary += getResult("3.6") + ",";
            DataSummary += getResult("3.7") + ",";
            DataSummary += getResult("3.8") + ",";
            DataSummary += getResult("3.9") + ",";
            DataSummary += "-" + ",";
            DataSummary += "-" + ",";
            DataSummary += "-" + ",";
            DataSummary += "-" + ",";
            DataSummary += "-" + ",";
            DataSummary += "-" + ",";
            DataSummary += "'" + getResult("3.15") + ",";
        }
        private void add_data3_FG_63T305_LF() {
            DataSummary += result_split_[0] + ",";
            DataSummary += result_split_[2] + ",";
            DataSummary += result_split_[3] + ",";
            DataSummary += result_split_[4] + ",";
            DataSummary += "'" + result_split_[22] + ",";
            DataSummary += "'" + result_split_[23] + ",";
            DataSummary += "'" + result_split_[24] + ",";
            DataSummary += result_split_[26] + ",";
            DataSummary += result_split_[27] + ",";
            DataSummary += result_split_[31] + ",";
            DataSummary += result_split_[32] + ",";
            DataSummary += result_split_[33] + ",";
            DataSummary += result_split_[34] + ",";
            DataSummary += result_split_[35] + ",";
            DataSummary += result_split_[38] + ",";
            DataSummary += result_split_[40] + ",";
            DataSummary += result_split_[42] + ",";
            DataSummary += result_split_[43] + ",";
            DataSummary += "'" + result_split_[44] + ",";
            DataSummary += "'" + result_split_[50] + ",";
            DataSummary += result_split_[51] + ",";
            DataSummary += result_split_[52] + ",";
            DataSummary += result_split_[53] + ",";
            DataSummary += result_split_[54] + ",";
            DataSummary += result_split_[55] + ",";
            DataSummary += result_split_[56] + ",";
            DataSummary += result_split_[57] + ",";
            DataSummary += result_split_[58] + ",";
            DataSummary += result_split_[59] + ",";
            DataSummary += result_split_[60] + ",";
            DataSummary += result_split_[61] + ",";
            DataSummary += result_split_[62] + ",";
            DataSummary += result_split_[63] + ",";
            DataSummary += result_split_[64] + ",";
            DataSummary += result_split_[65] + ",";
            DataSummary += "'" + result_split_[66] + ",";
        }
        private void add_data3_FG_63T306_LF() {
            DataSummary += result_split_[0] + ",";
            DataSummary += result_split_[2] + ",";
            DataSummary += result_split_[3] + ",";
            DataSummary += result_split_[4] + ",";
            DataSummary += "'" + result_split_[22] + ",";
            DataSummary += "'" + result_split_[23] + ",";
            DataSummary += "'" + result_split_[24] + ",";
            DataSummary += result_split_[26] + ",";
            DataSummary += result_split_[27] + ",";
            DataSummary += result_split_[31] + ",";
            DataSummary += result_split_[32] + ",";
            DataSummary += result_split_[33] + ",";
            DataSummary += result_split_[34] + ",";
            DataSummary += result_split_[35] + ",";
            DataSummary += result_split_[38] + ",";
            DataSummary += result_split_[40] + ",";
            DataSummary += result_split_[42] + ",";
            DataSummary += result_split_[43] + ",";
            DataSummary += "'" + result_split_[44] + ",";

            DataSummary += "'" + result_split_[46] + ",";
            DataSummary += result_split_[47] + ",";
            DataSummary += result_split_[48] + ",";

            DataSummary += "'" + result_split_[50] + ",";
            DataSummary += result_split_[51] + ",";
            DataSummary += result_split_[52] + ",";
            DataSummary += result_split_[53] + ",";
            DataSummary += result_split_[54] + ",";
            DataSummary += result_split_[55] + ",";
            DataSummary += result_split_[56] + ",";
            DataSummary += result_split_[57] + ",";
            DataSummary += result_split_[58] + ",";
            DataSummary += result_split_[59] + ",";
            DataSummary += result_split_[60] + ",";
            DataSummary += result_split_[61] + ",";
            DataSummary += result_split_[62] + ",";
            DataSummary += result_split_[63] + ",";
            DataSummary += result_split_[64] + ",";
            DataSummary += result_split_[65] + ",";
            DataSummary += "'" + result_split_[66] + ",";
        }
        private void add_data3_FG_63T307_LF() {
            DataSummary += result_split_[0] + ",";
            DataSummary += result_split_[2] + ",";
            DataSummary += result_split_[3] + ",";
            DataSummary += result_split_[4] + ",";
            DataSummary += "'" + result_split_[22] + ",";
            DataSummary += "'" + result_split_[23] + ",";
            DataSummary += "'" + result_split_[24] + ",";
            DataSummary += result_split_[26] + ",";
            DataSummary += result_split_[27] + ",";
            DataSummary += result_split_[31] + ",";
            DataSummary += result_split_[32] + ",";
            DataSummary += result_split_[33] + ",";
            DataSummary += result_split_[34] + ",";
            DataSummary += result_split_[35] + ",";
            DataSummary += result_split_[38] + ",";
            DataSummary += "'" + result_split_[39] + ",";
            DataSummary += result_split_[40] + ",";//Operator
            DataSummary += result_split_[42] + ",";
            DataSummary += result_split_[43] + ",";
            DataSummary += "'" + result_split_[44] + ",";

            DataSummary += "'" + result_split_[46] + ",";
            DataSummary += result_split_[47] + ",";
            DataSummary += result_split_[48] + ",";

            DataSummary += "'" + result_split_[50] + ",";
            DataSummary += result_split_[51] + ",";
            DataSummary += result_split_[52] + ",";
            DataSummary += result_split_[53] + ",";
            DataSummary += result_split_[54] + ",";
            DataSummary += result_split_[55] + ",";
            DataSummary += result_split_[56] + ",";
            DataSummary += result_split_[57] + ",";
            DataSummary += result_split_[58] + ",";
            DataSummary += result_split_[59] + ",";
            DataSummary += result_split_[60] + ",";
            DataSummary += result_split_[61] + ",";
            DataSummary += result_split_[62] + ",";
            DataSummary += result_split_[63] + ",";
            DataSummary += result_split_[64] + ",";
            DataSummary += result_split_[65] + ",";
            DataSummary += "'" + result_split_[66] + ",";
        }

        private static void DelaymS(int mS) {
            Stopwatch stopwatchDelaymS = new Stopwatch();
            stopwatchDelaymS.Restart();
            while (mS > stopwatchDelaymS.ElapsedMilliseconds) {
                if (!stopwatchDelaymS.IsRunning) stopwatchDelaymS.Start();
                Application.DoEvents();
                Thread.Sleep(50);
            }
            stopwatchDelaymS.Stop();
        }
        private string convert2time(int testTime) {
            int testTime_hh = 0;
            int testTime_mm = testTime / 60;
            int testTime_ss = testTime % 60;
            if (testTime_mm > 59) {
                testTime_hh = testTime_mm / 60;
                testTime_mm = testTime_mm % 60;
            }
            return testTime_hh.ToString("00") + ":" + testTime_mm.ToString("00") + ":" + testTime_ss.ToString("00");
        }
        private string getResult(string head) {
            string lkj = "";
            for (int kk = 0; kk < testResult.ResultString.Count; kk++) {
                if (testResult.ResultString[kk].Step != head) continue;
                lkj = testResult.ResultString[kk].Measured;
                testResult.ResultString.Remove(testResult.ResultString[kk]);
            }
            return lkj;
        }

        private bool flag_running = false;
        private void timer1_Tick(object sender, EventArgs e) {
            timer1.Enabled = false;
            if (!flag_running) { timer1.Enabled = true; return; }
            this.BackgroundImage = Properties.Resources.file_01;
            DelaymS(500);
            this.BackgroundImage = Properties.Resources.file_02;
            DelaymS(500);
            this.BackgroundImage = Properties.Resources.file_03;
            DelaymS(500);
            this.BackgroundImage = Properties.Resources.file_04;
            DelaymS(500);
            this.BackgroundImage = Properties.Resources.file_00;
            DelaymS(1000);
            timer1.Enabled = true;
        }
        List<string> head_all = new List<string>();
        private string FG_62T245_LF = "FG-62T245-LF";
        private string FG_63T305_LF = "FG-63T305-LF";
        private string FG_63T306_LF = "FG-63T306-LF";
        private string FG_63T307_LF = "FG-63T307-LF";
        private string FG_63T334_LF = "FG-63T334-LF";
        private string FG_63T335_LF = "FG-63T335-LF";
        private string FG_63T336_LF = "FG-63T336-LF";
        private string FG_64T236_LF = "FG-64T236-LF";
        private string FG_64T237_LF = "FG-64T237-LF";
        private void button1_Click(object sender, EventArgs e) {
            if (comboBox1.Text == null || comboBox1.Text == "") { MessageBox.Show("_กรุณาเลือก FG ก่อน"); return; }
            if(button1.Text == "RUN") {
                button1.Text = "STOP";
                button1.BackColor = Color.Red;
                flag_running = true;
            } else {
                button1.Text = "RUN";
                button1.BackColor = Color.Aqua;
                flag_running = false;
            }
            head_all.Clear();
            if (comboBox1.Text == FG_63T305_LF || comboBox1.Text == FG_62T245_LF || comboBox1.Text == FG_63T334_LF) {
                head_all.Add("Column1");
                head_all.Add("Firmware_CRC32");
                head_all.Add("Battery_Volt_Test");
                head_all.Add("Running_Curr_Test");
                head_all.Add("Modem_FW_Test");
                head_all.Add("Modem_IMEI_Test");
                head_all.Add("SIM_ICCID_Test");
                head_all.Add("Crystal_Test_ppm");
                head_all.Add("Crystal_Test_kHz");
                head_all.Add("Standby_Curr_Test");
                head_all.Add("Sleep_Curr_Test");
                head_all.Add("fail");
                head_all.Add("Final_Result");
                head_all.Add("DATE_TIME");
                head_all.Add("TESTER_ID");
                head_all.Add("Operator");
                head_all.Add("Test_Start_Time");
                head_all.Add("Test_Finish_Time");
                head_all.Add("Test_Total_Time");
                head_all.Add("Check_Version_Hardware");
                head_all.Add("3V7_LTE_Volt_Test");
                head_all.Add("Check_SW1_Test");
                head_all.Add("Check_SW2_Test");
                head_all.Add("Processor_Functional");
                head_all.Add("Measure_Light_Sensor_Dark");
                head_all.Add("Measure_Light_Sensor_Light");
                head_all.Add("Measure_Temp_Sensor");
                head_all.Add("Measure_Humidity_Sensor");
                head_all.Add("Check_Memory_EEPROM");
                head_all.Add("Led1_Red_On");
                head_all.Add("Led1_Green_On");
                head_all.Add("Led1_Blue_On");
                head_all.Add("Led2_Red_On");
                head_all.Add("Led2_Green_On");
                head_all.Add("Led2_Blue_On");
                head_all.Add("Check_Version_Application");
            }
            if (comboBox1.Text == FG_63T306_LF || comboBox1.Text == FG_63T335_LF) {
                head_all.Add("Column1");
                head_all.Add("Firmware_CRC32");
                head_all.Add("Battery_Volt_Test");
                head_all.Add("Running_Curr_Test");
                head_all.Add("Modem_FW_Test");
                head_all.Add("Modem_IMEI_Test");
                head_all.Add("SIM_ICCID_Test");
                head_all.Add("Crystal_Test_ppm");
                head_all.Add("Crystal_Test_kHz");
                head_all.Add("Standby_Curr_Test");
                head_all.Add("Sleep_Curr_Test");
                head_all.Add("fail");
                head_all.Add("Final_Result");
                head_all.Add("DATE_TIME");
                head_all.Add("TESTER_ID");
                head_all.Add("Operator");
                head_all.Add("Test_Start_Time");
                head_all.Add("Test_Finish_Time");
                head_all.Add("Test_Total_Time");
                head_all.Add("Check_Accelerometer");
                head_all.Add("NFC_ID_Test");
                head_all.Add("Measure_Pressure_Sensor");
                head_all.Add("Check_Version_Hardware");
                head_all.Add("3V7_LTE_Volt_Test");
                head_all.Add("Check_SW1_Test");
                head_all.Add("Check_SW2_Test");
                head_all.Add("Processor_Functional");
                head_all.Add("Measure_Light_Sensor_Dark");
                head_all.Add("Measure_Light_Sensor_Light");
                head_all.Add("Measure_Temp_Sensor");
                head_all.Add("Measure_Humidity_Sensor");
                head_all.Add("Check_Memory_EEPROM");
                head_all.Add("Led1_Red_On");
                head_all.Add("Led1_Green_On");
                head_all.Add("Led1_Blue_On");
                head_all.Add("Led2_Red_On");
                head_all.Add("Led2_Green_On");
                head_all.Add("Led2_Blue_On");
                head_all.Add("Check_Version_Application");
            }
            if (comboBox1.Text == FG_63T307_LF || comboBox1.Text == FG_63T336_LF) {
                head_all.Add("Column1");
                head_all.Add("Firmware_CRC32");
                head_all.Add("Battery_Volt_Test");
                head_all.Add("Running_Curr_Test");
                head_all.Add("Modem_FW_Test");
                head_all.Add("Modem_IMEI_Test");
                head_all.Add("SIM_ICCID_Test");
                head_all.Add("Crystal_Test_ppm");
                head_all.Add("Crystal_Test_kHz");
                head_all.Add("Standby_Curr_Test");
                head_all.Add("Sleep_Curr_Test");
                head_all.Add("fail");
                head_all.Add("Final_Result");
                head_all.Add("DATE_TIME");
                head_all.Add("TESTER_ID");
                head_all.Add("TEAM_SN");
                head_all.Add("Operator");
                head_all.Add("Test_Start_Time");
                head_all.Add("Test_Finish_Time");
                head_all.Add("Test_Total_Time");
                head_all.Add("Check_Accelerometer");
                head_all.Add("NFC_ID_Test");
                head_all.Add("Measure_Pressure_Sensor");
                head_all.Add("Check_Version_Hardware");
                head_all.Add("3V7_LTE_Volt_Test");
                head_all.Add("Check_SW1_Test");
                head_all.Add("Check_SW2_Test");
                head_all.Add("Processor_Functional");
                head_all.Add("Measure_Light_Sensor_Dark");
                head_all.Add("Measure_Light_Sensor_Light");
                head_all.Add("Measure_Temp_Sensor");
                head_all.Add("Measure_Humidity_Sensor");
                head_all.Add("Check_Memory_EEPROM");
                head_all.Add("Led1_Red_On");
                head_all.Add("Led1_Green_On");
                head_all.Add("Led1_Blue_On");
                head_all.Add("Led2_Red_On");
                head_all.Add("Led2_Green_On");
                head_all.Add("Led2_Blue_On");
                head_all.Add("Check_Version_Application");
            }
            if (comboBox1.Text == FG_64T236_LF || comboBox1.Text == FG_64T237_LF) {
                head_all.Add("Column1");
                head_all.Add("Firmware_CRC32");
                head_all.Add("Battery_Volt_Test");
                head_all.Add("Running_Curr_Test");
                head_all.Add("Modem_FW_Test");
                head_all.Add("Modem_IMEI_Test");
                head_all.Add("SIM_ICCID_Test");
                head_all.Add("Crystal_Test_ppm");
                head_all.Add("Crystal_Test_kHz");
                head_all.Add("Standby_Curr_Test");
                head_all.Add("Sleep_Curr_Test");
                head_all.Add("fail");
                head_all.Add("Final_Result");
                head_all.Add("DATE_TIME");
                head_all.Add("TESTER_ID");
                head_all.Add("Operator");
                head_all.Add("Test_Start_Time");
                head_all.Add("Test_Finish_Time");
                head_all.Add("Test_Total_Time");
                head_all.Add("Check_Version_Hardware");
                head_all.Add("3V7_LTE_Volt_Test");
                head_all.Add("Check_SW1_Test");
                head_all.Add("Check_SW2_Test");
                head_all.Add("Processor_Functional");
                head_all.Add("Measure_Light_Sensor_Dark");
                head_all.Add("Measure_Light_Sensor_Light");
                head_all.Add("Measure_Temp_Sensor");
                head_all.Add("Measure_Humidity_Sensor");
                head_all.Add("Check_Memory_EEPROM");
                head_all.Add("Led1_Red_On");
                head_all.Add("Led1_Green_On");
                head_all.Add("Led1_Blue_On");
                head_all.Add("Led2_Red_On");
                head_all.Add("Led2_Green_On");
                head_all.Add("Led2_Blue_On");
                head_all.Add("Check_Version_Application");
            }
        }
    }

    public class TestResult {
        public string Date { get; set; }
        public string Time { get; set; }
        public string LoginID { get; set; }
        public string VersionSW { get; set; }
        public string VersionFW { get; set; }
        public string VersionSpec { get; set; }
        public string TestTime { get; set; }
        public string LoadIn { get; set; }
        public string Mode { get; set; }
        public string Result { get; set; }
        public string SN { get; set; }
        public string Failure { get; set; }
        public List<ResultStepDetail> ResultString { get; set; }
    }
    public class ResultStepDetail {
        public string Step { get; set; }
        public string Description { get; set; }
        public string Tolerance { get; set; }
        public string Measured { get; set; }
        public string Result { get; set; }
    }
}
