using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.IO.Ports;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using MyActive;

namespace MyActiveX
{
    //不可改变
    [Guid("D5CDCDF1-5699-42E6-B81D-8FBB95092B99"), ComVisible(true)]
    //[Guid("073A987E-2A7C-4874-8BEE-321E04F4E84E"), ComVisible(true)]
    public partial class MyActiveXControl : UserControl, IObjectSafety
    {

        public MyActiveXControl()
        {
            InitializeComponent();
            MessageBox.Show("ActiveX test");
            this.MinimumSize = this.Size;
            init();
        }

        public delegate void HandleInterfaceUpdataDelegate(string text);//定义一个委托

        private HandleInterfaceUpdataDelegate interfaceUpdataHandle;//声明

        bool isClose = false;//是否关闭

        #region IObjectSafety 成员
        private const string _IID_IDispatch = "{00020400-0000-0000-C000-000000000046}";
        private const string _IID_IDispatchEx = "{a6ef9860-c720-11d0-9337-00a0c90dcaa9}";
        private const string _IID_IPersistStorage = "{0000010A-0000-0000-C000-000000000046}";
        private const string _IID_IPersistStream = "{00000109-0000-0000-C000-000000000046}";
        private const string _IID_IPersistPropertyBag = "{37D84F60-42CB-11CE-8135-00AA004BB851}";

        private const int INTERFACESAFE_FOR_UNTRUSTED_CALLER = 0x00000001;
        private const int INTERFACESAFE_FOR_UNTRUSTED_DATA = 0x00000002;
        private const int S_OK = 0;
        private const int E_FAIL = unchecked((int)0x80004005);
        private const int E_NOINTERFACE = unchecked((int)0x80004002);

        private bool _fSafeForScripting = true;
        private bool _fSafeForInitializing = true;

        public int GetInterfaceSafetyOptions(ref Guid riid, ref int pdwSupportedOptions, ref int pdwEnabledOptions)
        {
            int Rslt = E_FAIL;

            string strGUID = riid.ToString("B");
            pdwSupportedOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER | INTERFACESAFE_FOR_UNTRUSTED_DATA;
            switch (strGUID)
            {
                case _IID_IDispatch:
                case _IID_IDispatchEx:
                    Rslt = S_OK;
                    pdwEnabledOptions = 0;
                    if (_fSafeForScripting == true)
                        pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_CALLER;
                    break;
                case _IID_IPersistStorage:
                case _IID_IPersistStream:
                case _IID_IPersistPropertyBag:
                    Rslt = S_OK;
                    pdwEnabledOptions = 0;
                    if (_fSafeForInitializing == true)
                        pdwEnabledOptions = INTERFACESAFE_FOR_UNTRUSTED_DATA;
                    break;
                default:
                    Rslt = E_NOINTERFACE;
                    break;
            }

            return Rslt;
        }

        public int SetInterfaceSafetyOptions(ref Guid riid, int dwOptionSetMask, int dwEnabledOptions)
        {
            int Rslt = E_FAIL;
            string strGUID = riid.ToString("B");
            switch (strGUID)
            {
                case _IID_IDispatch:
                case _IID_IDispatchEx:
                    if (((dwEnabledOptions & dwOptionSetMask) == INTERFACESAFE_FOR_UNTRUSTED_CALLER) && (_fSafeForScripting == true))
                        Rslt = S_OK;
                    break;
                case _IID_IPersistStorage:
                case _IID_IPersistStream:
                case _IID_IPersistPropertyBag:
                    if (((dwEnabledOptions & dwOptionSetMask) == INTERFACESAFE_FOR_UNTRUSTED_DATA) && (_fSafeForInitializing == true))
                        Rslt = S_OK;
                    break;
                default:
                    Rslt = E_NOINTERFACE;
                    break;
            }

            return Rslt;
        }

        #endregion


        #region 初始化
        private void init()
        {
            #region 控件初始化
            btnSend.Enabled = false;
            cbbComList.Items.AddRange(SerialPort.GetPortNames());

            #endregion
            #region 设置保存初始化
            //cbbComList.DataBindings.Add(new System.Windows.Forms.Binding("Text", global::ZUART.Properties.Settings.Default, "cbbComList", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            //cbbBaudRate.DataBindings.Add(new System.Windows.Forms.Binding("Text", global::ZUART.Properties.Settings.Default, "cbbBaudRate", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            //cbbDataBits.DataBindings.Add(new System.Windows.Forms.Binding("Text", global::ZUART.Properties.Settings.Default, "cbbDataBits", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            //cbbStopBits.DataBindings.Add(new System.Windows.Forms.Binding("Text", global::ZUART.Properties.Settings.Default, "cbbStopBits", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            //cbbParity.DataBindings.Add(new System.Windows.Forms.Binding("Text", global::ZUART.Properties.Settings.Default, "cbbParity", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            //txtSendData.DataBindings.Add(new System.Windows.Forms.Binding("Text", global::ZUART.Properties.Settings.Default, "txtSendData", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            //txtAutoSendms.DataBindings.Add(new System.Windows.Forms.Binding("Text", global::ZUART.Properties.Settings.Default, "txtAutoSendms", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            //chkAutoLine.DataBindings.Add(new System.Windows.Forms.Binding("Checked", global::ZUART.Properties.Settings.Default, "chkAutoLine", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            //chkShowTime.DataBindings.Add(new System.Windows.Forms.Binding("Checked", global::ZUART.Properties.Settings.Default, "chkShowTime", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            //chkRecSend.DataBindings.Add(new System.Windows.Forms.Binding("Checked", global::ZUART.Properties.Settings.Default, "chkRecSend", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            //chkfromFileSend.DataBindings.Add(new System.Windows.Forms.Binding("Checked", global::ZUART.Properties.Settings.Default, "chkfromFileSend", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            //chkAutoAddSend.DataBindings.Add(new System.Windows.Forms.Binding("Checked", global::ZUART.Properties.Settings.Default, "chkAutoAddSend", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            //chkAutoCleanSend.DataBindings.Add(new System.Windows.Forms.Binding("Checked", global::ZUART.Properties.Settings.Default, "chkAutoCleanSend", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            //chkAutoSend.DataBindings.Add(new System.Windows.Forms.Binding("Checked", global::ZUART.Properties.Settings.Default, "chkAutoSend", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            //rbtnHex.DataBindings.Add(new System.Windows.Forms.Binding("Checked", global::ZUART.Properties.Settings.Default, "rbtnHex", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            //rbtnASCII.DataBindings.Add(new System.Windows.Forms.Binding("Checked", global::ZUART.Properties.Settings.Default, "rbtnASCII", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            //rbtnUTF8.DataBindings.Add(new System.Windows.Forms.Binding("Checked", global::ZUART.Properties.Settings.Default, "rbtnUTF8", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            //rbtnUnicode.DataBindings.Add(new System.Windows.Forms.Binding("Checked", global::ZUART.Properties.Settings.Default, "rbtnUnicode", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            //rbtnSendHex.DataBindings.Add(new System.Windows.Forms.Binding("Checked", global::ZUART.Properties.Settings.Default, "rbtnSendHex", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            //rbtnSendASCII.DataBindings.Add(new System.Windows.Forms.Binding("Checked", global::ZUART.Properties.Settings.Default, "rbtnSendASCII", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            //rbtnSendUTF8.DataBindings.Add(new System.Windows.Forms.Binding("Checked", global::ZUART.Properties.Settings.Default, "rbtnSendUTF8", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            //rbtnSendUnicode.DataBindings.Add(new System.Windows.Forms.Binding("Checked", global::ZUART.Properties.Settings.Default, "rbtnSendUnicode", true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));

            //for (int i = 0; i < ListSend_Count; i++)
            //{
            //    ListSendTextBox[i].DataBindings.Add(new System.Windows.Forms.Binding("Text", global::ZUART.Properties.Settings.Default, "ListSend_Text" + i, true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            //    ListSendCheckBox[i].DataBindings.Add(new System.Windows.Forms.Binding("Checked", global::ZUART.Properties.Settings.Default, "ListSend_Hex" + i, true, System.Windows.Forms.DataSourceUpdateMode.OnPropertyChanged));
            //}
            #endregion
        }
        #endregion


        #region 打开串口按钮
        private void btnOpen_Click(object sender, EventArgs e)
        {
            if (cbbComList.Items.Count <= 0)
            {
                //MessageBox.Show("没有发现串口,请检查线路！");
                Log.Text = "没有发现串口,请检查线路！";
                return;
            }
            //if (ComDevice.IsOpen == false)
            if (btnOpen.Text == "打开串口")
            {
                ComOpen(true);
            }
            else
            {
                ComOpen(false);
            }
        }

        #endregion
        #region 串口参数设置监听 打开串口后设置串口参数立即生效(关闭后重新打开串口)
        private void cbbComList_DropDown(object sender, EventArgs e)
        {
            cbbComList.Items.Clear();
            cbbComList.Items.AddRange(SerialPort.GetPortNames());
        }
        private void cbbComSetChange(object sender, EventArgs e)
        {
            if (ComDevice.IsOpen)
            {
                btnOpen_Click(null, null);
                btnOpen_Click(null, null);
            }

        }
        #endregion

        #region 打开/关闭串口
        public String ComOpen(bool b)
        {
            String str = "";
            try
            {
                if (b)
                {
                    ComDevice.PortName = cbbComList.Text.ToString();
                    ComDevice.BaudRate = Convert.ToInt32(cbbBaudRate.Text.ToString());
                    ComDevice.Parity = (Parity)Convert.ToInt32(cbbParity.SelectedIndex.ToString());
                    ComDevice.DataBits = Convert.ToInt32(cbbDataBits.Text.ToString());
                    ComDevice.StopBits = (StopBits)Convert.ToInt32(cbbStopBits.Text.ToString());
                    ComDevice.Open();
                }
                else
                {
                    if (ComDevice.IsOpen)
                        ComDevice.Close();
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                //AddContent("串口打开错误:" + ex.Message + "\r\n");
                str = "错误:" + ex.Message + "\r\n";

            }
            finally
            {
                if (ComDevice.IsOpen == false)
                {
                    btnSend.Enabled = false;
                    btnOpen.Text = "打开串口";
                    btnOpen.Image = MyActive.Properties.Resources.close;
                    Log.Text = "串口已关闭";
                }
                else
                {
                    btnSend.Enabled = true;
                    btnOpen.Text = "关闭串口";
                    btnOpen.Image = MyActive.Properties.Resources.open;
                    // 串口号,波特率,数据位,停止位.校验位
                    Log.Text = "串口已开启:" + cbbComList.Text + "," + cbbBaudRate.Text + "," + cbbDataBits.Text + "," + cbbStopBits.Text + "," + cbbParity.Text;
                }
            }
            return str;
        }
        #endregion

        #region 发送数据
        public bool SendData(byte[] data)
        {
            if (ComDevice.IsOpen)
            {
                try
                {
                    ComDevice.Write(data, 0, data.Length);//发送数据
                    return true;
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                //MessageBox.Show("串口未打开", "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                AddContent("串口未打开\r\n");
            }
            return false;
        }
        #endregion



        #region 发送数据button事件
        private void btnSend_Click(object sender, EventArgs e)
        {
            if (chkAutoSend.Checked)
            {
                if (txtSendData.Text.Length < 1)
                {
                    MessageBox.Show("发送数据为空!", "错误");
                    return;
                }
                if (txtAutoSendms.Text.Length < 1) txtAutoSendms.Text = "500";
                timerAutoSend.Interval = Convert.ToInt32(txtAutoSendms.Text);
                if (timerAutoSend.Enabled)
                {
                    timerAutoSend.Enabled = false;
                    groupBoxComSetting.Enabled = true;
                    groupboxSendSetting.Enabled = true;
                    btnSend.Text = "发送";
                }
                else
                {
                    timerAutoSend.Enabled = true;
                    groupBoxComSetting.Enabled = false;
                    groupboxSendSetting.Enabled = false;
                    btnSend.Text = "停止发送";
                }
            }
            else
            {
                if (SendStr(txtSendData.Text))
                {
                    if (chkAutoCleanSend.Checked)
                    {
                        txtSendData.Text = "";
                    }
                }
            }
        }
        #endregion


        #region 选择编码发送字符串
        private bool SendStr(String str)
        {
            return SendStr(txtSendData.Text, rbtnSendHex.Checked);
        }
        private bool SendStr(String str, bool hexbool)
        {

            byte[] sendData = null;

            if (hexbool)
            {
                try
                {
                    sendData = strToHexByte(str.Trim());
                }
                catch (Exception)
                {
                    //throw;
                    MessageBox.Show("字符串转十六进制有误,请检测输入格式.", "错误!");
                    return false;
                }
            }
            else if (rbtnSendASCII.Checked)
            {
                //sendData = Encoding.ASCII.GetBytes(str);
                sendData = Encoding.GetEncoding("GBK").GetBytes(str);
            }
            else if (rbtnSendUTF8.Checked)
            {
                sendData = Encoding.UTF8.GetBytes(str);
            }
            else if (rbtnSendUnicode.Checked)
            {
                sendData = Encoding.Unicode.GetBytes(str);
            }
            else
            {
                sendData = Encoding.GetEncoding("GBK").GetBytes(str);
            }

            if (this.SendData(sendData))//发送数据成功计数
            {
                lblSendCount.Invoke(new MethodInvoker(delegate
                {
                    lblSendCount.Text = "发送:" + (int.Parse(lblSendCount.Text.Substring(3)) + sendData.Length).ToString();
                }));
                return true;
            }

            return false;
        }
        #endregion

        #region 字符串转换16进制字节数组
        private byte[] strToHexByte(string hexString)
        {
            hexString = hexString.Replace(" ", "");
            if ((hexString.Length % 2) != 0) hexString += " ";
            byte[] returnBytes = new byte[hexString.Length / 2];
            for (int i = 0; i < returnBytes.Length; i++)
                returnBytes[i] = Convert.ToByte(hexString.Substring(i * 2, 2).Replace(" ", ""), 16);
            return returnBytes;
        }
        #endregion

        #region 接收数据监听
        private void Com_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            byte[] ReDatas = new byte[ComDevice.BytesToRead];
            ComDevice.Read(ReDatas, 0, ReDatas.Length);//读取数据
            this.AddData(ReDatas);//输出数据
        }
        #endregion
        #region 接收文本框,输入监听,供输入直接发送
        private void txtShowData_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (chkRecSend.Checked && ComDevice.IsOpen)
            {
                //if ((e.KeyChar >= ' ' && e.KeyChar <= '~') || e.KeyChar == '\r')//这是允许输入退格键允许输入0-9数字
                //{
                SendStr(e.KeyChar.ToString(), false);
                e.Handled = true;
                //}
            }
        }
        #endregion

        #region 接收文本框字符处理
        public void AddData(byte[] data)
        {
            if (rbtnHex.Checked)
            {
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < data.Length; i++)
                {
                    sb.AppendFormat("{0:x2}" + " ", data[i]);
                }
                AddContent(sb.ToString().ToUpper());
            }
            else if (rbtnUTF8.Checked)
            {
                AddContent(new UTF8Encoding().GetString(data));
            }
            else if (rbtnUnicode.Checked)
            {
                AddContent(new UnicodeEncoding().GetString(data));
            }
            else// if (rbtnASCII.Checked)
            {
                //AddContent(new ASCIIEncoding().GetString(data));
                AddContent(Encoding.GetEncoding("GBK").GetString(data));
            }

            lblRevCount.Invoke(new MethodInvoker(delegate
            {
                lblRevCount.Text = "接收:" + (int.Parse(lblRevCount.Text.Substring(3)) + data.Length).ToString();
            }));
        }




        //输入到显示区域
        private void AddContent(string content)
        {
            this.BeginInvoke(new MethodInvoker(delegate
            {
                if (chkAutoLine.Checked && txtShowData.Text.Length > 0)
                {
                    txtShowData.AppendText("\r\n");
                    if (chkShowTime.Checked)
                    {
                        txtShowData.AppendText("【" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss.fff") + "】" + "\r\n");
                    }
                }
                txtShowData.AppendText(content);
            }));
        }
        #endregion

        #region 清空接收区
        private void lkbClearRev_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            txtShowData.Clear();
        }
        #endregion
        #region 清空发送
        private void lkbClearSend_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            txtSendData.Clear();

        }
        #endregion

        #region 显示接受时间监听
        private void chkShowTime_CheckedChanged(object sender, EventArgs e)
        {
            if (chkShowTime.Checked)
            {
                chkAutoLine.Checked = true;
                chkAutoLine.Enabled = false;
                chkShowTime.Checked = true;
            }
            else
            {
                chkAutoLine.Enabled = true;
            }
        }
        #endregion


        #region 计数清零
        private void btnCleanCount_Click(object sender, EventArgs e)
        {
            lblSendCount.Text = "发送:0";
            lblRevCount.Text = "接收:0";
        }
        #endregion

        #region 自动定时发送
        #region 自动发送定时器函数
        private void timerAutoSend_Tick(object sender, EventArgs e)
        {
            SendStr(txtSendData.Text);
        }
        #endregion
        #region 自动发送间隔时间,只能输入数字
        private void txtAutoSendms_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar != '\b' && ((e.KeyChar < '0') || (e.KeyChar > '9')))//这是允许输入退格键允许输入0-9数字
            {
                e.Handled = true;
            }
        }
        #endregion

        #endregion

        #region 编码RadioButton控件单击监听
        private void rbtnSend_Click(object sender, EventArgs e)
        {
            rbtnSendHex.Checked = false;
            rbtnSendASCII.Checked = false;
            rbtnSendUTF8.Checked = false;
            rbtnSendUnicode.Checked = false;
            ((RadioButton)sender).Checked = true;
        }
        private void rbtn_Click(object sender, EventArgs e)
        {
            rbtnHex.Checked = false;
            rbtnASCII.Checked = false;
            rbtnUTF8.Checked = false;
            rbtnUnicode.Checked = false;
            ((RadioButton)sender).Checked = true;
        }
        #endregion
    }
}