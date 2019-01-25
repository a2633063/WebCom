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


        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                interfaceUpdataHandle = new HandleInterfaceUpdataDelegate(UpdateTextBox);//实例化委托对象 

                serialPort1.DataReceived += new SerialDataReceivedEventHandler(serialPort1_DataReceived);
                if (!serialPort1.IsOpen)
                {
                    serialPort1.Open();
                }
                button2.Enabled = true;
                button1.Enabled = false;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
            //timer1.Enabled = true;

        }
        /// <summary>
        /// 控件加载事件
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MyActiveXControl_Load(object sender, EventArgs e)
        {
            setOrgComb();
        }
        /// <summary>
        /// 初始化串口
        /// </summary>
        private void setOrgComb()
        {
            serialPort1.PortName = "COM1";          //端口名称
            serialPort1.BaudRate = 1200;            //波特率
            serialPort1.Parity = Parity.None;       //奇偶效验
            serialPort1.StopBits = StopBits.One;    //效验
            serialPort1.DataBits = 8;               //每个字节的数据位长度
        }
        /// <summary>
        /// 更新数据
        /// </summary>
        /// <param name="text"></param>
        private void UpdateTextBox(string text)
        {
            //richTextBox1.Text = text + "\n\t" + richTextBox1.Text;
            textBox1.Text = text;
        }

        /// <summary>
        /// 接收数据是发生
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void serialPort1_DataReceived(object sender, SerialDataReceivedEventArgs e)
        {
            //获取接收缓冲区中数据的字节数
            if (serialPort1.BytesToRead > 5)
            {

                string strTemp = serialPort1.ReadExisting();//读取串口
                double weight = -1;//获取到的重量
                foreach (string str in strTemp.Split('='))//获取稳定的值
                {
                    double flog = 0;
                    //数据是否正常
                    if (double.TryParse(str, out flog) && str.IndexOf('.') > 0 && str[str.Length - 1] != '.')
                    {
                        //数据转换   串口获取到的数据是倒叙的  因此进行反转
                        char[] charArray = str.ToCharArray();
                        Array.Reverse(charArray);
                        string left = new string(charArray).Split('.')[0];
                        string right = new string(charArray).Split('.')[1];
                        if (right.Length == 2)
                        {
                            weight = int.Parse(left) + int.Parse(right) / 100.0;
                        }
                    }
                }
                if (weight >= 0)
                {
                    //在拥有控件的基础窗口句柄的线程上，用指定的参数列表执行指定委托。                    
                    this.Invoke(interfaceUpdataHandle, weight.ToString());//取到数据   更新
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                button1.Enabled = true;
                button2.Enabled = false;
                serialPort1.Close();
                //timer1.Enabled = false;


            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (isClose)
            {
                return;
            }
            try
            {
                string send = "" + (char)(27) + 'p';
                send = serialPort1.ReadExisting();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                button2_Click(null, null);
            }
        }
    }
}