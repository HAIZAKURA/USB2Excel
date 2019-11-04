using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.IO;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace USB2Excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            base.Load += this.Form1_Load;
            List<WeakReference> _ENCList = Form1.__ENCList;
            lock (_ENCList)
            {
                Form1.__ENCList.Add(new WeakReference(this));
            }
            InitializeComponent();
        }

        private string openFileName;
        const string HZ_DATALOG = "DATALOG.HZ";
        const string HZ_EVENTLOG = "EVENTLOG.HZ";
        const string LZ_DATALOG = "DATALOG.LZ";
        const string LZ_EVENTLOG = "EVENTLOG.LZ";

        private void Button1_Click(object sender, EventArgs e)
        {
            MessageBox.Show("请注意：如果记录内容太多，该操作的耗时会很长！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
            this.openFileDialog1.ShowDialog();
        }

        private void OpenFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            this.openFileName = this.openFileDialog1.FileName;
            if (this.openFileName.EndsWith(HZ_DATALOG) || this.openFileName.EndsWith(LZ_DATALOG))
            {
                this.saveFileDialog1.FileName = string.Concat(new string[]
                {
                    "DATALOG_",
                    DateTime.Now.Year.ToString(),
                    "-",
                    DateTime.Now.Month.ToString(),
                    "-",
                    DateTime.Now.Day.ToString(),
                    "-",
                    DateTime.Now.Hour.ToString(),
                    "-",
                    DateTime.Now.Minute.ToString(),
                });
            }
            else if (this.openFileName.EndsWith(HZ_EVENTLOG) || this.openFileName.EndsWith(LZ_EVENTLOG))
            {
                this.saveFileDialog1.FileName = string.Concat(new string[]
                {
                    "EVENTLOG_",
                    DateTime.Now.Year.ToString(),
                    "-",
                    DateTime.Now.Month.ToString(),
                    "-",
                    DateTime.Now.Day.ToString(),
                    "-",
                    DateTime.Now.Hour.ToString(),
                    "-",
                    DateTime.Now.Minute.ToString(),
                });
            }
            this.saveFileDialog1.ShowDialog();
        }

        private void SaveFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            FileStream fileStream = new FileStream(this.openFileName, FileMode.Open);
            BinaryReader binaryReader = new BinaryReader(fileStream);
            object instance = RuntimeHelpers.GetObjectValue(Interaction.CreateObject("Excel.Application", ""));
            object obj = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(NewLateBinding.LateGet(instance, null, "workbooks", new object[0], null, null, null), null, "add", new object[0], null, null, null));
            object obj2 = RuntimeHelpers.GetObjectValue(NewLateBinding.LateGet(obj, null, "Sheets", new object[]
            {
                1
            }, null, null, null));
            int procType = 0;
            if (this.openFileName.EndsWith(HZ_DATALOG) || this.openFileName.EndsWith(LZ_DATALOG))
            {
                procType = 1;
            }
            else if (this.openFileName.EndsWith(HZ_EVENTLOG) || this.openFileName.EndsWith(LZ_EVENTLOG))
            {
                procType = 2;
            }
            bool flag = (procType == 1);
            try
            {
                object[] array;
                bool[] array2;
                if (flag)
                {
                    int num2 = 2;
                    int num3;
                    int num4;
                    do
                    {
                        object instance2 = obj2;
                        Type type = null;
                        string memberName = "Columns";
                        array = new object[]
                        {
                            num2
                        };
                        object[] arguments = array;
                        string[] argumentNames = null;
                        Type[] typeArguments = null;
                        array2 = new bool[]
                        {
                            true
                        };
                        object instance3 = NewLateBinding.LateGet(instance2, type, memberName, arguments, argumentNames, typeArguments, array2);
                        if (array2[0])
                        {
                            num2 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
                        }
                        NewLateBinding.LateSetComplex(instance3, null, "NumberFormatLocal", new object[]
                        {
                            "0.0"
                        }, null, null, false, true);
                        num2++;
                        num3 = num2;
                        num4 = 15;
                    }
                    while (num3 <= num4);
                    NewLateBinding.LateSet(obj2, null, "Cells", new object[]
                    {
                        1,
                        2,
                        "设定温度 Setting Temp. Value (ºC)"
                    }, null, null);
                    NewLateBinding.LateSet(obj2, null, "Cells", new object[]
                    {
                        1,
                        3,
                        "测量温度 Actual Temp. Value (ºC)"
                    }, null, null);
                    NewLateBinding.LateSet(obj2, null, "Cells", new object[]
                    {
                        1,
                        4,
                        "设定湿度 Setting Humidity Value (%RH)"
                    }, null, null);
                    NewLateBinding.LateSet(obj2, null, "Cells", new object[]
                    {
                        1,
                        5,
                        "测量湿度 Actual Humidity Value (%RH)"
                    }, null, null);
                    NewLateBinding.LateSet(obj2, null, "Cells", new object[]
                    {
                        1,
                        6,
                        "设定光照 Setting Lighting Value (%)"
                    }, null, null);
                    NewLateBinding.LateSet(obj2, null, "Cells", new object[]
                    {
                        1,
                        7,
                        "设定光照 Setting Lighting Value (K.LX)"
                    }, null, null);
                    NewLateBinding.LateSet(obj2, null, "Cells", new object[]
                    {
                        1,
                        8,
                        "测量光照 Actual Lighting Value (K.LX)"
                    }, null, null);
                    NewLateBinding.LateSet(obj2, null, "Cells", new object[]
                    {
                        1,
                        9,
                        "设定CO2浓度 Setting CO2 Value (%)"
                    }, null, null);
                    NewLateBinding.LateSet(obj2, null, "Cells", new object[]
                    {
                        1,
                        10,
                        "测量CO2浓度 Actual CO2 Value (%)"
                    }, null, null);
                    NewLateBinding.LateSet(obj2, null, "Cells", new object[]
                    {
                        1,
                        11,
                        "设定紫外强度 Setting UV Value (W/M2)"
                    }, null, null);
                    NewLateBinding.LateSet(obj2, null, "Cells", new object[]
                    {
                        1,
                        12,
                        "测量紫外强度 Actual UV Value (W/M2)"
                    }, null, null);
                }
                else
                {
                    flag = (procType == 2);
                    if (flag)
                    {
                        int num2 = 2;
                        int num4;
                        int num5;
                        do
                        {
                            object instance4 = obj2;
                            Type type2 = null;
                            string memberName2 = "Columns";
                            array = new object[]
                            {
                                num2
                            };
                            object[] arguments2 = array;
                            string[] argumentNames2 = null;
                            Type[] typeArguments2 = null;
                            array2 = new bool[]
                            {
                                true
                            };
                            object instance5 = NewLateBinding.LateGet(instance4, type2, memberName2, arguments2, argumentNames2, typeArguments2, array2);
                            if (array2[0])
                            {
                                num2 = (int)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(int));
                            }
                            NewLateBinding.LateSetComplex(instance5, null, "NumberFormatLocal", new object[]
                            {
                                "@"
                            }, null, null, false, true);
                            num2++;
                            num5 = num2;
                            num4 = 15;
                        }
                        while (num5 <= num4);
                        NewLateBinding.LateSet(obj2, null, "Cells", new object[]
                        {
                            1,
                            2,
                            "事件 Event                                                            "
                        }, null, null);
                    }
                }
                NewLateBinding.LateSet(obj2, null, "Cells", new object[]
                {
                    1,
                    1,
                    "时间 Time (年月日时分秒 YYYY-MM-DD HH:MM:SS)"
                }, null, null);
                NewLateBinding.LateCall(NewLateBinding.LateGet(NewLateBinding.LateGet(obj2, null, "Cells", new object[0], null, null, null), null, "EntireColumn", new object[0], null, null, null), null, "AutoFit", new object[0], null, null, null, true);
                flag = (procType == 0);
                if (flag)
                {
                    obj2 = null;
                    NewLateBinding.LateCall(obj, null, "Close", new object[0], null, null, null, true);
                    obj = null;
                    NewLateBinding.LateCall(instance, null, "Quit", new object[0], null, null, null, true);
                    instance = null;
                    binaryReader.Close();
                    fileStream.Close();
                }
                long num6 = 1L;
                try
                {
                    binaryReader.BaseStream.Seek(0L, SeekOrigin.Begin);
                    while (binaryReader.BaseStream.Position < binaryReader.BaseStream.Length)
                    {
                        try
                        {
                            flag = (procType == 1);
                            if (flag)
                            {
                                byte[] array3 = binaryReader.ReadBytes(32);
                                string text = string.Concat(new string[]
                                {
                                    (2000 + (int)array3[3]).ToString(),
                                    "/",
                                    array3[4].ToString(),
                                    "/",
                                    array3[5].ToString(),
                                    " ",
                                    array3[6].ToString(),
                                    ":",
                                    array3[7].ToString(),
                                    ":",
                                    array3[8].ToString()
                                });
                                num6 += 1L;
                                NewLateBinding.LateSet(obj2, null, "Cells", new object[]
                                {
                                    num6,
                                    1,
                                    text
                                }, null, null);
                                int num2 = 0;
                                int num4;
                                int num9;
                                do
                                {
                                    short num7 = (short)array3[9 + (num2 << 1)];
                                    double num8;
                                    unchecked
                                    {
                                        num7 = (short)(num7 << 8);
                                        num7 += (short)array3[checked(10 + (num2 << 1))];
                                        num8 = (double)num7;
                                        num8 /= 10.0;
                                    }
                                    NewLateBinding.LateSet(obj2, null, "Cells", new object[]
                                    {
                                        num6,
                                        num2 + 2,
                                        num8
                                    }, null, null);
                                    num2++;
                                    num9 = num2;
                                    num4 = 10;
                                }
                                while (num9 <= num4);
                            }
                            else
                            {
                                flag = (procType == 2);
                                if (flag)
                                {
                                    byte[] array3 = binaryReader.ReadBytes(16);
                                    string text = string.Concat(new string[]
                                    {
                                        (2000 + (int)array3[3]).ToString(),
                                        "/",
                                        array3[4].ToString(),
                                        "/",
                                        array3[5].ToString(),
                                        " ",
                                        array3[6].ToString(),
                                        ":",
                                        array3[7].ToString(),
                                        ":",
                                        array3[8].ToString()
                                    });
                                    num6 += 1L;
                                    NewLateBinding.LateSet(obj2, null, "Cells", new object[]
                                    {
                                        num6,
                                        1,
                                        text
                                    }, null, null);
                                    switch (array3[9])
                                    {
                                        case 0:
                                            text = "关机 Power Off";
                                            break;
                                        case 1:
                                            text = "箱温温度上限 Chamber Temperature Upper Limit";
                                            break;
                                        case 2:
                                            text = "箱温温度下限 Chamber Temperature Lower Limit";
                                            break;
                                        case 3:
                                            text = "湿度上限 Humidity Upper Limit";
                                            break;
                                        case 4:
                                            text = "湿度下限 Humidity Lower Limit";
                                            break;
                                        case 5:
                                            text = "低水位报警 Low Water";
                                            break;
                                        case 6:
                                            text = "开门报警 Door Open";
                                            break;
                                        case 7:
                                            text = "箱温温度传感器故障 Chamber Temperature Sensor Failure";
                                            break;
                                        case 8:
                                            text = "蒸发温度传感器故障 Evaporation Temperature Sensor Failure";
                                            break;
                                        case 9:
                                            text = "主电源故障 Main Power Failure";
                                            break;
                                        case 10:
                                            text = "电池故障 Battery Failure";
                                            break;
                                        case 11:
                                            text = "开机 Power On";
                                            break;
                                        case 12:
                                            text = "运行 Start";
                                            break;
                                        case 13:
                                            text = "停止 Stop";
                                            break;
                                        case 14:
                                            text = "开照明 Lighting On";
                                            break;
                                        case 15:
                                            text = "关照明 Lighting Off";
                                            break;
                                        case 16:
                                            text = "开紫外 UV On";
                                            break;
                                        case 17:
                                            text = "关紫外 UV Off";
                                            break;
                                        case 18:
                                            text = "打开排水阀 Open The Drain Valve";
                                            break;
                                        case 19:
                                            text = "关闭排水阀 Close The Drain Valve";
                                            break;
                                        case 20:
                                            text = "编程 Program";
                                            break;
                                        case 21:
                                            text = "进入用户菜单 Enter The User Menu";
                                            break;
                                        case 22:
                                            text = "进入二级菜单 Enter Second-Level Menu";
                                            break;
                                        case 23:
                                            text = "进入三级菜单 Enter Third-Level Menu";
                                            break;
                                        case 24:
                                            text = "进入四级菜单 Enter Fourth-Level Menu";
                                            break;
                                        case 25:
                                            text = "CO2上限 CO2 Upper Limit";
                                            break;
                                        case 26:
                                            text = "CO2下限 CO2 Lower Limit";
                                            break;
                                        case 27:
                                            text = "门温传感器故障 Door Temperature Sensor Failure";
                                            break;
                                        case 28:
                                            text = "独立限温报警 Independent Limit Temperature Alarm";
                                            break;
                                        case 29:
                                            text = "紫外上限 UV Upper Limit";
                                            break;
                                        case 30:
                                            text = "紫外下限 UV Lower Limit";
                                            break;
                                        default:
                                            text = "未定义 Undefined";
                                            break;
                                    }
                                    /*
                                    flag = (array3[9] == 0);
                                    if (flag)
                                    {
                                        text = "关机Power Off";
                                    }
                                    else
                                    {
                                        flag = (array3[9] == 1);
                                        if (flag)
                                        {
                                            text = "箱温温度上限Chamber Temperature Upper Limit";
                                        }
                                        else
                                        {
                                            flag = (array3[9] == 2);
                                            if (flag)
                                            {
                                                text = "箱温温度下限Chamber Temperature Lower Limit";
                                            }
                                            else
                                            {
                                                flag = (array3[9] == 3);
                                                if (flag)
                                                {
                                                    text = "湿度上限Humidity Upper Limit";
                                                }
                                                else
                                                {
                                                    flag = (array3[9] == 4);
                                                    if (flag)
                                                    {
                                                        text = "湿度下限Humidity Lower Limit";
                                                    }
                                                    else
                                                    {
                                                        flag = (array3[9] == 5);
                                                        if (flag)
                                                        {
                                                            text = "低水位报警Low Water";
                                                        }
                                                        else
                                                        {
                                                            flag = (array3[9] == 6);
                                                            if (flag)
                                                            {
                                                                text = "开门报警Door Open";
                                                            }
                                                            else
                                                            {
                                                                flag = (array3[9] == 7);
                                                                if (flag)
                                                                {
                                                                    text = "箱温温度传感器故障Chamber Temperature Sensor Failure";
                                                                }
                                                                else
                                                                {
                                                                    flag = (array3[9] == 8);
                                                                    if (flag)
                                                                    {
                                                                        text = "蒸发温度传感器故障Evaporation Temperature Sensor Failure";
                                                                    }
                                                                    else
                                                                    {
                                                                        flag = (array3[9] == 9);
                                                                        if (flag)
                                                                        {
                                                                            text = "主电源故障Main Power Failure";
                                                                        }
                                                                        else
                                                                        {
                                                                            flag = (array3[9] == 10);
                                                                            if (flag)
                                                                            {
                                                                                text = "电池故障Battery Failure";
                                                                            }
                                                                            else
                                                                            {
                                                                                flag = (array3[9] == 11);
                                                                                if (flag)
                                                                                {
                                                                                    text = "开机Power On";
                                                                                }
                                                                                else
                                                                                {
                                                                                    flag = (array3[9] == 12);
                                                                                    if (flag)
                                                                                    {
                                                                                        text = "运行Start";
                                                                                    }
                                                                                    else
                                                                                    {
                                                                                        flag = (array3[9] == 13);
                                                                                        if (flag)
                                                                                        {
                                                                                            text = "停止:Stop";
                                                                                        }
                                                                                        else
                                                                                        {
                                                                                            flag = (array3[9] == 14);
                                                                                            if (flag)
                                                                                            {
                                                                                                text = "开照明Lighting On";
                                                                                            }
                                                                                            else
                                                                                            {
                                                                                                flag = (array3[9] == 15);
                                                                                                if (flag)
                                                                                                {
                                                                                                    text = "关照明Lighting Off";
                                                                                                }
                                                                                                else
                                                                                                {
                                                                                                    flag = (array3[9] == 16);
                                                                                                    if (flag)
                                                                                                    {
                                                                                                        text = "开紫外UV On";
                                                                                                    }
                                                                                                    else
                                                                                                    {
                                                                                                        flag = (array3[9] == 17);
                                                                                                        if (flag)
                                                                                                        {
                                                                                                            text = "关紫外UV Off";
                                                                                                        }
                                                                                                        else
                                                                                                        {
                                                                                                            flag = (array3[9] == 18);
                                                                                                            if (flag)
                                                                                                            {
                                                                                                                text = "打开排水阀Open The Drain Valve";
                                                                                                            }
                                                                                                            else
                                                                                                            {
                                                                                                                flag = (array3[9] == 19);
                                                                                                                if (flag)
                                                                                                                {
                                                                                                                    text = "关闭排水阀Close The Drain Valve";
                                                                                                                }
                                                                                                                else
                                                                                                                {
                                                                                                                    flag = (array3[9] == 20);
                                                                                                                    if (flag)
                                                                                                                    {
                                                                                                                        text = "编程Program";
                                                                                                                    }
                                                                                                                    else
                                                                                                                    {
                                                                                                                        flag = (array3[9] == 21);
                                                                                                                        if (flag)
                                                                                                                        {
                                                                                                                            text = "进入用户菜单Enter The User Menu";
                                                                                                                        }
                                                                                                                        else
                                                                                                                        {
                                                                                                                            flag = (array3[9] == 22);
                                                                                                                            if (flag)
                                                                                                                            {
                                                                                                                                text = "进入二级菜单Enter Second-Level Menu";
                                                                                                                            }
                                                                                                                            else
                                                                                                                            {
                                                                                                                                flag = (array3[9] == 23);
                                                                                                                                if (flag)
                                                                                                                                {
                                                                                                                                    text = "进入三级菜单Enter Third-Level Menu";
                                                                                                                                }
                                                                                                                                else
                                                                                                                                {
                                                                                                                                    flag = (array3[9] == 24);
                                                                                                                                    if (flag)
                                                                                                                                    {
                                                                                                                                        text = "进入四级菜单Enter Fourth-Level Menu";
                                                                                                                                    }
                                                                                                                                    else
                                                                                                                                    {
                                                                                                                                        flag = (array3[9] == 25);
                                                                                                                                        if (flag)
                                                                                                                                        {
                                                                                                                                            text = "CO2上限CO2 Upper Limit";
                                                                                                                                        }
                                                                                                                                        else
                                                                                                                                        {
                                                                                                                                            flag = (array3[9] == 26);
                                                                                                                                            if (flag)
                                                                                                                                            {
                                                                                                                                                text = "CO2下限CO2 Lower Limit";
                                                                                                                                            }
                                                                                                                                            else
                                                                                                                                            {
                                                                                                                                                flag = (array3[9] == 27);
                                                                                                                                                if (flag)
                                                                                                                                                {
                                                                                                                                                    text = "门温传感器故障Door Temperature Sensor Failure";
                                                                                                                                                }
                                                                                                                                                else
                                                                                                                                                {
                                                                                                                                                    flag = (array3[9] == 28);
                                                                                                                                                    if (flag)
                                                                                                                                                    {
                                                                                                                                                        text = "独立限温报警Independent limit temperature alarm";
                                                                                                                                                    }
                                                                                                                                                }
                                                                                                                                            }
                                                                                                                                        }
                                                                                                                                    }
                                                                                                                                }
                                                                                                                            }
                                                                                                                        }
                                                                                                                    }
                                                                                                                }
                                                                                                            }
                                                                                                        }
                                                                                                    }
                                                                                                }
                                                                                            }
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    */
                                    NewLateBinding.LateSet(obj2, null, "Cells", new object[]
                                    {
                                        num6,
                                        2,
                                        text
                                    }, null, null);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                object instance6 = obj;
                Type type3 = null;
                string memberName3 = "SaveAs";
                object[] array4 = new object[1];
                object[] array5 = array4;
                int num10 = 0;
                SaveFileDialog saveFileDialog = this.saveFileDialog1;
                array5[num10] = saveFileDialog.FileName;
                array = array4;
                object[] arguments3 = array;
                string[] argumentNames3 = null;
                Type[] typeArguments3 = null;
                array2 = new bool[]
                {
                    true
                };
                NewLateBinding.LateCall(instance6, type3, memberName3, arguments3, argumentNames3, typeArguments3, array2, true);
                if (array2[0])
                {
                    saveFileDialog.FileName = (string)Conversions.ChangeType(RuntimeHelpers.GetObjectValue(array[0]), typeof(string));
                }
                obj2 = null;
                NewLateBinding.LateCall(obj, null, "Close", new object[0], null, null, null, true);
                obj = null;
                NewLateBinding.LateCall(instance, null, "Quit", new object[0], null, null, null, true);
                instance = null;
                binaryReader.Close();
                fileStream.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "错误", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            MessageBox.Show("转换成功！", "提示", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            string processName = Process.GetCurrentProcess().ProcessName;
            bool flag = Process.GetProcessesByName(processName).GetUpperBound(0) > 0;
            if (flag)
            {
                Interaction.MsgBox("程序已经在运行", MsgBoxStyle.Critical, null);
                ProjectData.EndApp();
            }
        }

        private static List<WeakReference> __ENCList = new List<WeakReference>();
    }
}
