
namespace 电脑桌面背景上显示姓名
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.IO;
    using System.Runtime.InteropServices;
    using System.Windows.Forms;
    using ExcelDataReader;
    using System.Text.RegularExpressions;

    public class Program
    {
        // 导入Windows API函数
        [DllImport("user32.dll")]
        public static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern int GetWindowLong(IntPtr hWnd, int nIndex);

        [DllImport("user32.dll")]
        public static extern bool SetLayeredWindowAttributes(IntPtr hWnd, uint crKey, byte bAlpha, uint dwFlags);

        // 定义常量
        public const int GWL_EXSTYLE = -20;
        public const int WS_EX_LAYERED = 0x80000;
        public const int LWA_COLORKEY = 0x1;
        static System.Windows.Forms.Label label;
        public static void Main()
        {
            // 创建一个窗体
            Form form = new Form
            {
                FormBorderStyle = FormBorderStyle.None,
                StartPosition = FormStartPosition.CenterScreen,

                // 设置窗体背景为透明
                BackColor = Color.Black,
                TransparencyKey = Color.Black,

                //隐藏任务栏上的图标
                ShowInTaskbar = false
            };
            // 创建一个标签来显示姓名或编号
               label = new System.Windows.Forms.Label
            {
                Text = GetTextToShow(), // 设置姓名或编号
                //Text = Application.ProductName,
                Font = new Font("微软雅黑", 48, FontStyle.Bold), // 设置字体样式和大小
                ForeColor = Color.White, // 设置字体颜色
                AutoSize = true, // 自动调整标签大小以适应文本内容
                Location = new Point(10, 10) // 设置标签位置
            };

            // 将标签添加到窗体中
            form.Controls.Add(label);

            //创建一个Timer对象
            var timer = new System.Timers.Timer
            {
                // 设置计时器的间隔为10分钟（以毫秒为单位）
                Interval = 10 * 60 * 1000
            };


            // 设置计时器的Elapsed事件处理程序
            //timer.Elapsed += (sender, e) => TimerElapsed(sender, e, label);
            timer.Elapsed += Timer_Tick;            // 计时器的Elapsed事件处理程序

            // 启动计时器
            timer.Start();


            // 设置窗体大小和位置
            form.Size = new Size(label.Width + 20, label.Height + 20);
            // 获取屏幕的工作区大小
            Rectangle workingArea = Screen.GetWorkingArea(form);

            // 设置窗体的位置为右上角
            form.Location = new Point(workingArea.Right - 2*form.Width, workingArea.Top+form.Height);
            // 设置窗体为分层窗口
            SetWindowLong(form.Handle, GWL_EXSTYLE, GetWindowLong(form.Handle, GWL_EXSTYLE) | WS_EX_LAYERED);

            // 设置窗体透明度
            SetLayeredWindowAttributes(form.Handle, 0, 128, LWA_COLORKEY);

            // 运行窗体
            Application.Run(form);

        }

        public static string GetTextToShow(string filePath = "课程表.xlsx", string sheetName = "Sheet1", string cellAddress = "1")
        {
            int columnIndex=-1, rowIndex=-1;
            string className=string.Empty;
            string studentName=string.Empty;
            int NO = 0;
            // 获取当前时间的周几
            DayOfWeek currentDayOfWeek = DateTime.Now.DayOfWeek;
            // 将英文周几转换为中文周几
            string chineseDayOfWeek = string.Empty;
            string chineseDayOfWeek2 = string.Empty;

            switch (currentDayOfWeek)
            {
                case DayOfWeek.Sunday:
                    chineseDayOfWeek = "星期日";
                    chineseDayOfWeek2 = "周日";
                    break;
                case DayOfWeek.Monday:
                    chineseDayOfWeek = "星期一";
                    chineseDayOfWeek2 = "周日";
                    break;
                case DayOfWeek.Tuesday:
                    chineseDayOfWeek = "星期二";
                    chineseDayOfWeek2 = "周二";
                    break;
                case DayOfWeek.Wednesday:
                    chineseDayOfWeek = "星期三";
                    chineseDayOfWeek2 = "周三";
                    break;
                case DayOfWeek.Thursday:
                    chineseDayOfWeek = "星期四";
                    chineseDayOfWeek2 = "周四";
                    break;
                case DayOfWeek.Friday:
                    chineseDayOfWeek = "星期五";
                    chineseDayOfWeek2 = "周五";
                    break;
                case DayOfWeek.Saturday:
                    chineseDayOfWeek = "星期六";
                    chineseDayOfWeek2 = "周六";
                    break;
            }
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var dataSet = reader.AsDataSet();
                    try
                    {
                        var dataTable = dataSet.Tables["课程表"];

                        for (int i = 0; i < dataTable.Rows.Count; i++)
                        {
                            // 遍历查找包含特定时间数据的列
                            for (int column = 0; column < dataTable.Columns.Count; column++)
                            {
                                var cellData = dataTable.Rows[i][column].ToString();

                                if (cellData != string.Empty && cellData.Contains("-"))
                                {
                                    if (IsTimeInRange(cellData))
                                    {
                                    rowIndex = i;
                                    }
                                }
                                // 判断单元格的值是否与当前时间的周几匹配
                                if (cellData.Equals(chineseDayOfWeek, StringComparison.OrdinalIgnoreCase) || cellData.Equals(chineseDayOfWeek2, StringComparison.OrdinalIgnoreCase))
                                {
                                    columnIndex = column;
                                }
                                if (rowIndex > -1 && columnIndex > -1)
                                {
                                    className=dataTable.Rows[rowIndex][columnIndex].ToString();
                                    break;
                                }
                            }
                            if (rowIndex > -1 && columnIndex > -1)
                            {
                                break;
                            }
                        }
                        try
                        {
                            if(className == string.Empty) return string.Empty;
                            dataTable = dataSet.Tables[className];
                            rowIndex =-1;
                            // 读取学生姓名和IP地址
                            var localIP = GetLocalIPAddress();
                            for (int i = 0; i < dataTable.Rows.Count; i++)
                            {
                                // 遍历查找包含本机IP地址数据的列
                                for (int column = 0; column < dataTable.Columns.Count; column++)
                                {
                                    var cellData = dataTable.Rows[i][column].ToString();

                                    if (cellData != string.Empty && cellData.Contains(localIP))
                                    {
                                        rowIndex = i;
                                        break;
                                    }
                                    if (rowIndex > -1) break;
                                }
                            }

                            // 遍历查找本机IP地址所在行的单元格，获取学生姓名
                            for (int column = 0; column < dataTable.Columns.Count; column++)
                            {
                                var cellData = dataTable.Rows[rowIndex][column].ToString();

                                if (Regex.IsMatch(cellData, @"^-?\d+$"))
                                {
                                    NO = int.Parse(cellData);
                                    continue;
                                }
                                if(Regex.IsMatch(cellData, @"^[\u4e00-\u9fa5]{2,4}$"))
                                {
                                    studentName = cellData;
                                }
                                if (NO>0 && studentName.Length>0) return NO.ToString() + " " + studentName;
                            }
                            return NO.ToString() + " " + studentName;
                        }
                        catch 
                        {
                            Console.WriteLine("表名不存在");
                            return string.Empty;
                        }
                    }
                    catch (KeyNotFoundException)
                    {
                        // 处理表名不存在的情况
                        // 可以输出错误信息或执行其他逻辑
                        Console.WriteLine("表名不存在");
                        return string.Empty;
                    }
                }
            }
        }
        private static void Timer_Tick(object sender, EventArgs e)
        {
            // 在标签上显示当前时间
            label.Invoke((MethodInvoker)(() =>
            {
                // 在标签上显示当前时间
                label.Text = GetTextToShow(); // 设置姓名或编号
            }));
        }

        private static bool IsTimeInRange(string timeData)
        {
            // 移除冒号和空格，统一为特定的时间格式
            timeData= timeData.Replace("：", ":").Replace(" ", "").Replace("—","-");
            var timeDatas=timeData.Split('-');
            TimeSpan startTime = TimeSpan.Parse(timeDatas[0]);
            TimeSpan endTime = TimeSpan.Parse(timeDatas[1]);
            TimeSpan currentTime = DateTime.Now.TimeOfDay;
            if (startTime <= endTime)
            {
                return currentTime >= startTime && currentTime <= endTime;
            }
            else
            {
                return currentTime >= startTime || currentTime <= endTime;
            }
        }



        static string GetLocalIPAddress()
        {
            var host = System.Net.Dns.GetHostEntry(System.Net.Dns.GetHostName());
            foreach (var ip in host.AddressList)
            {
                if (ip.AddressFamily == System.Net.Sockets.AddressFamily.InterNetwork)
                {
                    return ip.ToString();
                }
            }
            return "";
        }

    }


}



