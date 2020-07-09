using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Drawing;
using System.Runtime.InteropServices;

namespace AutoBook
{
    class MouseEvent
    {

        public const int SM_CXSCREEN = 0;
        public const int SM_CYSCREEN = 1;

        [DllImport("User32.dll")]
        public static extern int GetSystemMetrics(int nIndex);

        //移动鼠标 
        public const int MOUSEEVENTF_MOVE = 0x0001;
        //模拟鼠标左键按下 
        public const int MOUSEEVENTF_LEFTDOWN = 0x0002;
        //模拟鼠标左键抬起 
        public const int MOUSEEVENTF_LEFTUP = 0x0004;
        //模拟鼠标右键按下 
        public const int MOUSEEVENTF_RIGHTDOWN = 0x0008;
        //模拟鼠标右键抬起 
        public const int MOUSEEVENTF_RIGHTUP = 0x0010;
        //模拟鼠标中键按下 
        public const int MOUSEEVENTF_MIDDLEDOWN = 0x0020;
        //模拟鼠标中键抬起 
        public const int MOUSEEVENTF_MIDDLEUP = 0x0040;
        //标示是否采用绝对坐标 
        public const int MOUSEEVENTF_ABSOLUTE = 0x8000;
        [DllImport("user32.dll")]
        public static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);

        public static void EventMouseMove(int x, int y)
        {
            UtilsLog.Log("EventMouseMove ENTER " + "X=" + x.ToString() + " Y=" + y.ToString());
            //SetCursorPos(x, y);

            int mx = x * 65535 / GetSystemMetrics(SM_CXSCREEN);
            int my = y * 65535 / GetSystemMetrics(SM_CYSCREEN);
            mouse_event(MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_MOVE, mx, my, 0, 0);
            Thread.Sleep(100);
        }

        public static void EventMouseClick(int x, int y)
        {
            UtilsLog.Log("EventMouseClick ENTER " + "X=" + x.ToString() + "Y=" + y.ToString());

            int mx = x * 65535 / GetSystemMetrics(SM_CXSCREEN);
            int my = y * 65535 / GetSystemMetrics(SM_CYSCREEN);

            mouse_event(MOUSEEVENTF_ABSOLUTE | MOUSEEVENTF_MOVE, mx, my, 0, 0);
            Thread.Sleep(100);

            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_ABSOLUTE, mx, my, 0, 0);
            Thread.Sleep(100);
            mouse_event(MOUSEEVENTF_LEFTUP | MOUSEEVENTF_ABSOLUTE, mx, my, 0, 0);
        }
    }
}
