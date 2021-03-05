using System;
using System.Collections.Generic;
using System.Text;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Security.Principal;


namespace RoboVMonitorApp
{
    /// <summary>
    /// Delegate for retrieving the Browser window opened during click event NXTT Maintenace Test
    /// </summary>
    /// <param name="hwnd"></param>
    /// <param name="lParam"></param>
    /// <returns></returns>
    public delegate bool IECallBack(int hwnd, int lParam);

    public class Win32
    {

        //API Constants
        public const ushort WM_KEYDOWN = 0x0100;
        public const ushort WM_KEYUP = 0x0101;
        public const uint WM_CHAR = 0x0102;
        public const uint WM_SETTEXT = 12;
        public const uint WM_GETTEXT = 13;
        public const uint WM_GETTEXTLENGTH = 14;
        public const uint WM_COMMAND = 273;
        public const uint WM_ACTIVATE = 6;
        public const uint WM_SETFOCUS = 7;
        public const uint WM_CLOSE = 16;
        public const int BN_CLICKED = 245;

        public const uint BM_CLICK = 245;

        public const int WM_SYSCOMMAND = 274;
        public const uint SC_MAXIMIZE = 61488;

        public const int SW_HIDE = 0;
        public const int SW_SHOWNORMAL = 1;
        public const int SW_SHOWMINIMIZED = 2;
        public const int SW_SHOWMAXIMIZED = 3;
        public const int SW_SHOWNOACTIVATE = 4;
        public const int SW_SHOW = 5;
        public const int SW_RESTORE = 9;

        public const int SW_SHOWDEFAULT = 10;

        public const int GW_CHILD = 5;
        public const int GW_HWNDFIRST = 0;
        public const int GW_HWNDLAST = 1;
        public const int GW_HWNDNEXT = 2;
        public const int GW_HWNDPREV = 3;
        public const int GW_OWNER = 4;


        [DllImport("user32.Dll")]
        public static extern int EnumWindows(IECallBack x, int y);
        [DllImport("User32.dll")]
        public static extern Boolean EnumChildWindows(int hWndParent, Delegate lpEnumFunc, int lParam);

        [DllImport("User32.Dll")]
        public static extern void GetWindowText(int h, StringBuilder s, int nMaxCount);
        [DllImport("User32.Dll")]
        public static extern void GetClassName(int h, StringBuilder s, int nMaxCount);
        
        [DllImport("User32.Dll")]
        public static extern void GetWindowText(IntPtr h, StringBuilder s, int nMaxCount);
        
        [DllImport("User32.Dll")]
        public static extern void GetClassName(IntPtr h, StringBuilder s, int nMaxCount);
        
        [DllImport("User32.Dll")]
        public static extern IntPtr PostMessage(IntPtr hWnd, int msg, int wParam, int lParam);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool PostMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32", EntryPoint = "FindWindowEx")]
        public static extern IntPtr FindWindowExA(IntPtr hWnd1, IntPtr hWnd2, String lpsz1, String lpsz2);

        [DllImport("user32")]
        public static extern IntPtr GetWindow(IntPtr hwnd, int flag);

        [DllImport("user32")]
        public static extern IntPtr GetDesktopWindow();

        [DllImport("user32", EntryPoint = "GetWindowText")]
        public static extern IntPtr GetWindowTextA(IntPtr hwnd, string lpString, int cch);

        [DllImport("user32")]
        public static extern IntPtr GetParent(IntPtr hwnd);

        [DllImport("user32.dll")]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

        [DllImport("user32")]
        public static extern IntPtr ShowWindow(IntPtr hwnd, int cmdShow);

        [DllImport("user32")]
        public static extern IntPtr BringWindowToTop(IntPtr hwnd);

        [DllImport("user32")]
        public static extern int GetWindowTextLength(IntPtr hwnd);

        [DllImport("user32")]
        public static extern bool MoveWindow(IntPtr Hwnd, int x, int y, int cx, int cy, bool repaint);
        /// <summary>
        /// The FindWindow API
        /// </summary>
        /// <param name="lpClassName">the class name for the window to search for</param>
        /// <param name="lpWindowName">the name of the window to search for</param>
        /// <returns></returns>
        [DllImport("User32.dll")]
        public static extern IntPtr FindWindow(String lpClassName, String lpWindowName);
 

        [DllImport("user32.dll")]
        public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, uint Msg, int wParam, string lParam);
        
        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, uint Msg, int wParam, StringBuilder lParam);

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, uint Msg, string wParam, int lParam);

        [DllImport("user32.dll")] //sends a windows message to the specified window
        public static extern int SendMessage(IntPtr hWnd, int Msg, int wParam, int lParam);

        [DllImport("user32.dll")] //sends a windows message to the specified window
        public static extern int SendMessage(IntPtr hWnd, uint Msg, int wParam, int lParam);

        [DllImport("user32.dll")] //sends a windows message to the specified window
        public static extern int SendMessage(IntPtr hWnd, uint Msg, uint wParam, int lParam);

        [DllImport("user32", EntryPoint = "SetWindowText")]
        public static extern int SetWindowTextA(IntPtr hwnd, string lpString);

        [DllImport("user32")]
        public static extern int SetFocus(IntPtr hwnd);

        [DllImport("user32.dll")] //Set the active window
        public static extern IntPtr SetActiveWindow(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint SetParent(IntPtr hWndChild, IntPtr hWndNewParent);

        /// <summary>
        /// The FindWindowEx API
        /// </summary>
        /// <param name="parentHandle">a handle to the parent window </param>
        /// <param name="childAfter">a handle to the child window to start search after</param>
        /// <param name="className">the class name for the window to search for</param>
        /// <param name="windowTitle">the name of the window to search for</param>
        /// <returns></returns>
        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, string windowTitle);

        /// <summary>
        /// The SendMessage API
        /// </summary>
        /// <param name="hWnd">handle to the required window</param>
        /// <param name="msg">the system/Custom message to send</param>
        /// <param name="wParam">first message parameter</param>
        /// <param name="lParam">second message parameter</param>
        /// <returns></returns>
        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern int SendMessageA(int hWnd, int msg, int wParam, IntPtr lParam);


        [DllImport("user32")]
        public static extern int SetForegroundWindow(IntPtr hwnd);

        /// <summary>
        /// The FindWindow API
        /// </summary>
        /// <param name="lpClassName">the class name for the window to search for</param>
        /// <param name="lpWindowName">the name of the window to search for</param>
        /// <returns></returns>
        [DllImport("User32.dll")]
        public static extern Int32 FindWindowA(String lpClassName, String lpWindowName);

        [DllImport("user32.dll")]
        public static extern  bool BlockInput(bool fBlockIt);

        [DllImport("user32.dll")]
        public static extern int GetWindowThreadProcessId(int hWnd, ref int processID);

        /// <summary>
        /// Get window position
        /// </summary>
        /// <param name="hwnd">Handle of the required window</param>
        /// <param name="lpRect">Output parameter that will get position of the required window</param>
        /// <returns></returns>
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(HandleRef hwnd, out RECT lpRect);

        [StructLayout(LayoutKind.Sequential)]
        private struct RECT
        {
            public int Left;        // x position of upper-left corner
            public int Top;         // y position of upper-left corner
            public int Right;       // x position of lower-right corner
            public int Bottom;      // y position of lower-right corner
        }

        /// <summary>
        /// This method is used when automation of Java Applet, which is in tab, starts
        /// </summary>
        /// <param name="hwnd">Handle of the required window</param>
        /// <param name="currentObject">Pass "this" for this parameter</param>
        /// <returns></returns>
        public static bool StartTabAutomation(IntPtr hwnd, object currentObject)
        {
            RECT rect;
            try
            {
                GetWindowRect(new HandleRef(currentObject, hwnd), out rect);
                Win32.SetParent(hwnd, IntPtr.Zero);
                SetWindowPos(hwnd, IntPtr.Zero, rect.Left, rect.Top, rect.Right - rect.Left + 1, rect.Bottom - rect.Top + 1, 1);
                return true;
            }
            catch(Exception Ex)
            {
                //Logging.InsertLog(Logging.LogSeverity.Error, 3, "Error in SetWindowToCurrentPosition() method", Ex.Message, Logging.LogTarget.LogToFile);
                return false;
            }

        }

        /// <summary>
        /// This method is used when automation of Java Applet, which is in tab, is finished
        /// </summary>
        /// <param name="hwnd">Handle of the required window</param>
        /// <param name="parentHwnd">Handle of the parent control</param>
        public static void FinishTabAutomation(IntPtr hwnd, IntPtr parentHwnd)
        {
            Win32.SetParent(hwnd, parentHwnd);
        }

        /* -- End Added by Dharmesh Gadhiya -- */

        public static void SendKey(int key, IntPtr hWnd)
        {
            SetActiveWindow(hWnd);
            SendMessage(hWnd, WM_KEYDOWN, key, 0);
            SendMessage(hWnd, WM_KEYUP, key, 0);
        }

        public static bool ShowWindow(Process _Process, int nCmdShow)
        {
            return ShowWindowAsync(_Process.MainWindowHandle, nCmdShow);
        }

        /// <summary>
        /// Returns the required window handle
        /// </summary>
        /// <param name="windowHandle">The main window handle from which the node path is defined</param>
        /// <param name="controlPath">Path to the Control Node
        /// </param>
        /// <returns></returns>
        public static IntPtr GetControlHandle(IntPtr windowHandle, string controlPath)
        {
            IntPtr controlWindowHandle = windowHandle;
            IntPtr ptr = IntPtr.Zero;
            try
            {
                for (int nodeCount = 0; nodeCount < controlPath.Length; nodeCount++)
                {
                    switch (controlPath[nodeCount])
                    {
                        case 'f':
                            controlWindowHandle = Win32.GetWindow(controlWindowHandle, Win32.GW_CHILD);
                            break;
                        case 'n':
                            controlWindowHandle = Win32.GetWindow(controlWindowHandle, Win32.GW_HWNDNEXT);
                            break;
                        case 'p': //parent node
                            controlWindowHandle = Win32.GetWindow(controlWindowHandle, Win32.GW_OWNER);
                            break;
                        case 'l': //lastchild
                            controlWindowHandle = Win32.GetWindow(controlWindowHandle, Win32.GW_HWNDLAST);
                            break;
                        default:
                            break;
                    }
                }
            }
            catch(Exception Ex)
            {
                //Logging.InsertLog(Logging.LogSeverity.Error, 3, "To Extract the required node in the element", Ex.Message, Logging.LogTarget.LogToFile);
                return IntPtr.Zero;
            }
            return controlWindowHandle;
        }

        public static StringBuilder GetTextFromTextBox(IntPtr windowHandle)
        {
            StringBuilder text;
            try
            {
                int maxLength = Win32.SendMessage(windowHandle, Win32.WM_GETTEXTLENGTH, 0, 0);
                text = new StringBuilder(maxLength + 1);
                Win32.SendMessage(windowHandle, Win32.WM_GETTEXT, text.Capacity, text);
            }
            catch(Exception Ex)
            {
                //Logging.InsertLog(Logging.LogSeverity.Error, 3, "To Extract the required node in the element", Ex.Message, Logging.LogTarget.LogToFile);
                return null;
            }
            return text;
        }

        public static StringBuilder GetTextFromControl(IntPtr windowHandle)
        {
            StringBuilder text;
            try
            {

                int maxLength = Win32.GetWindowTextLength(windowHandle);
                text = new StringBuilder(maxLength + 1);
                Win32.GetWindowText(windowHandle, text, text.Capacity);
            }
            catch(Exception Ex)
            {
                //Logging.InsertLog(Logging.LogSeverity.Error, 3, "To Extract the required node in the element", Ex.Message, Logging.LogTarget.LogToFile);
                return null;
            }
            return text;
        }
    }
}


