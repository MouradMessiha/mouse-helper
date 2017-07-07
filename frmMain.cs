using System;
using System.Drawing;
using System.Windows.Forms;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;

namespace MouseHelper
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private struct MSLLHOOKSTRUCT
        {
            public Point pt;
            public Int32 mouseData;
            public Int32 flags;
            public Int32 time;
            public IntPtr extra;
        }
        
        // mouse hook variables
        private IntPtr _mouseHook;
        private int mintScreenHeight;
        private int mintScreenWidthsTotal;
        private const Int32 WH_MOUSE_LL = 14;
        private const Int32 MOUSEEVENTF_WHEEL = 0x0800;
        private const Int32 MOUSEEVENTF_LEFTDOWN = 0x02; 
        private const Int32 MOUSEEVENTF_LEFTUP = 0x04;
        private const Int32 MOUSEEVENTF_RIGHTDOWN = 0x08;
        private const Int32 MOUSEEVENTF_RIGHTUP = 0x10;
        private delegate Int32 CallBack(Int32 nCode, IntPtr wParam, ref MSLLHOOKSTRUCT lParam);
        [MarshalAs(UnmanagedType.FunctionPtr)] private CallBack _mouseProc;

        [DllImport("user32.dll")]
        static extern IntPtr SetWindowsHookExW(Int32 idHook, CallBack HookProc, IntPtr hInstance, Int32 wParam);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool UnhookWindowsHookEx(IntPtr hhk);
        [DllImport("user32.dll")]
        static extern Int32 CallNextHookEx(Int32 idHook, Int32 nCode, IntPtr wParam, MSLLHOOKSTRUCT lParam);
        [DllImport("kernel32.dll")]
        static extern IntPtr GetModuleHandleW(IntPtr fakezero);
        [DllImport("user32.dll")]
        static extern void mouse_event(int dwFlags, int dx, int dy, int dwData, int dwExtraInfo);
        [DllImport("user32.dll")]
        static extern int GetForegroundWindow();
        [DllImport("user32")]
        static extern UInt32 GetWindowThreadProcessId(Int32 hWnd, out Int32 lpdwProcessId);
        [DllImport("User32")]
        static extern int SetForegroundWindow(IntPtr hwnd);
        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        static extern bool IsIconic(IntPtr hWnd);
        
        // keyboard hook variables
        private const int WH_KEYBOARD_LL = 13;
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_KEYUP = 0x101;
        private static LowLevelKeyboardProc _proc = HookCallback;
        private static IntPtr _hookID = IntPtr.Zero;
        private const int KEYEVENTF_EXTENDEDKEY = 0x1;
        private const int KEYEVENTF_KEYUP       = 0x2;  

        private delegate IntPtr LowLevelKeyboardProc(int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern IntPtr SetWindowsHookEx(int idHook, LowLevelKeyboardProc lpfn, IntPtr hMod, uint dwThreadId);
        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        static extern IntPtr GetModuleHandle(string lpModuleName);
        [DllImport("user32.dll")]
        static extern void keybd_event(byte bVk, byte bScan, uint dwFlags, int dwExtraInfo);
        [DllImport("user32.dll")]
        static extern short GetKeyState(int keyCode);
        
        private bool mblnMouseToLeft = false;          // to avoid multiple runs when mouse gets to the left
        private bool mblnMouseToRight = false;         // to avoid multiple runs when mouse gets to the right
        private bool mblnMouseToTopLeft = false;       // to avoid multiple runs when mouse gets in the top left corner
        private static bool mblnMouseFunctionsActive = true;    // flag if all the functions should be active

        private bool mblnCapsLockSavedState;
        private bool mblnNumLockSavedState;
        private bool mblnScrollLockSavedState;

        static bool mblnSkipKeyboardEvent = false;
        static bool mblnKeyDown = false;
        static bool mblnMouseDown = false;

        private void frmMain_Load(object sender, EventArgs e)
        {
            if (GetKeyState((int)Keys.NumLock) == 0)   // turn num lock on at the startup, Windows 7 always starts with num lock off
            {
                // click num lock to switch it back
                keybd_event((byte)Keys.NumLock, 0x45, KEYEVENTF_EXTENDEDKEY, 0);                    // key down
                keybd_event((byte)Keys.NumLock, 0x45, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0);  // key up
            }

            // save lock key states
            mblnCapsLockSavedState = GetKeyState((int)Keys.CapsLock) != 0;
            mblnNumLockSavedState = GetKeyState((int)Keys.NumLock) != 0;
            mblnScrollLockSavedState = GetKeyState((int)Keys.Scroll) != 0;

            objNotifyIcon.Icon = Properties.Resources.CursorYellow;
            Hide();

            mintScreenHeight = Screen.AllScreens[0].Bounds.Height;
            mintScreenWidthsTotal = 0;
            for (int intScreenCount = 0; intScreenCount < Screen.AllScreens.Length; intScreenCount++)
            {
                mintScreenWidthsTotal += Screen.AllScreens[intScreenCount].Bounds.Width;
            }

            // mouse hook
            InstallHook();

            // keyboard hook
            _hookID = SetHook(_proc);
        }

        private void frmMain_FormClosed(object sender, FormClosedEventArgs e)
        {
            // remove mouse hook
            RemoveHook();

            // remove keyboard hook
            UnhookWindowsHookEx(_hookID);
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        // Mouse hook functions

        public bool InstallHook()
        {
            if (_mouseHook == IntPtr.Zero)
            {
                _mouseProc = new CallBack(MouseHookProc);
                _mouseHook = SetWindowsHookExW(WH_MOUSE_LL, _mouseProc, GetModuleHandleW(IntPtr.Zero), 0);
            }
            return _mouseHook != IntPtr.Zero;
        }

        public void RemoveHook()
        {
            if (_mouseHook == IntPtr.Zero) 
            {
                return;
            }
            UnhookWindowsHookEx(_mouseHook);
            _mouseHook = IntPtr.Zero;
        }

        private Int32 MouseHookProc(Int32 nCode,IntPtr wParam,ref MSLLHOOKSTRUCT lParam)
        {
            if (wParam.ToInt32() == 513)  // mouse down
            {
                mblnMouseDown = true;
            }
            else
            {
                if (wParam.ToInt32() == 514)  // mouse up
                {
                    mblnMouseDown = false;
                }
                else
                {
                    if (wParam.ToInt32() == 512)  // mouse move
                    {
                        if (lParam.pt.X <= 20 & lParam.pt.Y <= 20 & !mblnMouseToTopLeft) // top left corner of screen
                        {
                            mblnMouseToTopLeft = true;  // avoid multiple toggles simultaneously

                            // toggle keyboard clicks functionality flag
                            if (!mblnMouseFunctionsActive)
                            {
                                mblnMouseFunctionsActive = true;
                                objNotifyIcon.Icon = Properties.Resources.CursorYellow;
                                // save lock key states
                                mblnCapsLockSavedState = GetKeyState((int)Keys.CapsLock) != 0;
                                mblnNumLockSavedState = GetKeyState((int)Keys.NumLock) != 0;
                                mblnScrollLockSavedState = GetKeyState((int)Keys.Scroll) != 0;
                            }
                            else
                            {
                                mblnMouseFunctionsActive = false;
                                objNotifyIcon.Icon = Properties.Resources.CursorBlack;
                            }
                        }

                        // restore the lock key states if mouse functions active(because caps lock is used to simulate mouse clicks)
                        if (mblnMouseFunctionsActive)
                        {
                            if (mblnCapsLockSavedState != (GetKeyState((int)Keys.CapsLock) != 0))
                            {
                                // click caps lock to switch it back
                                mblnSkipKeyboardEvent = true;
                                keybd_event((byte)Keys.CapsLock, 0x45, KEYEVENTF_EXTENDEDKEY, 0);                    // key down
                                keybd_event((byte)Keys.CapsLock, 0x45, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0);  // key up
                                mblnSkipKeyboardEvent = false;
                            }

                            if (mblnNumLockSavedState != (GetKeyState((int)Keys.NumLock) != 0))
                            {
                                // click num lock to switch it back
                                mblnSkipKeyboardEvent = true;
                                keybd_event((byte)Keys.NumLock, 0x45, KEYEVENTF_EXTENDEDKEY, 0);                    // key down
                                keybd_event((byte)Keys.NumLock, 0x45, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0);  // key up
                                mblnSkipKeyboardEvent = false;
                            }

                            if (mblnScrollLockSavedState != (GetKeyState((int)Keys.Scroll) != 0))
                            {
                                // click scroll lock to switch it back
                                mblnSkipKeyboardEvent = true;
                                keybd_event((byte)Keys.Scroll, 0x45, KEYEVENTF_EXTENDEDKEY, 0);                    // key down
                                keybd_event((byte)Keys.Scroll, 0x45, KEYEVENTF_EXTENDEDKEY | KEYEVENTF_KEYUP, 0);  // key up
                                mblnSkipKeyboardEvent = false;
                            }
                        }

                        if (lParam.pt.X > 30 | lParam.pt.Y > 30 & mblnMouseToTopLeft) // allow another toggle now that it's outside the top left block
                        {
                            mblnMouseToTopLeft = false;
                        }

                        if (lParam.pt.Y <= 0 & lParam.pt.X > 20 & mblnMouseFunctionsActive) // mouse at the top of the screen
                        {
                            mouse_event(MOUSEEVENTF_WHEEL, 0, 0, 10, 0);  // scroll up
                        }
                        if (lParam.pt.Y >= mintScreenHeight - 1 & mblnMouseFunctionsActive) // mouse at the bottom of the screen   
                        {
                            mouse_event(MOUSEEVENTF_WHEEL, 0, 0, -10, 0);  // scroll down
                        }

                        if (lParam.pt.X <= 0 & lParam.pt.Y > 20 & !mblnMouseToLeft & mblnMouseFunctionsActive) // mouse at the left of the screen
                        {
                            mblnMouseToLeft = true;         // prevent multiple runs as mouse moves in the leftmost of the screen
                            if (!mblnMouseDown)
                            {
                                SwitchApplication(false);       // switch backward between non minimized applications
                            }
                        }

                        if (lParam.pt.X > 5)  // allow another mouse to left functionality
                        {
                            mblnMouseToLeft = false;
                        }

                        if (lParam.pt.X >= mintScreenWidthsTotal - 1 & !mblnMouseToRight & mblnMouseFunctionsActive) // mouse at the right of the screens
                        {
                            mblnMouseToRight = true;         // prevent multiple runs as mouse moves in the rightmost of the screens
                            if (!mblnMouseDown)
                            {
                                SwitchApplication(true);          // switch forward between non minimized applications
                            }
                        }

                        if (lParam.pt.X < mintScreenWidthsTotal - 6)  // allow another mouse to right functionality
                        {
                            mblnMouseToRight = false;
                        }
                    }
                }
            }
            
            return CallNextHookEx(WH_MOUSE_LL, nCode, wParam, lParam);
        }

        void SwitchApplication(bool pblnDirectionForward)
        {
            // switch between non minimized windows applications, forward or backward
            Process[] lstProcesses = Process.GetProcesses();   // get list of all processes
            int intProcessIndex = 0;
            bool blnFound = false;
            int intLoopCount;

            int hwnd;
            hwnd = GetForegroundWindow();
            Int32 intProcessID = 1;
            GetWindowThreadProcessId(hwnd, out intProcessID);

            while (intProcessIndex < lstProcesses.Length & !blnFound)
            {
                if (lstProcesses[intProcessIndex].Id == intProcessID)
                {
                    blnFound = true;
                }
                if (!blnFound)
                {
                    intProcessIndex++;
                }
            }

            if (blnFound)
            {
                if (pblnDirectionForward)
                {
                    intProcessIndex++;
                    if (intProcessIndex >= lstProcesses.Length)
                    {
                        intProcessIndex = 0;
                    }
                }
                else
                {
                    intProcessIndex--;
                    if (intProcessIndex < 0)
                    {
                        intProcessIndex = lstProcesses.Length - 1;
                    }
                }
                blnFound = false;
                intLoopCount = 0;

                while (!blnFound & intLoopCount < lstProcesses.Length)
                {
                    if (lstProcesses[intProcessIndex].MainWindowTitle != "")
                    {
                        if (!IsIconic(lstProcesses[intProcessIndex].MainWindowHandle))
                        {
                            blnFound = true;
                        }
                    }

                    if (!blnFound)
                    {
                        if (pblnDirectionForward)
                        {
                            intProcessIndex++;
                            if (intProcessIndex >= lstProcesses.Length)
                            {
                                intProcessIndex = 0;
                            }
                        }
                        else
                        {
                            intProcessIndex--;
                            if (intProcessIndex < 0)
                            {
                                intProcessIndex = lstProcesses.Length - 1;
                            }
                        }
                    }
                    intLoopCount++;
                }
                if (blnFound)
                {
                    SetForegroundWindow(lstProcesses[intProcessIndex].MainWindowHandle);
                }
            }
        }

        // Keyboard hook functions
        private IntPtr SetHook(LowLevelKeyboardProc proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        private static IntPtr HookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 & mblnMouseFunctionsActive & !mblnSkipKeyboardEvent)
            {
                int vkCode = Marshal.ReadInt32(lParam);
                if (wParam == (IntPtr)WM_KEYDOWN )
                {
                    if (!mblnKeyDown)   // to avoid key repeat when it's held down for a long time
                    {
                        mblnKeyDown = true;
                        if (vkCode == (Int32)Keys.CapsLock)
                        {
                            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                            mblnMouseDown = true;
                        }
                        if (vkCode == (Int32)Keys.Scroll)
                        {
                            mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0);
                            mblnMouseDown = true;
                        }
                        if (vkCode == (Int32)Keys.NumLock)
                        {
                            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                            mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                        }
                        if (vkCode == (Int32)Keys.PrintScreen)
                        {
                            mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                            mblnMouseDown = true;
                        }
                    }
                }
                if (wParam == (IntPtr)WM_KEYUP )
                {
                    mblnKeyDown = false;    
                    if (vkCode == (Int32)Keys.CapsLock)
                    {
                        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                        mblnMouseDown = false;
                    }
                    if (vkCode == (Int32)Keys.Scroll)
                    {
                        mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0);
                        mblnMouseDown = false;
                    }
                    if (vkCode == (Int32)Keys.PrintScreen)
                    {
                        mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                        mblnMouseDown = false;
                    }
                }
            }

            return CallNextHookEx(_hookID, nCode, wParam, lParam);
        }

        private void frmMain_Resize(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                Hide();
            }
        }

        private void objNotifyIcon_MouseClick(object sender, MouseEventArgs e)
        {
            switch (e.Button)
            {
                case MouseButtons.Right:
                    Show();
                    this.WindowState = FormWindowState.Normal;
                    break;
            }
        }
    }
}
