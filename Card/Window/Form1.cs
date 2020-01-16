using ConsoleApp4;

using System;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Window
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            this.Visible = false;
            this.ShowInTaskbar = false;
            Task.Run(() =>
            {
                Run();
            });
            this.Hide();
        }

        private void Hide()
        {
            while (true)
            {
                ShowWindow(this.Handle, 0);
                Thread.Sleep(1000 * 60);
            }
        }

        static async Task Run()
        {
            Boolean doit = false;
            while (true)
            {
                try
                {
                    if (Helper.Sleep(ref doit) == true)
                    {
                        await Helper.PostCard();
                        doit = true;
                    }
                }
                catch (Exception)
                {
                    //ToDo
                }
            }
        }

        [DllImport("user32.dll", EntryPoint = "ShowWindow", SetLastError = true)]
        private static extern Boolean ShowWindow(IntPtr hWnd, int nCmdShow);

    }
}
