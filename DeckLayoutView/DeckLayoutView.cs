using Hamilton.HxSys3DView;
using HxSysDeckLib;
using System;
using System.Drawing.Imaging;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using System.Threading.Tasks;

namespace Test
{
    [Guid("04988F73-2593-4D69-AFF5-40B6B7EC8440")]
    [ComVisible(true)]
    public interface IDeckLayoutView
    {
        void ShowDeck2(object deck);
        void ShowDeck(string file);
        void ConvertDeckToImage2(object deck, string imageFile);
        void ConvertDeckToImage(string deckFile, string imageFile);
    }
    [Guid("3B6FC438-9185-49C6-91F8-B47273DAE0C5")]
    [ClassInterface(ClassInterfaceType.None)]
    [ProgId("Test.DeckLayoutView")]
    [ComVisible(true)]
    public class DeckLayoutView : IDeckLayoutView
    {
        public void ShowDeck2(object deck)
        {
            Thread t = new Thread(() =>
            {
                var form = Show3DSystemViewForm(deck);
                form.ShowDialog();
            });
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();
        }
        public void ConvertDeckToImage2(object deck, string imageFile)
        {
            Thread t = new Thread(() =>
            {
                var form = new Form();
                form.Width = 800;
                form.Height = 600;
                form.Text = "3D system view";
                form.TopMost = true;
                var panel = new Panel();
                form.Controls.Add(panel);
                panel.Dock = DockStyle.Fill;
                HxInstrument3DView view = new HxInstrument3DView();
                view.Initialize2(panel.Handle.ToInt32(), deck);
                view.Mode = HxSys3DViewMode.RunVisualization;
                view.ModifyEnable = false;
                Task.Run(() =>
                {
                    Thread.Sleep(1000);
                    RECT rect = new RECT();
                    GetWindowRect(panel.Handle, ref rect);
                    double scale = GetWindowsScreenScalingFactor();
                    CaptureScreen(rect.Left * scale, rect.Top * scale, scale * (rect.Right - rect.Left), scale * (rect.Bottom - rect.Top), imageFile);
                    form.Close();
                });
                form.ShowDialog();
            });
            t.SetApartmentState(ApartmentState.STA);
            t.Start();
            t.Join();
        }
        public void ConvertDeckToImage(string deckFile, string imageFile)
        {
            var HxSystemDeck = new SystemDeck();
            HxSystemDeck.InitSystemFromFile(deckFile);
            var deck = HxSystemDeck.GetInstrumentLayout("ML_STAR", null);
            ConvertDeckToImage2(deck, imageFile);
        }
        public void ShowDeck(string file)
        {
            var HxSystemDeck = new SystemDeck();
            HxSystemDeck.InitSystemFromFile(file);
            var deck = HxSystemDeck.GetInstrumentLayout("ML_STAR", null);
            ShowDeck2(deck);
        }
        static Form Show3DSystemViewForm(object HxSystemDeck)
        {
            var form = new Form();
            form.Width = 800;
            form.Height = 600;
            form.Text = "3D system view";
            form.TopMost = true;
            var panel = new Panel();
            form.Controls.Add(panel);
            panel.Dock = DockStyle.Fill;
            HxInstrument3DView view = new HxInstrument3DView();
            view.Initialize2(panel.Handle.ToInt32(), HxSystemDeck);
            view.Mode = HxSys3DViewMode.RunVisualization;
            view.ModifyEnable = false;
            return form;
        }
        public void CaptureScreen(double x, double y, double width, double height, string imageFile)
        {
            int ix, iy, iw, ih;
            ix = Convert.ToInt32(x);
            iy = Convert.ToInt32(y);
            iw = Convert.ToInt32(width);
            ih = Convert.ToInt32(height);
            Bitmap image = new Bitmap(iw, ih, System.Drawing.Imaging.PixelFormat.Format32bppArgb);
            Graphics g = Graphics.FromImage(image);
            g.CopyFromScreen(ix, iy, 0, 0, new System.Drawing.Size(iw, ih), CopyPixelOperation.SourceCopy);
            image.Save(imageFile, ImageFormat.Png);
            g.Dispose();
            image.Dispose();
        }
        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }
        [DllImport("user32.dll")]
        public static extern IntPtr GetWindowRect(IntPtr hWnd, ref RECT rect);
        [DllImport("gdi32.dll", CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]
        public static extern int GetDeviceCaps(IntPtr hDC, int nIndex);
        public enum DeviceCap
        {
            VERTRES = 10,
            DESKTOPVERTRES = 117
        }
        static double GetWindowsScreenScalingFactor()
        {
            Graphics GraphicsObject = Graphics.FromHwnd(IntPtr.Zero);
            IntPtr DeviceContextHandle = GraphicsObject.GetHdc();
            int LogicalScreenHeight = GetDeviceCaps(DeviceContextHandle, (int)DeviceCap.VERTRES);
            int PhysicalScreenHeight = GetDeviceCaps(DeviceContextHandle, (int)DeviceCap.DESKTOPVERTRES);
            double ScreenScalingFactor = Math.Round((double)PhysicalScreenHeight / (double)LogicalScreenHeight, 2);
            GraphicsObject.ReleaseHdc(DeviceContextHandle);
            GraphicsObject.Dispose();
            return ScreenScalingFactor;
        }
    }
}
