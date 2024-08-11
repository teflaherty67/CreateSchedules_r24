using System.Windows.Media.Imaging;

namespace CreateSchedules_r24
{
    internal class clsButtonData
    {
        public PushButtonData Data { get; set; }
        public clsButtonData(string name, string text, string className,
            System.Drawing.Bitmap largeImage,
            System.Drawing.Bitmap smallImage,
            string toolTip)
        {
            Data = new PushButtonData(name, text, GetAssemblyName(), className);
            Data.ToolTip = toolTip;

            Data.LargeImage = BitmapToImageSource(largeImage);
            Data.Image = BitmapToImageSource(smallImage);

            // set command availability
            Data.AvailabilityClassName = "CreateSchedules_r24.Utils.CommandAvailability";
        }
        public clsButtonData(string name, string text, string className,
            System.Drawing.Bitmap largeImage,
            System.Drawing.Bitmap smallImage,
            System.Drawing.Bitmap largeImageDark,
            System.Drawing.Bitmap smallImageDark,
            string toolTip)
        {
            Data = new PushButtonData(name, text, GetAssemblyName(), className);
            Data.ToolTip = toolTip;

            Data.LargeImage = BitmapToImageSource(largeImage);
            Data.Image = BitmapToImageSource(smallImage);

            // set command availability
            Data.AvailabilityClassName = "CreateSchedules_r24.Utils.CommandAvailability";
        }
        public static Assembly GetAssembly()
        {
            return Assembly.GetExecutingAssembly();
        }
        public static string GetAssemblyName()
        {
            return GetAssembly().Location;
        }
        public static BitmapImage BitmapToImageSource(System.Drawing.Bitmap bm)
        {
            using (MemoryStream mem = new MemoryStream())
            {
                bm.Save(mem, System.Drawing.Imaging.ImageFormat.Png);
                mem.Position = 0;
                BitmapImage bmi = new BitmapImage();
                bmi.BeginInit();
                bmi.StreamSource = mem;
                bmi.CacheOption = BitmapCacheOption.OnLoad;
                bmi.EndInit();

                return bmi;
            }
        }
    }
}
