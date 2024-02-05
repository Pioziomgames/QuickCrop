using Microsoft.Win32;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace QuickCrop
{
    internal class Program
    {
        static void Main(string[] args)
        {
            if (args.Length > 0 && File.Exists(args[0]))
            {
                try
                {
                    Image ogFile = Image.FromFile(args[0]);

                    //Not much point in supporting anything else other than png since no one uses anything else anymore
                    //ImageFormat ogFormat = ogFile.RawFormat;

                    Bitmap output = CropToContent(ogFile);
                    ogFile.Dispose(); //Stop the lock on the original file

                    //output.Save(args[0], ogFormat);
                    output.Save(args[0], ImageFormat.Png);
                }
                catch (Exception ex)
                {
                    Console.WriteLine("An error occured while cropping the image:\n" + ex.Message);
                    Quit();
                }
            }
            else
            {
                ShowConsoleWindow(); //Show the console window (it will not appear otherwise 
                //because the application is set to Windows Application)
                if (args.Length > 0 && args[0].ToLower() == "-a")
                {
                    bool failed = false;
                    string EntryName = "Crop Image to Content";
                    if (args.Length > 1)
                        EntryName = args[1];

                    try
                    {
                        Registry.CurrentUser.CreateSubKey("Software\\Classes\\SystemFileAssociations\\.png");
                        using (RegistryKey? png = Registry.CurrentUser.OpenSubKey("Software\\Classes\\SystemFileAssociations\\.png", true))
                        {

                            if (png != null)
                            {
                                png.CreateSubKey("shell");
                                png.CreateSubKey("shell\\QuickCrop");
                                png.CreateSubKey("shell\\QuickCrop\\command");
                                using (RegistryKey? quickCrop = png.OpenSubKey("shell\\QuickCrop", true))
                                {
                                    if (quickCrop != null)
                                    {
                                        quickCrop?.SetValue(string.Empty, EntryName, RegistryValueKind.String);
                                        quickCrop?.Close();
                                    }
                                    else
                                        failed = true;
                                }
                                using (RegistryKey? command = png.OpenSubKey("shell\\QuickCrop\\command", true))
                                {
                                    if (command != null)
                                    {
                                        //Everything else return the path to the dll file so we're using this
                                        command?.SetValue(string.Empty, $"\"{System.Diagnostics.Process.GetCurrentProcess()?.MainModule?.FileName}\" \"%1\"", RegistryValueKind.String);
                                        command?.Close();
                                    }
                                    else
                                        failed = true;
                                }
                                png.Close();
                            }
                            else failed = true;
                        }
                    }
                    catch
                    {
                        failed = true;
                    }
                    if (failed)
                        Console.WriteLine("Unable to add QuickCrop to the context menu of png files");
                    else
                        Console.WriteLine("QuickCrop was added to the context menu of png files");
                    Quit();
                }
                else if (args.Length > 0 && args[0].ToLower() == "-r")
                {
                    try
                    {
                        Registry.CurrentUser.DeleteSubKeyTree("Software\\Classes\\SystemFileAssociations\\.png\\Shell\\QuickCrop");
                        Console.WriteLine("QuickCrop was removed from the context menu of png files");
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Unable to remove QuickCrop from the context menu of png files:\n" + ex.Message);
                    }
                    Quit();
                }
                else
                {
                    Console.WriteLine("QuickCrop by Pioziomgames");
                    Console.WriteLine("A program for that allows quick cropping to content of png files.");
                    Console.WriteLine("Cropping to content refers to removing of all fully transparent pixel lines surrounding the visible image data.)");
                    Console.WriteLine("Usage:\n");
                    Console.WriteLine("QuickCrop.exe {file.png} //Crop the file to content");
                    Console.WriteLine("QuickCrop.exe -a //Add QuickCrop to the context menu of png files (Can also be used to update the executable location)");
                    Console.WriteLine("QuickCrop.exe -a {Name of the Context Menu Option} //Add QuickCrop to the context menu but with a custom option name");
                    Console.WriteLine("QuickCrop.exe -r //Remove QuickCrop from the context menu of png files");
                    Quit();
                }
            }
        }
        private static Bitmap CropToContent(Image image)
        {
            Bitmap img = new Bitmap(image); //Make sure that the image is 32bit

            //Find the visible bounds of the image
            object lockObject = new object();
            int minX = img.Width;
            int minY = img.Height;
            int maxX = 0, maxY = 0;

            BitmapData bmpData = img.LockBits(new Rectangle(0, 0, img.Width, img.Height), ImageLockMode.ReadOnly, img.PixelFormat);

            try
            {
                int bytesPerPixel = Image.GetPixelFormatSize(img.PixelFormat) / 8;
                int heightInPixels = bmpData.Height;
                int widthInBytes = bmpData.Width * bytesPerPixel;

                //Use parallel to greatly speed up the process
                Parallel.For(0, heightInPixels, y =>
                {
                    //Get the pointer to the current line
                    IntPtr ptr = IntPtr.Add(bmpData.Scan0, y * bmpData.Stride);

                    for (int x = 0; x < widthInBytes; x += bytesPerPixel)
                    {
                        Color pixelColor = Color.FromArgb(Marshal.ReadInt32(ptr, x));
                        //Check if pixel is visible
                        if (pixelColor.A > 0)
                        {
                            lock (lockObject)
                            {
                                minX = Math.Min(x / bytesPerPixel, minX);
                                minY = Math.Min(y, minY);
                                maxX = Math.Max(x / bytesPerPixel, maxX);
                                maxY = Math.Max(y, maxY);
                            }
                        }
                    }
                });
            }
            finally
            {
                img.UnlockBits(bmpData);
            }

            int width = maxX - minX + 1;
            int height = maxY - minY + 1;
            if (width == img.Width && height == img.Height)
                return img;

            //Build the output image
            Bitmap croppedImage = new Bitmap(width, height);

            BitmapData srcData = img.LockBits(new Rectangle(minX, minY, width, height), ImageLockMode.ReadOnly, img.PixelFormat);
            BitmapData destData = croppedImage.LockBits(new Rectangle(0, 0, width, height), ImageLockMode.WriteOnly, croppedImage.PixelFormat);

            try
            {
                int bytesPerPixel = Image.GetPixelFormatSize(img.PixelFormat) / 8;
                int srcStride = srcData.Stride;
                int destStride = destData.Stride;

                Parallel.For(0, height, y =>
                {
                    IntPtr srcPtr = IntPtr.Add(srcData.Scan0, y * srcStride);
                    IntPtr destPtr = IntPtr.Add(destData.Scan0, y * destStride);
                    for (int x = 0; x < width; x++)
                    {
                        Color pixelColor = Color.FromArgb(Marshal.ReadInt32(srcPtr, x * bytesPerPixel));
                        Marshal.WriteInt32(destPtr, x * bytesPerPixel, pixelColor.ToArgb());
                    }
                });
            }
            finally
            {
                img.UnlockBits(srcData);
                croppedImage.UnlockBits(destData);
            }
            return croppedImage;
        }
        private static void Quit()
        {
            Console.WriteLine("\n\nPress Any Button to Exit");
            Console.ReadKey();
            Environment.Exit(0);
        }
        [DllImport("kernel32.dll", SetLastError = true)]
        static extern bool AllocConsole();

        [DllImport("kernel32.dll")]
        static extern IntPtr GetConsoleWindow();

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
        [DllImport("kernel32.dll")]
        static extern bool AttachConsole(int dwProcessId);
        public static void ShowConsoleWindow()
        {
            if (!AttachConsole(-1)) //Try to attach to an existing console first
                AllocConsole(); //If not allocate a new one

            var handle = GetConsoleWindow();
            if (handle != IntPtr.Zero)
                ShowWindow(handle, 5); //SW_SHOW
        }
    }
}