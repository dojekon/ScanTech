using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Xamarin.Essentials;
using Xamarin.Forms;
using ZXing.Mobile;
using ZXing.Net.Mobile.Forms;

namespace ScanTech {
    public partial class MainPage : ContentPage {

        System.Data.DataTable table;

        void Clean() {
            Invent.Text = "";
            Mol.Text = "";
            Place.Text = "";
            Naim.Text = "";
            Serial.Text = "";
            Message.Text = "";
        }

        public MainPage() {
            InitializeComponent();
            var byteArray = ScanTech.Properties.Resources._base;
            Stream stream = new MemoryStream(byteArray);

            using (var reader = ExcelReaderFactory.CreateReader(stream)) {
                    table = reader.AsDataSet().Tables["коротко"];
                }
            
        }

        private void Button_Clicked(object sender, EventArgs e) {
            


            var scan = new ZXingScannerPage();
            Navigation.PushAsync(scan);
            scan.OnScanResult += (scannedCode) => {
                Device.BeginInvokeOnMainThread(async () => {
                    Navigation.PopAsync();
                    Clean();
                    System.Data.DataRow result = table.Select("Column5 LIKE '%" + scannedCode + "%'").FirstOrDefault();

                    try {
                        Invent.Text = result.ItemArray[1].ToString();
                        Mol.Text = result.ItemArray[2].ToString().Trim();
                        Place.Text = result.ItemArray[3].ToString().Trim();
                        Naim.Text = result.ItemArray[4].ToString().Trim();
                        Serial.Text = result.ItemArray[5].ToString().Trim();
                    } catch { Message.Text = "Ничего не найдено :("; }
                    
                });
            };

        }

        public CameraResolution SelectLowestResolutionMatchingDisplayAspectRatio(List<CameraResolution> availableResolutions) {
            CameraResolution result = null;
            //a tolerance of 0.1 should not be visible to the user
            double aspectTolerance = 0.1;
            var displayOrientationHeight = DeviceDisplay.MainDisplayInfo.Orientation == DisplayOrientation.Portrait ? DeviceDisplay.MainDisplayInfo.Height : DeviceDisplay.MainDisplayInfo.Width;
            var displayOrientationWidth = DeviceDisplay.MainDisplayInfo.Orientation == DisplayOrientation.Portrait ? DeviceDisplay.MainDisplayInfo.Width : DeviceDisplay.MainDisplayInfo.Height;
            //calculatiing our targetRatio
            var targetRatio = displayOrientationHeight / displayOrientationWidth;
            var targetHeight = displayOrientationHeight;
            var minDiff = double.MaxValue;
            //camera API lists all available resolutions from highest to lowest, perfect for us
            //making use of this sorting, following code runs some comparisons to select the lowest resolution that matches the screen aspect ratio and lies within tolerance
            //selecting the lowest makes Qr detection actual faster most of the time
            foreach (var r in availableResolutions.Where(r => Math.Abs(((double)r.Width / r.Height) - targetRatio) < aspectTolerance)) {
                //slowly going down the list to the lowest matching solution with the correct aspect ratio
                if (Math.Abs(r.Height - targetHeight) < minDiff)
                    minDiff = Math.Abs(r.Height - targetHeight);
                result = r;
            }
            return result;
        }

       
    }
}
