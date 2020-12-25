using Amazon.Rekognition;
using Amazon.Rekognition.Model;
using Amazon.Runtime;
using Amazon.Runtime.CredentialManagement;
using Amazon.S3;
using Amazon.S3.Transfer;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Linq;
using GoogleHelper;

namespace DetectTextAWS
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        string filePath;
        Bitmap image;
        private void button1_Click(object sender, EventArgs e)
        {
            DetectTextAndPostToGGsheet(image);
        }

        void DetectTextAndPostToGGsheet(Bitmap image)
        {
            List<Bitmap> bitmaps = new List<Bitmap>();

            image = CropImage(image, 33, 257, 1195, 522);
            int height = image.Height / 6;
            int width = image.Width;
            for (int i = 0; i < 6; i ++)
            {
                int posY = i * height;
                bitmaps.Add(CropImage(image, new Rectangle(0, posY ,width , height)));
            }

            AmazonRekognitionClient rekognitionClient = new AmazonRekognitionClient("", "", Amazon.RegionEndpoint.EUWest2);

            int index = 3;
            foreach (Bitmap anh in bitmaps)
            {
                List<string> data = SplitStr(SendRequestDetect(ref rekognitionClient, anh));
                
                SendRequestToGGsheet("1qV_b1Sb9byNwOrAdMCv9HhF7u4AtQFuMcVl4Ay7yDqI", "sheet1!A" + index, data.ToList<object>());
                index++;
            }
        }

        void SendRequestToGGsheet(string spreadsheetId, string range, List<object> oblist)
        {
            //List<object> oblist = new List<object> { "1979-2007" };
            List<IList<object>> q = new List<IList<object>> { oblist };
            QHelper qHelper = new QHelper();
            qHelper.CreateService();
            qHelper.UpdateEntry(spreadsheetId, range, q);
        }

        string SendRequestDetect(ref AmazonRekognitionClient rekognitionClient, Bitmap b)
        {
            string res = "";
            DetectTextRequest detectText = new DetectTextRequest()
            {
                Image = new Amazon.Rekognition.Model.Image()
                {
                    Bytes = new MemoryStream(ImageToBytes(b))
                }
            };
            //DetectTextRequest detectTextRequest = new DetectTextRequest()
            //{
            //    Image = new Amazon.Rekognition.Model.Image()
            //    {
            //        S3Object = new S3Object()
            //        {
            //            Name = photo,
            //            Bucket = bucket
            //        }
            //    }
            //};

            try
            {
                DetectTextResponse detectTextResponse = rekognitionClient.DetectText(detectText);
                //TextDetection text = detectTextResponse.TextDetections[0];
                
                foreach (TextDetection text in detectTextResponse.TextDetections)
                {
                    if (text.Type == TextTypes.LINE)
                    {
                        res += text.DetectedText + " ";
                    }
                    //Console.WriteLine("Detected: " + text.DetectedText);
                    //Console.WriteLine("Confidence: " + text.Confidence);
                    //Console.WriteLine("Id : " + text.Id);
                    //Console.WriteLine("Parent Id: " + text.ParentId);
                    //Console.WriteLine("Type: " + text.Type);

                    //Console.WriteLine("------------------------------------");
                }
                return res;
            }
            catch
            {
                return "";
            }
        }

        List<string> SplitStr(string pattern)
        {
            bool validA = pattern.All(c => Char.IsDigit(c) || c.Equals('-') || c.Equals(' '));
            if (validA == true)
            {
                pattern = pattern.Trim();
                return new List<string>() { pattern };
            }
            Remove2WhiteSpaceConstant(ref pattern);
            List<string> listStr = new List<string>();
            string percent = "";
            int count = 0;
            foreach (char c in pattern)
            {
                if (!Char.IsLetter(c) && !Char.IsWhiteSpace(c))
                {
                    listStr.Add(pattern.Substring(0, count).TrimStart().TrimEnd());
                    percent = pattern.Substring(count);
                    break;
                }
                count++;
            }
            
            foreach (string s in percent.Split(' '))
            {
                listStr.Add(s);
            }

            Console.WriteLine(listStr.Count);
            return listStr;
        }

        void Remove2WhiteSpaceConstant(ref string h)
        {
            string res = "";
            string[] mang = h.Split(' ');
            foreach (string s in mang)
            {
                if (s != "")
                {
                    res += s + " ";
                }
            }
            h = res.TrimEnd();
        }

        Bitmap CropImage(Bitmap img, Rectangle cropRect)
        {
            Bitmap bitmap = new Bitmap(cropRect.Width, cropRect.Height);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.DrawImage(img, new Rectangle(0, 0, bitmap.Width, bitmap.Height), cropRect, GraphicsUnit.Pixel);
            }
            return bitmap;
        }

        Bitmap CropImage(Bitmap img, int x1, int y1, int x2, int y2)
        {
            int width = Math.Abs(x2 - x1);
            int height = Math.Abs(y2 - y1);
            Bitmap bitmap = new Bitmap(width, height);
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.DrawImage(img, new Rectangle(0, 0, bitmap.Width, bitmap.Height), new Rectangle(x1, y1,width, height), GraphicsUnit.Pixel);
            }
            img.Dispose();
            return bitmap;
        }

        byte[] ImageToBytes(Bitmap img)
        {
            return (byte[])(new ImageConverter()).ConvertTo(img, typeof(byte[]));
        }
        byte[] ImageToBytes(string filePath)
        {
            using (FileStream fs = new FileStream(filePath, FileMode.Open))
            {
                using (BinaryReader br = new BinaryReader(fs))
                {
                    return br.ReadBytes((Int32)fs.Length);
                }
            } 
        }
        void WriteProfile(string profileName, string keyId, string secret)
        {
            Console.WriteLine($"Create the [{profileName}] profile...");
            var options = new CredentialProfileOptions
            {
                AccessKey = keyId,
                SecretKey = secret
            };
            var profile = new CredentialProfile(profileName, options);
            var sharedFile = new SharedCredentialsFile();
            sharedFile.RegisterProfile(profile);
        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            using (OpenFileDialog fileDialog = new OpenFileDialog())
            {
                if (fileDialog.ShowDialog() == DialogResult.OK)
                {
                    filePath = fileDialog.FileName;
                    pictureBox1.Invoke(new MethodInvoker(() =>
                    {
                        image = (Bitmap)Bitmap.FromFile(filePath);
                        pictureBox1.Image = image;
                    }));
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Bitmap b = (Bitmap)Bitmap.FromFile(@"D:\Project\C#\DetectTextAWS\DetectTextAWS\bin\Debug\Photo\4.png");
            AmazonRekognitionClient rekognitionClient = new AmazonRekognitionClient("AKIAQM4PGWWXTMLCCSU5", "wLLqDZEActNaw1mcztP3O74s91SuNpwaHHl2Gl+h", Amazon.RegionEndpoint.EUWest2);
            SendRequestDetect(ref rekognitionClient, b);
        }
    }
}
