using OpenCvSharp;

using System;
using System.Collections.Generic;
using System.Configuration;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

using Tesseract;

namespace ConsoleApp4
{
    public class Helper
    {

        private static readonly String pattern1 = "<input[\\s]*type[\\s]*=[\\s]*\"text\"[\\s]*name[\\s]*=[\\s]*\"(KEY[\\d]+)\"[\\s]*/>";
        private static readonly String pattern2 = "<input[\\s]*type[\\s]*=[\\s]*\"hidden\"[\\s]*name[\\s]*=[\\s]*\"neusoft_key\"[\\s]*value[\\s]*=[\\s]*\"(ID[\\w]+)\"[\\s]*/>";
        private static readonly String pattern3 = "<input[\\s]*type[\\s]*=[\\s]*\"text\"[\\s]*class[\\s]*=[\\s]*\"textfield\"[\\s]*name[\\s]*=[\\s]*\"(ID[\\w|!]+)\"";
        private static readonly String pattern4 = "<input[\\s]*type[\\s]*=[\\s]*\"password\"[\\s]*class[\\s]*=[\\s]*\"textfield\"[\\s]*name[\\s]*=[\\s]*\"(KEY[\\w|!]+)\"[\\s]*/>";
        private static readonly String pattern5 = "<input[\\s]*type[\\s]*=[\\s]*\"text\"[\\s]*class[\\s]*=[\\s]*\"a\"[\\s]*style[\\s]*=[\\s]*\"[\\s]*width:[\\s]*93px;[\\s]*height:[\\s]*22px;[\\s]*vertical[\\s]*-[\\s]*align:middle;[\\s]*border:[\\s]*1px[\\s]*solid[\\s]*#707070;\"[\\s]*name[\\s]*=[\\s]*\"([\\w|!]+)\"[\\s]*/>";

        public static async Task PostCard()
        {
            Int32 retry = 3;
            while (retry > 0)
            {
                try
                {
                    using (var handler = new HttpClientHandler())
                    {
                        handler.UseCookies = true;
                        var cc = new CookieContainer();
                        handler.CookieContainer = cc;
                        using (HttpClient client = new HttpClient(handler))
                        {
                            client.BaseAddress = new Uri("http://kq.neusoft.com");
                            client.DefaultRequestHeaders.Add("Accept-Encoding", "gzip, deflate");
                            client.DefaultRequestHeaders.Add("Accept-Language", "zh-CN,zh;q=0.8,zh-TW;q=0.7,zh-HK;q=0.5,en-US;q=0.3,en;q=0.2");
                            client.DefaultRequestHeaders.Add("Connection", "keep-alive");
                            client.DefaultRequestHeaders.Add("Host", "kq.neusoft.com");
                            client.DefaultRequestHeaders.Add("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:68.0) Gecko/20100101 Firefox/68.0");
                            Match match1, match2, match3, match4, match5;
                            using (var message = new HttpRequestMessage(HttpMethod.Get, "index.jsp"))
                            {
                                message.Headers.Add("Upgrade-Insecure-Requests", @"1");
                                message.Headers.Add("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8");
                                var response = await client.SendAsync(message);
                                response.EnsureSuccessStatusCode();
                                string responseBody = await response.Content.ReadAsStringAsync();
                                var cookie = response.Headers.GetValues("Set-Cookie");
                                var temp = cookie.ToList()[0].Split(';')[0];
                                client.DefaultRequestHeaders.Add("Cookie", temp);
                                client.DefaultRequestHeaders.Add("Cookie", temp);
                                match1 = Regex.Match(responseBody, pattern1);
                                match2 = Regex.Match(responseBody, pattern2);
                                match3 = Regex.Match(responseBody, pattern3);
                                match4 = Regex.Match(responseBody, pattern4);
                                match5 = Regex.Match(responseBody, pattern5);
                            }
                            using (var codeMessage = new HttpRequestMessage(HttpMethod.Get, "imageRandeCode"))
                            {
                                codeMessage.Headers.Add("Accept", "image/webp,*/*");
                                codeMessage.Headers.Add("Referer", "http://kq.neusoft.com/index.jsp");
                                var image = await client.SendAsync(codeMessage);
                                image.EnsureSuccessStatusCode();
                                var reader = await image.Content.ReadAsStreamAsync();
                                using (FileStream write = new FileStream("imageRandeCode" + ".jpg", FileMode.OpenOrCreate, FileAccess.ReadWrite))
                                {
                                    Byte[] buffer = new Byte[4096];
                                    var read = reader.Read(buffer, 0, buffer.Length);
                                    while (read != 0)
                                    {
                                        write.Write(buffer, 0, buffer.Length);
                                        read = reader.Read(buffer, 0, buffer.Length);
                                    }
                                }
                            }
                            String code;
                            using (Mat src = new Mat("imageRandeCode.jpg", ImreadModes.AnyColor | ImreadModes.AnyDepth))
                            {
                                var dst = new Mat();
                                Cv2.CvtColor(src, dst, ColorConversionCodes.RGB2GRAY);
                                using (var threshold = dst.Threshold(200, 255, ThresholdTypes.BinaryInv | ThresholdTypes.Otsu))
                                {
                                    threshold.SaveImage("imageRandeCode.jpg");
                                    using (var engine = new TesseractEngine("tessdata", "num", EngineMode.TesseractOnly))
                                    {
                                        engine.SetVariable("tessedit_char_whitelist", "0123456789");
                                        using (var img = Pix.LoadFromFile("imageRandeCode.jpg"))
                                        {
                                            using (var page = engine.Process(img))
                                            {
                                                code = page.GetText();
                                            }
                                        }
                                    }
                                }
                            }
                            ConfigurationManager.RefreshSection("appSettings");
                            var name = ConfigurationManager.AppSettings["name"];
                            var password = ConfigurationManager.AppSettings["password"];
                            List<KeyValuePair<String, String>> nameValueCollection = new List<KeyValuePair<String, String>>();
                            nameValueCollection.Add(new KeyValuePair<String, String>("login", "true"));
                            nameValueCollection.Add(new KeyValuePair<String, String>("neusoft_attendance_online", ""));
                            nameValueCollection.Add(new KeyValuePair<String, String>(match1.Groups[1].ToString(), ""));
                            nameValueCollection.Add(new KeyValuePair<String, String>("neusoft_key", match2.Groups[1].ToString()));
                            nameValueCollection.Add(new KeyValuePair<String, String>(match3.Groups[1].ToString(), name));
                            nameValueCollection.Add(new KeyValuePair<String, String>(match4.Groups[1].ToString(), password));
                            nameValueCollection.Add(new KeyValuePair<String, String>(match5.Groups[1].ToString(), code.Trim()));
                            using (FormUrlEncodedContent content = new FormUrlEncodedContent(nameValueCollection))
                            {
                                using (var logMessage = new HttpRequestMessage(HttpMethod.Post, "login.jsp"))
                                {
                                    logMessage.Headers.Add("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8");
                                    logMessage.Headers.Add("Upgrade-Insecure-Requests", @"1");
                                    logMessage.Content = content;
                                    var loginResponse = await client.SendAsync(logMessage);
                                    loginResponse.EnsureSuccessStatusCode();
                                    var loginBodyMessage = await loginResponse.Content.ReadAsStringAsync();
                                    String loginPattern = "<input[\\s]*type[\\s]*=[\\s]*\"hidden\"[\\s]*name[\\s]*=[\\s]*\"currentempoid\"[\\s]*value[\\s]*=[\\s]*\"([\\d]+)\">";
                                    var loginMatch = Regex.Match(loginBodyMessage, loginPattern);
                                    nameValueCollection.Clear();
                                    nameValueCollection.Add(new KeyValuePair<string, string>("currentempoid", loginMatch.Groups[1].ToString()));
                                    nameValueCollection.Add(new KeyValuePair<string, string>("browser", "firefox"));
                                    using (FormUrlEncodedContent aaaaContent = new FormUrlEncodedContent(nameValueCollection))
                                    {
                                        using (var postCardMessage = new HttpRequestMessage(HttpMethod.Post, "record.jsp"))
                                        {
                                            postCardMessage.Headers.Add("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8");
                                            postCardMessage.Headers.Add("Upgrade-Insecure-Requests", @"1");
                                            postCardMessage.Headers.Add("Referer", @"http://kq.neusoft.com/attendance.jsp");
                                            postCardMessage.Content = aaaaContent;
                                            var aaaMessage = await client.SendAsync(postCardMessage);
                                            aaaMessage.EnsureSuccessStatusCode();
                                            var aaaReposone = await aaaMessage.Content.ReadAsStringAsync();
                                            retry = 0;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                catch (Exception)
                {
                    retry--;
                    Thread.Sleep(60 * 1000 * 1);
                }
            }
        }

        public static TimeSpan GetTime()
        {
            var seed = new Random(DateTime.Now.Millisecond);
            var result = new TimeSpan(0, seed.Next(25), seed.Next(59));
            return result;
        }

        public static Boolean Sleep(ref Boolean doit)
        {
            Boolean result = false;
            var dateTime = DateTime.Now;
            TimeSpan sleep = new TimeSpan(0, 1, 0);
            if (doit == true)
            {
                sleep = new TimeSpan(1, 0, 0);
                doit = false;
            }
            else if ((dateTime.DayOfWeek != DayOfWeek.Saturday) && 
                     (dateTime.DayOfWeek != DayOfWeek.Sunday))
            {
                if ((dateTime.Hour == 8) ||
                    (dateTime.Hour == 18))
                {
                    sleep = GetTime();
                    result = true;
                }
            }
            Thread.Sleep(sleep);
            return result;
        }

        public void Test(Object sender, EventArgs arg)
        {
            Console.WriteLine("1111");
            Console.Read();
        }
    }
}
