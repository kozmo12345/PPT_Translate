using Aspose.Slides;
using Aspose.Slides.Util;
using Microsoft.Win32;
using Realweb;
using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using realSystemEvent = Realweb.Translate;
using Json;
using System.Threading;

namespace CsConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            const string licFilePath = @"../../Aspose.Slides.lic";
            var license = new Aspose.Slides.License();
            license.SetLicense(licFilePath);
            var filePath = @"../../doc.pptx";

            Presentation presentation = new Presentation(filePath);

            ISlide slide = presentation.Slides[0];

            var textFrames = SlideUtil.GetAllTextFrames(presentation, true);

            string query = "";
            var papago = new Translate();

            string sourceLang = "ja";
            string targetLang = "ko";

            List<string> locales = new List<string>();
            locales.Add("ja");
            locales.Add("zh-cn");
            locales.Add("zh-tw");

            for (int i = 0; i < textFrames.Length; i++)
            {
                var textFrame = textFrames[i];
                string completeRate = i.ToString() + "/" + textFrames.Length.ToString() + "...";
                Console.WriteLine(completeRate);

                if (i == 100) break;

                foreach (var para in textFrame.Paragraphs)
                {
                    foreach (var port in para.Portions)
                    {
                        string res = string.Empty;
                        query = port.Text;
                        var nSec = DateTime.Now.Second;
                        res = papago.detect(query, locales, nSec) ? papago.translate(query, sourceLang, targetLang, nSec) : query;
                        IPortion np = new Portion();
                        np.Text = res;
                        port.Text = res;
                    }
                }

            }

            presentation.Save("../../translate.pptx", Aspose.Slides.Export.SaveFormat.Pptx);

        }

    }
}

namespace Realweb
{
    public class Translate
    {
        public int second { get; set; }

        private int ti = 0;
        private int di = 0;

        public string translate(string query, string sourceLang, string targetLang, int nSec)
        {
            string text = query;

            if (text.Trim().Length > 1)
            {
                try
                {
                    //if (nSec != second)
                    //{
                    //    ti = 0;
                    //}
                    //else
                    //{
                    //    ti++;
                    //}

                    //if (ti >= 6)
                    //{
                    //    Thread.Sleep(1000);
                    //}

                    string url = "https://openapi.naver.com/v1/papago/n2mt"; //ja => en 만 가능
                    //string url = "https://openapi.naver.com/v1/language/translate";

                    if (url.Contains("n2mt") && targetLang == "ko")
                    {
                        targetLang = "en";
                    }

                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                    request.Headers.Add("X-Naver-Client-Id", "id"); //use u are Id
                    request.Headers.Add("X-Naver-Client-Secret", "password"); //use u are password
                    request.Method = "POST";
                    byte[] byteDataParams = Encoding.UTF8.GetBytes("source=" + sourceLang + "&target=" + targetLang + "&text=" + query);
                    request.ContentType = "application/x-www-form-urlencoded";
                    request.ContentLength = byteDataParams.Length;
                    Stream st = request.GetRequestStream();
                    st.Write(byteDataParams, 0, byteDataParams.Length);
                    st.Close();
                    HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                    Stream stream = response.GetResponseStream();
                    StreamReader reader = new StreamReader(stream, Encoding.UTF8);
                    text = reader.ReadToEnd();

                    stream.Close();
                    response.Close();
                    reader.Close();

                    IDictionary<string, object> res = JsonParser.FromJson(text);

                    if (res.Keys.Contains("message"))
                    {
                        var message = (Dictionary<string, object>)res["message"];
                        var result = (Dictionary<string, object>)message["result"];
                        text = result["translatedText"].ToString();
                    }

                    second = nSec;

                    Console.WriteLine(text);

                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    throw ex;
                }

            }

            return text;
        }

        public bool detect(string query, List<string> locales, int nSec)
        {

            if (query.Trim().Length > 1)
            {
                try
                {
                    //if (nSec != second)
                    //{
                    //    di = 0;
                    //}
                    //else
                    //{
                    //    di++;
                    //}

                    //if (di >= 6)
                    //{
                    //    Thread.Sleep(1000);
                    //}

                    string text = "";
                    string url = "https://openapi.naver.com/v1/papago/detectLangs";
                    HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                    request.Headers.Add("X-Naver-Client-Id", "id"); //use u are Id
                    request.Headers.Add("X-Naver-Client-Secret", "password"); //use u are password
                    request.Method = "POST";
                    byte[] byteDataParams = Encoding.UTF8.GetBytes("query=" + query);
                    request.ContentType = "application/x-www-form-urlencoded";
                    request.ContentLength = byteDataParams.Length;
                    Stream st = request.GetRequestStream();
                    st.Write(byteDataParams, 0, byteDataParams.Length);
                    st.Close();
                    HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                    Stream stream = response.GetResponseStream();
                    StreamReader reader = new StreamReader(stream, Encoding.UTF8);
                    text = reader.ReadToEnd();
                    stream.Close();
                    response.Close();
                    reader.Close();
                    stream.Close();
                    response.Close();
                    reader.Close();

                    IDictionary<string, object> res = JsonParser.FromJson(text);

                    if (res.Keys.Contains("langCode"))
                    {
                        if(locales.Contains(res["langCode"].ToString()))
                        {
                            return true;
                        }
                    }

                    second = nSec;

                    Console.WriteLine(text);

                    
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                    throw ex;
                }
            }
            return false;
        }
    }

}