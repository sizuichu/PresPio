using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace PresPio
    {
    public class YouDao
        {
        public string YouDaos(string q, string from, string to)
            {
            string result = "";
            string url = "http://fanyi.youdao.com/translate_o?smartresult=dict&smartresult=rule/";
            string u = "fanyideskweb";
            string c = "Y2FYu%TNSbMCxc3t2u^XT";
            TimeSpan ts = (DateTime.UtcNow - new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc));
            long millis = (long)ts.TotalMilliseconds;
            string curtime = Convert.ToString(millis);
            Random rd = new Random();
            string f = curtime + rd.Next(0, 9);
            string signStr = u + q + f + c;
            string sign = GetMd5Str_32(signStr);
            Dictionary<String, String> dic = new Dictionary<String, String>();
            dic.Add("i", q);
            dic.Add("from", from);
            dic.Add("to", to);
            dic.Add("smartresult", "dict");
            dic.Add("client", "fanyideskweb");
            dic.Add("salt", f);
            dic.Add("sign", sign);
            dic.Add("lts", curtime);
            dic.Add("bv", GetMd5Str_32("5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.82 Safari/537.36"));
            dic.Add("doctype", "json");
            dic.Add("version", "2.1");
            dic.Add("keyfrom", "fanyi.web");
            dic.Add("action", "FY_BY_REALTlME");
            //dic.Add("typoResult", "false");

            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            req.Method = "POST";
            req.ContentType = "application/x-www-form-urlencoded; charset=UTF-8";
            req.Referer = "http://fanyi.youdao.com/";
            req.UserAgent = "Mozilla/5.0 (Windows NT 6.1; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/93.0.4577.82 Safari/537.36";
            req.Headers.Add("Cookie", "OUTFOX_SEARCH_USER_ID=-2030520936@111.204.187.35; OUTFOX_SEARCH_USER_ID_NCOO=798307585.9506682; UM_distinctid=17c2157768a25e-087647b7cf38e8-581e311d-1fa400-17c2157768b8ac; P_INFO=15711476666|1632647789|1|youdao_zhiyun2018|00&99|null&null&null#bej&null#10#0|&0||15711476666; JSESSIONID=aaafZvxuue5Qk5_d9fLWx; ___rl__test__cookies=" + curtime);
            StringBuilder builder = new StringBuilder();
            int i = 0;
            foreach (var item in dic)
                {
                if (i > 0)
                    builder.Append("&");
                builder.AppendFormat("{0}={1}", item.Key, item.Value);
                i++;
                }
            byte[] data = Encoding.UTF8.GetBytes(builder.ToString());
            req.ContentLength = data.Length;
            using (Stream reqStream = req.GetRequestStream())
                {
                reqStream.Write(data, 0, data.Length);
                reqStream.Close();
                }
            HttpWebResponse resp = (HttpWebResponse)req.GetResponse();
            Stream stream = resp.GetResponseStream();
            using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
                {
                JObject jo = (JObject)JsonConvert.DeserializeObject(reader.ReadToEnd());
                if (jo.Value<string>("errorCode").Equals("0"))
                    {
                    var tgtarray = jo.SelectToken("translateResult").First().Values<string>("tgt").ToArray();
                    result = string.Join("", tgtarray);
                    }
                }
            return result;
            }

        public static string GetMd5Str_32(string encryptString)
            {
            byte[] result = Encoding.UTF8.GetBytes(encryptString);
            MD5CryptoServiceProvider md5 = new MD5CryptoServiceProvider();
            byte[] output = md5.ComputeHash(result);
            string encryptResult = BitConverter.ToString(output).Replace("-", "");
            return encryptResult;
            }
        }
    }