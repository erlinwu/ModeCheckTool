using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;

using Log4NetHelper;
using Newtonsoft.Json.Linq;

namespace HttpRequestHelper
{
    public class HttpRequestEx
    {
        string sReqURL;
        string sReqSendData;

        //定义委托  
        public delegate void Asyncdelegate(WebProxy objName);

        //异步调用完成时，执行回调方法  
        private void CallbackMethod(IAsyncResult ar)
        {
            Asyncdelegate dlgt = (Asyncdelegate)ar.AsyncState;
            dlgt.EndInvoke(ar);
        }

        //异步调用 提交的方法  
        public  virtual  void HttpRequest_AsyncCall(string ReqURL,string ReqSendData)
        {
            if (sReqURL == "")//没有访问地址无法请求http服务
            {
                throw new ArgumentNullException("http request url is empty ");

            }
            sReqURL = ReqURL;
            sReqSendData = ReqSendData;
            Asyncdelegate isgt = new Asyncdelegate(Commit);
            IAsyncResult ar = isgt.BeginInvoke(null, new AsyncCallback(CallbackMethod), isgt);
        }

        //向APM接口提交数据  
        //为什么要用WebProxy，因为.Net 4.0以下没有Host属性，无法设置标头来做DNS重连  
        public virtual void Commit(WebProxy objName = null)
        {
            
            string ret = string.Empty;
            string ip = string.Empty;
            try
            {
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(sReqURL);
                request.Method = "POST";//  
                request.Timeout = 5000;
                request.ContentType = "application/x-www-form-urlencoded";
                request.ServicePoint.Expect100Continue = false;
                if (objName != null)
                {
                    request.Proxy = objName;
                }
                byte[] byteArray = Encoding.UTF8.GetBytes(sReqSendData);
                request.ContentLength = byteArray.Length;
                Stream dataStream = request.GetRequestStream();
                dataStream.Write(byteArray, 0, byteArray.Length);
                dataStream.Close();

                string resp;
                using (HttpWebResponse response = request.GetResponse() as HttpWebResponse)
                {
                    StreamReader reader = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
                    resp = reader.ReadToEnd();
                    reader.Close();
                }
                LogHelper.WriteLog(typeof(HttpRequestEx), "调用Http服务("+ sReqURL + ")返回结果：" + resp);
            }
            catch (WebException ex)
            {
                LogHelper.WriteLog(typeof(HttpRequestEx), "调用Http服务(" + sReqURL + ")失败：" + ex.Status.ToString());
               
            }

        }

        //同步方法调用http服务
        public string HttpRequest_Call(string urlstr, string json_ary)
        {

            try
            {
                if (urlstr == "")//没有访问地址无法请求http服务
                {
                    throw new ArgumentNullException("http request url is empty ");

                }
                string send_jsonstr = json_ary;
                HttpWebRequest request = WebRequest.Create(urlstr) as HttpWebRequest;//
                request.Timeout = 5111;//超时
                request.Method = "post";//
                request.KeepAlive = true;
                request.AllowAutoRedirect = false;
                request.ContentType = "application/x-www-form-urlencoded;charset=utf-8";
                //
                byte[] postdatabtyes = Encoding.UTF8.GetBytes(send_jsonstr);
                request.ContentLength = postdatabtyes.Length;
                Stream requeststream = request.GetRequestStream();
                requeststream.Write(postdatabtyes, 0, postdatabtyes.Length);
                requeststream.Close();
                //
                string resp;
                using (HttpWebResponse response = request.GetResponse() as HttpWebResponse)
                {
                    StreamReader sr = new StreamReader(response.GetResponseStream(), Encoding.UTF8);
                    resp = sr.ReadToEnd();
                }
                return resp;
            }
            catch (Exception ex)
            {
                LogHelper.WriteLog(typeof(HttpRequestEx), "调用HttpWeb服务(" + urlstr + ")出错:" + ex.Message);
                //return "";
                throw new Exception("调用HttpWeb服务(" + urlstr + ")出错:" + ex.Message);
            }

        }

    }
    
}
