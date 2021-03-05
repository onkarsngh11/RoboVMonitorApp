using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;

namespace RoboVMonitorApp
{
    class CreateJSON
    {
        string tabName = "CreateJSON";
        public string logPath = Environment.CurrentDirectory + "\\ExceptionHandling.log";
        #region GetJSON
        public string GetJson(string serverName, string portNumber, string indexName, string nameOfField, string valueOfField, int size)
        {

            try
            {
                string url = "http://" + serverName + ":" + portNumber + "/" + indexName + "/_search?size=" + Convert.ToString(size)+"&q="+valueOfField;//size=" + Convert.ToString(size);
                var JavaScriptSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();
                //string data = JavaScriptSerializer.Serialize(url);
                string data = string.Empty; //GeneratePostJSON(nameOfField, valueOfField);
                string ResponseJSON = PostJSON(url, data);
                return ResponseJSON;
            }
            catch (Exception ex)
            {
                return string.Empty;
            }

        }

        //private string GeneratePostJSON(string nameOfField, string valueOfField)
        //{

        //    try
        //    {
        //        string data = string.Empty;
        //        string[] value = valueOfField.Split(',');
        //        Issue issue = new Issue();
        //        foreach (var item in value)
        //        {
        //            issue.query= ;
        //            var JavaScriptSerializer = new System.Web.Script.Serialization.JavaScriptSerializer();

        //            if (string.IsNullOrEmpty(data))
        //                data = JavaScriptSerializer.Serialize(issue);
        //            else
        //                data = JavaScriptSerializer.Serialize(issue) + data;
        //        }
        //        return data;
        //    }
        //    catch (Exception ex)
        //    {
        //        Logger.LogFileWrite(tabName + "_" + ex.Message, logPath);
        //        throw;
        //    }

        //}

        private string PostJSON(string url, string json)
        {

            try
            {
                var httpWebRequest = (HttpWebRequest)WebRequest.Create(url);
                httpWebRequest.ContentType = "application/json";
                httpWebRequest.Method = "GET";
                string result = string.Empty;
                //using (var streamWriter = new StreamWriter(httpWebRequest.GetRequestStream()))
                //{

                //    streamWriter.Write(json);
                //}

                var httpResponse = (HttpWebResponse)httpWebRequest.GetResponse();
                using (var streamReader = new StreamReader(httpResponse.GetResponseStream()))
                {
                    result = streamReader.ReadToEnd();
                }
                return result;
            }
            catch (Exception ex)
            {
                return string.Empty;
            }

        }
        #endregion
    }

    public class Issue
    {
        public Query query { get; set; }
        //public Doc doc { get; set; }
        public Issue()
        {
            query = new Query();
        }
        //public Issue(string update)
        //{
        //    doc = new Doc();
        //}
    }

    public class Query
    {
        public Wildcard wildcard { get; set; }
        public Query()
        {
            wildcard = new Wildcard();
        }

    }

    public class Wildcard
    {
        public string Client { get; set; }
    }
}
