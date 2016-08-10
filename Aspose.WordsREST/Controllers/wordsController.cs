using Aspose.Words;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Runtime.Serialization.Formatters.Binary;
using System.Web.Http;
using System.Web.Script.Serialization;

namespace Aspose.WordsREST.Controllers
{
    public class wordsController : ApiController
    {
        // GET api/values
        public IEnumerable<string> Get()
        {
            return new string[] { "value1", "value2" };
        }

        // GET api/values/5
        public AsposeRest Get(int id)
        {
            return new AsposeRest();
        }
        public string Get(int id,int e)
        {
            return "value";
        }
        // POST api/values
        public string Post(HttpRequestMessage request)
        {
            MemoryStream ms = new MemoryStream();
            try
            {
                var content = request.Content;
                string resultedJosn = content.ReadAsStringAsync().Result;
                JavaScriptSerializer js = new JavaScriptSerializer();
                js.MaxJsonLength = Int32.MaxValue;
                if (!String.IsNullOrEmpty(resultedJosn))
                {
                    dynamic receivedData = js.Deserialize<object>(resultedJosn);
                    string fileformat = receivedData["Fileformate"];
                    string file = receivedData["File"];
                    byte[] bytUpfile = Convert.FromBase64String(file);
                    Stream stream = new MemoryStream(bytUpfile);

                    MemoryStream memStream = new MemoryStream();
                    BinaryFormatter binForm = new BinaryFormatter();

                    memStream.Write(bytUpfile, 0, bytUpfile.Length);
                    memStream.Seek(0, SeekOrigin.Begin);

                    Document doc = new Document(memStream);
                    //doc.OriginalLoadFormat = LoadFormat.Docx;

                    doc.Save(ms, getFormate(fileformat));
                    
                }
                
            }
            catch (Exception e)
            { 
            }
            byte[] ar = ms.ToArray();
            return Convert.ToBase64String(ar);
           
        }

        // PUT api/values/5
        public void Put(int id, [FromBody]string value)
        {
        }

        // DELETE api/values/5
        public void Delete(int id)
        {
        }

        private SaveFormat getFormate(string formate)
        {
            SaveFormat s = SaveFormat.Doc;
            if (string.Equals("html", formate.Trim().ToLower()))
            {
                s = SaveFormat.Html;
            }
            else if (string.Equals("bmp", formate.Trim().ToLower()))
            {
                s = SaveFormat.Bmp;
            }
            else if (string.Equals("doc", formate.Trim().ToLower()))
            {
                s = SaveFormat.Doc;
            }
            else if (string.Equals("docm", formate.Trim().ToLower()))
            {
                s = SaveFormat.Docm;
            }
            else if (string.Equals("docx", formate.Trim().ToLower()))
            {
                s = SaveFormat.Docx;
            }
            else if (string.Equals("dot", formate.Trim().ToLower()))
            {
                s = SaveFormat.Dot;
            }
            else if (string.Equals("dotm", formate.Trim().ToLower()))
            {
                s = SaveFormat.Dotm;
            }
            else if (string.Equals("dotx", formate.Trim().ToLower()))
            {
                s = SaveFormat.Dotx;
            }
            else if (string.Equals("emf", formate.Trim().ToLower()))
            {
                s = SaveFormat.Emf;
            }
            else if (string.Equals("epub", formate.Trim().ToLower()))
            {
                s = SaveFormat.Epub;
            }
            else if (string.Equals("flatopc", formate.Trim().ToLower()))
            {
                s = SaveFormat.FlatOpc;
            }
            else if (string.Equals("jpeg", formate.Trim().ToLower()))
            {
                s = SaveFormat.Jpeg;
            }
            else if (string.Equals("pdf", formate.Trim().ToLower()))
            {
                s = SaveFormat.Pdf;
            }
            else
            {
                s = SaveFormat.Docx;
            }


            return s;
        }


    }
}
