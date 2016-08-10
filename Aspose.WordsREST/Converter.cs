using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;

namespace Aspose.WordsREST
{
    public class Converter

    {
        private string fileformate = string.Empty;
        private string file;

        public string File
        {
            get { return file; }
            set { file = value; }
        }

        public string Fileformate
        {
            get { return fileformate; }
            set { fileformate = value; }
        }
    }
}