using System;
using System.IO;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace bruteCSharp
{
    class Excel
    {
        string path = "";
        _Application excel = new _Excel.Application();
        Workbook wb;
        bool open = false;
        int total = 0;
        string beg;

        public Excel(string path, int Sheet, string beg)
        {
            this.path = path;
            this.beg = beg;
            //wb = excel.Workbooks.Open(path, ReadOnly: true, Password: b);
            //wb = excel.Workbooks.Open(path);
            //ws = (Worksheet)wb.Worksheets[Sheet];
        }

        public void tryFile(FileStream f)
        {
            StreamReader stream = new StreamReader(f);
            string s = stream.ReadLine();
            bool open = false;
            int i = 0;
            while(s != null && !open)
            {
                try
                {
                    wb = excel.Workbooks.Open(path, ReadOnly: true, Password: beg + s);
                    Console.WriteLine("password: " + beg + s);
                    //close();
                    wb.Close();
                    open = true;
                }
                catch (Exception e)
                {
                    
                    //Console.WriteLine(b);
                }
                //Console.WriteLine(s);
                s = stream.ReadLine();
                
            }
            close();
        }

        public void setBeg(string b)
        {
            this.beg = b;
        }

        public void close()
        {            
            excel.Quit();
            
            System.Runtime.InteropServices.Marshal.ReleaseComObject(excel);
            excel = null;
        }
    }
}
