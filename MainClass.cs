using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;

namespace bruteCSharp
{
    class MainClass
    {
        public static int t = 0;
        public static int index = 0;

        public static void brute(int charLeft, char[] b, FileStream[] files, int[] split, char[] valid)
        {
            if (charLeft == 0)
            {
                t++;
                if(t > split[index])
                {
                    t = 1;
                    index++;
                    //Console.WriteLine(index);
                }
                files[index].Write(Encoding.GetEncoding("UTF-8").GetBytes(b));
                files[index].Write(Encoding.GetEncoding("UTF-8").GetBytes("\n".ToCharArray()));
            }
            else
            {
                for (int i = 0; i < valid.Length; i++)
                {
                    b[b.Length - charLeft] = valid[i];
                    brute(charLeft - 1, b, files, split, valid);
                }
            }
        }

        public static void Main(string[] args)
        {
            string filepath;
            while(true)
            {
                Console.WriteLine("Enter File Path, (ex. C:\\Users\\JZ\\Desktop\\pass.xlsx)");
                filepath = Console.ReadLine();
                if(File.Exists(filepath))
                {
                    break;
                }
             
            }
            
            Console.Write("Enter beginning string : ");
            string bef = Console.ReadLine();


            string variation = "";

            Console.WriteLine("-----Password Variation-----");
            Console.Write("Add numbers? (ex. 1,2,3 ...)      y/n : " );
            string ans = Console.ReadLine();
            while(!ans.Equals("y") && !ans.Equals("n")) {
                Console.Write("Add numbers? (ex. 1,2,3 ...)      y/n : ");
                ans = Console.ReadLine();
            }
            if(ans.Equals("y"))
            {
                variation = variation + "1234567890";
            }

            Console.Write("Add lowercase letters? (ex. a,b,c ...)      y/n : ");
            ans = Console.ReadLine();
            while (!ans.Equals("y") && !ans.Equals("n"))
            {
                Console.Write("Add lowercase letters? (ex. a,b,c ...)      y/n : ");
                ans = Console.ReadLine();
            }
            if (ans.Equals("y"))
            {
                variation = variation + "abcdefghijklmnopqrstuvwxyz";
            }

            Console.Write("Add uppercase letters? (ex. A,B,C ...)      y/n : ");
            ans = Console.ReadLine();
            while (!ans.Equals("y") && !ans.Equals("n"))
            {
                Console.Write("Add uppercase letters? (ex. A,B,C ...)      y/n : ");
                ans = Console.ReadLine();
            }
            if (ans.Equals("y"))
            {
                variation = variation + "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            }

            Console.Write("Type all special characters you want to test for : ");
            variation = variation + Console.ReadLine();

            Console.WriteLine("\\/ Checking all combinations of these characters \\/\n");
            Console.WriteLine(variation);

            

            Console.Write("Enter length of characters to try at the end (ex. passwordXXXX = 4, passwordXXX = 3) : ");
            int lentotry = int.Parse(Console.ReadLine());

            double combinations = Math.Pow(variation.Length, lentotry);
            Console.WriteLine("Combinations = " + combinations + "\n");


            Console.Write("Enter number of threads : ");
            int threadslen = int.Parse(Console.ReadLine());

            int perThread = (int)((combinations / threadslen));
            int additional = (int)((((combinations / threadslen) - perThread) * threadslen) + 0.1);
            Console.WriteLine("_______");
            int[] pieces = new int[threadslen];
            for(int i = 0; i < threadslen; i++)
            {
                if(i < additional) {
                    pieces[i] = 1 + perThread;
                } else
                {
                    pieces[i] = perThread;
                }
                //Console.WriteLine(pieces[i]);
            }

            FileStream[] files = new FileStream[threadslen];
            for(int i = 0; i < threadslen; i++)
            {
                files[i] = File.Open(i + ".txt", FileMode.OpenOrCreate, FileAccess.Write);
                files[i].SetLength(0);
            }

            brute(lentotry, new char[lentotry], files, pieces, variation.ToCharArray());

            for (int i = 0; i < threadslen; i++)
            {
                files[i].Close();
            }

            Thread[] threads = new Thread[threadslen];
            for (int i = 0; i < threadslen; i++)
            {
                files[i] = File.Open(i + ".txt", FileMode.OpenOrCreate, FileAccess.Read);
            }
            var watch = new System.Diagnostics.Stopwatch();
            watch.Start();
            for (int i = 0; i < threadslen; i++)
            {
                var f = files[i];
                threads[i] = new Thread(() => new Excel(filepath, 1, bef).tryFile(f));
                threads[i].Start();
            }

            for (int i = 0; i < threadslen; i++)
            {
                threads[i].Join();
                files[i].Close();
            }
            Console.WriteLine("Complete.");

            watch.Stop();
            Console.WriteLine(watch.ElapsedMilliseconds);
            System.Diagnostics.Process[] process = System.Diagnostics.Process.GetProcessesByName("Excel");
            foreach (System.Diagnostics.Process p in process)
            {
                if (!string.IsNullOrEmpty(p.ProcessName))
                {
                    try
                    {
                        p.Kill();
                    }
                    catch { }
                }
            }
            Console.ReadLine();
        }
    }
}