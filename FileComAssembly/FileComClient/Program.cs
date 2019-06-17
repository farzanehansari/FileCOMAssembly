using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FileComAddIn;

namespace FileComClient
{
    class Program
    {
        static void Main(string[] args)
        {
            IFileHelper fileObj = new FileHelper(@"C:\temp", "testfile.txt");

            fileObj.Write_Line("File is Created and written by Farzaneh");

            Console.WriteLine(fileObj.DataFile);

            Console.ReadKey();
        }
    }
}
