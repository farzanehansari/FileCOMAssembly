using NLog;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace FileComAddIn
{
    [ComVisible(true)]
    [Guid("1F10BBFF-50B9-45B1-94CE-8B94C44A749C")]
    public class FileHelper : IFileHelper
    {
        private  Logger logger = LogManager.GetCurrentClassLogger();

        /// <summary>
        /// get,set Filename
        /// </summary>
        public string Filename { get; set; }

        /// <summary>
        ///  get,set Pathdir
        /// </summary>
        public string Pathdir { get; set; }

        /// <summary>
        /// full path dir of file : filedirectory + filename
        /// </summary>
        public string FullPathFile =>  $@"{Pathdir}\{Filename}";

        public string DataFile => ReadAllTextFile();

        public FileHelper()
        {
            try
            {               
                logger.Info("File Helper Load ...");

            }
            catch (Exception e)
            {
                logger.Error(e);
            }
        }
       
        public FileHelper(string filedir, string filename, bool overwrite = true)
        {
           
            try
            {
                logger.Info("File Helper Load ...");

                Pathdir = filedir;
                Filename = filename;
                if (!File.Exists(FullPathFile) || overwrite)
                {
                    using (StreamWriter sw = File.CreateText(FullPathFile))
                    {
                        logger.Info("File \"{0}\" is Created", filename);
                    }
                }
                else
                {
                    logger.Info("File \"{0}\" already exists.", filename);
                }

            }
            catch (Exception e)
            {
                logger.Error(e);
            }

        }

        /// <summary>
        /// Create File
        /// </summary>
        /// <param name="filedir"></param>
        /// <param name="filename"></param>
        /// <param name="overwrite">Optinal true</param>
        public void Init_File(string filedir, string filename, bool overwrite = true)
        {

            Filename = filename;
            Pathdir = filedir;

            if (!File.Exists(FullPathFile) || overwrite)
            {
                using (StreamWriter sw = File.CreateText(FullPathFile))
                {
                    logger.Info("File \"{0}\" is Created", filename);
                }
            }
            else
            {
                logger.Info("File \"{0}\" already exists.", filename);
            }
        }

        private string ReadAllTextFile()
        {
            if (File.Exists(FullPathFile))
            {
                return File.ReadAllText(FullPathFile);
            }
            return string.Empty;
        }

        /// <summary>
        /// Creates a new file,writes the specified string to file,and then closes the file.if the target file already exist,it is overwritten        
        /// </summary>
        /// <param name="text"></param>
        public void Write_Line(string text)
        {
            File.AppendAllText(FullPathFile, text);
        }

        /// <summary>
        /// Creates a new file,writes the specified string array to file,and then closes the file.
        /// </summary>
        /// <param name="lines"></param>
        public void Write_Line(string[] lines)
        {
            File.WriteAllLines(FullPathFile, lines);
        }

        /// <summary>
        ///  Delete File
        /// </summary>
        public void Delete()
        {
            if (File.Exists(FullPathFile))
            {
                File.Delete(FullPathFile);
            }
        }

        /// <summary>
        /// Delete line with specified line number
        /// </summary>
        /// <param name="num"></param>
        public void Delete_Line(int num)
        {
            if (File.Exists(FullPathFile))
            {
                string[] lines = File.ReadAllLines(FullPathFile);

                if (lines.Count() > 0 && num > 0)
                {
                    var newlines = lines.Where((line, index) => index != num - 1).ToArray();
                    Write_Line(newlines);
                }
            }
        }

        /// <summary>
        /// Delete lines between from - to sepcified lines
        /// </summary>
        /// <param name="from"></param>
        /// <param name="to"></param>
        public void Delete_Lines_Between(int from, int to)
        {
            if (File.Exists(FullPathFile))
            {
                string[] lines = File.ReadAllLines(FullPathFile);
                int linesCount = lines.Count();

                if (linesCount > 0 && (to - from) >= 0)
                {
                    var newlines = lines.Where((line, index) => index < (from - 1) | index > (to - 1)).ToArray();
                    Write_Line(newlines);
                }
            }
        }

        /// <summary>
        /// Delete n lines from end of file
        /// </summary>
        /// <param name="num"></param>
        public void Delete_Lines_From_End(int num)
        {
            if (File.Exists(FullPathFile))
            {
                string[] lines = File.ReadAllLines(FullPathFile);
                int linesCount = lines.Count();
                num = linesCount - num;
                if (linesCount > 0 && num >= 0 && num <= linesCount)
                {
                    var newlines = lines.Take(num).ToArray();
                    Write_Line(newlines);
                }
            }
        }

        /// <summary>
        /// Delete n lines from first of file
        /// </summary>
        /// <param name="num"></param>
        public void Delete_Lines_From_Start(int num)
        {
            if (File.Exists(FullPathFile))
            {
                string[] lines = File.ReadAllLines(FullPathFile);
                int linesCount = lines.Count();
                if (linesCount > 0 && num <= linesCount)
                {
                    var newlines = lines.Skip(num).ToArray();
                    Write_Line(newlines);
                }
            }
        }

        /// <summary>
        ///  Delete lines start with specified text
        /// </summary>
        /// <param name="text"></param>
        public void Delete_Lines_StartWith_Text(string text)
        {
            if (File.Exists(FullPathFile))
            {
                string[] lines = File.ReadAllLines(FullPathFile);

                if (lines.Count() > 0)
                {
                    var newlines = lines.Where(l => !(l.Trim().ToUpper().StartsWith(text.ToUpper()))).ToArray();
                    Write_Line(newlines);
                }
            }
        }

        /// <summary>
        ///  Replace Specified line text with new text
        /// </summary>
        /// <param name="num"></param>
        /// <param name="text"></param>
        public void Replace_Line_WithText(int num, string text)
        {
            if (File.Exists(FullPathFile))
            {
                string[] lines = File.ReadAllLines(FullPathFile);

                if (lines.Count() > 0 && num > 0)
                {
                    var newlines = lines.Select((line, index) => { if (index == num - 1) line = line.Replace(line, text); return line; }).ToArray();
                    Write_Line(newlines);
                }
            }
        }

      
    }
}
