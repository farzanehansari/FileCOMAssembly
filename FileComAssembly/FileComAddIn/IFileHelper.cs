namespace FileComAddIn
{
    public interface IFileHelper
    {
        string Filename { get; set; }
        string Pathdir { get; set; }
        string FullPathFile { get; }
        string DataFile { get; }

        void Init_File(string filedir, string filename, bool overwrite = true);
        void Write_Line(string text);
        void Write_Line(string[] lines);
        void Delete();
        void Delete_Lines_From_Start(int num);
        void Delete_Lines_From_End(int num);
        void Delete_Lines_Between(int from, int to);
        void Delete_Lines_StartWith_Text(string text);
        void Delete_Line(int num);
        void Replace_Line_WithText(int num, string text);
    }
}