Dim fileObj

Set fileObj =   CreateObject("FileComAddIn.FileHelper")



fileObj.Init_File "C:\temp", "filetest.txt", false

msgbox fileObj.DataFile

fileObj.Write_Line("Test FileHelper by VB Script Edit2" & vbNewLine)
fileObj.Write_Line("Test FileHelper by VB Script Edit2 Append")
msgbox fileObj.Filename
msgbox fileObj.DataFile


