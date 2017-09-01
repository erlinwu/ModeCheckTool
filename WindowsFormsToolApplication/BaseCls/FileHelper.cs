using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace FileHelper
{
    class FileHelper
    {
    
    }
    public class DirectoryAllFiles
    {
        static List<FileInformation> FileList = new List<FileInformation>();
        //遍历文件夹下所有文件（含子目录）的文件信息 递归方式
        public static List<FileInformation> GetAllFiles(DirectoryInfo dir, string filetype)
        { 
            FileInfo[] allFile = dir.GetFiles();
            foreach (FileInfo fi in allFile)
            {
                if (string.IsNullOrEmpty(filetype)|| filetype==fi.Extension)//不需要匹配类型 或者 类型和filetype参数一致
                {
                    FileList.Add(new FileInformation { FileName = fi.Name, FilePath = fi.FullName });
                }

                
            }
            DirectoryInfo[] allDir = dir.GetDirectories();
            foreach (DirectoryInfo d in allDir)
            {
                GetAllFiles(d,filetype);
            }
            return FileList;
        }

        //根据指定文件类型 遍历给定文件夹路径下的文件，获取文件名列表(不包含子文件夹）
        private List<string> GetFilesNameListByType(string filepath, string filetype)
        {

            try
            {
                if (filetype == null || filetype == "") filetype = "*";//返回全部

                List<string> ret_filenames = new List<string>();
                DirectoryInfo mydir = new DirectoryInfo(filepath);
                foreach (FileSystemInfo fsi in mydir.GetFileSystemInfos())
                {
                    if (fsi is FileInfo)
                    {

                        FileInfo fi = (FileInfo)fsi;
                        string n = Path.GetFullPath(fi.FullName);//获取文件的全路径-->C:\JiYF\BenXH\BenXHCMS.xml
                        string m = Path.GetFileName(fi.FullName);//获取文件的名称含有后缀-->BenXHCMS.xml
                        string s = Path.GetExtension(fi.FullName);//获取路径的后缀扩展名称 ".xml .jpg 文件后缀名"
                        string t = Path.GetPathRoot(fi.FullName);//获取路径的根目录-->C:\
                        string x = Path.GetDirectoryName(fi.FullName);//获取文件所在的目录 -->C:\JiYF\BenXH
                        string y = Path.GetFileNameWithoutExtension(fi.FullName);//获取文件的名称没有后缀-->BenXHCMS

                        if (s == "." + filetype)
                        {
                            ret_filenames.Add(fi.FullName);
                        }
                    }
                }
                return ret_filenames;
            }
            catch (Exception ex)
            {
                return null;
                throw new Exception(ex.Message);
                
            }
        }
    }

    public class FileInformation
    {
        public string FileName { get; set; }
        public string FilePath { get; set; }
    }
}
