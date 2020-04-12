using System;
using System.IO;
using System.IO.Compression;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Collections.Generic;

namespace AutoMail_v1
{
    class FFOperation
    {
       public string folderpath = System.IO.Directory.GetCurrentDirectory();
        public string host;
        public string from;
        public string cc;
        public string to;
        public string subject;
        public string imageComment;
        public string imagePATH;
        public string attPath;
        public string userMacro;
        public List<string> ReParameters(FFOperation NewFF,List<string> a)
        {
            a.Add(host);

            return a;
        }

        public FFOperation()
        { folderpath = System.IO.Directory.GetCurrentDirectory(); }
        ~FFOperation() { }


        [DllImport("User32.dll", CharSet = CharSet.Auto)]
        private static extern int GetWindowThreadProcessId(IntPtr hwnd, out int ID);
        /// <summary>
        /// 杀掉指定进程
        /// </summary>
        /// <param name="m_objExcel">进程对象名称</param>
        private void KillSpecialExcel(Excel.ApplicationClass m_objExcel)
        {
            try
            {
                if (m_objExcel != null)
                {
                    int lpdwProcessId;
                    GetWindowThreadProcessId(new IntPtr(m_objExcel.Hwnd), out lpdwProcessId);

                    System.Diagnostics.Process.GetProcessById(lpdwProcessId).Kill();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Delete Excel Process Error:" + ex.Message);
            }
        }
       /// <summary>
       /// 打开excel对象-->运行指定宏-->杀掉excel进程
       /// </summary>
       /// <param name="excelFilePath">excel存放路径</param>
       /// <param name="isShowExcel">运行excel时是否显示</param>
        public void RunExcelMacro(string excelFilePath, bool isShowExcel)
        {
            try
            {
                // 检查文件是否存在
                if (!File.Exists(excelFilePath))
                {
                    throw new System.Exception(excelFilePath + " 文件不存在");
                }            
                // 准备打开Excel文件时的缺省参数对象
                object oMissing = System.Reflection.Missing.Value;
                          
                // 创建Excel对象示例
                Excel.ApplicationClass oExcel = new Excel.ApplicationClass();

                // 判断是否要求执行时Excel可见
                if (isShowExcel)
                {
                    // 使创建的对象可见
                    oExcel.Visible = true;
                }
                else
                {
                    oExcel.Visible = false;
                }
                // 创建Workbooks对象
                Excel.Workbooks oBooks = oExcel.Workbooks;
                // 创建Workbook对象
                Excel._Workbook oBook = null;
                // 打开指定的Excel文件
                oBook = oBooks.Open(excelFilePath);
                //提取excel自定义参数用于发送mail
                foreach (Worksheet sheet in oBook.Worksheets)
                    if (sheet.Name == "setting")
                    {
                        //设定excel文档传参 
                       
                         host = (string)sheet.Range["B1"].Value2.ToString(); 
                         from = (string)sheet.Range["B2"].Value2.ToString();

                        if (sheet.Range["B3"].Value2 == null)
                        { cc = ""; }
                        else
                        { cc = (string)sheet.Range["B3"].Value2.ToString(); }
                        
                         to = (string)sheet.Range["B4"].Value2.ToString();

                         subject = (string)sheet.Range["B5"].Value2.ToString()+"--"+ DateTime.Now.ToLongTimeString();

                        if (sheet.Range["B7"].Value2==null)
                        {
                            imageComment ="";
                            imagePATH = "";
                         }
                        else
                        {
                            imageComment = (string)sheet.Range["B6"].Value2.ToString();
                            imagePATH = (string)sheet.Range["B7"].Value2.ToString();
                        }

                        if (sheet.Range["B8"].Value2 == null)
                        { attPath = ""; }
                        else
                        { attPath = (string)sheet.Range["B8"].Value2.ToString(); }
                       

                        if (sheet.Range["B9"].Value2 == null)
                        { userMacro = ""; }
                        else
                        { userMacro = (string)sheet.Range["B9"].Value2.ToString(); }

                        break;
                    }
                   
                // 执行Excel中的宏
                    if(userMacro!="N")
                    {
                        object[] paraObjects;
                        if (userMacro == "")
                        {
                            paraObjects = new object[] { "AutoRun" };
                        }
                        else
                        {
                            paraObjects = new object[] { userMacro };
                        }

                        RunMacro(oExcel, paraObjects);

                        oBook.SaveAs(folderpath + @"\ZIP\RefreshDone.xlsm");
                        oBook.SaveAs(folderpath + @"\ZIP\test.zip");
                    }
                


                oBook.Close(false, oMissing, oMissing);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBook);
                oBook = null;
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oBooks);
                oBooks = null;
                //oExcel.Quit();
                KillSpecialExcel(oExcel);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oExcel);
                oExcel = null;
                // 调用垃圾回收
                GC.Collect();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
     
        /// <summary>
        /// 运行宏
        /// </summary>
        /// <param name="oApp">打开的excel对象</param>
        /// <param name="arg">参数列表（第一个参数为要运行的宏名称，后面为指定宏的参数list）</param>
        private void RunMacro(object oApp, object[] arg)
        {
            try
            {
                oApp.GetType().InvokeMember("Run", System.Reflection.BindingFlags.Default |
                                                   System.Reflection.BindingFlags.InvokeMethod,
                                                   null, oApp, arg);
            }
            catch (Exception ex)
            {
                if (ex.InnerException.Message.ToString().Length > 0)
                {
                    throw ex.InnerException;
                }
                else
                {
                    throw ex;
                }
            }
        }


      //文件操作 的相关方法

        public void CheckFolder()
        {
            if (!Directory.Exists(folderpath + @"\ZIP"))
            {
                Directory.CreateDirectory(folderpath + @"\ZIP");
            }
        }

        public void DeleteFolder()
        {
            if (Directory.Exists(folderpath + @"\ZIP"))
            {
                Directory.Delete(folderpath + @"\ZIP", true);
            }
        }

        /// <summary>
        /// 提取出指定路径下的所有excel文件
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public List<string> getExcelName(List<string> filename)
        {
            DirectoryInfo files = new DirectoryInfo(folderpath);
            FileInfo[] file = files.GetFiles("*.xls*");
            for (int i = 0; i < file.Length; i++)
            {
                filename.Add(file[i].Name);
            }
            return filename;
        }

        /// <summary>
        /// 解压缩
        /// </summary>
        public void unzip()
        {
            ZipFile.ExtractToDirectory(folderpath + @"\ZIP\test.zip", folderpath + @"\ZIP");
        }

        /// <summary>
        /// 删除指定文件夹下的所有文件和文件夹
        /// </summary>
        /// <param name="path">文件夹路径</param>
        public void DeleteFile(string path)
        {
            foreach (string d in Directory.GetFileSystemEntries(path))
            {
                if (File.Exists(d))
                {
                    FileInfo fi = new FileInfo(d);
                    if (fi.Attributes.ToString().IndexOf("ReadOnly") != -1)
                        fi.Attributes = FileAttributes.Normal;
                    File.Delete(d);//直接删除其中的文件           
                }
                else
                {
                    DirectoryInfo d1 = new DirectoryInfo(d);
                    if (d1.GetFiles().Length != 0)
                    {
                        DeleteFile(d1.FullName);////递归删除子文件夹               
                    }
                    Directory.Delete(d);
                }
            }
        }

    }
}
