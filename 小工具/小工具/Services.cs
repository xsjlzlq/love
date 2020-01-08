using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using OfficeOpenXml;
using OfficeOpenXml.Packaging;
using OfficeOpenXml.Style;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Globalization;
using System.Text.RegularExpressions;
namespace 小工具
{
   public static class Services
    {
       //Excel文件
       public static List<Qlrs> GetQlrInfo(this string fileName)
       {
           if (string.IsNullOrWhiteSpace(fileName)) return null;
           List<Qlrs> qlr_list = new List<Qlrs>();
        if (fileName.IndexOf(".xlsx") > 0)
        {
            using (ExcelPackage excel = new ExcelPackage(new FileInfo(fileName)))
            {
                var sheet = excel.Workbook.Worksheets[1];
                int endRow = sheet.Dimension.End.Row;
                for (int i = 1; i < endRow + 1; i++)
                {
                    try
                    {
                        Qlrs qlr = new Qlrs() 
                        { 
                         ID=sheet.Cells[i,1].Value.ToString().Trim(),
                         QLRMC=sheet.Cells[i,2].Value.ToString().Trim()
                        };
                        qlr_list.Add(qlr);
                    }
                    catch (Exception) { continue; }
                }
            }
        }
            //文本文件
        else
        {
            using (FileStream fs = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                StreamReader sr = new StreamReader(fs,Encoding.Default);
                while (!sr.EndOfStream)
                {
                    try
                    {
                        string[] str = sr.ReadLine().Split(',');
                        Qlrs qlr = new Qlrs()
                        {
                            ID = str[0].Trim(),
                            QLRMC = str[1].Trim()
                        };
                        qlr_list.Add(qlr);
                    }
                    catch (Exception) { continue; }
                }
                sr.Close();
                fs.Close();
            }
        }
        return qlr_list;
       }
       public static DateTime? GetImagePSSJ(this string fileName)
       {
          
           try
           {
               if (!fileName.Contains(".jpg")) return null;
               Image theImage = Image.FromFile(fileName);
               PropertyItem[] propItems = theImage.PropertyItems;
               PropertyItem propItem = theImage.GetPropertyItem(0x9003);
               Byte[] propItemValue = propItem.Value;
               string dateTimeStr = System.Text.Encoding.ASCII.GetString(propItemValue).Trim('\0');
               DateTime? dt = DateTime.ParseExact(dateTimeStr, "yyyy:MM:dd HH:mm:ss", CultureInfo.InvariantCulture);
               if (dt.HasValue)
                   return dt;
               else
                   return null;
           }
           catch (Exception)
           { return null; }
           
               

       }
       public static List<Images> GetImageInfo(this string path,bool _check)
       {
           if (string.IsNullOrWhiteSpace(path)) return null;
           List<Images> img_List = new List<Images>();
           DirectoryInfo dir = new DirectoryInfo(path);
           foreach(FileInfo file in dir.GetFiles())
           {
                if (_check == true&&file.Extension != ".jpg" && file.Extension != ".JPG" &&
                     file.Extension != ".png" && file.Extension != ".PNG" &&
                     file.Extension != ".tif" && file.Extension != ".TIF" &&
                     file.Extension != ".bmp" && file.Extension != ".BMP") continue;
             
            Images img=new Images()
            {
             MC=file.Name,
             PSSJ = file.CreationTime,
             CreateTime =file.LastWriteTime,
             FullName=file.FullName
            };
            img_List.Add(img);
           }
           return img_List;
           
       }
       /// <summary>
       /// 
       /// </summary>
       /// <param name="path">目录文件夹</param>
       /// <param name="qlr_list">权利人列表</param>
       /// <param name="image_list">照片集合</param>
       /// /// <param name="k">比例</param>
       public static void Move(this string path,List<Qlrs> qlr_list,List<Images> image_list,int k)
       {
           int start = 0;
           foreach (var qlr in qlr_list)
           {
               string pathName = path + "\\" + qlr.ID+"、"+qlr.QLRMC;
               if(!Directory.Exists(pathName))
                   Directory.CreateDirectory(pathName);
               int max = start + k > image_list.Count ? image_list.Count : start + k;
               for (int i =start; i <max; i++)
               {
                   try
                   {
                       File.Move(image_list[i].FullName, pathName + "\\" + image_list[i].MC);
                   }
                   catch (Exception) { break; }
               }
               start+=k;
           }
       }
    }

}
