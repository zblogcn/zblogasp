﻿<%@ WebHandler Language="C#" Class="imageManager" %>
/**
 * Created by visual studio2010
 * User: xuheng
 * Date: 12-3-7
 * Time: 下午16:29
 * To change this template use File | Settings | File Templates.
 */
using System;
using System.Web;
using System.IO;

public class imageManager : IHttpHandler {
    
    public void ProcessRequest (HttpContext context) {
        context.Response.ContentType = "text/plain";

        string path = context.Server.MapPath("upload/");                  //最好使用缩略图地址，否则当网速慢时可能会造成严重的延时
        string[] filetype = { ".gif", ".png", ".jpg", ".jpeg", ".bmp" };                //文件允许格式
        
        string action = context.Server.HtmlEncode(context.Request["action"]);
       
        if(action == "get")
        {
            String str=String.Empty;
            DirectoryInfo info = new DirectoryInfo(path);
            
            //目录验证
            if (info.Exists)
            {
                foreach (FileInfo fi in info.GetFiles())
                {
                    if (Array.IndexOf(filetype, fi.Extension) != -1)
                    {
                        str += "upload/" + fi.Name + "ue_separate_ue";
                    }
                }
            }
            context.Response.Write(str);
        }
    }

 
    public bool IsReusable {
        get {
            return false;
        }
    }

}