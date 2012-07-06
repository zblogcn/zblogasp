<%@ WebHandler Language="C#" Class="fileUp" %>
/**
 * Created by visual studio 2010
 * User: xuheng
 * Date: 12-3-9
 * Time: 下午13:53
 * To change this template use File | Settings | File Templates.
 */
using System;
using System.Web;
using System.IO;

public class fileUp : IHttpHandler {
    
    public void ProcessRequest (HttpContext context) {
        context.Response.ContentType = "text/plain";
        
        //上传配置
        String pathbase = "upload/";                                      //保存路径
        string[] filetype = { ".rar", ".doc", ".docx", ".zip", ".pdf", ".txt", ".swf", ".wmv" };    //文件允许格式
        int size = 100;   //文件大小限制,单位MB,同时在web.config里配置环境默认为100MB
        
        //文件上传状态,当成功时返回SUCCESS,其余值将直接返回对应字符串
        String state = "SUCCESS";
        
        String title = String.Empty;
        String filename = String.Empty;
        String url = String.Empty;
        String currentType = String.Empty;
        String uploadpath = String.Empty;

        uploadpath = context.Server.MapPath(pathbase);

        HttpPostedFile uploadFile = null;

        try
        {
            uploadFile = context.Request.Files["upfile"];
            title = uploadFile.FileName;

            //目录验证
            if (!Directory.Exists(uploadpath))
            {
                Directory.CreateDirectory(uploadpath);
            }
            if (uploadFile == null)
            {
                context.Response.Write("{'state':'文件大小可能超出服务器环境配置！','url':'null','fileType':'null'}");
            }

            //格式验证
            string[] temp = uploadFile.FileName.Split('.');
            currentType = "." + temp[temp.Length - 1].ToLower();
            if (Array.IndexOf(filetype, currentType) == -1)
            {
                state = "不支持的文件类型！";
            }

            //大小验证
            if ((uploadFile.ContentLength / 1024) / 1024 > size)
            {
                state = "文件大小超出限制！";
            }

            //保存图片
            if (state == "SUCCESS")
            {
                filename = DateTime.Now.ToString("yyyy-MM-dd-ss") + System.Guid.NewGuid() + currentType;
                uploadFile.SaveAs(uploadpath + filename);
                url = pathbase + filename;
            }
        }
        catch (Exception)
        {
            state = "文件保存失败";
        }
        //向浏览器返回数据json数据
        context.Response.Write("{'state':'" + state + "','url':'" + url + "','fileType':'" + currentType + "','original':'" + uploadFile.FileName + "'}");
    }
 
    public bool IsReusable {
        get {
            return false;
        }
    }

}