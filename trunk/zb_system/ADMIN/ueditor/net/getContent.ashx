﻿<%@ WebHandler  Language="C#"  Class="getContent" %>
/**
 * Created by visual studio 2010
 * User: xuheng
 * Date: 12-3-6
 * Time: 下午21:23
 * To get the value of editor and output the value .
 */
using System;
using System.Web;

public class getContent : IHttpHandler {
    
    public void ProcessRequest (HttpContext context) {
        context.Response.ContentType = "text/html";

        //获取数据
        string content = context.Server.HtmlEncode(context.Request.Form["myEditor"]);
        string content1 = context.Server.HtmlEncode(context.Request.Form["myEditor1"]);
        //存入数据库或者其他操作
        //-------------

        //显示
        context.Response.Write("第1个编辑器的值;");
        context.Response.Write(context.Server.HtmlDecode(content));
        context.Response.Write("第2个编辑器的值;");
        context.Response.Write(context.Server.HtmlDecode(content1));
        context.Response.Write("<input type='button' value='点击返回' onclick='window.history.go(-1)' />");
    }
 
    public bool IsReusable {
        get {
            return false;
        }
    }

}