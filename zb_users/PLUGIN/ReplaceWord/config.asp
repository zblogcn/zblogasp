<?xml version="1.0" encoding="utf-8"?>
<words>
  <antidownload>
    <![CDATA[<% 'On Error Resume Next %>
      <% Response.Charset="UTF-8" %>
      <!-- #include file="..\..\c_option.asp" -->
      <!-- #include file="..\..\..\zb_system\function\c_function.asp" -->
      <!-- #include file="..\..\..\zb_system\function\c_system_lib.asp" -->
      <!-- #include file="..\..\..\zb_system\function\c_system_base.asp" -->
      <!-- #include file="..\..\..\zb_system\function\c_system_event.asp" -->
      <!-- #include file="..\..\..\zb_system\function\c_system_manage.asp" -->
      <!-- #include file="..\..\..\zb_system\function\c_system_plugin.asp" -->
      <!-- #include file="..\p_config.asp" -->
      <%System_Initialize:If BlogUser.Level>1 Then Call ShowError(6)%>]]>
  </antidownload>
  <word user="1" regexp="False">
    <str>fuck</str>
    <replace>**</replace>
    <description>脏话</description>
  </word>
  <word user="1" regexp="False">
    <str>你妈逼</str>
    <replace>**</replace>
    <description>脏话</description>
  </word>
</words>
