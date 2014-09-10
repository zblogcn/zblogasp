<%@ LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<% Option Explicit %>
<% 'On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<!-- #include file="..\..\c_option.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_function.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_lib.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_base.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_event.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_manage.asp" -->
<!-- #include file="..\..\..\zb_system\function\c_system_plugin.asp" -->
<!-- #include file="..\p_config.asp" -->
<!-- #include file="ASPJSON.asp" -->

<%
Call System_Initialize()
If BlogUser.Level > 1 Then Call ShowError(6)
If Not CheckPluginState("A2P") Then Call ShowError(48)
BlogTitle = "转换工具"

Dim aryTable, aryTables
Set aryTable = A2P_jsObject()
aryTables = Array("category", "article", "comment", "tag", "upload", "member")
aryTable("category") = "cate_ID"
aryTable("article")  = "log_ID"
aryTable("comment")  = "comm_ID"
aryTable("tag")      = "tag_ID"
aryTable("upload")   = "ul_ID"
aryTable("member")   = "mem_ID"

%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<style type="text/css">
tr {
	height: 32px
}
#show_result {
	cursor: wait
}
</style>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
        <div id="divMain">
          <div id="ShowBlogHint">
            <%Call GetBlogHint()%>
          </div>
          <div class="divHeader"><%=BlogTitle%></div>
          <div class="SubMenu"></div>
          <div id="divMain2"> 
            <script type="text/javascript">ActiveTopMenu("aPlugInMng");</script>
            <%ExportPage%>
          </div>
        </div>
        <!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->

<%Call System_Terminate()%>
<%
Sub ExportPage

%>

<div id="div-result" style="display:none">
  <p>进度：<span id="span-progress">0</span>%，已用时<span id="span-second">0</span>秒。</p>
  <progress id="progress" max="100" min="0" value="0" style="width: 100%"></progress>
  <ul id="ul-result">
  </ul>
</div>
<div id="div-input">
  <p>点击按钮，即开始转换所有的文章。</p>
  <p>转换完成后，请手动下载FTP内<b>zb_users/plugin/A2P/output</b>全部数据并上传到Z-BlogPHP的<b>zb_users/plugin/PDC/input</b>文件夹内</p>
  <form method="post" action="">
    <p>
      <input name="" id="button-submit" type="submit" class="button" value="开始转换" />
    </p>
  </form>
</div>
<script type="text/javascript">

(function(){
	window.A2P = {
		time_second: 0,
		config: {
			rebuild_count: 200,
		},
		data: <%=A2P_toJson(GetDatabaseCount())%>,
		step: 0,
		pos: 0,
		queue: [],
		build: function() {
			for(var name in this.data) {
				var value = this.data[name];
				for(var i = value.min - 1; i <= value.max; i += this.config.rebuild_count) {
					this.queue.push({"type": name, "min": i + 1, "max": i + this.config.rebuild_count});
				}
			}
		},
		log: function(data) {
			$("<li>【" + new Date().getFullDate() + "】 " + data + "</li>").appendTo("#ul-result");
		},
		run: function() {
			if (this.pos >= this.queue.length) {
				clearInterval(this.time_second);
				this.log("转换完成！请手动下载FTP内zb_users\/plugin\/A2P\/output全部文件并上传到Z-BlogPHP的zb_users\/plugin\/PDC\/intput文件夹内");
				alert("转换完成！请手动下载FTP内zb_users\/plugin\/A2P\/output全部文件并上传到Z-BlogPHP的zb_users\/plugin\/PDC\/input文件夹内");
				return;
			} else {
				var that = this;
				return $.ajax({
					url: "post.asp",
					type: "POST",
					data: $.extend(that.queue[that.pos], {"count": that.pos + 1, "sum": that.queue.length}),
					dataType: "json",
					error: function(xhr){
						that.log("任务" + (that.pos + 1) + "执行失败，出现" + xhr.status + "错误，具体错误为" + xhr.responseText);
					},
					success: function(data) {
						var progress = ((that.pos / (that.queue.length - 1)) * 100).toFixed(2);
						$("#span-progress").html(progress);
						$("#progress").val(progress);
						that.log("任务" + (that.pos + 1) + "返回：" + data.message);
						that.pos++;
						that.run();
					}
				});
			}
		}
	};

	Date.prototype.getFullDate = function() {
		var o = this;
		return o.getFullYear() + "-" + (o.getMonth() + 1) + "-" + o.getDate() + " " + o.getHours() + ":" + o.getMinutes() + ":" + o.getSeconds();
	}
	
	$("#button-submit").click(function() {
		A2P.time_second = setInterval(function() {
			var obj = $("#span-second");
			obj.html(parseInt(obj.html()) + 1);
		}, 1000);
		$("#div-result").show();
		$("#div-input").hide();
		A2P.build();
		A2P.log("分割为" + (A2P.queue.length) + "份子任务");
		A2P.run();
		return false;
	});
	
})();



</script>
<%
End Sub

Function GetDatabaseCount()

	Dim i
	Dim strSQL, objJSON, aryJSON, objRs
	Dim strTable, strField
	Set aryJSON = A2P_jsObject
	
	For i = 0 To Ubound(aryTables)
		strTable = aryTables(i)
		strField = aryTable(strTable)
		Set aryJSON(strTable) = A2P_jsObject
		strSQL = "SELECT COUNT(" & strField & ") AS icount, MAX(" & strField & ") AS imax, MIN(" & strField & ") AS imin "
		strSQL = strSQL & " FROM [blog_" & strTable & "]"
		Set objRs = objConn.Execute(strSQL)
		aryJSON(strTable)("count") = objRs("icount")
		aryJSON(strTable)("max"  ) = objRs("imax"  )
		aryJSON(strTable)("min"  ) = objRs("imin"  )
	Next 
	
	Set GetDatabaseCount = aryJSON

End Function
%>