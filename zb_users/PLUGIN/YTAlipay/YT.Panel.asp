<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001"%>
<%
'///////////////////////////////////////////////////////////////////////////////
'// 作	 者:    	瑜廷(YT.Single)
'// 技术支持:    33195@qq.com
'// 程序名称:    	Content Manage System
'// 开始时间:    	2011.03.26
'// 最后修改:    2012-08-08
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<%' On Error Resume Next %>
<% Response.Charset="UTF-8" %>
<% Response.Buffer=True %>
<!-- #INCLUDE FILE="../../C_OPTION.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_FUNCTION.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_SYSTEM_LIB.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_SYSTEM_BASE.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_SYSTEM_EVENT.ASP" -->
<!-- #INCLUDE FILE="../../../ZB_SYSTEM/FUNCTION/C_SYSTEM_PLUGIN.ASP" -->
<!-- #INCLUDE FILE="../../PLUGIN/P_CONFIG.ASP" -->
<%
Call System_Initialize()
'检查非法链接
Call CheckReference("")
'检查权限
If BlogUser.Level>1 Then Call ShowError(6) 
If CheckPluginState("YTAlipay") = False Then Call ShowError(48)
If CheckPluginState("YTCMS") = False Then
	Response.Write("您没有安装YT.CMS插件,无法进行配置管理")
	Response.End()
End If
%>
<!--#include file="..\..\..\zb_system\admin\admin_header.asp"-->
<script language="javascript" type="text/javascript">
	var ZC_BLOG_HOST='<%=ZC_BLOG_HOST%>';
	var YT_CMS_XML_URL='<%=ZC_BLOG_HOST&"ZB_USERS/THEME/"&ZC_BLOG_THEME&"/"%>';
	var isAlipay = <%=LCase(CheckPluginState("YTAlipay"))%>
</script>
<script language="javascript" src="../YTCMS/SCRIPT/YT.Lib.js" type="text/javascript"></script>
<!--#include file="..\..\..\zb_system\admin\admin_top.asp"-->
<div id="divMain"><div id="ShowBlogHint">
      <%Call GetBlogHint()%>
    </div>
  <div class="divHeader">支付宝</div>
  <div class="SubMenu"> <a href="YT.Panel.asp"><span class="m-left m-now">订单管理</span></a><a href="YT.Config.asp"><span class="m-left">系统配置</span></a>
  </div>
  <div id="divMain2">
  <div id="Panel">数据加载中...</div>
<div id="aListView" style="display:none">
      <table width="100%" style="margin-top:0;" cellspacing="0" cellpadding="0" border="0">
        <tr class="color2">
          <td height="25" align="center">订单号</td>
          <td align="center">订单名称</td>
          <td align="center">交易状态</td>
          <td align="center">创建时间</td>
          <td align="center">操作</td>
        </tr>
        <tr class="color1">
          <td align="center"></td>
          <td align="center"></td>
          <td align="center"></td>
          <td align="center"></td>
          <td align="center"></td>
        </tr>
        <tr>
        	<td colspan="5">
            	<input type="text" value="" />
                <select>
                <option value="0">订单号</option>
                <option value="1">订单名称</option>
                </select><input type="button" value="查询" />
                <select></select><font></font>
            </td>
        </tr>
      </table>
</div>
<div id="aDetails" style="display:none">
<table width="100%" style="margin-top:0;" cellspacing="0" cellpadding="0" border="0">
<tr class="color1">
<td width="25%" align="right"></td>
<td width="75%"></td>
</tr>
<tr>
  <td align="right">&nbsp;</td>
  <td>
    <input type="button" value="发货" /><p style="display:none;">
支付宝交易号：<em></em><br />
物流公司名称：<input type="text" /><br />
物流发货单号：<input type="text" /><br />
物流发货类型：<select>
  <option value="EMS">EMS</option>
  <option value="POST">平邮</option>
  <option value="EXPRESS" selected="selected">快递</option>
</select>
    </p><input type="button" value="返回" />
  </td>
</tr>
</table>
</div>
</div>
</div>
<script type="text/javascript">ActiveLeftMenu("aYTAlipayMng");</script>
<script language="javascript" type="text/javascript">
$(document).ready(function(){
	YT.Panel.Alipay.M(({ Action: 'GetJsonList',OrderID:'',OrderName:'',intPage:1,t:Math.random()}));					   
});
YT.Panel.Alipay={
	YT_ALIPAY_TRANSACTION_TYPE:['担保交易','即时交易','双功能']	,
	TRADE_STATUS:[[{Text:'交易创建',
					   Intro:'等待买家付款。注：当卖家修改了价格时，也是发送该通知'},
					   {Text:'买家已付款，等待卖家发货',
					   Intro:'买家已付款，等待卖家发货'},
					   {Text:'卖家已发货，等待买家确认收货',
					   Intro:'卖家已发货，等待买家确认收货'},
					   {Text:'交易成功',
					   Intro:'买家已确认收货，交易结束'},
					   {Text:'交易关闭',
					   Intro:'出现情况： 1、 卖家或系统关闭了买家还没有付款的交易 2、 买家申请退款成功'},
					   {Text:'买家申请退款',
					   Intro:'进入退款流程'},
					   {Text:'卖家拒绝退款',
					   Intro:'卖家拒绝买家的申请退款，此时买家可再申请退款也可继续走正常的交易流程'},
					   {Text:'卖家同意退款，等待买家退货',
					   Intro:'等待买家把货寄回给卖家，当买家选择有收到货时'},
					   {Text:'买家已退货，等待卖家收到退货',
					   Intro:'买家已把货寄回给卖家，等待卖家收到退回的货，当买家选择有收到货时'},
					   {Text:'退款成功',
					   Intro:'交易完成'},
					   {Text:'退款关闭',
					   Intro:'买卖双方终止了退款操作，并走正常交易流程完成了交易'}],
					  [{Text:'交易创建',
					   Intro:'等待买家付款。该状态不会发送通知且不开通。'},
					   {Text:'交易成功',
					   Intro:'没有开通高级即时到帐的成功状态'},
					   {Text:'交易成功',
					   Intro:'开通了高级即时到帐的成功状态'},
					   {Text:'交易关闭（默认不开通）',
					   Intro:'全额退款时反馈、卖家手动关闭交易状态是“等待买家付款状态”或买家逾期不付款系统自动关闭交易。该状态开通条件：开通高级即时到帐功能权限，且非常需要在通知里判断订单退款同步，可申请开通。 '},
					   {Text:'退款成功',
					   Intro:'进入退款流程后，交易状态与退款状态会同时存在。 全额退款情况时：trade_status= TRADE_CLOSED，而refund_status=REFUND_SUCCESS；不是全额退款情况时：trade_status= TRADE_SUCCESS，而refund_status=REFUND_SUCCESS； '},,
					   {Text:'退款关闭',
					  Intro:'进入退款流程后，交易状态与退款状态会同时存在。 全额退款情况时：trade_status= TRADE_CLOSED，而refund_status=REFUND_SUCCESS；不是全额退款情况时：trade_status= TRADE_SUCCESS，而refund_status=REFUND_SUCCESS； '}]],
	M:function(n){
		$('.readyAlipay').remove();
		$.ajax({
			url: 'YT.Ajax.asp',
			type: 'POST',
			dataType: 'html',
			data: n,
			success: function(json) {
				try{
					if(json!=''){d=eval('('+json+')');}
					var d = d||{intPage:0,pageCount:0,objRow:[]};
					var t = $('#aListView').find('table').clone();
						if(n.OrderID != ''){
							t.find('tr:last').find('input').eq(0).val(n.OrderID);
							t.find('tr:last').find('select').eq(0).val(0);
						}
						if(n.OrderName != ''){
							t.find('tr:last').find('input').eq(0).val(n.OrderName);
							t.find('tr:last').find('select').eq(0).val(1);
						}
						t.find('tr:last').find('input').eq(1).click(function(){
							switch(parseInt(t.find('tr:last').find('select').eq(0).val())){
								case 0:
								n.OrderID = t.find('tr:last').find('input').eq(0).val();
								n.OrderName = '';
								break;
								case 1:
								n.OrderID = '';
								n.OrderName = t.find('tr:last').find('input').eq(0).val();
								break;
							}
							n.t=Math.random();
							YT.Panel.Alipay.M(n);
						});
						t.find('tr:last').find('font').text(d.intPage+'/'+d.pageCount);
						for(var i=0;i<d.pageCount;i++){
							t.find('tr:last').find('select').eq(1).append('<option value="'+(i+1)+'">第'+(i+1)+'页</option>');
						}
						t.find('tr:last').find('select').eq(1).val(d.intPage).change(function(){
							n.intPage = $(this).val();
							YT.Panel.Alipay.M(n);
						});
						var YT_ALIPAY_TYPE = 1;
						for(var i=0;i<d.objRow.length;i++){
							YT_ALIPAY_TYPE = 1;
							if(d.objRow[i].Service==0||d.objRow[i].Service==2){YT_ALIPAY_TYPE = 0;}
							var r = t.find('tr').eq(1).clone();
								r.find('td').eq(0).attr('align','').text(d.objRow[i].OrderID);
								r.find('td').eq(1).attr('align','').html('<a href="'+ZC_BLOG_HOST+'zb_system/view.asp?id='+d.objRow[i].log_ID+'" target="_blank">'+unescape(d.objRow[i].OrderName)+'</a>');
								r.find('td').eq(2).text(YT.Panel.Alipay.TRADE_STATUS[YT_ALIPAY_TYPE][d.objRow[i].Status].Text).attr('title',YT.Panel.Alipay.TRADE_STATUS[YT_ALIPAY_TYPE][d.objRow[i].Status].Intro);
								r.find('td').eq(3).text(d.objRow[i].Time);
								r.find('td').eq(4).html('<a href="javascript:void(0)" rel="'+d.objRow[i].ID+'">删除</a> | <a href="javascript:void(0)" rel="'+i+'">详情</a>');
								r.find('td').eq(4).find('a:first').click(function(){
									if(confirm('您确定删除此订单吗?')){
										var ID = $(this).attr('rel');
										$.get('YT.Ajax.asp', { Action: 'DelAlipayOrder',ID:ID,t:Math.random() },
										function(){
											YT.Panel.Alipay.M(n);
										});
									}				  
								});
								r.find('td').eq(4).find('a:last').click(function(){
									YT_ALIPAY_TYPE = 1;
									if(d.objRow[$(this).attr('rel')].Service==0||d.objRow[$(this).attr('rel')].Service==2){YT_ALIPAY_TYPE = 0;}
									var _t = $('#aDetails').find('table').clone();
									var _d = [{Text:'订单号',Value:d.objRow[$(this).attr('rel')].OrderID}];
										_d.push({Text:'订单名称',Value:unescape(d.objRow[$(this).attr('rel')].OrderName)});
										_d.push({Text:'交易状态',Value:YT.Panel.Alipay.TRADE_STATUS[YT_ALIPAY_TYPE][d.objRow[$(this).attr('rel')].Status].Text});
										_d.push({Text:'创建时间',Value:d.objRow[$(this).attr('rel')].Time});
									var jsonBody = unescape(d.objRow[$(this).attr('rel')].Body);
										jsonBody = eval('('+jsonBody+')');
										for(var j=0;j<jsonBody.length;j++){
											_d.push(jsonBody[j]);	
										}
										for(var j=0;j<_d.length;j++){
											_r = _t.find('tr').eq(0).clone();
											_r.find('td:first').html(unescape(_d[j].Text));
											if(unescape(_d[j].Text) == '交易类型'){
												_r.find('td:last').html(YT.Panel.Alipay.YT_ALIPAY_TRANSACTION_TYPE[d.objRow[$(this).attr('rel')].Service]);
											}else{
												_r.find('td:last').html(unescape(_d[j].Value));	
											}
											_r.attr('class','ready').insertBefore(_t.find('.color3')[0]);
										}
										if(YT_ALIPAY_TYPE==0&&d.objRow[$(this).attr('rel')].Status==1){
											_t.find('p').find('em').text(d.objRow[$(this).attr('rel')].trade_no);
											var out_trade_no = d.objRow[$(this).attr('rel')].OrderID;
											var service = d.objRow[$(this).attr('rel')].Service;
											_t.find('input:first').show().click(function(){
												if($(this).val()=='发货'){
													$(this).unbind('click');
													$(this).val('确认').click(function(){	 
														var e=this,j={};
															j.Action='Delivery';
															j.out_trade_no=out_trade_no;
															j.service=service;
															j.trade_no=$(e).parent().find('p').find('em').text();
															j.logistics_name=$(e).parent().find('p').find('input').eq(0).val();	
															j.invoice_no=$(e).parent().find('p').find('input').eq(1).val();
															j.transport_type=$(e).parent().find('p').find('select').val();
															j.node=$.toJSONString(['response/tradeBase/trade_status']);
															j.t=Math.random();
															if(j.logistics_name!=''&&j.invoice_no!=''){
																$.ajax({
																	url: 'YT.Ajax.asp',
																	type: 'POST',
																	dataType: 'html',
																	data: j,
																	success: function(s) {
																		s=s.split('|');
																		if(s[0]=='WAIT_BUYER_CONFIRM_GOODS'){
																			_t.find('tr.ready').eq(2).find('td').eq(1).text(YT.Panel.Alipay.TRADE_STATUS[YT_ALIPAY_TYPE][2].Text).attr('title',YT.Panel.Alipay.TRADE_STATUS[YT_ALIPAY_TYPE][2].Intro);
																			
																			$(e).hide().parent().find('p').hide();
																		}
																	}
																});	
															}
													}).parent().find('p').show();	
												}
											});
										}else{
											_t.find('input:first').hide();
										}
										_t.find('input:last').click(function(){YT.Panel.Alipay.M(n);}); 
										$('#Panel').html(_t);
								});
								r.attr('class','readyAlipay').hover(function() {
									$(this).addClass('color1')
								}, function() {
									$(this).removeClass('color1')
								}).insertBefore(t.find('.color3')[0]);
						}
						$('#Panel').html(t);
				}catch(e){}
			}
		});
	}
}
</script>
<!--#include file="..\..\..\zb_system\admin\admin_footer.asp"-->
<%
Call System_Terminate()

If Err.Number<>0 then
  Call ShowError(0)
End If
%>