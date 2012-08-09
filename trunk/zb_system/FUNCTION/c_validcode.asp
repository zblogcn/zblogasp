<%@ CODEPAGE=65001 %>
<%
'///////////////////////////////////////////////////////////////////////////////
'//              Z-Blog
'// 作    者:    朱煊(zx.asd)
'// 版权所有:    RainbowSoft Studio
'// 技术支持:    rainbowsoft@163.com
'// 程序名称:    
'// 程序版本:    
'// 单元名称:    c_validcode.asp
'// 开始时间:    2005.02.18
'// 最后修改:    
'// 备    注:    校验码生成 其中校验码类库引用并修改了网络上的代码
'///////////////////////////////////////////////////////////////////////////////
%>
<% Option Explicit %>
<% On Error Resume Next %>
<!-- #include file="../../zb_users/c_option.asp" -->
<!-- #include file="c_function.asp" -->
<!-- #include file="c_system_lib.asp" -->
<!-- #include file="c_system_base.asp" -->
<!-- #include file="c_system_plugin.asp" -->
<!-- #include file="../../zb_users/plugin/p_config.asp" -->

<%
Call ActivePlugin()
'//////' QQ17862153/////ASP无组件验证码程序开始///////////
Class Com_GifCode_Class
Public Noisy, Count, Width, Height, Angle, Offset, Border
Private Graph(), Margin(3)
Private Sub Class_Initialize()
	Randomize 
	Noisy = 6 + Fix(Rnd()*5) ' 干扰点出现的概率
	Count = 5		' 字符数量
	Width = ZC_VERIFYCODE_WIDTH	    ' 图片宽度
	Height = ZC_VERIFYCODE_HEIGHT		' 图片高度
	Angle = 2		' 角度随机变化量
	Offset = 5		' 偏移随机变化量
	Border = 1		' 边框大小
End Sub 
Public Function Create(str)
	Dim i
	Dim vIndex
	ReDim Graph(Width-1, Height-1)
	For i = 0 To Count - 1
		vIndex=Asc(CStr(Mid(str,i+1,1)))
		SetDraw vIndex, i
	Next
End Function

Sub SetDot(pX, pY)
	If pX * (Width-pX-1) >= 0 And pY * (Height-pY-1) >= 0 Then
		Graph(pX, pY) = 1
	End If
End Sub

Public Sub SetDraw(pIndex, pNumber)
	Dim DotData(20000)
	If pIndex<0 Then pIndex=0-pIndex
	DotData(48) = Array(100, 20, 70, 1, 20, 1, 1, 30, 1, 80, 30, 100, 70, 100, 100, 80, 100, 60, 90, 20, 80,3)
	DotData(49) = Array(30, 15, 50, 1, 50, 100)
	DotData(50) = Array(1 ,34 ,30 ,1 ,71, 1, 100, 34, 1, 100, 93, 100, 100, 86)
	DotData(51) = Array(1, 1, 100, 1, 42, 42, 100, 70, 50, 100, 1, 70)
	DotData(52) = Array(100, 73, 6, 73, 75, 6, 75, 100)
	DotData(53) = Array(100, 1, 1, 1, 1, 50, 50, 35, 100, 55, 100, 80, 50, 100, 1, 95)
	DotData(54) = Array(100, 20, 70, 1, 20, 1, 1, 30, 1, 80, 30, 100, 70, 100, 100, 80, 100, 60, 70, 50, 30, 50, 1, 60)
	DotData(55) = Array(6, 26, 6, 6, 100, 6, 53, 100)
	DotData(56) = Array(100, 30, 100, 20, 70, 1, 30, 1, 1, 20, 1, 30, 100, 70, 100, 80, 70, 100, 30, 100, 1, 80, 1, 70, 100, 30)
	DotData(57) = Array(1, 80, 30, 100, 80, 100, 100, 70, 100, 20, 70, 1, 30, 1, 1, 20, 1, 40, 30, 50, 70, 50, 100, 40)
	DotData(79) = Array(45.9999,9,63.9181,8.4192,77.272,10.6847,82.9999,22,96.2945,48.2625,74.768,81.3012,54.9999,86,32.0808,91.4477,10.0894,69.101,15.9999,45,21.5835,22.2322,29.6815,20.2174,45.9999,9,48.9999,19,32.2322,31.0691,21.2245,39.5137,28.9999,67,59.8162,88.8911,81.6997,60.7469,72.9999,26,66.0566,21.2553,61.8256,18.8626,48.9999,19)
	DotData(76) = Array(36.9999,77,51.9628,77.1729,60.12,71.9557,72.9999,71,74.7325,73.7401,75.455,73.5108,75.9999,78,66.7703,86.098,46.1487,88.175,30.9999,87,30.3333,86,29.6665,84.9998,28.9999,84,25.4086,78.9341,27.9511,70.7807,27.9999,63,28.6665,45.0018,29.3333,26.9981,29.9999,9,32.235,8.0077,32.8591,7.5191,35.9999,7,36.9998,7.6665,38,8.3333,38.9999,9,41.4547,29.1112,37.755,56.3367,36.9999,77)
	DotData(86) = Array(27.9999,8,29.6664,8.3332,31.3334,8.6666,32.9999,9,37.3328,23.9985,41.667,39.0014,45.9999,54,47.7791,59.0391,46.8741,64.371,50.9999,67,52.2686,57.4296,69.0215,12.0699,75,9,76.2879,8.2145,77.4851,8.1402,79.9999,8,81.3654,10.5599,81.8783,11.4986,81.9999,16,73.4381,23.8316,64.9999,56.7227,59.9999,69,56.7271,77.0363,57.7136,86.214,46.9999,87,40.4323,76.1471,22.9803,23.8171,25,10,25.9998,9.3333,27,8.6665,27.9999,8)
	DotData(68) = Array(26.9999,2,50.9768,2.6929,92.2812,31.5136,78.9999,63,73.2446,76.6443,55.3819,87.391,35.9999,81,32.5037,79.8471,24.458,78.94,22.9999,76,22.9999,66.0009,22.9999,55.9989,22.9999,46,27.8806,37.0149,18.3765,17.2066,23.9999,7,24.9307,4.3191,25.6945,4.05,26.9999,2,32.9999,15,37.8489,30.1114,33.139,52.8678,32.9999,69,34.3331,69.9998,35.6667,71,36.9999,72,47.2517,72.4768,53.7032,71.3397,61.9999,69,81.9931,40.5378,60.1951,22.295,32.9999,15)
	DotData(15883) = Array(73.9999,6,77.9134,6.7076,81.4513,7.7678,82.9999,11,84.3318,14.7139,82.9101,28.7949,82.9999,36,83.1996,52.0179,85.16,73.4579,81.9999,88,79.5362,89.5329,79.7243,90.2913,75.9999,91,74.2064,82.366,69.0111,81.8282,64.9999,76,66.3331,76,67.6667,76,68.9999,76,70.9927,77.4707,73.6494,77.8116,76.9999,78,78.4047,73.8943,78.0742,66.6562,77.9999,61,83.6494,52.2813,75.3178,13.591,73.9999,6,27.9999,12,36.5875,11.6456,41.8333,13.7264,41.9999,22,40.3965,22.5576,39.5963,22.8004,36.9999,23,33.5356,19.4039,29.7513,17.4548,27.9999,12,39.9999,34,30.0572,34.2798,22.5554,37.7905,12.9999,36,12.6666,35,12.3332,33.9998,11.9999,33,26.1998,32.0241,42.7713,22.4775,54.9999,27,54.3692,28.754,54.8071,27.9714,53.9999,29,51.4162,30.9099,46.5323,31.0172,43.9999,33,50.7433,37.7499,42.8145,48.9461,40.9999,55,46.2943,58.3804,51.0354,64.3078,51.9999,72,51.1957,73.021,51.645,72.274,50.9999,74,50,74,48.9998,74,47.9999,74,44.6669,69.667,41.3329,65.3329,37.9999,61,37.6666,61,37.3332,61,36.9999,61,36.9999,61.3332,36.9999,61.6667,36.9999,62,33.3336,65.9996,29.6662,70.0003,25.9999,74,20.8992,77.7937,15.093,80.2029,8.9999,83,8.9999,82.3333,8.9999,81.6665,8.9999,81,17.5531,76.0962,31.2644,66.291,32.9999,55,28.5983,52.2125,22.1221,45.8459,20.9999,40,21.3332,40,21.6666,40,21.9999,40,24.272,42.3248,34.4441,50.0717,35.9999,49,38.5339,45.7617,39.6687,39.122,39.9999,34,59.9999,25,64.1142,25.2747,64.761,26.1384,67.9999,27,67.9999,28.9998,67.9999,31.0001,67.9999,33,63.9614,39.6027,67.4411,57.3111,64.9999,64,64.3333,64,63.6665,64,62.9999,64,62.9999,63.3333,62.9999,62.6665,62.9999,62,58.7562,55.0532,63.5271,33.1475,59.9999,25)
	DotData(80) = Array(37.9999,61,37.9999,70.3323,37.9999,79.6676,37.9999,89,35.5625,90.297,35.1672,90.7496,30.9999,91,26.7436,82.6999,28.8606,63.7873,28.9999,52,29.9998,39.6678,31,27.332,31.9999,15,33.9997,14.6666,36.0001,14.3332,37.9999,14,53.2504,4.6792,77.8893,24.6734,72.9999,42,69.1131,55.7739,56.8159,61.784,37.9999,61,39.9999,23,39.6666,32.3323,39.3332,41.6676,38.9999,51,48.6364,51.8924,58.9518,50.4888,61.9999,42,63.7308,39.676,63.9681,37.2739,63.9999,33,57.4466,25.6531,54.9042,22.6702,39.9999,23)
	DotData(89) = Array(22.9999,9,24.6664,9.3332,26.3334,9.6667,27.9999,10,30.7444,15.7097,48.6494,47.258,52.9999,49,54.8517,36.0172,64.1003,17.1817,71.9999,9,73.641,9.6231,73.0482,9.2978,73.9999,10,78.0128,11.4825,78.0481,12.5275,77.9999,18,66.6545,31.201,60.3765,58.4757,50.9999,74,47.172,80.3377,49.163,86.5549,38.9999,87,38,85.6667,36.9998,84.3331,35.9999,83,39.6662,75.334,43.3336,67.6658,46.9999,60,42.9626,56.9521,21.2224,20.6492,19.9999,15,19.9387,14.1764,19.737,12.8051,20.9999,11,21.6665,10.3333,22.3333,9.6665,22.9999,9)
	DotData(85) = Array(25.9999,10,28.5719,11.3742,28.7558,11.1054,29.9999,14,35.454,24.1899,28.2559,43.6324,30.9999,56,33.2526,66.153,37.0839,69.9659,42.9999,76,46.9995,76.3332,51.0003,76.6666,54.9999,77,76.1264,57.2784,65.9261,44.1635,71.9999,12,74.235,11.0077,74.8591,10.5191,77.9999,10,82.4715,17.2537,81.1735,33.1879,80.9999,45,73.3806,59.5706,79.3753,75.7049,62.9999,84,58.3931,86.3336,50.2188,86.4069,44.9999,85,37.9333,83.0949,31.7061,81.9954,27.9999,77,18.4366,64.1098,18.1418,28.7395,22.9999,11,24.998,10.6036,24.9781,10.6616,25.9999,10)
	DotData(83) = Array(51.9999,8,62.1804,7.7812,67.9337,9.97,75,12,75.3332,13.9998,75.6666,16.0001,75.9999,18,75,18.3332,73.9998,18.6667,72.9999,19,61.7165,23.9368,40.7955,9.1162,38.9999,32,52.6651,39.7278,90.9915,37.144,78.9999,66,73.0743,80.2592,31.0839,93.5272,22.9999,72,22.0373,70.4209,21.9909,67.9352,21.9999,65,24.4374,63.7028,24.8327,63.2502,28.9999,63,36.7613,75.8312,47.7708,74.0665,62.9999,70,66.2949,65.5413,69.3432,64.628,69.9999,57,67.282,53.5364,66.6163,50.7243,61.9999,49,52.3598,42.6192,34.6279,48.5444,29.9999,37,18.577,21.3489,44.2309,12.7827,51.9999,8)
	DotData(71) = Array(46.9999,18,59.051,17.7531,67.2173,19.7858,75.9999,22,75.9999,23.9998,75.9999,26.0001,75.9999,28,76.9625,29.5789,77.0089,32.0647,76.9999,35,76.3333,35,75.6665,35,75,35,72.677,25.2899,56.2233,12.7333,42.9999,22,13.0545,30.105,21.7002,89.2512,64.9999,78,66.3331,76.3334,67.6667,74.6664,68.9999,73,68.6666,67.6671,68.3332,62.3327,67.9999,57,64.6811,55.5114,61.5252,54.9515,55.9999,55,55.9999,54.3333,55.9999,53.6665,55.9999,53,66.3322,53,76.6676,53,86.9999,53,86.9999,53.3332,86.9999,53.6667,86.9999,54,86.6666,54,86.3332,54,85.9999,54,84.4136,54.8884,79.5736,55.0762,77.9999,56,75.8226,59.7055,76.8589,69.4885,76.9999,75,53.3629,90.4995,2.7447,78.4096,16.9999,40,23.4542,22.6095,32.8401,25.8259,46.9999,18)
	'http://www.dc9.cn/ SIPO，ASP无组件验证码 sipo1209@gmail.com QQ17862153
	DotData(74) = Array(63.9999,16,64.5076,36.3185,63.9001,76.0158,52.9999,84,40.0418,93.871,18.8161,73.3038,21.9999,59,23.8743,58.3738,25.9175,58.0678,28.9999,58,33.3584,66.7231,35.4308,69.0521,41.9999,75,44.333,75,46.6668,75,48.9999,75,48.9999,74.3334,48.9999,73.6665,48.9999,73,56.5104,62.3346,53.1217,33.8014,52.9999,17,48.41,17.1664,40.9172,17.7576,37.9999,16,33.9871,14.5174,33.9518,13.4724,33.9999,8,34.9998,7.6666,36,7.3332,36.9999,7,39.7433,5.294,75.5131,5.2024,77.9999,7,80.0244,8.4633,79.7933,10.5283,79.9999,14,77.2598,15.7325,77.489,16.455,72.9999,17,71.105,15.849,67.341,15.9197,63.9999,16)
	DotData(78) = Array(19.9999,8,21.9997,8.3332,24.0001,8.6666,25.9999,9,35.0868,24.5012,61.2034,62.2072,75.9999,70,75.9999,61.3341,75.9999,52.6658,75.9999,44,74.6667,33.001,73.3331,21.9988,71.9999,11,74.5599,9.6345,75.4985,9.1216,79.9999,9,88.7868,24.0613,88.2169,63.2779,83.9999,84,81.7648,84.9922,81.1408,85.4808,77.9999,86,60.6741,69.8191,40.4222,51.6208,28.9999,30,28.6666,30,28.3332,30,27.9999,30,27.9999,31.3331,27.9999,32.6667,27.9999,34,20.7458,47.9427,37.632,85.4178,19.9999,88,19,87.3334,17.9998,86.6665,16.9999,86,15.3918,71.1446,14.2388,24.3459,16.9999,10,17.9998,9.3333,19,8.6665,19.9999,8)
	DotData(67) = Array(71.9999,4,73.6664,4.3332,75.3334,4.6667,76.9999,5,76.9999,5.6665,76.9999,6.3334,76.9999,7,79.0824,10.2442,79.0952,15.6094,78.9999,21,76.0652,22.865,76.0911,23.6704,70.9999,24,68.2559,18.7966,67.4936,15.9051,58.9999,16,45.5572,26.433,30.388,39.7546,35.9999,63,40.7478,66.6703,42.2995,69.7766,50.9999,70,54.1359,67.8421,71.7951,59.8092,71.9999,60,75.8212,61.4102,75.8429,61.9191,75.9999,67,75.3056,67.9395,75.6322,67.3903,75,69,68.1001,71.7188,62.9298,77.6621,54.9999,80,30.2846,87.2864,19.0248,59.421,25.9999,40,30.387,27.785,41.4382,12.0973,52.9999,7,60.1034,3.8681,67.7463,7.7669,71.9999,4)
	DotData(73) = Array(53.9999,22,53.6666,38.9982,53.3332,56.0016,52.9999,73,63.3196,72.3667,72.5279,70.9107,72.9999,81,71.6667,81.9998,70.3331,83,68.9999,84,54.9104,84.0337,39.6243,86.8666,28.9999,85,24.406,74.8936,33.439,75.277,42.9999,74,42.9999,67.3339,42.9999,60.666,42.9999,54,39.3264,47.7461,42.9352,28.4857,43.9999,21,37.3453,21.2568,31.0837,22.0646,27.9999,20,25.6947,17.2803,26.2347,15.7743,26.9999,12,27.6665,12,28.3333,12,28.9999,12,36.9544,7.3318,67.1381,12.3804,73.9999,14,74.2415,17.2206,74.9164,17.1667,73.9999,19,72.1364,23.8742,67.7523,23.3978,60.9999,23,59.4209,22.0373,56.9352,21.9909,53.9999,22)
	DotData(12089) = Array(51.9999,70,51.6666,73.9995,51.3332,78.0003,50.9999,82,66.0994,81.7565,81.27,79.8435,91.9999,85,91.3692,86.754,91.8071,85.9714,90.9999,87,89.927,87.7516,90.0543,87.6013,87.9999,88,83.891,85.4516,75.6811,88.1575,69.9999,87,60.5394,85.0723,41.4198,86.5052,30.9999,89,26.5173,90.0732,16.0209,93.0477,10.9999,90,8.3157,89.1087,8.4106,88.9716,6.9999,87,19.9986,85.6667,33.0012,84.3331,45.9999,83,45.9999,79.3336,45.9999,75.6662,45.9999,72,39.1736,72.4297,34.99,73.8531,30.9999,70,35.9994,69,41.0004,67.9998,45.9999,67,45.9999,63.6669,45.9999,60.3329,45.9999,57,45.3333,57,44.6665,57,43.9999,57,39.2385,60.0176,31.5213,56.5169,26.9999,58,23.9139,59.0122,18.6513,67.1084,11.9999,69,11.9999,68.3334,11.9999,67.6665,11.9999,67,16.1428,63.0815,22.0279,57.8489,25,53,25.9998,50.0002,27,46.9996,27.9999,44,30.4403,46.5185,32.6595,47.1643,33.9999,51,31.9942,52.3966,31.9343,52.346,30.9999,55,32.3331,55,33.6667,55,34.9999,55,37.319,53.3707,41.1914,53.0811,44.9999,53,45.599,51.2754,45.8903,49.8478,45.9999,47,44.6113,44.8642,44.5129,40.9355,42.9999,39,41.1126,40.3902,36.0496,42.0967,35.9999,42,31.6009,36.9691,32.6932,29.0094,29.9999,22,28.8655,19.0473,26.169,17.3427,25,14,29.3574,12.5116,37.0715,12.9246,42.9999,13,46.6275,10.6965,64.8215,8.2415,69.9999,8,73.1512,10.8728,76.7414,11.0113,77.9999,16,72.2674,20.909,73.4521,35.2481,66.9999,39,64.0853,36.2946,49.8956,37.1211,45.9999,40,50.1083,40.1995,50.5686,40.1072,51.9999,43,52.4268,43.6501,51.6673,48.9106,52.9999,51,56.5266,48.442,63.1544,47.9928,68.9999,48,69.8674,49.1913,69.8224,49.1246,70.9999,50,70.9999,50.3332,70.9999,50.6666,70.9999,51,70.3333,51,69.6665,51,68.9999,51,65.2337,53.9092,57.2345,54.9024,51.9999,56,51.9999,59.3329,51.9999,62.6669,51.9999,66,52.3332,66,52.6666,66,52.9999,66,56.0968,63.8329,61.7212,63.903,66.9999,64,67.3332,64.6665,67.6666,65.3333,67.9999,66,68.2512,66.9678,68.7839,66.3792,67.9999,67,64.7694,69.5268,57.4497,69.965,51.9999,70,35.9999,28,36.3332,30.6663,36.6666,33.3336,36.9999,36,46.5426,36.0791,58.3228,34.8378,65.9999,32,66.4103,25.8016,67.8823,22.1088,67.9999,15,66.743,14.3496,65.9292,13.9105,64.9999,13,54.9233,14.8631,43.9128,16.6455,32.9999,17,32.9999,17.3332,32.9999,17.6667,32.9999,18,34.7625,20.6264,34.0691,23.3331,36.9999,25,41.4856,21.9631,51.3802,21.1207,57.9999,21,58.9614,22.766,59.2571,22.6804,59.9999,25,51.831,25.9051,45.4793,27.8786,35.9999,28)
	DotData(84) = Array(56.9999,23,57.1909,41.7537,64.1799,68.9633,57.9999,85,56.0001,85.3332,53.9997,85.6667,51.9999,86,51.0604,85.3056,51.6096,85.6322,50,85,50,84.3333,50,83.6665,50,83,49,63.0019,47.9998,42.998,46.9999,23,46.6666,23,46.3332,23,45.9999,23,41.7073,25.4316,29.7486,22.6049,23.9999,22,23.6666,21.3333,23.3332,20.6665,22.9999,20,21.2126,17.3007,21.6592,17.1565,21.9999,14,25.2638,12.2683,29.2529,11.8851,34.9999,12,51.6649,12.6665,68.3349,13.3333,84.9999,14,85.9998,15.3331,87,16.6667,87.9999,18,87.3768,19.641,87.7021,19.0482,86.9999,20,84.7089,25.7263,74.8757,24.4023,66.9999,24,64.9377,22.7515,60.5478,22.9001,56.9999,23)
	DotData(70) = Array(38.9999,19,38.9999,25.9992,38.9999,33.0006,38.9999,40,39.3332,40,39.6666,40,39.9999,40,47.3178,35.1233,59.7337,37.7426,69.9999,38,71.7325,40.7401,72.4549,40.5108,72.9999,45,71.2111,46.5686,71.6519,46.9038,68.9999,48,61.8778,51.4628,46.2643,45.8352,38.9999,51,34.6453,58.8118,38.6332,79.1483,36.9999,88,35.1255,88.6261,33.0824,88.9321,29.9999,89,24.1845,77.5791,27.8619,50.3751,27.9999,35,31.4325,28.9265,27.8434,16.5358,28.9999,10,31.4374,8.7028,31.8327,8.2502,35.9999,8,41.5151,10.1975,87.3588,1.5276,76.9999,18,76.0604,18.6943,76.6096,18.3677,75,19,72.0694,20.2432,65.9749,17.8126,61.9999,17,53.831,15.3298,45.7561,18.2258,38.9999,19)
	DotData(66) = Array(37.9999,5,58.2773,4.4844,62.766,10.2632,71.9999,20,73.6251,37.4202,68.5071,36.932,63.9999,48,75.2834,50.991,82.144,63.5204,72.9999,73,64.3543,81.9627,49.2563,84.6557,30.9999,84,25.904,75.7006,27.8303,56.1049,27.9999,43,28.9998,31.0011,30,18.9987,30.9999,7,33.7952,6.7155,36.2301,6.2393,37.9999,5,39.9999,15,39.6666,24.3323,39.3332,33.6675,38.9999,43,54.0706,43.6534,63.0841,39.4991,62.9999,25,56.6105,18.5733,53.6019,14.8134,39.9999,15,37.9999,74,49.3179,74.1524,63.843,71.1994,66.9999,63,67.3332,63,67.6666,63,67.9999,63,67.6666,62.6667,67.3332,62.3332,66.9999,62,61.1381,54.2008,49.7284,55.1329,37.9999,53,37.9999,59.9992,37.9999,67.0007,37.9999,74)
	DotData(90) = Array(22.9999,10,42.0086,9.8574,65.9682,8.6015,81.9999,12,82.3332,13.6665,82.6666,15.3334,82.9999,17,63.9226,35.4193,44.7956,52.1961,31.9999,77,38.9992,77,46.0006,77,52.9999,77,59.0103,73.1753,79.8464,75.4587,81.9999,80,83.1705,82.7599,82.8743,83.4268,79.9999,85,75.1775,87.7979,21.7972,86.2825,19.9999,85,19.6666,85,19.3332,85,18.9999,85,18.7584,81.7793,18.0834,81.8332,18.9999,80,20.8119,68.6866,55.0229,26.0304,63.9999,20,63.9999,19.6666,63.9999,19.3332,63.9999,19,63,19,61.9998,19,60.9999,19,53.2462,23.424,31.2811,20.1472,20.9999,20,20.6666,19.6666,20.3332,19.3332,19.9999,19,19.7584,15.7793,19.0834,15.8332,19.9999,14,21.0684,11.3158,21.4341,11.7957,22.9999,10)
	DotData(72) = Array(29.9999,48,44.4451,46.9927,55.7736,42.2725,71.9999,42,71.4596,29.214,69.5933,10.255,79.9999,8,80.9395,8.6943,80.3903,8.3677,81.9999,9,82.6854,11.0503,82.9793,13.6749,82.9999,17,80.45,21.5277,82.096,32.5932,81.9999,39,81.7928,52.8249,81.7617,68.9876,81.9999,83,79.0652,84.865,79.0911,85.6704,73.9999,86,73.6666,85,73.3332,83.9998,72.9999,83,69.2173,76.3886,70.8649,61.3021,70.9999,52,55.2955,51.8622,44.0036,56.646,29.9999,58,30.5558,69.3513,32.104,87.3601,20.9999,88,18.7393,83.6845,18.8573,76.1087,18.9999,69,22.8424,62.0395,19.958,19.867,21.9999,9,23.8743,8.3738,25.9175,8.0678,28.9999,8,29.3332,8.9998,29.6666,10,29.9999,11,32.7233,15.7391,30.1556,39.3588,29.9999,48)
	DotData(77) = Array(30.9999,8,34.774,8.9951,35.6621,9.5069,36.9999,13,43.979,23.2245,44.2811,50.3585,48.9999,63,49.6943,62.0604,49.3677,62.6096,50,61,57.1334,50.3572,59.7006,8.9146,72.9999,8,81.4426,21.6654,79.9861,44.729,84.9999,62,86.5597,67.3727,91.9999,78.0087,89.9999,83,88.5897,86.8212,88.0808,86.8429,82.9999,87,74.7076,73.4379,75.7592,52.862,69.9999,36,69.6666,36,69.3332,36,68.9999,36,67.8084,47.0856,56.5005,84.3752,45.9999,85,40.3822,75.4754,37.1766,60.7671,33.9999,49,32.6173,43.8784,33.5577,39.5218,30.9999,36,27.6669,51.9983,24.3329,68.0016,20.9999,84,18.5625,85.297,18.1672,85.7496,13.9999,86,11.8386,82.61,10.9628,81.3106,10.9999,75,21.7893,55.574,16.4122,23.3222,30.9999,8)
	DotData(75) = Array(33.9999,7,35.6664,7,37.3334,7,38.9999,7,41.5454,11.8831,41.1429,21.2035,40.9999,29,39.5513,31.411,39.8778,37.029,39.9999,41,40.3332,41,40.6666,41,40.9999,41,45.4471,34.2721,66.5522,8.6557,75,8,75.9998,9.3331,77,10.6667,77.9999,12,72.4273,25.2494,58.3676,40.3192,46.9999,48,47.3332,48.3332,47.6666,48.6667,47.9999,49,53.8975,57.3138,79.5885,74.0798,79.9999,82,77.5625,83.297,77.1672,83.7496,72.9999,84,68.2129,80.422,62.6664,78.5466,57.9999,75,52.0005,69.0005,45.9993,62.9993,39.9999,57,39.9999,66.3323,39.9999,75.6675,39.9999,85,39.3056,85.9395,39.6322,85.3903,38.9999,87,37.0001,87,34.9997,87,32.9999,87,27.1652,77.4632,29.8338,53.6205,29.9999,39,33.5956,32.5629,30.6601,15.2882,31.9999,8,33.641,7.3768,33.0482,7.7021,33.9999,7)
	DotData(82) = Array(27.9999,11,54.5626,10.2461,75.4785,19.0861,75,45,72.6668,48.3329,70.333,51.6669,67.9999,55,64.281,57.7515,58.5238,58.0666,55.9999,62,64.5462,64.2939,77.4474,73.6906,78.9999,83,78.3333,83.9998,77.6665,85,76.9999,86,73.9452,86.0471,74.2702,86.7018,72.9999,86,58.8157,81.1074,52.3813,66.0979,33.9999,64,33.9999,64.6665,33.9999,65.3333,33.9999,66,36.1773,69.7055,35.141,79.4885,34.9999,85,32.4333,86.4942,31.8396,87.3328,27.9999,88,25.5709,84.1703,24.8476,81.1688,25,74,23.0246,70.6024,23.9231,62.1242,23.9999,57,24.1678,45.7989,24.0697,17.7959,27.9999,11,35.9999,22,35.6666,31.9989,35.3332,42.0009,34.9999,52,48.3983,52.5596,60.5774,51.3703,63.9999,42,64.7053,40.8536,64.7391,40.269,64.9999,38,57.662,28.4937,53.3897,21.7896,35.9999,22)
	DotData(88) = Array(23.9999,8,34.8399,11.042,46.1974,28.7296,50.9999,38,51.6665,38,52.3333,38,52.9999,38,59.6419,26.7089,68.0983,17.0412,76.9999,8,78.641,8.6231,78.0482,8.2978,78.9999,9,81.8777,10.2563,80.7073,9.1689,81.9999,12,82.6481,13.0106,82.581,13.0318,82.9999,15,77.0985,20.7653,60.8929,40.8633,58.9999,49,63.791,52.9293,81.2106,77.2487,81.9999,84,79.0652,85.865,79.0911,86.6704,73.9999,87,67.1359,76.6873,58.6414,66.8324,51.9999,56,51.6666,56.3332,51.3332,56.6666,50.9999,57,40.8047,64.2409,34.182,85.2711,18.9999,86,18.6666,85,18.3332,83.9998,17.9999,83,17.2946,81.8536,17.2608,81.269,16.9999,79,24.2646,71.9989,41.4587,55.1327,44.9999,46,44.6666,46,44.3332,46,43.9999,46,39.1686,33.305,21.4315,25.8532,19.9999,11,21.3331,10,22.6667,8.9998,23.9999,8)
	DotData(69) = Array(36.9999,20,37.284,27.5957,37.0012,34.4745,36.9999,44,48.6654,43,60.3344,41.9998,71.9999,41,80.5869,55.4215,50.0986,54.4631,35.9999,54,36.3332,62.6658,36.6666,71.3341,36.9999,80,47.3322,80,57.6676,80,67.9999,80,68.9395,79.8873,70.8221,80.1732,72.9999,81,73.3332,82.9998,73.6666,85.0001,73.9999,87,70.7552,89.0697,69.9658,89.955,63.9999,90,59.6264,92.5605,46.5145,91.0772,41.9999,90,35.8641,88.5359,30.6567,89.5393,27.9999,85,22.3756,75.3902,24.9494,52.025,26.9999,43,29.4304,32.3025,25.3841,19.2117,26.9999,11,29.4374,9.7028,29.8327,9.2502,33.9999,9,37.7916,10.7787,45.5297,7.918,50,7,59.8547,4.976,70.754,8.519,76.9999,11,77.6261,12.8744,77.9321,14.9175,77.9999,18,75.44,19.3654,74.5013,19.8783,69.9999,20,62.0934,14.9645,45.732,19.1252,36.9999,20)
	DotData(81) = Array(43.9999,1,72.7507,0.3488,81.0931,8.9525,86.9999,31,88.5752,36.8798,90.4303,46.8693,87.9999,54,85.2002,62.2143,77.0936,65.8972,73.9999,73,79.8256,76.8223,90.5964,85.0334,91.9999,93,89.0652,94.865,89.0911,95.6704,83.9999,96,77.3339,90.0005,70.6659,83.9993,63.9999,78,57.3233,78.7352,50.9526,82.7317,42.9999,81,23.1694,76.6816,0.5361,54.7604,11.9999,27,19.1544,9.6749,29.7711,9.9376,43.9999,1,47.9999,11,37.0963,17.2473,27.056,18.1937,21.9999,30,12.1183,53.0741,30.7588,70.607,52.9999,70,52.9999,69.6666,52.9999,69.3332,52.9999,69,48.4944,66.0269,44.9273,61.5918,43.9999,55,44.6665,54,45.3333,52.9998,45.9999,52,50.1126,51.793,48.8832,51.731,52.9999,52,55.8556,57.4382,60.6404,60.1768,63.9999,65,71.2821,61.3495,72.9878,56.6564,77.9999,51,78.5625,28.1852,73.2984,22.1402,61.9999,12,57.3337,11.6667,52.6661,11.3332,47.9999,11)
	DotData(19521) = Array(65.9999,30,56.5118,29.3371,47.9417,30.898,36.9999,31,36.172,34.2451,37.49,32.5278,34.9999,34,34.9999,33.3333,34.9999,32.6665,34.9999,32,32.3962,29.0824,25.2919,14.6141,23.9999,11,27.9995,11,32.0003,11,35.9999,11,39.834,8.5444,60.9933,5.5227,65.9999,5,69.0578,7.7546,72.3687,7.5618,73.9999,12,69.7457,15.0236,66.8592,23.9157,65.9999,30,33.9999,19,42.3482,18.8416,47.754,17.0884,56.9999,17,57.9087,18.0836,56.9219,17.0846,57.9999,18,57.9999,18.6665,57.9999,19.3333,57.9999,20,50.334,20.6665,42.6658,21.3333,34.9999,22,34.9999,22.3332,34.9999,22.6667,34.9999,23,35.6486,23.8603,36.3375,26.1443,36.9999,27,38.0729,27.7516,37.9456,27.6013,39.9999,28,44.6258,25.0639,57.0332,25.6392,62.9999,25,63.3332,20.0005,63.6666,14.9994,63.9999,10,63.3333,10,62.6665,10,61.9999,10,58.1743,12.3966,38.8253,13.8808,32.9999,14,32.9999,14.6665,32.9999,15.3333,32.9999,16,33.7516,17.0729,33.6013,16.9456,33.9999,19,31.9999,42,31.6666,47.666,31.3332,53.3338,30.9999,59,31.6665,59,32.3333,59,32.9999,59,37.1464,56.4287,72.0111,52.4844,76.9999,53,76.9999,53.9998,76.9999,55,76.9999,56,76.3333,56,75.6665,56,75,56,67.4916,60.6274,53.0534,56.9413,45.9999,62,46.6665,62,47.3333,62,47.9999,62,50.2795,65.3468,54.7825,68.6717,58.9999,70,60.8985,67.4472,65.4463,66.5552,66.9999,64,68.256,61.934,64.4818,61.5302,66.9999,60,67.9998,60,69,60,69.9999,60,72.5251,62.5972,74.1164,62.3786,75,67,70.857,68.5461,64.9672,70.1103,61.9999,73,62.3332,73,62.6666,73,62.9999,73,69.0329,80.0798,95.461,84.6739,97.9999,89,97.3333,89,96.6665,89,95.9999,89,73.1258,105.356,55.094,64.211,37.9999,62,37.9999,62.3332,37.9999,62.6667,37.9999,63,40.642,67.1094,38.1716,79.1675,37.9999,86,44.568,85.0849,47.0933,81.7413,53.9999,81,53.9999,81.6665,53.9999,82.3333,53.9999,83,45.2409,85.9695,40.5285,95.7491,30.9999,97,30.6666,96,30.3332,94.9998,29.9999,94,28.2606,90.4187,32.0253,87.7626,32.9999,84,32.9999,77.0006,32.9999,69.9992,32.9999,63,31.6667,63,30.3331,63,28.9999,63,26.1964,76.2513,14.7064,90.7737,1.9999,94,1.9999,93.3334,1.9999,92.6665,1.9999,92,17.2129,82.0541,28.2665,60.2234,25,38,25.3332,38,25.6666,38,25.9999,38,25.9999,37.6667,25.9999,37.3332,25.9999,37,42.0068,40.0118,59.5691,30.5318,71.9999,35,71.9999,35.6665,71.9999,36.3333,71.9999,37,58.3875,37.8177,46.8885,41.7479,31.9999,42,60.9999,44,62.3331,44.6665,63.6667,45.3333,64.9999,46,64.0912,47.0836,65.078,46.0846,63.9999,47,58.7897,51.0274,46.5308,51.9816,37.9999,52,37.3333,51,36.6665,49.9998,35.9999,49,42.9255,48.7395,56.013,47.4567,60.9999,44)
	DotData(65) = Array(58.9999,6,62.4385,6.9529,62.6574,6.8801,63.9999,10,67.6662,26.6649,71.3336,43.3349,75,60,75.9867,65.3457,80.4458,72.01,78.9999,76,77.5897,79.8212,77.0808,79.8429,71.9999,80,67.7806,73.0809,67.266,61.7808,62.9999,55,61.712,54.2144,60.5147,54.1402,57.9999,54,51.3339,55.6664,44.6659,57.3334,37.9999,59,32.5533,65.9389,33.727,79.2356,21.9999,80,20.6345,77.44,20.1216,76.5013,19.9999,72,22.2207,68.9152,26.1043,61.4453,26.9999,58,26.6666,56.0001,26.3332,53.9998,25.9999,52,27.9997,50.6667,30.0001,49.3331,31.9999,48,36.8577,41.7464,38.7782,33.7159,42.9999,27,47.5909,19.6967,54.1333,12.9874,58.9999,6,55.9999,26,53.6455,33.24,47.8586,39.7779,43.9999,46,45.6664,46,47.3334,46,48.9999,46,52.9995,45,57.0003,43.9998,60.9999,43,60.9999,42.3333,60.9999,41.6665,60.9999,41,58.9645,37.762,58.3098,31.585,57.9999,27,57.3333,26.6666,56.6665,26.3332,55.9999,26)
	DotData(87) = Array(10.9999,17,13.333,17.3332,15.6668,17.6667,17.9999,18,20.8848,31.0306,26.1337,52.5296,30.9999,66,31.6943,65.0604,31.3677,65.6096,31.9999,64,39.6243,52.8092,38.9012,25.2929,50.9999,19,52.0106,18.3518,52.0318,18.4188,53.9999,18,54.9998,18.6665,56,19.3333,56.9999,20,57.8942,28.85,63.5161,68.8549,67.9999,72,69.6456,59.0699,78.1593,46.8328,81.9999,35,83.7586,29.5814,82.8729,23.1054,85.9999,19,87.6346,16.8538,90.1765,17.102,93.9999,17,94.6665,17.9998,95.3333,19,95.9999,20,92.2272,38.1864,82.6933,55.1065,75.9999,71,72.4638,79.3967,74.3005,87.9257,61.9999,88,61.6666,87,61.3332,85.9998,60.9999,85,56.1023,78.1315,52.2516,51.0807,51.9999,40,51.6666,40,51.3332,40,50.9999,40,50.9999,40.3332,50.9999,40.6667,50.9999,41,42.8827,53.8462,41.1228,69.6496,35.9999,85,33.7648,85.9922,33.1408,86.4808,29.9999,87,29,86.3333,27.9998,85.6665,26.9999,85,23.0094,79.9843,23.9053,71.1535,21.9999,64,19.096,53.0975,7.2732,24.9211,10.9999,17)
	'http://www.dc9.cn/ SIPO，ASP无组件验证码 sipo1209@gmail.com QQ17862153

	Dim vExtent : vExtent = Width / Count
	Margin(0) = Border + vExtent * (Rnd * Offset) / 100 + Margin(1)
	Margin(1) = vExtent * (pNumber + 1) - Border - vExtent * (Rnd * Offset) / 100
	Margin(2) = Border + Height * (Rnd * Offset) / 100
	Margin(3) = Height - Border - Height * (Rnd * Offset) / 100
	
	Dim vStartX, vEndX, vStartY, vEndY
	Dim vWidth, vHeight, vDX, vDY, vDeltaT
	Dim vAngle, vLength
	
	vWidth =Int(Margin(1) - Margin(0))
	vHeight =Int(Margin(3) - Margin(2))
	vStartX = Int((DotData(pIndex)(0)-1) * vWidth / 100)
	vStartY = Int((DotData(pIndex)(1)-1) * vHeight / 100)
	
	Dim i, j
	For i = 1 To UBound(DotData(pIndex), 1)/2
		If DotData(pIndex)(2*i-2) <> 0 And DotData(pIndex)(2*i) <> 0 Then
			vEndX = (DotData(pIndex)(2*i)-1) * vWidth / 100
			vEndY = (DotData(pIndex)(2*i+1)-1) * vHeight / 100
			vDX = vEndX - vStartX
			vDY = vEndY - vStartY
			If vDX = 0 Then
				vAngle = Sgn(vDY) * 3.14/2
			Else
				vAngle = Atn(vDY / vDX)
			End If
			If Sin(vAngle) = 0 Then
				vLength = vDX
			Else
				vLength = vDY / Sin(vAngle)
			End If
			vAngle = vAngle + (Rnd - 0.5) * 2 * Angle * 3.14 * 2 / 100
			vDX = Int(Cos(vAngle) * vLength)
			vDY = Int(Sin(vAngle) * vLength)
			If Abs(vDX) > Abs(vDY) Then vDeltaT = Abs(vDX) Else vDeltaT = Abs(vDY)
			For j = 1 To vDeltaT
				SetDot Margin(0) + vStartX + j * vDX / vDeltaT, Margin(2) + vStartY + j * vDY / vDeltaT
			Next
			vStartX = vStartX + vDX
			vStartY = vStartY + vDY
		End If
	Next
End Sub

Public Sub Output()
	Response.Expires = -9999
	Response.AddHeader "pragma", "no-cache"
	Response.AddHeader "cache-ctrol", "no-cache"
	Response.ContentType = "image/gif"
	Response.BinaryWrite ChrB(Asc("G")) & ChrB(Asc("I")) & ChrB(Asc("F"))
	Response.BinaryWrite ChrB(Asc("8")) & ChrB(Asc("9")) & ChrB(Asc("a"))
	Response.BinaryWrite ChrB(Width Mod 256) & ChrB((Width \ 256) Mod 256)
	Response.BinaryWrite ChrB(Height Mod 256) & ChrB((Height \ 256) Mod 256)
	Response.BinaryWrite ChrB(128) & ChrB(0) & ChrB(0)
	Dim R,G,B,Rf,Gf,Bf
	Randomize
	R=255-Int(Rnd*255/4): G=255-Int(Rnd*255/4): B=255-Int(Rnd*255/4): Rf=255-R: Gf=255-G: Bf=255-B
	Response.BinaryWrite ChrB(R) & ChrB(G) & ChrB(B)
	Response.BinaryWrite ChrB(Rf) & ChrB(Gf) & ChrB(Bf)
	Response.BinaryWrite ChrB(Asc(","))
	Response.BinaryWrite ChrB(0) & ChrB(0) & ChrB(0) & ChrB(0)
	Response.BinaryWrite ChrB(Width Mod 256) & ChrB((Width \ 256) Mod 256)
	Response.BinaryWrite ChrB(Height Mod 256) & ChrB((Height \ 256) Mod 256)
	Response.BinaryWrite ChrB(0) & ChrB(7) & ChrB(255)

	Dim x, y, i : i = 0
	For y = 0 To Height - 1
		For x = 0 To Width - 1
			If Rnd < Noisy / 100 Then
				Response.BinaryWrite ChrB(1-Graph(x, y))
			ElseIf x * (x-Width) = 0 Or y * (y-Height) = 0 Then
				Response.BinaryWrite ChrB(Graph(x, y))
			ElseIf Graph(x-1, y) = 1 Or Graph(x, y) Or Graph(x, y-1) = 1 Then
				Response.BinaryWrite ChrB(1)
			Else
				Response.BinaryWrite ChrB(0)
			End If

			If (y * Width + x + 1) Mod 126 = 0 Then
				Response.BinaryWrite ChrB(128)
				i = i + 1
			End If
			If (y * Width + x + i + 1) Mod 255 = 0 Then
				If (Width*Height - y * Width - x - 1) > 255 Then
					Response.BinaryWrite ChrB(255)
				Else
					Response.BinaryWrite ChrB(Width * Height Mod 255)
				End If
			End If
		Next
	Next
	Response.BinaryWrite ChrB(128) & ChrB(0) & ChrB(129) & ChrB(0) & ChrB(59)
End Sub
End Class
For Each sAction_Plugin_ExportValidCode_Begin in Action_Plugin_ExportValidCode_Begin
	If Not IsEmpty(sAction_Plugin_ExportValidCode_Begin) Then Call Execute(sAction_Plugin_ExportValidCode_Begin)
Next

Dim mCode
Dim code
Set mCode = New Com_GifCode_Class
mCode.Create(GetVerifyNumber)
mCode.Output()
Set mCode = Nothing
For Each sAction_Plugin_ExportValidCode_End in Action_Plugin_ExportValidCode_End 
	If Not IsEmpty(sAction_Plugin_ExportValidCode_End) Then Call Execute(sAction_Plugin_ExportValidCode_End)
Next
%>
