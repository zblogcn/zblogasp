<?xml version="1.0" encoding="utf-8" ?> 
<xsl:stylesheet
     version="1.0"
     xmlns:xsl="http://www.w3.org/1999/XSL/Transform">
     <xsl:output method="html" indent="yes" doctype-system="http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd" doctype-public="-//W3C//DTD XHTML 1.0 Transitional//EN" />
     <xsl:template match="rss/channel">
     <html>
          <head>
               <title><xsl:value-of select="title" /></title>
               <style media="all" type="text/css">
                    .ChannelTitle
                    {
                         font-family:  Verdana;
                         font-weight:  bold;
                         text-align:  center;
                         font-size:  1.5em;
                         padding-bottom:  15px;
                    }
                    .ChannelTitle span
                    {		
			 padding:0 2em;
                         font-weight: normal ;
                         font-size:  50%;
                    }
                    .ArticleEntry
                    {
                         padding:  20px;
                    }
                    .ArticleTitle
                    {
                         font-weight:  bold;
                         padding-left:  5px;
                         padding-top:  5px;
                         padding-bottom:  5px;
			 border-bottom:  1px  solid #3165c6 ;
                    }
	
                    .ArticleHeader
                    {
                         padding-left:  5px;
                         padding-top:  5px;
                         padding-bottom:  5px;
			 color:#666666;
			 font-size:  80%;
                    }
                    .ArticleDescription
                    {
                         padding-left:  5px;
                         padding-top:  5px;
                         padding-bottom:  5px;
                         padding-right:  5px;
                    }
                    .ArticleFootor
                    {  
                         padding-top:  5px;
                         padding-bottom: 5px;                    
			 font-size:  90%;
                    }
		    body{
			background-color:white;
			font-size:  100%;
			color:#333333;
		    }
		    a:link {
			color:#3165C6;
			text-decoration:none;
		    }
		    a:hover, a:active {
			color:#027AC6;
			text-decoration:none;
		    }
		    a:visited {
			text-decoration:none;
		    }
		    ul{
			margin:0.3em 0.5em;
			padding:0pt;
		    }
		    ul li{
			background-image:none;
			margin-left:0pt;
			margin-right:0pt;
			padding-left:0pt;
			padding-right:0.5em;
			display:inline;
			list-style-type:none;
	           }
               </style>
          </head>     
          <body>
		<div class="ChannelTitle">
		     <a href="{link}"><xsl:value-of select="title" /></a><span> - <xsl:value-of select="description" /></span>
		</div>
               <xsl:apply-templates select="item" />
          </body>
     </html>
     </xsl:template>
     <xsl:template match="item">
          <div class="ArticleEntry">
               <div class="ArticleTitle">
                    <a href="{link}" target="_blank"><xsl:value-of select="title" /></a>
               </div>
               <div class="ArticleHeader">
                    <xsl:value-of select="pubDate" />  |  <xsl:value-of select="author" />
               </div>
               <div class="ArticleDescription">
                    <xsl:value-of select="description" disable-output-escaping="yes" />
               </div>
		<div class="ArticleFootor">
		<ul>
			<li class="comments"><a href="{comments}" target="_blank">Comments</a></li>
		</ul>
		</div>
          </div>
          <br />
     </xsl:template>
</xsl:stylesheet>
