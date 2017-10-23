<!-- #include file="../../Inc/Common.asp"--> <%'추가파일 %>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		
<TITLE>
</TITLE>

<script language="VBscript">



Sub QueryDetailA(ByVal TA)
    Dim iXmlHttp  
    
    On Error Resume Next  

    Set iXmlHttp = CreateObject("Msxml2.XMLHTTP.4.0")	
    
    parent.T000001A2.location.href = "T000002A.asp?TA="& TA
  
End Sub

</script>
<body scroll=no>
<form method=post action="CheckTableBiz.asp" name="form" target=MyBizASP >
<center>
<%
   Dim pRec
   Dim AStr
   Dim BStr
   

    MetaConnString = MakeConnString(GetGlobalInf("gDBServerIP") ,GetGlobalInf("gDBLoginID"),GetGlobalInf("gDBSAPwd"),gDataBase)      
   
   Set pRec  = Server.CreateObject("ADODB.RecordSet")
   pRec.Open "select distinct substring(table_id,1,1) A from dbo.Z_TABLE_LIST order by A ",  MetaConnString
  
   
   
   Do While Not ( pRec.EOF Or pRec.BOF)

      AStr =  AStr & "<td bgcolor=#ffffff align=center><A href=T000002A.asp?TA=" & pRec(0) & " target = T000001A2>" & UCASE(pRec(0)) & "[ID]</A></td>"
      pRec.MoveNext

   Loop   

   pRec.Close
   
   Response.Write "<table border=0 width=98% cellpadding=1 cellspacing=1 bgcolor=#b0c4de>" 
   Response.Write "<tr bgcolor=#cccccc>" & AStr & "</tr>"
   Response.Write "</table>"
    
%>

</center>
</body>


