<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #include file="../../Inc/Common.asp"--> <%'추가파일 %>
<LINK REL="stylesheet" TYPE="Text/css" HREF="../SheetStyle.css">			
<TITLE></TITLE>
<script language = vbs>
Sub dbSave()
    form.submit
End Sub   
</script>
<BODY topmargin="0" leftmargin="0" marginwidth="0" marginheight="0" border=1 scroll=yes>
<br>
<br>
<center>

<%
'select spid,status,hostname from master..sysprocesses where  dbid = (select dbid from master..sysdatabases  where name ='bb')

    Dim pConn
    Dim pRec  
    
	Call LoadBasisGlobalInf()

    MetaConnString   = MakeConnString(GetGlobalInf("gDBServerIP"),GetGlobalInf("gDBLoginID"),GetGlobalInf("gDBSAPwd"),gDataBase    )   

	Set pConn = Server.CreateObject("ADODB.Connection")
	Set pRec = Server.CreateObject("ADODB.RecordSet")
	
	
    pConn.Open MetaConnString
    
    pRec.Open "select table_id from z_table_list where table_id like '" & Request.Form("TA") & "%' " ,MetaConnString
    
    If Err.number = 0 Then
    
       Do While  Not (pRec.EOF Or pRec.BOF)
          pConn.Execute "update z_table_list set kind = '" & Request.Form(pRec("table_id")) & "' where table_id = '" & pRec("table_id") & "'"

          pRec.MoveNext
       Loop
 
    End If
    

    Call OkProcess()
   

    pConn.Close
    
    Set pConn = Nothing

Sub OkProcess()
%>
    <SCRIPT LANGUAGE="VBS">
       Call parent.OkProcess()
    </SCRIPT>
<%
End Sub
%>


</center>
</form>
</BODY>