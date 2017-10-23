<%
'**********************************************************************************************
'*  1. Module Name          : �������� ���/���� 
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/04/06
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'*                            -1999/09/12 : ..........
'**********************************************************************************************

Response.Expires = -1								'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True								'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.


'�� : �׻� ���� ���̵� ������ �������� �²���(<)% �� %�첩��(>)�� New Line�� ��ġ�Ͽ� 
'	  ���� ���̵� ������ Ŭ���̾�Ʈ ���̵� ������ ��ġ�� ������ �� �ֵ��� �Ѵ�.
'�� : �Ʒ� HTML ������ ����Ǿ�� �ȵȴ�. 
%>

<!-- #Include file="../inc/IncServer.asp" -->
<!--#Include file="../inc/incServerAdoDb.asp" -->
<!-- #Include file="../inc/UNI2KCMCom.inc" -->	

<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Event.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Operation.vbs"></SCRIPT>

<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next														'��: 

Dim strTable, strStatus, strKeyNo, gstrSQL, ErrMsg
Dim strSubject, strWriter, strAuth , strContents, strPasswd, strUsrId
Dim strMode	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
dim strRec


strMode = "" & Request("txtMode")												'�� : ���� ���¸� ���� 

strKeyNo = "" & Request("n") 
strTable = "B_NOTICE"

strUsrId = Replace(gUsrId, "'", "''")

Dim objConn
Dim objRec

	Set objConn = Server.CreateObject("ADODB.Connection")
	Set objRec = Server.CreateObject("ADODB.RecordSet")
	objConn.ConnectionString = gADODBConnString
	objConn.Open

Select Case CStr(strMode)	
	
Case "1"																	'��: �ű� ���� ��û�� ���� 

	objRec.Open "select pwd from Z_USR_MAST_REC where  Usr_Id = '" & strUsrId & "'", objConn

	strPasswd = objRec("pwd")

	Call GetRequest
    
	' SQL �� �����		
	Call GetSQLForInsert()		

    On Error Resume Next 
	objconn.execute gstrSQL
   
	objRec.Close
	objConn.Close

	set objRec = nothing
	set objConn = nothing	
	
	Call DisplayMsgBox("210030", vbOKOnly, "", "", I_MKSCRIPT)	  '��ϵǾ����ϴ�!		
	
	Response.Write "<Script Language=vbscript>"			& vbCr
	Response.Write "Parent.DbSaveOk "			& vbCr
	Response.Write "</Script>" & vbCr		
	Response.End																	'��: Process End



Case "2"																		'��: ���� ���� ��û�� ���� 

	gstrSQL = "select * from B_NOTICE where  NoticeNum = " & strKeyNo
	objRec.Open gstrSQL, objConn

	If UCASE(gUsrId) <> UCASE(objRec("usr_id")) Then
		Call DisplayMsgBox("210033", vbOKOnly, "", "", I_MKSCRIPT) '������ �����ϴ�!
	Else 

		Err.Clear																		'��: Protect system from crashing

		Call GetRequest

		Call GetSQLForUpdate()
		
		objconn.Execute gstrSQL	
		
		Call DisplayMsgBox("210031", vbOKOnly, "", "", I_MKSCRIPT) '�����Ǿ����ϴ�!
	
	End If

	objRec.Close
	objConn.Close

	set objRec = nothing
	set objConn = nothing

	Response.Write "<Script Language=vbscript>"			& vbCr
	Response.Write "Parent.DbSaveOk "			& vbCr
	Response.Write "</Script>" & vbCr		
	Response.End																	'��: Process End
	
Case "3"	    
	
	gstrSQL = "select * from B_NOTICE where  NoticeNum = " & strKeyNo
	objRec.Open gstrSQL, objConn

	strRec = objRec("Usr_id")

	If UCASE(gUsrId) <> UCASE(objRec("usr_id")) Then

		Call DisplayMsgBox("210033", vbOKOnly, "", "", I_MKSCRIPT) '������ �����ϴ�!
	Else 
		
		Err.Clear                                                               '��: Protect system from crashing	
		
		Call GetRequest

		gstrSql = "Delete from B_NOTICE Where NoticeNum=" & strKeyNo 
		
		objconn.Execute gstrSQL	

		Call DisplayMsgBox("210032", vbOKOnly, "", "", I_MKSCRIPT)  '�����Ǿ����ϴ�!
		
	End If

	objRec.Close
	objConn.Close

	set objRec = nothing
	set objConn = nothing

	Response.Write "<Script Language=vbscript>"			& vbCr
	Response.Write "location.href = ""notice1.asp"""			& vbCr
	Response.Write "</Script>" & vbCr				
	Response.End	'��: Process End

End Select

'==============================================================================
' ����� ���� ���� �Լ� 
'==============================================================================
%>
<!--Script Language="VBScript" RUNAT=Server-->
<%
Sub GetRequest()
	strSubject = Replace(Request("subject"), "'", "''")
	strSubject = Replace(strSubject, "<", "&lt;")
	strSubject = Replace(strSubject, ">", "&gt;")

	strWriter = Replace(Request("Writer"), "'", "''")
	strWriter = Replace(strWriter, "<", "&lt;")
	strWriter = Replace(strWriter, ">", "&gt;")

	strContents = Replace(Request("txtContent"), "'", "''")
	strContents = Replace(strContents, "<", "&lt;")
	strContents = Replace(strContents, ">", "&gt;")

End Sub

Sub GetSQLForInsert()
	Dim strField, strValue
		
    gstrSQL = " INSERT INTO " & strTable
    strField = " ("
    strValue = " Values("
    
    If strWriter <> "" Then
        strField = strField & "Writer"
        strValue = strValue & "'" & strWriter & "'"
    End If
    
    If strUsrId <> "" Then
        strField = strField & ",Usr_Id"
        strValue = strValue & ",'" & strUsrId & "'"
    End If
    
    If strPasswd <> "" Then
        strField = strField & ",Pwd"
        strValue = strValue & ",'" & strPasswd & "'"
    End If
    
    If strSubject <> "" Then
        strField = strField & ",Subject"
        strValue = strValue & ",'" & strSubject & "'"
    End If
    
    If strContents <> "" Then
        strField = strField & ",Contents"
        strValue = strValue & ",'" & strContents & "'"
    End If
	
    strField = strField & ",RegDate" 
    strValue = strValue & ", getdate()"  '& date()  
    
    strField = strField & ")"
    strValue = strValue & ")"
    
    gstrSQL = gstrSQL & strField & strValue

End Sub

Sub GetSQLForUpdate()
	Dim strField
		
    gstrSQL = " UPDATE " & strTable
    strField = " Set "
    
    If strWriter <> "" Then
        strField = strField & "Writer='" & strWriter & "'"
    End If
    
    If strUsrId <> "" Then
        strField = strField & ",Usr_Id='" & strUsrId & "'"
    End If
    
    If strPasswd <> "" Then
        strField = strField & ", Pwd='" & strPasswd & "'"
    End If
    
    If strSubject <> "" Then
        strField = strField & ", Subject='" & strSubject & "'"
    End If
    
    If strContents <> "" Then
        strField = strField & ", Contents='" & strContents & "'"
    End If

	If strContents <> "" Then
        strField = strField & ", RegDate= getdate()" 
    End If
    
    strField = strField & "  WHERE NoticeNum = " & strKeyNo
    
    gstrSQL = gstrSQL & strField

End Sub
%>