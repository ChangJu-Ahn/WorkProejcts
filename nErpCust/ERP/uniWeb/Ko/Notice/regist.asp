<%
'**********************************************************************************************
'*  1. Module Name          : 공지사항 등록/수정 
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
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -1999/09/12 : ..........
'**********************************************************************************************

Response.Expires = -1								'☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.


'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>

<!-- #Include file="../inc/IncServer.asp" -->
<!--#Include file="../inc/incServerAdoDb.asp" -->
<!-- #Include file="../inc/UNI2KCMCom.inc" -->	

<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Event.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Variables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../inc/Operation.vbs"></SCRIPT>

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next														'☜: 

Dim strTable, strStatus, strKeyNo, gstrSQL, ErrMsg
Dim strSubject, strWriter, strAuth , strContents, strPasswd, strUsrId
Dim strMode	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
dim strRec


strMode = "" & Request("txtMode")												'☜ : 현재 상태를 받음 

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
	
Case "1"																	'☜: 신규 저장 요청을 받음 

	objRec.Open "select pwd from Z_USR_MAST_REC where  Usr_Id = '" & strUsrId & "'", objConn

	strPasswd = objRec("pwd")

	Call GetRequest
    
	' SQL 문 만들기		
	Call GetSQLForInsert()		

    On Error Resume Next 
	objconn.execute gstrSQL
   
	objRec.Close
	objConn.Close

	set objRec = nothing
	set objConn = nothing	
	
	Call DisplayMsgBox("210030", vbOKOnly, "", "", I_MKSCRIPT)	  '등록되었습니다!		
	
	Response.Write "<Script Language=vbscript>"			& vbCr
	Response.Write "Parent.DbSaveOk "			& vbCr
	Response.Write "</Script>" & vbCr		
	Response.End																	'☜: Process End



Case "2"																		'☜: 수정 저장 요청을 받음 

	gstrSQL = "select * from B_NOTICE where  NoticeNum = " & strKeyNo
	objRec.Open gstrSQL, objConn

	If UCASE(gUsrId) <> UCASE(objRec("usr_id")) Then
		Call DisplayMsgBox("210033", vbOKOnly, "", "", I_MKSCRIPT) '권한이 없습니다!
	Else 

		Err.Clear																		'☜: Protect system from crashing

		Call GetRequest

		Call GetSQLForUpdate()
		
		objconn.Execute gstrSQL	
		
		Call DisplayMsgBox("210031", vbOKOnly, "", "", I_MKSCRIPT) '수정되었습니다!
	
	End If

	objRec.Close
	objConn.Close

	set objRec = nothing
	set objConn = nothing

	Response.Write "<Script Language=vbscript>"			& vbCr
	Response.Write "Parent.DbSaveOk "			& vbCr
	Response.Write "</Script>" & vbCr		
	Response.End																	'☜: Process End
	
Case "3"	    
	
	gstrSQL = "select * from B_NOTICE where  NoticeNum = " & strKeyNo
	objRec.Open gstrSQL, objConn

	strRec = objRec("Usr_id")

	If UCASE(gUsrId) <> UCASE(objRec("usr_id")) Then

		Call DisplayMsgBox("210033", vbOKOnly, "", "", I_MKSCRIPT) '권한이 없습니다!
	Else 
		
		Err.Clear                                                               '☜: Protect system from crashing	
		
		Call GetRequest

		gstrSql = "Delete from B_NOTICE Where NoticeNum=" & strKeyNo 
		
		objconn.Execute gstrSQL	

		Call DisplayMsgBox("210032", vbOKOnly, "", "", I_MKSCRIPT)  '삭제되었습니다!
		
	End If

	objRec.Close
	objConn.Close

	set objRec = nothing
	set objConn = nothing

	Response.Write "<Script Language=vbscript>"			& vbCr
	Response.Write "location.href = ""notice1.asp"""			& vbCr
	Response.Write "</Script>" & vbCr				
	Response.End	'☜: Process End

End Select

'==============================================================================
' 사용자 정의 서버 함수 
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