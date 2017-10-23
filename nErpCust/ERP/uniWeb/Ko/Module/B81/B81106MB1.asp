<% 
'======================================================================================================
'*  1. Module Name          : CIS
'*  2. Function Name        : 문서관리 등록 
'*  3. Program ID           : B81106Mb1.asp
'*  4. Program Name         : B81106Mb1.asp
'*  5. Program Desc         :
'*  6. Modified date(First) : 2005/01/25
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Lee Wol san
'*  9. Modifier (Last)      :
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

%>
<% Option Explicit %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

Call HideStatusWnd									'☜: 모든 작업 완료후 작업진행중 표시창을 Hide

Dim pB1A011											'입력/수정용 ComProxy Dll 사용 변수 
Dim strMode											'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 
Dim strSpread
Dim StrNextKey		' 다음 값 
Dim lgStrPrevKey	' 이전 값 
Dim LngMaxRow		' 현재 그리드의 최대Row
Dim LngRow         
Dim RowData(5)
Dim RowDataPre
Dim lgSelectList
call LoadBasisGlobalInf()

    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    'Multi SpreadSheet
    strSpread = Request("txtSpread")
    lgLngMaxRow       = Request("txtMaxRows")    
   

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)    
                                                             '☜: Query
             Call SubBizQueryMulti()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
             
             Call SubBizSaveMulti()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
        Case "view"  
        
           Call SubBizQueryView()
    End Select

Sub SubBizSave()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub

Sub SubBizDelete()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
End Sub
'-----------------------------------------------------------------------------------------
Sub SubBizQueryMulti()
'-----------------------------------------------------------------------------------------
    Dim strData
    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    DIM arr,arrCnt
    dim i,j
    'BlankchkFlg = False
    
   'On Error Resume Next  
	Err.Clear                                                                  '☜: Clear Error status
 	
		 
	LngRow = 0
	Call SubMakeSQLStatements()
	
	Call SubOpenDB(lgObjConn)   
		adoRec.Open lgStrSQL, lgObjConn,adOpenStatic, adLockReadOnly 
		arrCnt = adoRec.RecordCount 
	if arrCnt > 0 then 		arr=adoRec.GetRows 
	Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

	if arrCnt>0 then
		for i=0 to arrCnt-1
		   LngRow = LngRow + 1
			for j=0 to uBound(arr,1)
			if j=2 or j=3 then
				strData = strData & Chr(11) & UniConvDateDbToCompany(arr(j,i),"")
			else
				strData = strData & Chr(11) & arr(j,i)
			
			end if
				
			next 
			strData =  strData & Chr(11) & LngRow &  Chr(11) & Chr(12) 
		next 
	else
	
	Call DisplayMsgBox("900014", vbOKOnly, "자료", "", I_MKSCRIPT)
	'Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
	
	'Response.End 	
    end if
%>

<Script Language=vbscript>
    Dim LngRow          
    Dim strTemp
    Dim strData
	With parent																	'☜: 화면 처리 ASP 를 지칭함 
	.frm1.vspdData.ReDraw = False
	 strData = "<%=ConvSPChars(strData)%>"
    .ggoSpread.Source = .frm1.vspdData 
  	.ggoSpread.SSShowData strData
	.frm1.vspdData.ReDraw = True
	.DbQueryOk
	End With
</Script>	
<% 

End Sub    

'========================================================================
Sub SubBizQueryView() '상세화면에 내용보여주ㄱ;ㅣ 
'========================================================================
  
 Dim strData
    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    DIM arr,arrCnt
    dim i,j,strFile_no
    'BlankchkFlg = False
    
 '   On Error Resume Next  
	Err.Clear                                                                  '☜: Clear Error status
 	strFile_no=Request("file_no")
	lgStrSQL = "SELECT ins_person, title, insrt_user_id,ins_dt, file_desc  FROM B_CIS_FILE_HEAD "
	lgStrSQL =lgStrSQL & " WHERE FILE_NO=" & strFile_no
	lgStrSQL =lgStrSQL & " ORDER BY FILE_NO DESC"
	
	LngRow = 0
	'Call SubMakeSQLStatements()
	Call SubOpenDB(lgObjConn)   
	Set adoRec = Server.CreateObject("ADODB.Recordset")
		adoRec.Open lgStrSQL, lgObjConn,adOpenStatic, adLockReadOnly 
		arrCnt = adoRec.RecordCount 
		arr=adoRec.GetRows 
	Call SubCloseDB(lgObjConn) 
	  
	if arrCnt>0 then %>
	
		<Script Language=vbscript>
		With parent.frames(1).frm1
		    .ins_person.value	="<%=arr(0,0)%>"
			.title.value		="<%=arr(1,0)%>"
			.insrt_user_id.value="<%=arr(2,0)%>"
			.insrt_dt.value     ="<%=arr(3,0)%>"
			.file_desc.value     ="<%=replace(arr(4,0),chr(13),"<br>")%>" 
	
		End With
	</Script>	
<%
	
	else%>

	
			
    <%end if%>
    
<% 

End Sub    

'========================================================================
'SubBizDelete(삭제)
'========================================================================

Sub SubBizSaveMulti()
	
	Dim strSql
	Dim Temp,file_no,i
	Dim arrFile_id
	
	temp = split(strSpread,chr(12))
	
	Call SubOpenDB(lgObjConn)
	
	lgObjConn.beginTrans()
	
	for i=0 to  uBound(temp)-1
		file_no=split(temp(i),chr(11))(2)
		
		lgStrSQL="SELECT FILE_ID FROM B_CIS_FILE_DETAIL WHERE FILE_NO='"&file_no&"'"
		adoRec.Open lgStrSQL, lgObjConn,adOpenStatic, adLockReadOnly 
		'================
		'FILE DELETE ADD
		'================
		if not adoRec.eof then 'file list가 있으면 file_Id배열에 담은후 B_CIS_FILE_DETAIL delete
			arrFile_id= adoRec.getRows()
			Call FileDelete(arrFile_id)
		end if
		adoRec.Close 
		
		strSql="DELETE FROM B_CIS_FILE_HEAD WHERE FILE_NO="&file_no 
		lgObjConn.execute strSql 
		strSql="DELETE FROM B_CIS_FILE_DETAIL WHERE FILE_NO="&file_no
		lgObjConn.execute strSql
		
	
	next
	If CheckSYSTEMError(Err,True) = True Then                                              
		lgObjConn.rollbacktrans()
		Call SubCloseDB(lgObjConn) 
		Response.End 
	else
		lgObjConn.committrans()
		Call SubCloseDB(lgObjConn) 
		Call DisplayMsgBox("210032", vbOKOnly, "", "", I_MKSCRIPT)  '삭제되었습니다!
    End If
  
       
%>
<Script Language=vbscript>
	With parent																	    '☜: 화면 처리 ASP 를 지칭함 
		CALL .DbSaveOk()
		parent.MyBizASP1.location.reload
	End With
</Script>
<%					

End Sub

%>

<%


'========================================================================
'SubMakeSQLStatements
'========================================================================
Function FileDelete(byVal pArr )
 	on Error Resume Next
	Dim filePath
 	Dim i
	
	filePath=server.MapPath (".")&"\files\"

	for i=0 to uBound(pArr,2)
		Call pfile.fileDelete(replace(filePath & pArr(0,i),"\","/"))   
	next
End Function

'========================================================================
'SubMakeSQLStatements
'========================================================================
Sub SubMakeSQLStatements()
	Dim file_no
	Dim strUsrId
	Dim strtitle
	DIM strIns_person
	Dim FromInsrtDt,ToInsrtDt
	Dim FromUseDt,ToUseDt
	
   
	strtitle		= FilterVar(Request("txtTITLE")&"%", "", "S")
	strIns_person   = FilterVar(Request("txtINS_PERSON")&"%", "", "S")
	
	FromInsrtDt		= Request("txtFromInsrtDt") 
	ToInsrtDt		= Request("txtToInsrtDt")	
	FromUseDt		= Request("txtFromUseDt")	
	ToUseDt			= Request("txtToUseDt")
	
	lgStrSQL =  " SELECT FILE_NO,INS_PERSON,USE_DT,INS_DT,TITLE,FILE_ABBR,INSRT_USER_ID "
	lgStrSQL = lgStrSQL & " FROM B_CIS_FILE_HEAD "
	lgStrSQL = lgStrSQL & " WHERE insrt_dt BETWEEN  '"&uniconvdate(FromInsrtDt)&"' AND '"&uniconvdate(ToInsrtDt)&"'"
	lgStrSQL = lgStrSQL & " AND USE_DT BETWEEN '"&uniconvdate(FromUseDt)&"' AND '"&uniconvdate(ToUseDt)&"'"
	lgStrSQL = lgStrSQL & " AND INS_PERSON LIKE " & strIns_person
	lgStrSQL = lgStrSQL & " AND TITLE LIKE " & strtitle
	lgStrSQL = lgStrSQL & " ORDER BY FILE_NO DESC"



End Sub


%>















<OBJECT RUNAT=server PROGID="ADODB.Recordset" id=adoRec></OBJECT>
<OBJECT RUNAT=server PROGID="PuniFile.CTransfer" id=pfile></OBJECT>