<% 
'======================================================================================================
'*  1. Module Name          : CIS
'*  2. Function Name        : 품의서 문서관리(S) 
'*  3. Program ID           : S3322MB1_KO412.asp
'*  4. Program Name         : S3322MB1_KO412.asp
'*  5. Program Desc         :
'*  6. Modified date(First) : 2005/01/25
'*  7. Modified date(Last)  : 2007/07/06
'*  8. Modifier (First)     : Lee Wol san
'*  9. Modifier (Last)      : Lee Ho Jun
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

Call LoadBasisGlobalInf()

    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    'Multi SpreadSheet
    strSpread		  = Request("txtSpread")
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
    DIm arr,arrCnt
    Dim i,j

    'BlankchkFlg = False
   'On Error Resume Next  
	Err.Clear                                                                  '☜: Clear Error status
		 
	LngRow = 0
	Call SubMakeSQLStatements()
	
	Call SubOpenDB(lgObjConn)   
		adoRec.Open lgStrSQL, lgObjConn,adOpenStatic, adLockReadOnly 
		arrCnt = adoRec.RecordCount 
	If arrCnt > 0 then 		arr=adoRec.GetRows
	Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

	If arrCnt>0 then
		For i=0 to arrCnt-1
		   LngRow = LngRow + 1
			For j=0 to uBound(arr,1)
'				If j=2 or j=3 then
'					strData = strData & Chr(11) & UniConvDateDbToCompany(arr(j,i),"")
'				Else
					strData = strData & Chr(11) & arr(j,i)
'				End if
			Next 
			strData =  strData & Chr(11) & LngRow &  Chr(11) & Chr(12) 
		Next 

	Else
		Call DisplayMsgBox("900014", vbOKOnly, "자료", "", I_MKSCRIPT)
    End if
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
Sub SubBizQueryView() '상세화면에 내용보여주기 
'========================================================================
  
 Dim strData
    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    DIM arr,arrCnt
    dim i,j,strProjectCode,strReportNo
    'BlankchkFlg = False
    
 '   On Error Resume Next  
	Err.Clear                                                                  '☜: Clear Error status
 	strProjectCode=FilterVar(Request("txtProjectCode"), "", "S")
 	strReportNo=FilterVar(Request("sReportNo"), "", "S")
 	
	lgStrSQL = "SELECT project_code, report_text  FROM S_Prj_Report_Hdr_ko412 "
	lgStrSQL =lgStrSQL & " WHERE PROJECT_CODE=" & strProjectCode
	lgStrSQL =lgStrSQL & " AND Report_NO=" & strReportNo
	lgStrSQL =lgStrSQL & " ORDER BY PROJECT_CODE DESC"

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
		    .txtprojectcode.value	="<%=arr(0,0)%>"
		    .txtReportText.value	= "<%=replace(arr(1,0),chr(13),"<br>")%>" 
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

	On Error Resume Next
	
	Dim strSql
	Dim Temp,strproject_code,strreport_no, i
	Dim arrFile_id
	
	temp = split(strSpread,chr(12))
	
	Call SubOpenDB(lgObjConn)

	lgObjConn.beginTrans()

	for i = 0 to  UBound(temp)-1
	
		strproject_code = split(temp(i),chr(11))(2)
		strreport_no = split(temp(i),chr(11))(3)
		strreport_no = split(temp(i),chr(11))(3)
		
		lgStrSQL="SELECT REPORT_ID FROM S_PRJ_REPORT_DTL_KO412 WHERE PROJECT_CODE='"&strproject_code&"' and REPORT_NO = '"&strreport_no &"' "

		adoRec.Open lgStrSQL, lgObjConn,adOpenStatic, adLockReadOnly 
		'================
		'FILE DELETE ADD
		'================
		if not adoRec.eof then 'file list가 있으면 file_Id배열에 담은후 B_CIS_FILE_DETAIL delete
			arrFile_id= adoRec.getRows()
			Call FileDelete(arrFile_id)
		end if
		adoRec.Close 
		
		strSql="DELETE FROM S_PRJ_REPORT_HDR_KO412 WHERE PROJECT_CODE='"&strproject_code &"' and REPORT_NO = '"&strreport_no &"'"
		Response.Write strSQL
		lgObjConn.execute strSql

		
		strSql="DELETE FROM S_PRJ_REPORT_DTL_KO412 WHERE PROJECT_CODE='"&strproject_code &"' and REPORT_NO = '"&strreport_no &"'"
		lgObjConn.execute strSql
		Response.Write strSQL		
	
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
'FileDelete
'========================================================================
Function FileDelete(byVal pArr )
 	on Error Resume Next
	
	Dim filePath
 	Dim i
 	DIm WshNetwork'
	DIm NetworkPath, oDrives
	
	set WshNetwork = Server.CreateObject("WScript.Network")

	Set oDrives = WshNetwork.EnumNetworkDrives

	if oDrives.Count-1 >0 then
		WshNetwork.RemoveNetworkDrive"G:"
'		Set WshNetwork = Nothing
'		Set oDrives= nothing
	end if

	WshNetwork.MapNetworkDrive "G:", "\\192.168.10.99\d$\Unierp_File",false,"administrator","nepes123"

	NetworkPath="G:\"

	if err.number <> 0 then
		if oDrives.Count-1 >0 then
			WshNetwork.RemoveNetworkDrive"G:"
'			Set WshNetwork = Nothing
'			Set oDrives= nothing
		end if

		Call DisplayMessageBox("파일저장이 되지 않았습니다.")
		Response.End 
	end if
	
	'filePath=server.MapPath (".")&"\files\"
	filePath = NetworkPath

	For i=0 to uBound(pArr,2)
		Call pfile.fileDelete(replace(filePath & pArr(0,i),"\","/"))   
	Next
	
	Set WshNetwork = Nothing
	Set oDrives= nothing
	
	
End Function

'========================================================================
'SubMakeSQLStatements
'========================================================================
Sub SubMakeSQLStatements()
	Dim file_no
	Dim strUsrId
	Dim strtitle
	DIM strProjectCode
	Dim FromInsrtDt,ToInsrtDt
	Dim FromUseDt,ToUseDt
   
	'strtitle		= FilterVar(Request("txtTITLE")&"%", "", "S")
	strProjectCode   = FilterVar(Request("txtProjectCode")&"%", "''", "S")
	
	lgStrSQL =  " SELECT PROJECT_CODE,REPORT_NO, 'TEST_REPORT_NAME' AS REPORT_NM,ins_user,INS_DT,REPORT_ABBR"	'INSRT_USER_ID "
	lgStrSQL = lgStrSQL & " FROM S_PRJ_REPORT_HDR_KO412 "
	lgStrSQL = lgStrSQL & " WHERE PROJECT_CODE LIKE " & strProjectCode
	'lgStrSQL = lgStrSQL & " AND TITLE LIKE " & strtitle
	lgStrSQL = lgStrSQL & " ORDER BY PROJECT_CODE DESC"
response.write lgStrSQL
End Sub

%>

<OBJECT RUNAT=server PROGID="ADODB.Recordset" id=adoRec></OBJECT>
<OBJECT RUNAT=server PROGID="PuniFile.CTransfer" id=pfile></OBJECT>