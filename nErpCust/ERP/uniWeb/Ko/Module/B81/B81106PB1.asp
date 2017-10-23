<% Option Explicit%>
<!--
'**********************************************************************************************
'*  1. Module Name          : CIS
'*  2. Function Name        : 
'*  3. Program ID           : B81106MB
'*  4. Program Name         : 
'*  5. Program Desc         : Create, update and delete notice.
'*  6. Comproxy List        : ADO Query Program
'*  7. Modified date(First) : 2005/01/30
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : lee wol san
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/UNI2KCMCom.inc" -->	
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="B81COMM.ASP" -->

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
'On Error Resume Next														'☜: 

'---------------------------------------Common-----------------------------------------------------------
Call LoadBasisGlobalInf()    


Dim file_no, lgStrSQL, lgStrSQL2
Dim strtitle, strIns_person, strFile_desc, strPasswd, strUsrId
Dim strMode	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim lgObjConn
Dim lgObjComm
Dim lgObjRs

Dim iStrFileInfo
Dim arrTemp,arr, tempFileId
Dim iStrYYYYMM
Dim iPos,iFilePath, iFileId, iNoticenum, i
Dim iCntNoticeNum		'글삭제시 Flag
Dim strFile_abbr
Dim strUse_dt
Dim strInsrt_dt
Dim delTemp


Const C_FileName = 0
Const C_FileId = 1	
Const C_FileSize = 2
Const C_FileCDate = 3

strMode			= Request("txtMode")												'☜ : 현재 상태를 받음 
file_no			= Request("file_no")
strUsrId		= Replace(gUsrId, "'", "''")
strtitle		= FilterVar(Request("txtTITLE"), "", "S")
strIns_person   = FilterVar(Request("INS_PERSON"), "", "S")
strFile_desc	= FilterVar(Request("FILE_DESC"), "", "S")
strFile_abbr	= FilterVar(Request("FILE_ABBR"), "", "S")

strinsrt_dt		= filterVar(UniConvDate(Request("Insrt_dt")),"''","S")
strUse_dt		= filterVar(UniConvDate(Request("Use_dt")),"''","S")
delTemp         = Request("delTemp") '삭제할파일 


Call SubOpenDB(lgObjConn)    
Call SubCreateCommandObject(lgObjComm)

'FILES폴더 없으면 생성 
		Dim fso
		iFilePath = SERVER.MapPath (".") & "\files\" 
		
		Set fso = CreateObject("Scripting.FileSystemObject")   
		If Not fso.FolderExists(iFilePath) then		
		 	   fso.CreateFolder(iFilePath)
		End If	
		Set fso = Nothing 
		
		
Select Case CStr(strMode)
	
    Case CStr(UID_M0001)   
			
		
			'=================================================
			' INsert
			'=================================================
			
		    lgObjConn.beginTrans()
		    
			lgStrSQL = "INSERT INTO B_CIS_FILE_HEAD(INS_PERSON,   TITLE,FILE_ABBR, FILE_DESC,INSRT_USER_ID, INSRT_DT,INS_DT,USE_DT) "			
			lgStrSQL = lgStrSQL & "VALUES(" & strIns_person & "," & strtitle & "," & strFile_abbr & "," & strFile_desc & ",'"&strUsrId&"', "&strinsrt_dt&","&strinsrt_dt&","&strUse_dt&")"

			lgObjConn.execute lgStrSQL

			If Trim(Request("txtFileInf")) <> "" Then '파일첨부		
				Call RunFileAttachForINSERT()
			End If
			
			if lgObjConn.errors.count<>0 then
				lgObjConn.rollBackTrans()
				Call DisplayMsgBox("800407", vbOKOnly, "", "", I_MKSCRIPT)	  '작업실행중 에러입니다.
				Response.End 
			else
				lgObjConn.commitTrans()

				Call DisplayMsgBox("210030", vbOKOnly, "", "", I_MKSCRIPT)	  '등록되었습니다!	
			end if
			'=================================================
				

			'등록후 화면 다시 로딩 
			Response.Write "<Script Language=vbscript>"			& vbCr
			Response.Write "Parent.dbsaveok "			& vbCr
			Response.Write "</Script>" & vbCr		
			Response.End 

    Case CStr(UID_M0002)   
			'=================================================
			' UPDATE
			'=================================================                                                      '☜: 글수정 

			lgStrSQL =  "SELECT insrt_user_Id  FROM B_CIS_FILE_HEAD WHERE FILE_NO ="&file_no
		
			If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists	
				Call DisplayMsgBox("210037", vbInformation, "", "", I_MKSCRIPT)      '☜ : 게시물에 해당하는 자료가 존재하지 않습니다.
			Else
			
				If UCASE(gUsrId) <> UCASE(lgObjRs("insrt_user_Id")) Then
					Call DisplayMsgBox("210033", vbOKOnly, "", "", I_MKSCRIPT) '권한이 없습니다!					
					Response.Write "<Script Language=vbscript>"			& vbCr
					Response.Write "Self.Close()"
					Response.Write "</Script>" & vbCr		
					Response.End															
				Else 
					Err.Clear	
					
				'=========================
				'  Delete 
				'=========================
					
				lgStrSQL  = "DELETE FROM B_CIS_FILE_DETAIL WHERE FILE_NO = '" & file_no & "' AND FILE_iD IN  ('"&Replace(delTemp,",","','")&"')"																					
				 call FileDelete(split(delTemp,","))
				lgObjConn.execute lgStrSQL														
				'=========================
				'  update
				'=========================
																					'☜: Protect system from crashing
					lgStrSQL =" UPDATE B_CIS_FILE_HEAD " 
					lgStrSQL = lgStrSQL & " SET TITLE ="&strtitle&","
					lgStrSQL = lgStrSQL & " INS_PERSON="&strIns_person&","
					lgStrSQL = lgStrSQL & " FILE_ABBR="&strFile_abbr&","
					lgStrSQL = lgStrSQL & " FILE_DESC="&strFile_desc&","
					lgStrSQL = lgStrSQL & "UPDT_USER_ID="&filtervar(strUsrId,"''","S")&","
					lgStrSQL = lgStrSQL & "USE_DT="&strUse_dt&","
					lgStrSQL = lgStrSQL & "INS_DT="&strinsrt_dt&","
					lgStrSQL = lgStrSQL & "UPDT_DT=GETDATE()"
					lgStrSQL = lgStrSQL & " WHERE FILE_NO="&file_no
					lgObjConn.Execute lgStrSQL	


					If Trim(Request("txtFileInf")) <> "" Then '파일첨부 
						
						Call RunFileAttachForUPDATE()
					End If
					Call DisplayMsgBox("210031", vbOKOnly, "", "", I_MKSCRIPT) '수정되었습니다!						
								
				End If				
		    End If
		    
			'수정후 화면 다시 로딩 
			Response.Write "<Script Language=vbscript>"			& vbCr
			Response.Write "Parent.dbsaveok() "			& vbCr
			Response.Write "</Script>" & vbCr		
			Response.End 
			


End Select

Call SubCloseCommandObject(lgObjComm)    
Call SubCloseDB(lgObjConn)      	

Response.End	'☜: Process End

'==============================================================================
' 사용자 정의 서버 함수 
'==============================================================================
Function RunFileAttachForINSERT()
'on error resume next
	dim file_Id

	iStrFileInfo = Request("txtFileInf")	
	iStrFileInfo = FilterVar(iStrFileInfo, "", "S")           '☜For Single Quotation			
	arrTemp = Split(Request("txtFileInf"), gRowSep)
				
	lgStrSQL = ""

					
	For i = 0 To UBound(arrTemp) - 1
	
		arr  = Split(arrTemp(i), gColSep)								
		iPos      = InStrRev(arr(C_FileId), "/", -1)
		iFileId   = Right(arr(C_FileId), len(arr(C_FileId)) - iPos)

		Call FileTransfer(arr(C_FileId), iFilePath, iFileId)
		'arrTemp
		'(0) FILE NAME
		'(1) FILE ID
		'(2) FILE SIZE
		'(3) FILECDATE
		file_Id= split(arr(1),"/")(ubound(split(arr(1),"/")))
		lgStrSQL ="INSERT INTO B_CIS_FILE_DETAIL"
		lgStrSQL = lgStrSQL & " (FILE_NO,SEQ_NO,FILE_NM,FILE_ID,FILE_SIZE,INSRT_DT,INSRT_USER_ID) "
		lgStrSQL = lgStrSQL & " SELECT ( SELECT MAX( FILE_NO) FROM B_CIS_FILE_HEAD),"
		lgStrSQL = lgStrSQL & " "&i&","&filterVar(arr(0),"''","S")&","&filterVar(file_Id,"''","S")&","&arr(2)&",GETDATE(),"&filterVar(strUsrId,"''","S")&""
		
		lgObjConn.execute lgStrSQL
					
	Next

End Function

function kk(str)%>
<script language="javaScript">
	parent.frm1.file_desc.value="<%=str%>"
	</script>
	
<%end function


'=======================================================================
'RunFileAttachForUPDATE
'=======================================================================
Function RunFileAttachForUPDATE()

					
	tempFileId = ""

	iStrFileInfo = Request("txtFileInf")
	iStrFileInfo = FilterVar(iStrFileInfo, "''", "S")           '☜For Single Quotation			
	arrTemp = Split(Request("txtFileInf"), gRowSep)
						
	lgStrSQL = ""
	'(0) FILE NAME
		'(1) FILE ID
		'(2) FILE SIZE
		'(3) FILECDATE					
						
	For i = 0 To UBound(arrTemp) - 1
		
		'Response.Write arrTemp(i)
		arr  = Split(arrTemp(i), gColSep)								
		iPos      = InStrRev(arr(C_FileId), "/", -1)
		iFileId   = Right(arr(C_FileId), len(arr(C_FileId)) - iPos)

							
		lgStrSQL2  = "SELECT 1 FROM B_CIS_FILE_DETAIL WHERE '" & iFileId & "' IN (SELECT FiLE_ID FROM B_CIS_FILE_DETAIL WHERE FILE_no = '" & file_no & "') " 			
		
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL2,"X","X") = True Then                    '기존파일 
				'기존파일은 신규로 해줄 작업이 없다.							
				
		Else	'신규 파일은 파일복사, INSERT 
				Call FileTransfer(arr(C_FileId), iFilePath, iFileId)
							
					lgStrSQL ="INSERT INTO B_CIS_FILE_DETAIL"
					lgStrSQL = lgStrSQL & " (FILE_NO,SEQ_NO,FILE_NM,FILE_ID,FILE_SIZE,INSRT_DT,INSRT_USER_ID) "
					lgStrSQL = lgStrSQL & " SELECT '"&file_no&"',"
					lgStrSQL = lgStrSQL & " (SELECT ISNULL(MAX(SEQ_NO)+1,1) FROM B_CIS_FILE_DETAIL WHERE FILE_NO="&FILE_NO&"),"
					lgStrSQL = lgStrSQL & " "&filterVar(arr(0),"''","S")&","&filterVar(iFileId,"''","S")&","&arr(2)&",GETDATE(),"&filterVar(strUsrId,"''","S")&""

	
			lgObjConn.execute lgStrSQL
			
		End If														
							
		tempFileId = tempFileId & "'" & iFileId & "',"
	Next
'Response.End 
End Function


'=======================================================================
'FileTransfer
'=======================================================================
Function FileTransfer(SourceFilePath,TargetPath, TargetFileName)
	
   Dim pfile
   Set pfile = CreateObject("PuniFile.CTransfer")

   Call pfile.FileCopy(SourceFilePath,TargetPath, TargetFileName)
   
   Set pfile = Nothing
   
End Function


'=======================================================================
'FileDelete
'=======================================================================
Function FileDelete(byVal pArr )
 	'on Error Resume Next
	Dim filePath
 	Dim i
	
	filePath=server.MapPath (".")&"\files\"

	for i=0 to uBound(pArr,1)
		if len(pArr(i)) > 1 then Call pfile.fileDelete(replace(filePath & pArr(i),"\","/"))   
		
	next
End Function

'=======================================================================
'FileTransfer
'=======================================================================
Function DisplayMessageBox(temp)
	Response.Write "<Script Language=vbscript>"            & vbCr
	Response.Write " msgbox """       &  temp   & """" & vbCr
	Response.Write "</Script>"                             & vbCr
End Function


%>


































<OBJECT RUNAT=server PROGID="PuniFile.CTransfer" id=pfile></OBJECT>