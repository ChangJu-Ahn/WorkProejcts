<% Option Explicit%>
<!--
'**********************************************************************************************
'*  1. Module Name          : CIS
'*  2. Function Name        : 
'*  3. Program ID           : S3322PB_KO412
'*  4. Program Name         : 
'*  5. Program Desc         : Create, update and delete notice.
'*  6. Comproxy List        : ADO Query Program
'*  7. Modified date(First) : 2005/01/30
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : lee wol san
'* 10. Modifier (Last)      : lee ho jun
'* 11. Comment              :
'**********************************************************************************************-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/UNI2KCMCom.inc" -->	
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="B81COMM.ASP" -->

<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 
On Error Resume Next														'☜: 
'---------------------------------------Common-----------------------------------------------------------
Call LoadBasisGlobalInf()    

Dim project_code, lgStrSQL, lgStrSQL2
Dim strReportNo, strIns_person, strReport_abbr, strReport_text, strPasswd, strUsrId
Dim strMode	'☜: 현재 MyBiz.asp 의 진행상태를 나타냄 

Dim lgObjConn
Dim lgObjComm
Dim lgObjRs

Dim iStrFileInfo
Dim arrTemp,arr, tempFileId
Dim iStrYYYYMM
Dim iPos,iFilePath, iFileId, iNoticenum, i
Dim iCntNoticeNum		'글삭제시 Flag
Dim strProjectCode
Dim strUse_dt
Dim strInsrt_dt
Dim delTemp

Const C_FileName = 0
Const C_FileId   = 1	
Const C_FileSize = 2
Const C_FileCDate = 3

strMode			= Request("txtMode")												'☜ : 현재 상태를 받음 
strProjectCode	= FilterVar(Request("project_code"), "", "S")
strUsrId		= Replace(gUsrId, "'", "''")
strReportNo		= FilterVar(Request("txtTITLE"), "", "S")
strIns_person   = FilterVar(Request("INS_PERSON"), "", "S")
strReport_text	= FilterVar(Request("txtReportText"), "", "S")
strReport_abbr	= FilterVar(Request("report_abbr"), "", "S")

'	Call ServerMesgBox(strMode, vbInformation, I_MKSCRIPT)

strinsrt_dt		= filterVar(UniConvDate(Request("Insrt_dt")),"''","S")
strUse_dt		= filterVar(UniConvDate(Request("Use_dt")),"''","S")
delTemp         = Request("delTemp") '삭제할파일 

Call SubOpenDB(lgObjConn)    
Call SubCreateCommandObject(lgObjComm)

	'FILES폴더 없으면 생성 
	Dim fso
	iFilePath = SERVER.MapPath (".") & "\files\"		'

	Set fso = CreateObject("Scripting.FileSystemObject")   
	If Not fso.FolderExists(iFilePath) then		
	 	   fso.CreateFolder(iFilePath)
	End If	
	Set fso = Nothing
	

	DIm WshNetwork'
	DIm NetworkPath, oDrives
	
	set WshNetwork = Server.CreateObject("WScript.Network")

	Set oDrives = WshNetwork.EnumNetworkDrives

	if oDrives.Count-1 >0 then
		WshNetwork.RemoveNetworkDrive"Y:"
'		Set WshNetwork = Nothing
'		Set oDrives= nothing
	end if

	WshNetwork.MapNetworkDrive "Y:", "\\192.168.10.99\d$\Unierp_File",false,"administrator","nepes123"

	NetworkPath="Y:\"

	if err.number <> 0 then
		if oDrives.Count-1 >0 then
			WshNetwork.RemoveNetworkDrive"Y:"
			Set WshNetwork = Nothing
			Set oDrives= nothing
		end if

		Call DisplayMessageBox("파일저장이 되지 않았습니다.")
		Response.End 
	end if

	'iFilePath = NetworkPath


Select Case CStr(strMode)
	
    Case CStr(UID_M0001)   
		
			'=================================================
			' INsert
			'=================================================

		    lgObjConn.beginTrans()

			lgStrSQL = "INSERT INTO S_PRJ_REPORT_HDR_KO412 (PROJECT_CODE, REPORT_NO,INS_USER,REPORT_ABBR, INS_DT,REPORT_TEXT,INSRT_USER_ID,INSRT_DT,UPDT_USER_ID,UPDT_DT) "
			lgStrSQL = lgStrSQL & "VALUES(" & strProjectCode & "," & strReportNo & "," & strIns_person & "," & strReport_abbr & ","&strinsrt_dt&"," & strReport_text & ",'"&strUsrId&"', "&strinsrt_dt&",'"&strUsrId&"',"&strinsrt_dt&")"

			lgObjConn.execute lgStrSQL
			'responee.end
			
			If Trim(Request("txtFileInf")) <> "" Then '파일첨부		
				Call RunFileAttachForINSERT()
			End If

			if lgObjConn.errors.count<>0 then
				Response.Write lgObjConn.errors.description 
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

			lgStrSQL =  "SELECT insrt_user_Id  FROM S_PRJ_REPORT_HDR_KO412 WHERE PROJECT_CODE ="& strprojectCode &" AND REPORT_NO = "& strReportNo

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
					
				lgStrSQL  = "DELETE FROM S_PRJ_REPORT_DTL_KO412 WHERE PROJECT_CODE = " & strProjectCode & " AND REPORT_ID IN  ('"&Replace(delTemp,",","','")&"')"
				 call FileDelete(split(delTemp,","))
				lgObjConn.execute lgStrSQL														
				'=========================
				'  update
				'=========================
																					'☜: Protect system from crashing
					lgStrSQL =" UPDATE S_PRJ_REPORT_HDR_KO412 " 
					lgStrSQL = lgStrSQL & " SET INS_USER="&strIns_person&","
					lgStrSQL = lgStrSQL & " REPORT_ABBR="&strReport_abbr&","
					lgStrSQL = lgStrSQL & " REPORT_TEXT="&strReport_text&","
					lgStrSQL = lgStrSQL & " UPDT_USER_ID="&Filtervar(strUsrId,"''","S")&","
				   'lgStrSQL = lgStrSQL & "USE_DT="&strUse_dt&","
					lgStrSQL = lgStrSQL & " INS_DT="&strinsrt_dt&","
					lgStrSQL = lgStrSQL & " UPDT_DT=GETDATE()"
					lgStrSQL = lgStrSQL & " WHERE PROJECT_CODE="& strProjectCode
					lgStrSQL = lgStrSQL & " AND REPORT_NO="& strReportNo
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

	Dim Report_Id

	iStrFileInfo = Request("txtFileInf")	
	iStrFileInfo = FilterVar(iStrFileInfo, "", "S")           '☜For Single Quotation			
	arrTemp = Split(Request("txtFileInf"), gRowSep)
				
	lgStrSQL = ""
	
	For i = 0 To UBound(arrTemp) - 1
	
		arr  = Split(arrTemp(i), gColSep)								
		iPos      = InStrRev(arr(C_FileId), "/", -1)
		iFileId   = Right(arr(C_FileId), len(arr(C_FileId)) - iPos)
'		iFilePath = Request.ServerVariables("APPL_PHYSICAL_PATH") & glang & "\Notice\files\" & iStrYYYYMM & "\"
		Call FileTransfer(arr(C_FileId), iFilePath, iFileId)
		'Call FileAttachCopy(iFilePath, "",	iFileId )	
		
		'arrTemp
		'(0) FILE NAME
		'(1) FILE ID
		'(2) FILE SIZE
		'(3) FILECDATE
		
		'*******************************************************************			
		'파일 저장은 서버에다 해준다 204 스토리지 서버에다 해준다...'
		'*******************************************************************				
		'iFilePath = "Y:\"

		Report_Id= split(arr(1),"/")(ubound(split(arr(1),"/")))
		lgStrSQL ="INSERT INTO S_PRJ_REPORT_DTL_KO412 (PROJECT_CODE, REPORT_NO, REPORT_SEQ_NO, REPORT_NM, REPORT_ID, REPORT_SIZE, REPORT_PATH, "
		lgStrSQL = lgStrSQL & " insrt_dt,insrt_user_id,updt_dt,updt_user_id) "
		lgStrSQL = lgStrSQL & " VALUES (" & strProjectCode & "," & strReportNo & ","
		lgStrSQL = lgStrSQL & " "&i&","&filterVar(arr(0),"''","S")&","&filterVar(Report_Id,"''","S")&","&arr(2)&","& filterVar(iFilePath,"''","S") &",GETDATE(),"& filterVar(strUsrId,"''","S") &",GETDATE(),"&filterVar(strUsrId,"''","S")&")"
		
		lgObjConn.execute lgStrSQL 
					
	Next

End Function

function kk(str)%>
<script language="javaScript">
	parent.frm1.txtReportText.value="<%=str%>"
	</script>
	
<%end function


'=======================================================================
'RunFileAttachForUPDATE
'=======================================================================
Function RunFileAttachForUPDATE()

	On Error Resume Next
					
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
		
		arr  = Split(arrTemp(i), gColSep)								
		iPos      = InStrRev(arr(C_FileId), "/", -1)
		iFileId   = Right(arr(C_FileId), len(arr(C_FileId)) - iPos)

'		iFilePath = Request.ServerVariables("APPL_PHYSICAL_PATH") & glang & "\Notice\files\" & iStrYYYYMM & "\"

		'FILES폴더 없으면 생성 
'		Dim fso
'		Set fso = CreateObject("Scripting.FileSystemObject")   
'		If Not fso.FolderExists(Request.ServerVariables("APPL_PHYSICAL_PATH") & glang & "\module\s3\files\") then
'		 	   fso.CreateFolder(Request.ServerVariables("APPL_PHYSICAL_PATH") & glang & "\module\s3\files\")
'		End If		
'		Set fso = Nothing
							
		lgStrSQL2  = "SELECT 1 FROM S_PRJ_REPORT_DTL_KO412 WHERE '" & iFileId & "' IN (SELECT REPORT_ID FROM S_PRJ_REPORT_DTL_KO412 WHERE project_Code = " & strProjectCode & " AND REPORT_no = " & strReportNo & ") "

		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL2,"X","X") = True Then                    '기존파일 
				'기존파일은 신규로 해줄 작업이 없다.							
		Else	'신규 파일은 파일복사, INSERT 
				Call FileTransfer(arr(C_FileId), iFilePath, iFileId)
				'Call FileAttachCopy(iFilePath, "",	iFileId )
				

				lgStrSQL ="INSERT INTO S_PRJ_REPORT_DTL_KO412"
				lgStrSQL = lgStrSQL & " (PROJECT_CODE,REPORT_NO,REPORT_SEQ_NO,REPORT_NM,REPORT_ID,REPORT_SIZE,REPORT_PATH, UPDT_DT,UPDT_USER_ID, INSRT_DT,INSRT_USER_ID) "
				lgStrSQL = lgStrSQL & " VALUES("& strprojectcode &","& strReportNo &", " & i & ", "& FilterVar(arr(0),"''","S") &","& filterVar(iFileId,"''","S") &","& arr(2) &",'" & iFilePath & "', GETDATE(),"&filterVar(strUsrId,"''","S")&", GETDATE(),"&filterVar(strUsrId,"''","S")&")"
				
				'lgStrSQL = "INSERT INTO B_NOTICE_FILE(NOTICENUM,FLE_ID,FLE_NAME, FLE_SIZE,FLE_PATH,INSRT_USER_ID,INSRT_DT,UPDT_USER_ID,UPDT_DT) "
				'lgStrSQL = lgStrSQL & "VALUES(" & strKeyNo & ", '" & iFileId & "', '" & arrTemp2(C_FileName) & "', '" & arrTemp2(C_FileSize) & "', '" & iFilePath & "', '" & gUsrId & "', Getdate(), '" & gUsrId & "', Getdate())  " 
	
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

   On error resume next
   
   Dim pfile
   Set pfile = CreateObject("PuniFile.CTransfer")

   Call ServerMesgBox(SourceFilePath, vbInformation, I_MKSCRIPT)
   Call ServerMesgBox(TargetPath, vbInformation, I_MKSCRIPT)
   Call ServerMesgBox(TargetFileName, vbInformation, I_MKSCRIPT)

   Call pfile.FileCopy(SourceFilePath,TargetPath, TargetFileName)

   'Fso.CopyFile File1,File2,true
   Set pfile = Nothing
   
End Function


'=======================================================================
'FileDelete
'=======================================================================
Function FileDelete(byVal pArr )
 	on Error Resume Next
	Dim filePath
 	Dim i
	
'	filePath=server.MapPath (".")&"\files\"
	filePath = NetworkPath

	for i=0 to uBound(pArr,1)
		if len(pArr(i)) > 1 then Call pfile.fileDelete(replace(filePath & pArr(i),"\","/"))
	next
	
End Function

'=======================================================================
'3
'=======================================================================
Function DisplayMessageBox(temp)
	Response.Write "<Script Language=vbscript>"            & vbCr
	Response.Write " msgbox """       &  temp   & """" & vbCr
	Response.Write "</Script>"                             & vbCr
End Function

'============================================================================================================
' Name : FileAttachCopy
' Desc :204번서버에다 첨부파일을 복사해 준다.
'============================================================================================================
Sub FileAttachCopy(Byval Path, byval Folder, Byval Filename)

	On Error Resume Next
	DIm FSO, File1, File2
	
   Set FSO = CreateObject("Scripting.FileSystemObject")
  
	If FSO.FolderExists("Y:\" & Folder) then
	Else
	   fso.CreateFolder("Y:\" & Folder)
	End if
	
	File1= Path & Filename
	File2= NetworkPath & Folder & "\" & Filename

	FSO.CopyFile File1 , File2 ,true

	Call pfile.fileDelete(File1)
	'FSO.DeleteFile(File1)
   
End Sub

	Set FSO = Nothing
	Set WshNetwork= nothing
	Set oDrives= nothing   

%>


<OBJECT RUNAT=server PROGID="PuniFile.CTransfer" id=pfile></OBJECT>