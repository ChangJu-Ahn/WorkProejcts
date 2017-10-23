<% Option Explicit%>
<!--
'**********************************************************************************************
'*  1. Module Name          : Notice
'*  2. Function Name        : 
'*  3. Program ID           : FRWriteBiz.asp
'*  4. Program Name         : 
'*  5. Program Desc         : Create, update and delete notice.
'*  6. Comproxy List        : ADO Query Program
'*  7. Modified date(First) : 2002/10/09
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Park Sang Hoon
'* 10. Modifier (Last)      : Park Sang Hoon
'* 11. Comment              :
'**********************************************************************************************-->

<!-- #Include file="../inc/IncSvrMain.asp" -->
<!-- #Include file="../inc/adovbs.inc" -->
<!-- #Include file="../inc/incServerAdoDb.asp" -->
<!-- #Include file="../inc/UNI2KCMCom.inc" -->	


<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
'On Error Resume Next														'��: 

'---------------------------------------Common-----------------------------------------------------------
Call LoadBasisGlobalInf()    


Dim strKeyNo, lgStrSQL, lgStrSQL2
Dim strSubject, strWriter, strContents, strPasswd, strUsrId
Dim strMode	'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

Dim lgObjConn
Dim lgObjComm
Dim lgObjRs

Dim iStrFileInfo
Dim arrTemp,arrTemp2, tempFileId
Dim iStrYYYYMM
Dim iPos,iFilePath, iFileId, iNoticenum, i
Dim iCntNoticeNum		'�ۻ����� Flag

Const C_FileName = 0
Const C_FileId = 1	
Const C_FileSize = 2
Const C_FileCDate = 3


strMode  = Request("txtMode")												'�� : ���� ���¸� ���� 
strKeyNo = Request("txtKeyNo") 

strUsrId = Replace(gUsrId, "'", "''")

strSubject  = FilterVar(Request("SUBJECT"), "", "SNM")
strWriter   = FilterVar(Request("WRITER"), "", "SNM")
strContents = FilterVar(Request("TXTCONTENT"), "", "SNM")

Call SubOpenDB(lgObjConn)    
Call SubCreateCommandObject(lgObjComm)

Select Case CStr(strMode)
	
    Case CStr(UID_M0001)                                                         '��: �۵�� 
		
			lgStrSQL = "SELECT PWD FROM Z_USR_MAST_REC WHERE USR_ID = '" & strUsrId & "'"
			
			If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists	
				Call DisplayMsgBox("210100", vbInformation, "", "", I_MKSCRIPT)      '�� : ����� ���� ������ �ش��ϴ� �ڷᰡ �������� �ʽ��ϴ�.
			Else
				strPasswd = lgObjRs(0)
		    End If
		    
			lgStrSQL = "INSERT INTO B_NOTICE(WRITER, USR_ID, PWD, SUBJECT, CONTENTS, REGDATE) "			
			lgStrSQL = lgStrSQL & "VALUES('" & strWriter & "','" & strUsrId & "','" & strPasswd & "','" & strSubject & "','" & strContents & "', GETDATE())"
			'lgStrSQL = "INSERT INTO B_NOTICE(NOTICENUM, WRITER, USR_ID, PWD, SUBJECT, CONTENTS0, REGDATE) "	'FOR HERMES
			'lgStrSQL = lgStrSQL & "VALUES(b_notice_no_seq.nextval, '" & strWriter & "','" & strUsrId & "','" & strPasswd & "','" & strSubject & "','" & strContents & "', SYSDATE)" 'FOR HERMES

			lgObjConn.execute lgStrSQL

			If Trim(Request("txtFileInf")) <> "" Then '����÷��			
				Call RunFileAttachForINSERT()
			End If

			Call DisplayMsgBox("210030", vbOKOnly, "", "", I_MKSCRIPT)	  '��ϵǾ����ϴ�!		

			'����� ȭ�� �ٽ� �ε� 
			Response.Write "<Script Language=vbscript>"			& vbCr
			Response.Write "Parent.DbSaveOk "			& vbCr
			Response.Write "</Script>" & vbCr		

    Case CStr(UID_M0002)                                                         '��: �ۼ��� 

			lgStrSQL = "select * from B_NOTICE where  NoticeNum = " & strKeyNo
			
			If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists	
				Call DisplayMsgBox("210037", vbInformation, "", "", I_MKSCRIPT)      '�� : �Խù��� �ش��ϴ� �ڷᰡ �������� �ʽ��ϴ�.
			Else
			
				If UCASE(gUsrId) <> UCASE(lgObjRs("usr_id")) Then
					Call DisplayMsgBox("210033", vbOKOnly, "", "", I_MKSCRIPT) '������ �����ϴ�!					
					Response.Write "<Script Language=vbscript>"			& vbCr
					Response.Write "Self.Close()"
					Response.Write "</Script>" & vbCr		
					Response.End															
				Else 

					Err.Clear																		'��: Protect system from crashing

					lgStrSQL = "UPDATE B_NOTICE SET WRITER = '" & strWriter & "', USR_ID = '" & strUsrId & "', SUBJECT = '" & strSubject & "', CONTENTS = '" & strContents & "', REGDATE = GETDATE() WHERE NOTICENUM = "	& strKeyNo		
					'lgStrSQL = "UPDATE B_NOTICE SET WRITER = '" & strWriter & "', USR_ID = '" & strUsrId & "', SUBJECT = '" & strSubject & "', CONTENTS0 = '" & strContents & "', REGDATE = SYSDATE WHERE NOTICENUM = "	& strKeyNo		'FOR HERMES
					lgObjConn.Execute lgStrSQL	

					If Trim(Request("txtFileInf")) <> "" Then '����÷�� 
						Call RunFileAttachForUPDATE()
					End If

					'������ ���� ó�� 
					If tempFileId <> "" Then
						tempFileId = mid(tempFileId, 1, len(tempFileId) - 1)
					Else
						tempFileId = "'1'"	'Nulló�� 
					End If

					lgStrSQL2  = "SELECT FLE_PATH, FLE_ID FROM B_NOTICE_FILE WHERE NOTICENUM = '" & strKeyNo & "' AND  FLE_ID NOT IN (" & tempFileId & ")"
					
					If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL2,"X","X") = False Then                    'If data not exists '���� ���� 
					Else						
						Do While Not lgObjRs.EOF
							lgStrSQL  = "DELETE B_NOTICE_FILE WHERE NOTICENUM = '" & strKeyNo & "' AND  FLE_ID NOT IN (" & tempFileId & ")"																					
							Call FileDelete(CStr(lgObjRs(0)) & CStr(lgObjRs(1)))														
							lgObjConn.execute lgStrSQL														
							lgObjRs.MoveNext							
						Loop
					End If	

					Call DisplayMsgBox("210031", vbOKOnly, "", "", I_MKSCRIPT) '�����Ǿ����ϴ�!						
								
				End If				
		    End If
		    
			'������ ȭ�� �ٽ� �ε� 
			Response.Write "<Script Language=vbscript>"			& vbCr
			Response.Write "Parent.DbSaveOk "			& vbCr
			Response.Write "</Script>" & vbCr		

    Case CStr(UID_M0003)                                                         '��: �ۻ��� 

			lgStrSQL = "select * from B_NOTICE where  NoticeNum = " & strKeyNo
			
			If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists	
				Call DisplayMsgBox("210037", vbInformation, "", "", I_MKSCRIPT)      '�� : �Խù��� �ش��ϴ� �ڷᰡ �������� �ʽ��ϴ�.
			Else
			
				If UCASE(gUsrId) <> UCASE(lgObjRs("usr_id")) Then
					Call DisplayMsgBox("210033", vbOKOnly, "", "", I_MKSCRIPT)		 '�� : ������ �����ϴ�!
					Response.End					
				Else 
					
					Err.Clear                                                               '��: Protect system from crashing	

					lgStrSQL2  = "SELECT FLE_PATH, FLE_ID FROM B_NOTICE_FILE WHERE NOTICENUM = '" & strKeyNo & "' "
						
					If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL2,"X","X") = False Then                    'If data not exists '���� ������ �״��						
					Else						
						Do While Not lgObjRs.EOF							
							Call FileDelete(CStr(lgObjRs(0)) & CStr(lgObjRs(1)))							
							lgObjRs.MoveNext
						Loop
					End If
					
					lgStrSQL = "DELETE B_NOTICE WHERE NOTICENUM = " & strKeyNo 
					
					lgObjConn.Execute lgStrSQL	
					
					Call DisplayMsgBox("210032", vbOKOnly, "", "", I_MKSCRIPT)  '�����Ǿ����ϴ�!

				End If	
			End If	
			
			'������ ȭ�� �ٽ� �ε� 
			lgStrSQL = "SELECT COUNT(*) FROM B_NOTICE "

			If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = True Then                    'If data not exists						
				iCntNoticeNum = lgObjRs(0)
			End If

			Response.Write "<Script Language=vbscript>"			& vbCr
						
			Response.Write "If ((" & iCntNoticeNum & " Mod parent.MyBizAsp.frTitle.intPageSize ) = 0 ) And (parent.MyBizAsp.frTitle.intNowPage = parent.MyBizAsp.frTitle.intTotalPage) Then " & vbCr
			Response.write "	Parent.MyBizASP.frames(""frTitle"").document.URL = ""frtitle.asp?page=""" & "& CStr(parent.MyBizAsp.frTitle.intNowPage) - 1" & vbCr						
			Response.Write "Else " & vbCr						
			Response.write "	Parent.MyBizASP.frames(""frTitle"").document.URL = ""frtitle.asp?page=""" & "& CStr(parent.MyBizAsp.frTitle.intNowPage)" & vbCr
			Response.Write "End If " & vbCr						
			Response.Write "</Script>" & vbCr															
			


End Select

Call SubCloseCommandObject(lgObjComm)    
Call SubCloseDB(lgObjConn)      	

Response.End	'��: Process End

'==============================================================================
' ����� ���� ���� �Լ� 
'==============================================================================
Function RunFileAttachForINSERT()

	lgStrSQL = "SELECT TOP 1 NOTICENUM, CONVERT(CHAR(7),GETDATE(),20) FROM B_NOTICE ORDER BY 1 DESC"
	'lgStrSQL = "SELECT NOTICENUM, TO_CHAR(SYSDATE,'YYYY-MM') FROM B_NOTICE WHERE ROWNUM = 1 ORDER BY 1 DESC"    'For HERMES
				  
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists	
		'Call DisplayMsgBox("210301", vbInformation, "", "", I_MKSCRIPT)      '�� : Login History Management : Cannot find the data.. 
	Else
		iNoticenum = lgObjRs(0)
		iStrYYYYMM = lgObjRs(1)				
	End If

	iStrFileInfo = Request("txtFileInf")	
	iStrFileInfo = FilterVar(iStrFileInfo, "", "SNM")           '��For Single Quotation			
	arrTemp = Split(Request("txtFileInf"), gRowSep)
				
	lgStrSQL = ""
				
	For i = 0 To UBound(arrTemp) - 1
				
		arrTemp2  = Split(arrTemp(i), gColSep)								
		iPos      = InStrRev(arrTemp2(C_FileId), "/", -1)
		iFilePath = Request.ServerVariables("APPL_PHYSICAL_PATH") & glang & "\Notice\files\" & iStrYYYYMM & "\"
		
		'FILES���� ������ ���� 
		Dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")   
		If Not fso.FolderExists(Request.ServerVariables("APPL_PHYSICAL_PATH") & glang & "\Notice\files\") then		
		 	   fso.CreateFolder(Request.ServerVariables("APPL_PHYSICAL_PATH") & glang & "\Notice\files\")
		End If	
		Set fso = Nothing   
							
		iFileId   = Right(arrTemp2(C_FileId), len(arrTemp2(C_FileId)) - iPos)

		Call FileTransfer(arrTemp2(C_FileId), iFilePath, iFileId)
					
		lgStrSQL = "INSERT INTO B_NOTICE_FILE(NOTICENUM,FLE_ID,FLE_NAME, FLE_SIZE,FLE_PATH,INSRT_USER_ID,INSRT_DT,UPDT_USER_ID,UPDT_DT) "
		lgStrSQL = lgStrSQL & "VALUES(" & iNoticenum & ", '" & iFileId & "', '" & arrTemp2(C_FileName) & "', '" & arrTemp2(C_FileSize) & "', '" & iFilePath & "', '" & gUsrId & "', Getdate(), '" & gUsrId & "', Getdate())  "
		'lgStrSQL = "INSERT INTO B_NOTICE_FILE(NOTICENUM,FLE_ID,FLE_NAME, FLE_SIZE,FLE_PATH,INSRT_USER_ID,INSRT_DT,UPDT_USER_ID,UPDT_DT) "	'FOR HERMES
		'lgStrSQL = lgStrSQL & "VALUES(" & iNoticenum & ", '" & iFileId & "', '" & arrTemp2(C_FileName) & "', '" & arrTemp2(C_FileSize) & "', '" & iFilePath & "', '" & gUsrId & "', Sysdate, '" & gUsrId & "', Sysdate)  " 'FOR HERMES

		lgObjConn.execute lgStrSQL
					
	Next

End Function

Function RunFileAttachForUPDATE()

	lgStrSQL2 = "SELECT CONVERT(CHAR(7),GETDATE(),20) "					
	'lgStrSQL2 = "SELECT TO_CHAR(SYSDATE,'YYYY-MM') FROM DUAL"   'FOR HERMES					
					
	If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL2,"X","X") = False Then                    'If data not exists	
	Else
		iStrYYYYMM = lgObjRs(0)
	End If
					
	tempFileId = ""

	iStrFileInfo = Request("txtFileInf")
						
	iStrFileInfo = FilterVar(iStrFileInfo, "", "SNM")           '��For Single Quotation			
	arrTemp = Split(Request("txtFileInf"), gRowSep)
						
	lgStrSQL = ""
						
						
	For i = 0 To UBound(arrTemp) - 1
							
		arrTemp2  = Split(arrTemp(i), gColSep)			

		'Call DisplayMessageBox("�Ķ����: " & arrTemp(i))
							
		'From arrTemp2(C_FileId),arrTemp2(C_FileCDate)  >> iFilePath, iFileId

		iPos      = InStrRev(arrTemp2(C_FileId), "/", -1)
		'iFilePath = Left(arrTemp2(C_FileId), iPos)	    'temp���� 
		iFilePath = Request.ServerVariables("APPL_PHYSICAL_PATH") & glang & "\Notice\files\" & iStrYYYYMM & "\"							

		'FILES���� ������ ���� 
		Dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")   
		If Not fso.FolderExists(Request.ServerVariables("APPL_PHYSICAL_PATH") & glang & "\Notice\files\") then		
		 	   fso.CreateFolder(Request.ServerVariables("APPL_PHYSICAL_PATH") & glang & "\Notice\files\")
		End If		
		Set fso = Nothing   
							
		iFileId   = Right(arrTemp2(C_FileId), len(arrTemp2(C_FileId)) - iPos)
							
							
		lgStrSQL2  = "SELECT 1 FROM B_NOTICE_FILE WHERE '" & iFileId & "' IN (SELECT FLE_ID FROM B_NOTICE_FILE WHERE NOTICENUM = '" & strKeyNo & "') " 			
		
		If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL2,"X","X") = True Then                    '�������� 
				'���������� �űԷ� ���� �۾��� ����.							
				
		Else	'�ű� ������ ���Ϻ���, INSERT 
				Call FileTransfer(arrTemp2(C_FileId), iFilePath, iFileId)
							
				lgStrSQL = "INSERT INTO B_NOTICE_FILE(NOTICENUM,FLE_ID,FLE_NAME, FLE_SIZE,FLE_PATH,INSRT_USER_ID,INSRT_DT,UPDT_USER_ID,UPDT_DT) "
				lgStrSQL = lgStrSQL & "VALUES(" & strKeyNo & ", '" & iFileId & "', '" & arrTemp2(C_FileName) & "', '" & arrTemp2(C_FileSize) & "', '" & iFilePath & "', '" & gUsrId & "', Getdate(), '" & gUsrId & "', Getdate())  " 
				'lgStrSQL = lgStrSQL & "INSERT INTO B_NOTICE_FILE(NOTICENUM,FLE_ID,FLE_NAME, FLE_SIZE,FLE_PATH,INSRT_USER_ID,INSRT_DT,UPDT_USER_ID,UPDT_DT) "	'FOR HERMES
				'lgStrSQL = lgStrSQL & "VALUES(" & strKeyNo & ", '" & iFileId & "', '" & arrTemp2(C_FileName) & "', '" & arrTemp2(C_FileSize) & "', '" & iFilePath & "', '" & gUsrId & "', SYSDATE, '" & gUsrId & "', SYSDATE)  " 		'FOR HERMES
				lgObjConn.execute lgStrSQL
				'Call WriteLog(lgStrSQL)
		End If														
							
		tempFileId = tempFileId & "'" & iFileId & "',"
	Next

End Function


Function FileTransfer(SourceFilePath,TargetPath, TargetFileName)
	
   Dim pfile
   Set pfile = CreateObject("PuniFile.CTransfer")

   Call pfile.FileCopy(SourceFilePath,TargetPath, TargetFileName)
   
   Set pfile = Nothing
   
End Function


Function FileDelete(TargetFilePath)

   Dim pfile
   Set pfile = CreateObject("PuniFile.CTransfer")

   Call pfile.fileDelete(TargetFilePath)   

   Set pfile = Nothing

End Function

Function DisplayMessageBox(temp)
	Response.Write "<Script Language=vbscript>"            & vbCr
	Response.Write " msgbox """       &  temp   & """" & vbCr
	Response.Write "</Script>"                             & vbCr
End Function


%>
