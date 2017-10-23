<% Option Explicit%>
<!--
'**********************************************************************************************
'*  1. Module Name          : CIS
'*  2. Function Name        : 
'*  3. Program ID           :
'*  4. Program Name         : 
'*  5. Program Desc         : 도면파일관리(저장)
'*  6. Comproxy List        :
'*  7. Modified date(First) :
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     :
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'**********************************************************************************************-->

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/UNI2KCMCom.inc" -->    

<%    

Call LoadBasisGlobalInf()    

Dim lgStrSQL
Dim lgStrInternalCd
Dim lgStrItemCd,lgStrReq_no

Dim strMode	

Dim lgObjConn
Dim lgObjComm
Dim lgObjRs

Dim iStrFileInfo

Dim arrFile
Dim iPos, istrFilePath, istrFileId

Const C_FileName  = 0
Const C_FileId    = 1	
Const C_FileSize  = 2
Const C_FileCDate = 3

strMode         = Request("txtMode")
lgStrInternalCd = Request("txtInternalCd")
lgStrItemCd     = Request("txtItemCd")
lgStrReq_no = Request("txtarReqNo")
 

'FILES폴더 없으면 생성 

istrFilePath = SERVER.MapPath (".") & "\FILES\" 
		   
Dim fso

Set fso = CreateObject("Scripting.FileSystemObject")
				 
If Not fso.FolderExists(istrFilePath) then        
   fso.CreateFolder(istrFilePath)
End If

Set fso = Nothing 

Call SubOpenDB(lgObjConn)
Call SubCreateCommandObject(lgObjComm)


    
If CStr(strMode) > CStr(UID_M0001) Then
   '수정/삭제 
   
   'lgStrSQL  = "SELECT FILE_PATH, ID_FILE FROM B_CIS_DOCUMENT_FILE WHERE INTERNAL_CD = '" & lgStrReq_no & "'" 
	lgStrSQL  = "SELECT FILE_PATH, ID_FILE FROM B_CIS_DOCUMENT_FILE WHERE REQ_NO = '" & lgStrReq_no & "'" 						
   If FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
   Else   					
	  Call FileDelete(CStr(lgObjRs(0)) & CStr(lgObjRs(1)))
   End If			
   	
   lgStrSQL = "DELETE FROM B_CIS_DOCUMENT_FILE WHERE REQ_NO = '" & lgStrReq_no & "'" 
                                
   lgObjConn.execute lgStrSQL
                      
   If lgObjConn.errors.count <> 0 then
      Call DisplayMsgBox("800407", vbOKOnly, "", "", I_MKSCRIPT) '작업실행중 에러입니다.
      Response.End
   Else 
      If CStr(strMode) = CStr(UID_M0003) Then 
         Call DisplayMsgBox("210032", vbOKOnly, "", "", I_MKSCRIPT) '삭제되었습니다!
      End If
   End If
       
End If

If CStr(strMode) < CStr(UID_M0003) And Trim(Request("txtFileInf")) <> "" Then '파일첨부 
   '입력/수정이면 
                     
   iStrFileInfo = Split(Request("txtFileInf") , gRowSep)               
				        
   arrFile    = Split(iStrFileInfo(0) , gColSep)
   iPos       = InStrRev(arrFile(C_FileId), "/", -1)
   istrFileId = Right(arrFile(C_FileId), Len(arrFile(C_FileId)) - iPos)
				                 
   Call FileTransfer(arrFile(C_FileId), istrFilePath, istrFileId)
				                   
   lgStrSQL = "INSERT INTO B_CIS_DOCUMENT_FILE ( "
   lgStrSQL = lgStrSQL & " ID_FILE, REQ_NO,FILE_NM, FILE_SIZE, FILE_PATH, ITEM_CD, INTERNAL_CD, INSRT_USER_ID, UPDT_USER_ID) "
   lgStrSQL = lgStrSQL & " VALUES ( " 
   lgStrSQL = lgStrSQL & FilterVar(istrFileId,"''","S") & ","
   lgStrSQL = lgStrSQL & FilterVar(lgStrReq_no,"''","S") & ","
   lgStrSQL = lgStrSQL & FilterVar(arrFile(0),"''","S") & ","
   lgStrSQL = lgStrSQL & FilterVar(arrFile(2),"''","S") & ","
   lgStrSQL = lgStrSQL & FilterVar(istrFilePath,"''","S") & ","
   lgStrSQL = lgStrSQL & FilterVar(lgStrItemCd,"''","S") & ","                   
   lgStrSQL = lgStrSQL & FilterVar(lgStrInternalCd,"''","S") & ","
   lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & "," 
   lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & ")" 
				    
   lgObjConn.execute lgStrSQL
		                  
   If lgObjConn.errors.count <> 0 then
      Call DisplayMsgBox("800407", vbOKOnly, "", "", I_MKSCRIPT)  '작업실행중 에러입니다.
      Response.End 
   Else
      If CStr(strMode) = CStr(UID_M0001) Then
         Call DisplayMsgBox("210030", vbOKOnly, "", "", I_MKSCRIPT) '등록되었습니다!
      Else   
         Call DisplayMsgBox("210031", vbOKOnly, "", "", I_MKSCRIPT) '수정되었습니다!
      End If   
   End if
       
End If

Call SubCloseCommandObject(lgObjComm)    
Call SubCloseDB(lgObjConn)          

Response.Write "<Script Language=vbscript>"    & vbCr
Response.Write "call Parent.DBSaveOk() "              & vbCr
Response.Write "call parent.FncClose() "              & vbCr
Response.Write "</Script>"                     & vbCr        
Response.End    '☜: Process End

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
Function FileDelete(TargetFilePath)

   Dim pfile
   Set pfile = CreateObject("PuniFile.CTransfer")

   Call pfile.fileDelete(TargetFilePath)   

   Set pfile = Nothing

End Function


%>