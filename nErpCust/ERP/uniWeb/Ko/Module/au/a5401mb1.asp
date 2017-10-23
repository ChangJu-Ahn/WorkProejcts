 
<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Account
'*  2. Function Name        : 
'*  3. Program ID           : b2112mb1
'*  4. Program Name         : 미결관리등록 
'*  5. Program Desc         : 세금신고사업장등록,수정,삭제 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/11/08
'*  8. Modified date(Last)  : 2002/11/08
'*  9. Modifier (First)     : Jung Sung Ki
'* 10. Modifier (Last)      : Jung Sung Ki
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                           
'**********************************************************************************************
Response.Expires = -1                                                                '☜ : will expire the response immediately
Response.Buffer = True                                                               '☜ : The server does not send output to the client until all of the ASP scripts on the current page have been processed

   On Error Resume Next                                                             '☜: Protect system from crashing
   Err.Clear                                                                        '☜: Clear Error status
	Dim lgErrorStatus, lgErrorPos, lgObjConn,lgObjRs
    Dim lgtxtAcctBaseNo
    Dim lgtxtCashAmt
    Dim lgcboCardMM
    Dim lgcboCardDD
    Dim txtAcctBaseNo
    Dim txtAcctBaseNm
    Dim lgOpModeCRUD

%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->


<%

	Call LoadBasisGlobalInf() 
	Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB") 	
	Call HideStatusWnd
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    
    lgtxtAcctBaseNo		= FilterVar(UCase(Request("txtAcctBaseNo")), "''", "S")
    lgtxtCashAmt		= UNIConvNum(Request("txtCashAmt"),0)
    lgcboCardMM			= FilterVar(UCase(Request("cboCardMM")), "''", "S")
    lgcboCardDD			= FilterVar(UCase(Request("cboCardDD")), "''", "S")

	'------ Developer Coding part (Start ) ------------------------------------------------------------------


	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
    

		

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection
%>
<Script Language="VBScript">
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"		
          If Trim("<%=lgErrorStatus%>") = "NO" Then			
             Parent.DBQueryOk        
          End If   
       Case "<%=UID_M0002%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          End If   
    End Select    
       
</Script>	

<%

response.end

Sub SubBizQuery()
    Dim iKey1
	Dim lgCARD_DD
	Dim lgStrSQL
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    ' 
    ' 
	lgStrSQL = "Select Top 1 ACCT_BASE_NO,ACCT_BASE_NM, CASH_AMT, CARD_MM,CARD_DD" 
	lgStrSQL = lgStrSQL & " From	A_OPEN_ACCT_BASE "
	'lgStrSQL = lgStrSQL & " WHERE	ACCT_BASE_NO = " & lgtxtAcctBaseNo 	
'Response.Write lgStrSQL

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then                    'If data not exists
       If lgPrevNext = "" Then
          Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
          Call SetErrorStatus()
       End If
       
    Else

		Response.Write " <Script Language=vbscript>	" & vbCr
		Response.Write " With parent.frm1           " & vbCr														
		Response.Write ".txtAcctBaseNo.value		= """ & ConvSPChars(lgObjRs("ACCT_BASE_NO"))	& """" & vbCr
		Response.Write ".txtAcctBaseNm.value		= """ & ConvSPChars(lgObjRs("ACCT_BASE_NM"))	& """" & vbCr
		Response.Write ".txtCashAmt.value			= """ & UNINumClientFormat(lgObjRs("CASH_AMT"), ggExchRate.DecPoint, 0)	& """" & vbCr
		Response.Write ".cboCardMM.value			= """ & ConvSPChars(lgObjRs("CARD_MM"))	& """" & vbCr

		if len(ConvSPChars(lgObjRs("CARD_DD"))) = 1 then
			Response.Write ".cboCardDD.value			= """ & "0" & ConvSPChars(lgObjRs("CARD_DD"))	& """" & vbCr
		else
			Response.Write ".cboCardDD.value			= """ & ConvSPChars(lgObjRs("CARD_DD"))	& """" & vbCr
		end if
		Response.Write ".hAcctBaseNo.value			= """ & ConvSPChars(lgObjRs("ACCT_BASE_NO"))	& """" & vbCr


		Response.write "End With	      " & vbCr
		Response.write "parent.DbQueryOk  " & vbCr
		Response.write " </Script>        " & vbCr

   End If


   Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet
End Sub	

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSave()
	Dim lgStrSQL
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------



	lgStrSQL = "UPDATE  A_OPEN_ACCT_BASE"
	lgStrSQL = lgStrSQL & " SET " 
    
	lgStrSQL = lgStrSQL & " CASH_AMT = " & lgtxtCashAmt & ","	   
	lgStrSQL = lgStrSQL & " CARD_MM  = " & lgcboCardMM  & ","			 
	lgStrSQL = lgStrSQL & " CARD_DD  = " & lgcboCardDD  & " "
    
'	lgStrSQL = lgStrSQL & " UPDT_USER_ID = " & FilterVar(gUsrId, "''","S") &  ","   
'	lgStrSQL = lgStrSQL & " UPDT_DT =" & FilterVar(GetSvrDateTime,"''","S")
	lgStrSQL = lgStrSQL & " WHERE ACCT_BASE_NO = " & lgtxtAcctBaseNo
  
	'---------- Developer Coding part (End  ) ---------------------------------------------------------------
	lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords
'	Call SubHandleError(,lgObjConn,lgObjRs,Err)
	
	Call SubCloseRs(lgObjRs)                                                    '☜ : Release RecordSSet
		
End Sub


%>


