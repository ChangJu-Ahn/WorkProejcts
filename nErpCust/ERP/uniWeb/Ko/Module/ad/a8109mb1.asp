<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

'On Error Resume Next
'Err.Clear

Call LoadBasisGlobalInf() 

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 ,rs1, rs2                    '☜ : DBAgent Parameter 선언 
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

Dim txtFromGlDt
Dim txtToGlDt
Dim txtPreFromGlDt
Dim txtPreToGlDt
Dim txtBizAreaCd
Dim txtClassType
Dim lgStrHqBrchFg
Dim strZeroFg
Dim lgStrUserId

Dim lgBizAreaCd
Dim lgBizAreaNm
Dim lgClassType
Dim lgClassNm 

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 

    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
    lgDataExist    = "No"
    
	txtFromGlDt		=Trim(Request("txtFromGlDt"))
	txtToGlDt		=Trim(Request("txtToGlDt"))
	txtPreFromGlDt	=Trim(Request("txtPreFromGlDt"))
	txtPreToGlDt	=Trim(Request("txtPreToGlDt"))
	txtBizAreaCd	=Trim(Request("txtBizAreaCd"))
	txtClassType	=Trim(Request("txtClassType"))
	lgStrHqBrchFg	=Trim(Request("strHqBrchFg"))
	strZeroFg		=Trim(Request("strZeroFg"))
	lgStrUserId		=Trim(Request("strUserId"))
    
    
    '처음조회시에만 sp를 호출하도록한다.------
    If CDbl(lgPageNo) < 1 Then		
		Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
		Call SubBizBatch()    
		Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection   
    End If
    '------------------------------------------
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
	Const C_SHEETMAXROWS_D  = 100 
    
    lgDataExist    = "Yes"
    lgstrData      = ""

    If CDbl(lgPageNo) > 0 Then
       rs0.Move     = CDbl(C_SHEETMAXROWS_D) * CDbl(lgPageNo) 
    End If

    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < C_SHEETMAXROWS_D Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < C_SHEETMAXROWS_D Then                                   '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
    
    If Not( rs1.EOF OR rs1.BOF) Then	
		
   		lgBizAreaCd = rs1(0)
		lgBizAreaNm = rs1(1)		
    End IF
        
    rs1.Close
    Set rs1= Nothing
    
    
    If Not( rs2.EOF OR rs2.BOF) Then	
		
   		lgClassType = rs2(0)
		lgClassNm = rs2(1)		
    End IF        
    rs2.Close
    Set rs2= Nothing
    
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(2)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(2,1)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 

    UNISqlId(0) = "a5108QA101"
    UNISqlId(1) = "A_GetBiz"
    UNISqlId(2) = "A_CLSTYPE"

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    If txtBizAreaCd = "" Then	 	 	
	 	UNIValue(1,0)  = ""	 	
	Else				
		UNIValue(1,0)  =  FilterVar(txtBizAreaCd, "''", "S") 
	End If    
	UNIValue(2,0)  =  FilterVar(txtClassType, "''", "S") 
	
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
    Dim iStr
    Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If

End Sub

'----------------------------------------------------------------------------------------------------------
' For Sp
'----------------------------------------------------------------------------------------------------------
Sub SubBizBatch()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call SubCreateCommandObject(lgObjComm)
    Call SubBizBatchMulti()                            '☜: Run Batch
    Call SubCloseCommandObject(lgObjComm)


    If lgErrorStatus    = "YES" Then
       'lgErrorPos = lgErrorPos & arrColVal(1) & gColSep         
    End If
    
    If lgErrorStatus = "NO"	Then
    	'Call DisplayMsgBox("183114", vbInformation, "", "", I_MKSCRIPT)		'작업이 완료되었습니다 
	End If

End Sub

'============================================================================================================
' Name : SubBizBatchMulti
' Desc : Batch Multi Data
'============================================================================================================
Sub SubBizBatchMulti()

	On Error Resume Next
	Err.Clear
	     
	Dim IntRetCD
	Dim strMsg_cd, strMsg_text
	Dim strSp    
	    
	strSp = "usp_a_bs"
	 
	With lgObjComm
	   .CommandText = strSp
	   .CommandType = adCmdStoredProc
				    
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",	 adInteger,	adParamReturnValue)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@this_start_dt",	 adWChar,	adParamInput,		8,	txtFromGlDt)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@this_end_dt",	 adWChar,	adParamInput,		8,	txtToGlDt)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@pre_start_dt",	 adWChar,	adParamInput,		8,	txtPreFromGlDt)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@pre_end_dt",		 adWChar,	adParamInput,		8,	txtPreToGlDt)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@class_type",		 adVarWChar,	adParamInput,		4, txtClassType)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@biz_area_cd",	 adVarWChar,	adParamInput,		10, txtBizAreaCd)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@hq_brch_fg",		 adWChar,	adParamInput,		1,	lgStrHqBrchFg)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@zero_fg",		 adWChar,	adParamInput,		1,	strZeroFg)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id",			 adVarWChar,	adParamInput,		13,	lgStrUserId)	   
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd",		     adVarWChar,	adParamOutput,		6)	   		  
		   
	   lgObjComm.Execute ,, adExecuteNoRecords	
	End With
   
	If Err.number = 0 Then
	   IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value		
	   If IntRetCD <> 1 then
	      strMsg_cd = lgObjComm.Parameters("@msg_cd").Value
	      If strMsg_Cd <> "" Then
		       Call DisplayMsgBox(strMsg_cd, vbInformation, "", "", I_MKSCRIPT)
		  End If
	      Response.end
	   End If
	        
	Else    
	  lgErrorStatus     = "YES"
	  Call DisplayMsgBox(Err.Description, vbInformation, "", "", I_MKSCRIPT)
	  Call SubHandleError(lgObjComm.ActiveConnection,lgObjRs,Err)
	End if

End Sub

'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
    lgErrorStatus     = "YES"                                                         '☜: Set error status
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)

    On Error Resume Next                                                              '☜: Protect system from crashing
    Err.Clear                                                                         '☜: Clear Error status

    If CheckSYSTEMError(pErr,True) = True Then
       ObjectContext.SetAbort
       Call SetErrorStatus
    Else
       If CheckSQLError(pConn,True) = True Then
          ObjectContext.SetAbort
          Call SetErrorStatus
       End If
   End If

End Sub

%>

<Script Language=vbscript>
 
	With Parent
	
		If "<%=lgDataExist%>" = "Yes" Then
		   'Show multi spreadsheet data from this line		   
		   .Frm1.txtBizAreaNm.Value		  = "<%=ConvSPChars(lgBizAreaNm)%>"    
		   .Frm1.txtClassTypeNm.Value	  = "<%=ConvSPChars(lgClassNm)%>"    		   
		   .ggoSpread.Source  = Parent.frm1.vspdData
		   .ggoSpread.SSShowData "<%=lgstrData%>"                  '☜ : Display data
		   .lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
		   .DbQueryOk
		End If   
    
    End With

</Script>
