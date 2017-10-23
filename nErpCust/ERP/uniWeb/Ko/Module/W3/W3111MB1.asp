<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->

<%   

  Dim varFISC_YEAR
  Dim varRep_type 
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)
    
    varFISC_YEAR      = Trim(UCase(lgKeyStream(1)))
    varRep_type       = Trim(UCase(lgKeyStream(2)))
    Const BIZ_MNU_ID = "W3111MA1"
    
    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection
     Call CheckVersion(varFISC_YEAR   ,varRep_type)	' 2005-03-11 버전관리기능 추가 
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)
             Call SubBizQuery()
        Case CStr(UID_M0002)
       
             Call SubBizSave()
             if lgErrorStatus <> "YES" then
                Call SubBizSave2()
             end if   
        Case CStr(UID_M0003)
             Call SubBizDelete()
    End Select

    Call SubCloseDB(lgObjConn)

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Dim iKey1
    dim strWhere

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    strWhere = " co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")  
	strWhere = strWhere & "  and fisc_year =" & FilterVar(Trim(lgKeyStream(1)),"","S")
	strWhere = strWhere & "  and rep_type =" &  FilterVar(Trim(lgKeyStream(2)),"","S")
	

    Call SubMakeSQLStatements("R",strWhere)                                       '☜ : Make sql statements

    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
          'Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
         
				%>
				<Script Language=vbscript>
				     
				          Parent.FNCnew
						  Parent.DBQueryfalse
				  </Script>       
				<%        
              Call SetErrorStatus()
    Else 

%>
<Script Language=vbscript>
       With Parent	
                .Frm1.txtW1.Value    = "<%=ConvSPChars(lgObjRs("w1"))%>"
                .Frm1.txtW1_A.Value  = "<%=ConvSPChars(lgObjRs("w1_A"))%>"
                .Frm1.txtW1_B.Value  = "<%=ConvSPChars(lgObjRs("w1_B"))%>"
                .Frm1.txtW1_C.Value  = "<%=ConvSPChars(lgObjRs("w1_C"))%>"
                .Frm1.txtW2.Value    = "<%=ConvSPChars(lgObjRs("w2"))%>"
                .Frm1.txtW1_D.Value  = "<%=ConvSPChars(lgObjRs("W1_D"))%>"
                .Frm1.txtW1_E.Value  = "<%=ConvSPChars(lgObjRs("W1_E"))%>"
                .Frm1.txtW1_F.Value  = "<%=ConvSPChars(lgObjRs("W1_F"))%>"
                .Frm1.txtW3.Value    = "<%=ConvSPChars(lgObjRs("w3"))%>"
                .Frm1.txtW4.Value    = "<%=ConvSPChars(lgObjRs("w4"))%>"
                .Frm1.txtW5.Value    = "<%=ConvSPChars(lgObjRs("w5"))%>"
                .Frm1.txtW6.Value    = "<%=ConvSPChars(lgObjRs("w6"))%>"
                .Frm1.txtW7.Value    = "<%=ConvSPChars(lgObjRs("w7"))%>"
                .Frm1.txtW8.Value    = "<%=ConvSPChars(lgObjRs("w8"))%>"
                .Frm1.txtW9.Value    = "<%=ConvSPChars(lgObjRs("w9"))%>"
                .Frm1.txtW10.Value   = "<%=ConvSPChars(lgObjRs("w10"))%>"
		        .DBQueryok
                
                
       End With          
</Script>       
<%     
    End If
    
    Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs)
    
End Sub	
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()
    On Error Resume Next
    Err.Clear
    
    lgIntFlgMode = CInt(Request("txtFlgMode"))                                       '☜: Read Operayion Mode (CREATE, UPDATE)

    Select Case lgIntFlgMode
        Case  OPMD_CMODE        
                                                '☜ : Create
              Call SubBizSaveSingleCreate()  
        Case  OPMD_UMODE           
              Call SubBizSaveSingleUpdate()
    End Select
End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()
    On Error Resume Next
    Err.Clear

    lgStrSQL = "DELETE From  TB_23A"

    lgStrSQL = lgStrSQL & " where co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S") 
    lgStrSQL = lgStrSQL & "  and fisc_year =" & FilterVar(Trim(UCase(lgKeyStream(1))),"","S") 
    lgStrSQL = lgStrSQL & "  and rep_type =" &FilterVar(Trim(UCase(lgKeyStream(2))),"","S") 
   
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
	 Call TB_15_DeleData("1", unicdbl(Request("txtW11"),0) )
	Call TB_15_DeleData("2", unicdbl(Request("txtW19"),0 ))
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
    On Error Resume Next
    Err.Clear
    
    lgStrSQL = "INSERT INTO TB_23A("
    lgStrSQL = lgStrSQL & " co_cd, "
    lgStrSQL = lgStrSQL & " FISC_YEAR, "
    lgStrSQL = lgStrSQL & " rep_type, "
    lgStrSQL = lgStrSQL & " W1, "
    lgStrSQL = lgStrSQL & " W1_A, "
    lgStrSQL = lgStrSQL & " W1_B, "
    lgStrSQL = lgStrSQL & " W1_C, "
    lgStrSQL = lgStrSQL & " W2, "
    lgStrSQL = lgStrSQL & " W1_D, "
    lgStrSQL = lgStrSQL & " W1_E, "
    lgStrSQL = lgStrSQL & " W1_F, "
    lgStrSQL = lgStrSQL & " W3, "
    lgStrSQL = lgStrSQL & " W4, "
    lgStrSQL = lgStrSQL & " W5, "
    lgStrSQL = lgStrSQL & " W6, "
    lgStrSQL = lgStrSQL & " W7, "
    lgStrSQL = lgStrSQL & " W8, "
    lgStrSQL = lgStrSQL & " W9, "
    lgStrSQL = lgStrSQL & " W10, "
    lgStrSQL = lgStrSQL & " INSRT_USER_ID, "
    lgStrSQL = lgStrSQL & " INSRT_DT, "
    lgStrSQL = lgStrSQL & " UPDT_USER_ID, "
    lgStrSQL = lgStrSQL & " UPDT_DT)"
    lgStrSQL = lgStrSQL & " VALUES("
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")     & ","                      
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(1))),"","S")     & ","           
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")     & ","             
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW1"),0)     & ","      
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW1_A"),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW1_B"),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW1_C"),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW2"),0)     & ","    
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW1_D"),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW1_E"),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW1_F"),0)     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW3"),0)     & ","    
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW4"),0)     & ","   
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW5"),0)     & ","   
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW6"),0)     & ","   
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW7"),0)     & ","   
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW8"),0)     & ","   
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW9"),0)     & ","   
    lgStrSQL = lgStrSQL & UNIConvNum(Request("txtW10"),0)     & ","   
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","                       
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","                                     
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","                     
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S")                                            
    lgStrSQL = lgStrSQL & ")" 


    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords

	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    On Error Resume Next
    Err.Clear

    lgStrSQL = lgStrSQL & " Update TB_23A set"

    lgStrSQL = lgStrSQL & " W1  = " & UNIConvNum(Request("txtW1"),0)   & ","  
    lgStrSQL = lgStrSQL & " W1_A  = " & UNIConvNum(Request("txtW1_A"),0)   & "," 
    lgStrSQL = lgStrSQL & " W1_B  = " & UNIConvNum(Request("txtW1_B"),0)   & ","  
    lgStrSQL = lgStrSQL & " W1_C  = " & UNIConvNum(Request("txtW1_C"),0)   & ","  
    lgStrSQL = lgStrSQL & " W2  = " & UNIConvNum(Request("txtW2"),0)   & ","  
    lgStrSQL = lgStrSQL & " W1_D  = " & UNIConvNum(Request("txtW1_D"),0)   & "," 
    lgStrSQL = lgStrSQL & " W1_E  = " & UNIConvNum(Request("txtW1_E"),0)   & ","  
    lgStrSQL = lgStrSQL & " W1_F  = " & UNIConvNum(Request("txtW1_F"),0)   & ","  
    lgStrSQL = lgStrSQL & " W3  = " & UNIConvNum(Request("txtW3"),0)   & ","  
    lgStrSQL = lgStrSQL & " W4		 = " & UNIConvNum(Request("txtW4"),0) & ","  
    lgStrSQL = lgStrSQL & " W5		 = " & UNIConvNum(Request("txtW5"),0)  & ","  	
    lgStrSQL = lgStrSQL & " W6		 = " & UNIConvNum(Request("txtW6"),0)   & ","  
    lgStrSQL = lgStrSQL & " W7		 = " & UNIConvNum(Request("txtW7"),0)  & ","  
    lgStrSQL = lgStrSQL & " W8		 = " & UNIConvNum(Request("txtW8"),0)  & ","  
    lgStrSQL = lgStrSQL & " W9       = " & UNIConvNum(Request("txtW9"),0)   & ","  
    lgStrSQL = lgStrSQL & " W10		 = " & UNIConvNum(Request("txtW10"),0)  & ","  
    lgStrSQL = lgStrSQL & " UPDT_USER_ID   = " & FilterVar(gUsrId,"''","S")   & ","  
    lgStrSQL = lgStrSQL & " UPDT_DT   = " &  FilterVar(GetSvrDateTime,"","S")   
    lgStrSQL = lgStrSQL & " where co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S") 
    lgStrSQL = lgStrSQL & "  and fisc_year =" & FilterVar(Trim(UCase(lgKeyStream(1))),"","S") 
    lgStrSQL = lgStrSQL & "  and rep_type =" &FilterVar(Trim(UCase(lgKeyStream(2))),"","S") 
  

    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords

	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode,pCode)


    Select Case pMode 
      Case "R"
            lgStrSQL = "    Select TOP 1  "
            lgStrSQL = lgStrSQL & " W1,       W1_A,		W1_B,		W1_C,       W2,		"
            lgStrSQL = lgStrSQL & " W1_D,     W1_E,		W1_F,		W3,         W4,		"
            lgStrSQL = lgStrSQL & " W5,       W6,		W7,		    W8,         W9,		"
            lgStrSQL = lgStrSQL & " W10	"
            lgStrSQL = lgStrSQL & " FROM  TB_23A "
            lgStrSQL = lgStrSQL & " where "
            lgStrSQL = lgStrSQL &   pCode  
 

 
    End Select
End Sub




Sub TB_15_PushData(Byval pSeqNo, Byval pAmt)
	On Error Resume Next 
	Err.Clear  
	Dim  wgCO_CD, sFISC_YEAR,sREP_TYPE
	
	wgCO_CD    = Trim(UCase(lgKeyStream(0)))
	sFISC_YEAR = Trim(lgKeyStream(1))
	sREP_TYPE  =Trim(lgKeyStream(2))

	Select Case pSeqNo
		Case "1"
		
		'접대비 6
			
				
			 lgStrSQL = "EXEC usp_TB_15_PushData "
			 lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ", "		' 법인 
			 lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ", "	' 사업연도 
			 lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ", "		' 신고구분 
			 lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(BIZ_MNU_ID)),"''","S") & ", "	' 전송자의 프로그램 
			 lgStrSQL = lgStrSQL & FilterVar(UNICDbl(pSeqNo, "0"),"0","D") & ", "		' 전송자의 순번 
			 lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("1")),"''","S") & ", "			' 1호/2호 
			 lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("2301")),"''","S") & ", "		' 과목 코드 '접대비 
			 lgStrSQL = lgStrSQL & FilterVar(pAmt,"0","D")  & ", "			' 금액 
			 lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("500")),"''","S") & ", "			' 기타사유유출 
			 lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("접대비중 5만원 초과분 중 신용카드 미사용액을 손금불산입하고 기타 사유출 처분함.")),"''","S") & ", "			' 조정내용 
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & " "
		Case "2"
			 lgStrSQL = "EXEC usp_TB_15_PushData "
			 lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ", "		' 법인 
			 lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ", "	' 사업연도 
			 lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ", "		' 신고구분 
			 lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(BIZ_MNU_ID)),"''","S") & ", "	' 전송자의 프로그램 
			 lgStrSQL = lgStrSQL & FilterVar(UNICDbl(pSeqNo, "0"),"0","D") & ", "		' 전송자의 순번 
			 lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("1")),"''","S") & ", "			' 1호/2호 
			 lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("2301")),"''","S") & ", "		' 과목 코드 '접대비 
			 lgStrSQL = lgStrSQL & FilterVar(pAmt,"0","D")  & ", "			' 금액 
			 lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("500")),"''","S") & ", "			' 기타사유유출 
			 lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("접대비 한도초과액글 손금불산입하고 기타사외유출로 처분함.")),"''","S") & ", "			' 조정내용 
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & " "

		
	End Select

	PrintLog "TB_15_PushData = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
		
End Sub





Sub TB_15_DeleData(Byval pSeqNo, Byval pAmt)
	On Error Resume Next 
	Err.Clear  
   	Dim  wgCO_CD, sFISC_YEAR,sREP_TYPE
	
 
	wgCO_CD    = Trim(UCase(lgKeyStream(0)))
	sFISC_YEAR = Trim(lgKeyStream(1))
	sREP_TYPE  =Trim(lgKeyStream(2))
	
	Select Case pSeqNo
		Case "1" 
			lgStrSQL = "EXEC usp_TB_15_DeleData "
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ", "		' 법인 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ", "	' 사업연도 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ", "		' 신고구분 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(BIZ_MNU_ID)),"''","S") & ", "	' 전송자의 프로그램 
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(pSeqNo, "0"),"0","D") & ", "		' 전송자의 순번 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("1")),"''","S") & ", "			' 1호/2호 
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & " "
			
		
			
		Case "2"
			lgStrSQL = "EXEC usp_TB_15_DeleData "
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(wgCO_CD)),"''","S") & ", "		' 법인 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sFISC_YEAR)),"''","S") & ", "	' 사업연도 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(sREP_TYPE)),"''","S") & ", "		' 신고구분 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(BIZ_MNU_ID)),"''","S") & ", "	' 전송자의 프로그램 
			lgStrSQL = lgStrSQL & FilterVar(UNICDbl(pSeqNo, "0"),"0","D") & ", "		' 전송자의 순번 
			lgStrSQL = lgStrSQL & FilterVar(Trim(UCase("1")),"''","S") & ", "			' 1호/2호 
			lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S") & " "
	End Select

	PrintLog "TB_15_DeleData = " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MU",lgObjConn,lgObjRs,Err)
		
End Sub


Sub  SubBizSave2()

		Call TB_15_DeleData("1", unicdbl(Request("txtW7"),"0" ))
		if  unicdbl(Request("txtW7"),"0") > 0 then
			Call TB_15_PushData("1", unicdbl(Request("txtW7"),"0"))
		end if	
        Call TB_15_DeleData("2", unicdbl(Request("txtW9"),"0" ))
    if  unicdbl(Request("txtW9"),"0") > 0 then
		Call TB_15_PushData("2", unicdbl(Request("txtW9"),"0"))
   end if			
	
	
	
   
   
End Sub
'============================================================================================================
' Name : CommonOnTransactionCommit
' Desc : This Sub is called by OnTransactionCommit Error handler
'============================================================================================================
Sub CommonOnTransactionCommit()
End Sub

'============================================================================================================
' Name : CommonOnTransactionAbort
' Desc : This Sub is called by OnTransactionAbort Error handler
'============================================================================================================
Sub CommonOnTransactionAbort()
    lgErrorStatus    = "YES"
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()

    lgErrorStatus     = "YES"
End Sub

'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next
    Err.Clear

    Select Case pOpCode
        Case "MC"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       
                       Call SetErrorStatus
                    End If
                 End If
        Case "MD"
        Case "MR"
        Case "MU"
                 If CheckSYSTEMError(pErr,True) = True Then
                    ObjectContext.SetAbort
                    Call SetErrorStatus
                 Else
                    If CheckSQLError(pConn,True) = True Then
                       ObjectContext.SetAbort
                       Call SetErrorStatus
                    End If
                 End If
    End Select
End Sub

%>

<Script Language="VBScript">
    Select Case "<%=lgOpModeCRUD %>"
       Case "<%=UID_M0001%>"

          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBQueryOk        
          Else
			Parent.DBQueryFalse        
          End If   
       Case "<%=UID_M0002%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DBSaveOk
          End If   
       Case "<%=UID_M0003%>"
          If Trim("<%=lgErrorStatus%>") = "NO" Then
             Parent.DbDeleteOk
          End If   
    End Select    
       
</Script>	
