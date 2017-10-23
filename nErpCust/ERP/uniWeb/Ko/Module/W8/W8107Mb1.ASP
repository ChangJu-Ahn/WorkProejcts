<%@ LANGUAGE=VBSCript TRANSACTION=Required%>
<% Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/lgSvrVariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../wcm/inc_SvrDebug.asp" -->
<!-- #Include file="../wcm/inc_SvrOperation.asp" -->
<%    
    Call HideStatusWnd                                                               '☜: Hide Processing message
    Call LoadBasisGlobalInf()  
    Call LoadInfTB19029B("I", "H","NOCOOKIE","MB")

    lgErrorStatus     = "NO"
    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    Call SubOpenDB(lgObjConn)                                                        '☜: Make a DB Connection

	Call CheckVersion(lgKeyStream(1), lgKeyStream(2))	' 2005-03-11 버전관리기능 추가 
	
    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)
             Call SubBizQuery()
        Case CStr(UID_M0002)
       
             Call SubBizSave()
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
	 		    
				  </Script>       
				<%        
              Call SetErrorStatus()
    Else 
				%>
				<Script Language=vbscript>
				       Dim i
				       Function Rdo_value_set(byval obj,byval Data)
				        Dim i
				          for i= 0 to obj.length - 1
				                     
	           			             if obj(i).value     =   Data then
	           			                 obj(i).checked = true
	           			                 Exit for
	           			             else
	           			                 obj(i).checked = false  
	           			             end if
				            Next
				       End Function 
				       With Parent	
				       
				              Call Rdo_value_set (.Frm1.rdoW1, "<%=ConvSPChars(lgObjRs("w1"))%>")
				              .frm1.txtw1_Rate.text    =   "<%=ConvSPChars(lgObjRs("w1_rate_view"))%>" 
				              Call Rdo_value_set (.Frm1.rdoW2, "<%=ConvSPChars(lgObjRs("w2"))%>")
				              Call Rdo_value_set (.Frm1.rdoW3, "<%=ConvSPChars(lgObjRs("w3"))%>")
				              Call Rdo_value_set (.Frm1.rdoW4, "<%=ConvSPChars(lgObjRs("w4"))%>")
				             
				                 .frm1.txtw5.text    =   "<%=ConvSPChars(lgObjRs("w5"))%>" 
				                 .frm1.txtw6.text    =   "<%=ConvSPChars(lgObjRs("w6"))%>" 
				              Call Rdo_value_set (.Frm1.rdoW8, "<%=ConvSPChars(lgObjRs("w8"))%>")   
	                          Call Rdo_value_set (.Frm1.rdoW9, "<%=ConvSPChars(lgObjRs("w9"))%>")
	                          Call Rdo_value_set (.Frm1.rdoW10, "<%=ConvSPChars(lgObjRs("w10"))%>")   			             			          
	                          Call Rdo_value_set (.Frm1.rdoW11, "<%=ConvSPChars(lgObjRs("w11"))%>")
	                                 
				                 .frm1.txtw12_1.text    =   "<%=ConvSPChars(lgObjRs("w12_A"))%>" 
				                 .frm1.txtw12_2.text    =   "<%=ConvSPChars(lgObjRs("w12_B"))%>" 
	                          Call Rdo_value_set (.Frm1.rdoW13, "<%=ConvSPChars(lgObjRs("w13"))%>")   			             			          
	                          Call Rdo_value_set (.Frm1.rdoW14, "<%=ConvSPChars(lgObjRs("w14"))%>")   			          
	                          Call Rdo_value_set (.Frm1.rdoW15, "<%=ConvSPChars(lgObjRs("w15"))%>")   
	                            .frm1.txtw16.text    =   "<%=UNIConvNum(lgObjRs("w16"),0)%>" 		
	                            .frm1.txtw17_1.text    =   "<%=UNIConvNum(lgObjRs("w17_1"),0)%>"
	                            .frm1.txtw17_2.text    =   "<%=UNIConvNum(lgObjRs("w17_2"),0)%>" 			          
	                            .frm1.txtw17_sum.text    =   "<%=UNIConvNum(lgObjRs("w17_sum"),0)%>"
	                            .frm1.txtw18_1.text    =   "<%=UNIConvNum(lgObjRs("w18_1"),0)%>"
	                            .frm1.txtw18_2.text    =   "<%=UNIConvNum(lgObjRs("w18_2"),0)%>" 			          
	                            .frm1.txtw18_sum.text    =   "<%=UNIConvNum(lgObjRs("w18_sum"),0)%>" 			          
	                            .frm1.txtw19_1.text    =   "<%=UNIConvNum(lgObjRs("w19_1"),0)%>"
	                            .frm1.txtw19_2.text    =   "<%=UNIConvNum(lgObjRs("w19_2"),0)%>" 			          
	                            .frm1.txtw19_sum.text    =   "<%=UNIConvNum(lgObjRs("w19_sum"),0)%>" 			           			          
	                            .frm1.txtw20_1.text    =   "<%=UNIConvNum(lgObjRs("w20_1"),0)%>"
	                            .frm1.txtw20_2.text    =   "<%=UNIConvNum(lgObjRs("w20_2"),0)%>" 			          
	                            .frm1.txtw20_sum.text    =   "<%=UNIConvNum(lgObjRs("w20_sum"),0)%>" 			          
	                            .frm1.txtw21_1.text    =   "<%=UNIConvNum(lgObjRs("w21_1"),0)%>"
	                            .frm1.txtw21_2.text    =   "<%=UNIConvNum(lgObjRs("w21_2"),0)%>" 			          
	                            .frm1.txtw21_sum.text    =   "<%=UNIConvNum(lgObjRs("w21_sum"),0)%>" 			          
	                            .frm1.txtw22_1.text    =   "<%=UNIConvNum(lgObjRs("w22_1"),0)%>"
	                            .frm1.txtw23_1.value    =   "<%=UNIConvNum(lgObjRs("w23_1"),0)%>"
	                            
	                            ' -- 200603 개정
	                            .frm1.txtW_TYPE.value    =   "<%=ConvSPChars(lgObjRs("W_TYPE"))%>"
	                            .frm1.txtW_TYPE_NM.value    =   "<%=ConvSPChars(lgObjRs("W_TYPE_NM"))%>"
	                             			          
				               
				                
				                 
				                 
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

    lgStrSQL = "DELETE From  TB_1"

    lgStrSQL = lgStrSQL & " where co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S") 
    lgStrSQL = lgStrSQL & "		  and fisc_year =" & FilterVar(Trim(UCase(lgKeyStream(1))),"","S") 
    lgStrSQL = lgStrSQL & "		  and rep_type =" &FilterVar(Trim(UCase(lgKeyStream(2))),"","S") 

	PrintLog "SubBizDelete.. : " & lgStrSQL
    lgObjConn.Execute lgStrSQL,,adCmdText+adExecuteNoRecords
	Call SubHandleError("MC",lgObjConn,lgObjRs,Err)
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
    On Error Resume Next
    Err.Clear
    
    lgStrSQL = "INSERT INTO TB_1("
    lgStrSQL = lgStrSQL & " co_cd, "
    lgStrSQL = lgStrSQL & " FISC_YEAR, "
    lgStrSQL = lgStrSQL & " rep_type, " 
    lgStrSQL = lgStrSQL & " W1, "
    lgStrSQL = lgStrSQL & " W1_Rate_View, "
    lgStrSQL = lgStrSQL & " W1_Rate, " & vbCrLf
    lgStrSQL = lgStrSQL & " W2, "
    lgStrSQL = lgStrSQL & " W3, "
    lgStrSQL = lgStrSQL & " W4, " 
    lgStrSQL = lgStrSQL & " W5, "
    lgStrSQL = lgStrSQL & " W6, " & vbCrLf
    lgStrSQL = lgStrSQL & " W8, "
    lgStrSQL = lgStrSQL & " W9, "
    lgStrSQL = lgStrSQL & " W10, "
    lgStrSQL = lgStrSQL & " W11, "
    lgStrSQL = lgStrSQL & " W12_A, " & vbCrLf
    lgStrSQL = lgStrSQL & " W12_B, "
    lgStrSQL = lgStrSQL & " W13, "
    lgStrSQL = lgStrSQL & " W14, "           
    lgStrSQL = lgStrSQL & " W15, "                          
    lgStrSQL = lgStrSQL & " W16, " & vbCrLf
    lgStrSQL = lgStrSQL & " W17_1, "
    lgStrSQL = lgStrSQL & " W17_2, "    
    lgStrSQL = lgStrSQL & " W17_Sum, " 
    lgStrSQL = lgStrSQL & " W18_1, "
    lgStrSQL = lgStrSQL & " W18_2, " & vbCrLf
    lgStrSQL = lgStrSQL & " W18_Sum, "
    lgStrSQL = lgStrSQL & " W19_1, "
    lgStrSQL = lgStrSQL & " W19_2, "
    lgStrSQL = lgStrSQL & " W19_sum, "
    lgStrSQL = lgStrSQL & " W20_1, " & vbCrLf
    lgStrSQL = lgStrSQL & " W20_2, "
    lgStrSQL = lgStrSQL & " W20_Sum, "
    lgStrSQL = lgStrSQL & " W21_1, "
    lgStrSQL = lgStrSQL & " W21_2, "
    lgStrSQL = lgStrSQL & " W21_Sum, " & vbCrLf
    lgStrSQL = lgStrSQL & " W22_1, "
    lgStrSQL = lgStrSQL & " W23_1, "
    ' -- 200603 : 법인유형별구분 추가
    lgStrSQL = lgStrSQL & " W_TYPE, "
    
    lgStrSQL = lgStrSQL & " INSRT_USER_ID, "
    lgStrSQL = lgStrSQL & " INSRT_DT, "
    lgStrSQL = lgStrSQL & " UPDT_USER_ID, "
    lgStrSQL = lgStrSQL & " UPDT_DT)"
    lgStrSQL = lgStrSQL & " VALUES("
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(0))),"","S")     & ","                      
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(1))),"","S")     & ","           
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(2))),"","S")     & ","   
    lgStrSQL = lgStrSQL & FilterVar(Trim(UCase(lgKeyStream(3))),"","S")     & ","
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(4),0)     & ","       
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(5),0)     & ","   & vbCrLf       
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(6),"''","S")     & "," 
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(7),"''","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(8),"''","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(9),"''","S")     & ","
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(10),"''","S")     & ","  & vbCrLf           
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(11),"''","S")     & ","      
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(12),"''","S")     & ","      
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(13),"''","S")     & ","      
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(14),"''","S")     & ","      
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(15),"Null","S")     & ","  & vbCrLf     
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(16),"Null","S")     & ","      
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(17),"''","S")     & ","      
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(18),"''","S")     & ","      
    lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(19),"''","S")     & ","      

    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(20),0)     & ","  & vbCrLf 
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(21),0)     & ","    
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(22),0)     & ","       
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(23),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(24),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(25),0)     & ","   & vbCrLf
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(26),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(27),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(28),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(29),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(30),0)     & ","  & vbCrLf 
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(31),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(32),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(33),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(34),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(35),0)     & ","  & vbCrLf 
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(36),0)     & ","  
    lgStrSQL = lgStrSQL & UNIConvNum(lgKeyStream(37),0)     & ","  

	' -- 200603 : 법인유형별구분 추가
	lgStrSQL = lgStrSQL & FilterVar(lgKeyStream(38),"''","S")     & ","  
 
    
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","                       
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S") & ","                                     
    lgStrSQL = lgStrSQL & FilterVar(gUsrId,"''","S")                        & ","                     
    lgStrSQL = lgStrSQL & FilterVar(GetSvrDateTime,"''","S")                                            
    lgStrSQL = lgStrSQL & ")" 

	PrintLog "SubBizSaveSingleCreate.. : " & lgStrSQL
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

    lgStrSQL = lgStrSQL & " Update TB_1 set"


     lgStrSQL = lgStrSQL & " W1 =  "& FilterVar(Trim(UCase(lgKeyStream(3))),"","S")      & ","
    lgStrSQL = lgStrSQL & " W1_Rate_View  =  "& UNIConvNum(lgKeyStream(4),0)             & ","       
    lgStrSQL = lgStrSQL & " W1_Rate       =  "&  UNIConvNum(lgKeyStream(5),0)            & ","     
    lgStrSQL = lgStrSQL & " W2			  =  "&  FilterVar(lgKeyStream(6),"''","S")      & "," 
    lgStrSQL = lgStrSQL & " W3			  =  "& FilterVar(lgKeyStream(7),"''","S")       & ","
    lgStrSQL = lgStrSQL & " W4			  =  "& FilterVar(lgKeyStream(8),"''","S")       & ","
    lgStrSQL = lgStrSQL & " W5			  =  "& FilterVar(lgKeyStream(9),"''","S")       & ","
    lgStrSQL = lgStrSQL & " W6			  =  "& FilterVar(lgKeyStream(10),"''","S")      & ","            
    lgStrSQL = lgStrSQL & " W8			  =  "&  FilterVar(lgKeyStream(11),"''","S")     & ","      
    lgStrSQL = lgStrSQL & " W9			  =  "& FilterVar(lgKeyStream(12),"''","S")      & ","      
    lgStrSQL = lgStrSQL & " W10			  =  "&	 FilterVar(lgKeyStream(13),"''","S")     & ","      
    lgStrSQL = lgStrSQL & " W11		      =  "&  FilterVar(lgKeyStream(14),"''","S")     & ","      
    lgStrSQL = lgStrSQL & " W12_A		  =  "& FilterVar(lgKeyStream(15),"Null","S")    & ","      
    lgStrSQL = lgStrSQL & " W12_B		  =  "&  FilterVar(lgKeyStream(16),"Null","S")   & ","      
    lgStrSQL = lgStrSQL & " W13			  =  "&   FilterVar(lgKeyStream(17),"''","S")    & ","         
    lgStrSQL = lgStrSQL & " W14			  =  "&    FilterVar(lgKeyStream(18),"''","S")   & ","                    
    lgStrSQL = lgStrSQL & " W15			  =  "&  FilterVar(lgKeyStream(19),"''","S")     & ","      
    lgStrSQL = lgStrSQL & " W16			  =  "&	 UNIConvNum(lgKeyStream(20),0)           & ","  
    lgStrSQL = lgStrSQL & " W17_1	      =  "&	  UNIConvNum(lgKeyStream(21),0)          & ","    
    lgStrSQL = lgStrSQL & " W17_2		  =  "&  UNIConvNum(lgKeyStream(22),0)           & ","       
    lgStrSQL = lgStrSQL & " W17_Sum		  =  "&   UNIConvNum(lgKeyStream(23),0)          & ","  
    lgStrSQL = lgStrSQL & " W18_1		  =  "&  UNIConvNum(lgKeyStream(24),0)           & ","  
    lgStrSQL = lgStrSQL & " W18_2	      =  "& UNIConvNum(lgKeyStream(25),0)            & ","  
    lgStrSQL = lgStrSQL & " W18_Sum	      =  "&	 UNIConvNum(lgKeyStream(26),0)           & ","  
    lgStrSQL = lgStrSQL & " W19_1	      =  "&  UNIConvNum(lgKeyStream(27),0)           & ","  
    lgStrSQL = lgStrSQL & " W19_2		  =  "&  UNIConvNum(lgKeyStream(28),0)           & ","  
    lgStrSQL = lgStrSQL & " W19_sum	      =  "& UNIConvNum(lgKeyStream(29),0)            & ","  
    lgStrSQL = lgStrSQL & " W20_1	      =  "& UNIConvNum(lgKeyStream(30),0)            & ","  
    lgStrSQL = lgStrSQL & " W20_2	      =  "&  UNIConvNum(lgKeyStream(31),0)           & ","  
    lgStrSQL = lgStrSQL & " W20_Sum	      =  "&   UNIConvNum(lgKeyStream(32),0)          & ","  
    lgStrSQL = lgStrSQL & " W21_1	      =  "&   UNIConvNum(lgKeyStream(33),0)          & ","  
    lgStrSQL = lgStrSQL & " W21_2	      =  "& UNIConvNum(lgKeyStream(34),0)            & ","  
    lgStrSQL = lgStrSQL & " W21_Sum	      =  "& UNIConvNum(lgKeyStream(35),0)            & ","  
    lgStrSQL = lgStrSQL & " W22_1	      =  "& UNIConvNum(lgKeyStream(36),0)            & ","  
    lgStrSQL = lgStrSQL & " W23_1 	      =  "& UNIConvNum(lgKeyStream(37),0)            & ","  

	' -- 200603 : 법인유형별구분 추가
	lgStrSQL = lgStrSQL & " W_TYPE		  =  "&	 FilterVar(lgKeyStream(38),"''","S")     & ","   
	
     lgStrSQL = lgStrSQL & " UPDT_DT   = " &  FilterVar(GetSvrDateTime,"","S")   
    lgStrSQL = lgStrSQL & " where co_cd = " & FilterVar(Trim(UCase(lgKeyStream(0))),"","S") 
    lgStrSQL = lgStrSQL & "  and fisc_year =" & FilterVar(Trim(UCase(lgKeyStream(1))),"","S") 
    lgStrSQL = lgStrSQL & "  and rep_type =" &FilterVar(Trim(UCase(lgKeyStream(2))),"","S") 

	PrintLog "SubBizSaveSingleUpdate.. : " & lgStrSQL
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
            lgStrSQL = lgStrSQL & "  W1,				W1_RATE_VIEW,		W2,				     "
            lgStrSQL = lgStrSQL & "  W3,				W4,					W5,					W6,	      W8,"
            lgStrSQL = lgStrSQL & "  W9,				W10,				W11,				W12_A ,	  W12_B , "
            lgStrSQL = lgStrSQL & "  W13,				W14,				w15,				w16,				 "
            lgStrSQL = lgStrSQL & "  W17_1,				W17_2,				w17_Sum,			w18_1,				 "	
            lgStrSQL = lgStrSQL & "  W18_2,				W18_Sum,			w19_1,				w19_2,				 "
            lgStrSQL = lgStrSQL & "  W19_Sum,			W20_1,   			w20_2,				w20_Sum,				 "
            lgStrSQL = lgStrSQL & "  W21_1,			    W21_2,				w21_sum ,  			w22_1,				w23_1	"
            lgStrSQL = lgStrSQL & " ,  W_TYPE , ( select MINOR_NM from dbo.ufn_TB_MINOR('W1093', '" & C_REVISION_YM & "') where minor_cd = TB_1.W_TYPE) W_TYPE_NM "
            lgStrSQL = lgStrSQL & " FROM  TB_1"
            lgStrSQL = lgStrSQL & " where "
            lgStrSQL = lgStrSQL &   pCode  
 
 
    End Select
    
    PrintLog "SubMakeSQLStatements.. : " & lgStrSQL
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
