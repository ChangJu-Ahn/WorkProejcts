<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->
<%

On Error Resume Next
Err.Clear

Call LoadBasisGlobalInf() 

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 ,rs1, rs2                    '�� : DBAgent Parameter ���� 
'Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
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

Dim lgSp_Id

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ�		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)				
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL					

Const C_SHEETMAXROWS_D  = 100

    Call HideStatusWnd 


    lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value
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
 	
	' ���Ѱ��� �߰� 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))		
	lgInternalCd		= Trim(Request("lgInternalCd"))	
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))	
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))   
    
    'ó����ȸ�ÿ��� sp�� ȣ���ϵ����Ѵ�.------
    If CDbl(lgPageNo) < 1 Then		
		Call SubOpenDB(lgObjConn)                                                        '��: Make a DB Connection
		Call SubBizBatch()    
		Call SubCloseDB(lgObjConn)                                                       '��: Close DB Connection   
	Else
		lgSp_Id			=Trim(Request("lgSp_Id"))
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
    
    lgDataExist    = "Yes"
    lgstrData      = ""

    If CDbl(lgPageNo) > 0 Then
       rs0.Move     = CDbl(C_SHEETMAXROWS_D) * CDbl(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
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

    If  iLoopCount < C_SHEETMAXROWS_D Then                                            '��: Check if next data exists
        lgPageNo = ""                                                  '��: ���� ����Ÿ ����.
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

    Redim UNISqlId(2)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
	
    Redim UNIValue(2,2)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 

    UNISqlId(0) = "a5109QA101"
    UNISqlId(1) = "A_GetBiz"
    UNISqlId(2) = "A_CLSTYPE"
    UNIValue(0,1) = FilterVar(lgSp_Id,"","S")

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    If txtBizAreaCd = "" Then	 	 	
	 	UNIValue(1,0)  = FilterVar("", "''", "S")	 	
	Else				
		UNIValue(1,0)  =  FilterVar(txtBizAreaCd, "''", "S") 
	End If    

	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then			
		UNIValue(1,0)  = UNIValue(1,0) & " AND BIZ_AREA_CD LIKE " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	UNIValue(2,0)  =  FilterVar(txtClassType, "''", "S") 
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))    
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
    Dim iStr
    Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    
    Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
    
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
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    Call SubCreateCommandObject(lgObjComm)
    Call SubBizBatchMulti()                            '��: Run Batch
    Call SubCloseCommandObject(lgObjComm)


    If lgErrorStatus    = "YES" Then
       'lgErrorPos = lgErrorPos & arrColVal(1) & gColSep         
    End If
    
    IF lgErrorStatus = "NO"	Then
    	'Call DisplayMsgBox("183114", vbInformation, "", "", I_MKSCRIPT)		'�۾��� �Ϸ�Ǿ����ϴ� 
	END IF
End Sub



'============================================================================================================
' Name : SubBizBatchMulti
' Desc : Batch Multi Data
'============================================================================================================
Sub SubBizBatchMulti()
	On Error Resume NEXT
	Err.Clear
	     
	Dim IntRetCD
	Dim strMsg_cd, strMsg_text
	Dim strSp    
	
	strSp = "usp_a_mcs_np06"
	
	'���� ���� �߰� 
	Dim BizAreaCd
	
	BizAreaCd = txtBizAreaCd
	
	If txtBizAreaCd = "" Then	 	 	
		If lgAuthBizAreaCd <> "" Then			
			BizAreaCd = lgAuthBizAreaCd
		End If
	Else
		If lgAuthBizAreaCd <> "" Then
			If UCASE(lgAuthBizAreaCd) <> UCASE(txtBizAreaCd) Then
		        Call DisplayMsgBox("124200", vbInformation, "", "", I_MKSCRIPT)
				Response.end
			End If
		End If
	End If
			 
	With lgObjComm
	   .CommandText = strSp
	   .CommandType = adCmdStoredProc
				    
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",	 adInteger,	adParamReturnValue)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@this_from_mnth",	 adWChar,	adParamInput,		6,	txtFromGlDt)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@this_to_mnth",	 adWChar,	adParamInput,		6,	txtToGlDt)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@pre_from_mnth",	 adWChar,	adParamInput,		6,	txtPreFromGlDt)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@pre_to_mnth",	 adWChar,	adParamInput,		6,	txtPreToGlDt)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@class_type",		 adVarWChar,	adParamInput,		20, txtClassType)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@biz_area_cd",	 adVarWChar,	adParamInput,		10, BizAreaCd)	   
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@hq_brch_fg",		 adWChar,	adParamInput,		1,	lgStrHqBrchFg)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@zero_fg",		 adWChar,	adParamInput,		1,	strZeroFg)
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@usr_id",			 adVarWChar,	adParamInput,		13,	lgStrUserId)	   
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@msg_cd",		     adVarWChar,	adParamOutput,		6)	   		  
	   lgObjComm.Parameters.Append lgObjComm.CreateParameter("@sp_id",		     adVarWChar,	adParamOutput,		13)	   		  
		   
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
		Else
			lgSp_Id = lgObjComm.Parameters("@sp_id").Value
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
    lgErrorStatus     = "YES"                                                         '��: Set error status
End Sub
'============================================================================================================
' Name : SubHandleError
' Desc : This Sub handle error
'============================================================================================================
Sub SubHandleError(pOpCode,pConn,pRs,pErr)
    On Error Resume Next                                                              '��: Protect system from crashing
    Err.Clear                                                                         '��: Clear Error status
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
		   
			Parent.strSp_Id		= "<%=lgSp_Id%>"

		   .Frm1.txtBizAreaCd.Value		  = "<%=ConvSPChars(lgBizAreaCd)%>"    
		   .Frm1.txtBizAreaNm.Value		  = "<%=ConvSPChars(lgBizAreaNm)%>"    
		   .Frm1.txtClassNm.Value			  = "<%=ConvSPChars(lgClassNm)%>"    		   
		   .ggoSpread.Source  = Parent.frm1.vspdData
		   .ggoSpread.SSShowData "<%=lgstrData%>"                  '�� : Display data
		   .lgPageNo      =  "<%=lgPageNo%>"               '�� : Next next data tag
		   .DbQueryOk
		End If   
    
    End With

</Script>	

