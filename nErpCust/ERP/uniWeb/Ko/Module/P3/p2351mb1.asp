<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->	
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!--
======================================================================================================
*  1. Module Name          : Template
*  2. Function Name        : 
*  3. Program ID           : p2351mb1.asp
*  4. Program Name         : MRP예시전개전환 
*  5. Program Desc         : 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/10/09
*  8. Modified date(Last)  : 
*  9. Modifier (First)     : Im Hyun Soo
* 10. Modifier (Last)      : Jung Yu Kyung	
* 11. Comment              :
=======================================================================================================-->
<% 

Call LoadBasisGlobalInf
Call HideStatusWnd
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")  

    On Error Resume Next
    Err.Clear
	
	Const C_SHEETMAXROWS_D  = 100
	
    '---------------------------------------Common-----------------------------------------------------------
    lgErrorStatus     = "NO"
    lgErrorPos        = ""
    lgOpModeCRUD      = Request("txtMode")
    
    lgKeyStream       = Split(Request("txtKeyStream"),gColSep)

    lgLngMaxRow       = Request("txtMaxRows") 
    lgMaxCount        = C_SHEETMAXROWS_D
    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0) 

	Dim IntRetCd
	Dim CurDate
	Dim lgPos
	Dim iNfFlg
	
	lgPos = Request("txtSpreadPos")
	
	CurDate = GetSvrDate

    Call SubOpenDB(lgObjConn) 

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)      			'☜: 전체Query
			If lgPos = "0" Then
				Call SubBizQuery("P")
				If IntRetCd <> -1 Then
					Call SubBizQuery("I")
					Call SubBizQuery("M")
					Call SubBizQueryMulti()			

					If iNfFlg = True Then
						Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
					End If
				End If
			Else
				Call SubBizQueryMulti()
				
			End If
		Case CStr(UID_M0002)      
             Call SubBizSaveMulti(lgOpModeCRUD)
             
        Case CStr(UID_M0003)   
			Call SubCreateCommandObject(lgObjComm)
			Call SubBizSaveMulti(lgOpModeCRUD)
			
			Call SubBizBatch()
			
			Call SubCloseCommandObject(lgObjComm)
    End Select
         
    Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery(pQryMode)

	On Error Resume Next
    Err.Clear
    	
	Dim strPlantCd
	Dim iKey1, iKey2
		
	iKey1 = FilterVar(lgKeyStream(0),"''","S")
	iKey2 = FilterVar(lgKeyStream(1),"''","S")
	
	Select Case pQryMode
		Case "P"
				
			'--------------
			'공장 체크		
			'--------------	
			lgStrSQL = ""
			Call SubMakeSQLStatements("P",iKey1,"","","","")

			If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then 

				Call DisplayMsgBox("125000", vbInformation, "", "", I_MKSCRIPT)
				Call SetErrorStatus()
				IntRetCd = -1
				
				Response.Write "<Script Language=VBScript>" & vbCrLf
					Response.Write "parent.frm1.txtPlantNm.value = """"" & vbCrLf
					Response.Write "parent.frm1.txtPlantCd.focus" & vbCrLf
				Response.Write "</Script>" & vbCrLf

			Else
				IntRetCd = 1
				
				Response.Write "<Script Language=VBScript>" & vbCrLf
					Response.Write "parent.frm1.txtPlantNm.value = """ & ConvSPChars(lgObjRs(1)) & """" & vbCrLf
				Response.Write "</Script>" & vbCrLf

			End If
		Case "I"
			IntRetCd = 1
				
			'--------------
			'품목명 
			'--------------	
			lgStrSQL = ""
			Call SubMakeSQLStatements("I",iKey2,iKey1,"","","") 

			If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then 
				
				Response.Write "<Script Language=VBScript>" & vbCrLf
					Response.Write "parent.frm1.txtItemNm.value = """"" & vbCrLf
				Response.Write "</Script>" & vbCrLf
			Else
				Response.Write "<Script Language=VBScript>" & vbCrLf
					Response.Write "parent.frm1.txtItemNm.value = """ & ConvSPChars(lgObjRs(1)) & """" & vbCrLf
				Response.Write "</Script>" & vbCrLf

			End If
		Case "M"
			IntRetCd = 1
				
			'--------------
			'공장정보 
			'--------------	
			lgStrSQL = ""
			Call SubMakeSQLStatements("M",iKey1,"","","","") 

			If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
%>
				<Script Language=vbscript>

					With Parent	
						.Frm1.txtFixExecFromDt.text		= ""
						.frm1.txtFixExecToDt.text		= ""
						.frm1.txtPlanExecToDt.text		= ""
						.frm1.rdoAvailInvFlg1.checked	= True
						.frm1.rdoSafeInvFlg1.checked	= True
						.frm1.rdoMpsConfirmFlg1.checked = True
				    End With 

				</Script>       
<%		
			Else
%>
				<Script Language=vbscript>

					With Parent	
						.Frm1.txtFixExecFromDt.text = "<%= UNIDateClientFormat(lgObjRs(9)) %>"
						.frm1.txtFixExecToDt.text	= "<%= UNIDateClientFormat(lgObjRs(10)) %>"
						.frm1.txtPlanExecToDt.text	= "<%= UNIDateClientFormat(lgObjRs(11)) %>"
						
						If "<%= UCase(lgObjRs(6)) %>" = "Y" Then
							.frm1.rdoSafeInvFlg1.checked	= True
						Else
							.frm1.rdoSafeInvFlg2.checked	= True
						End If
						
						If "<%= UCase(lgObjRs(7)) %>" = "Y" Then
							.frm1.rdoAvailInvFlg1.checked	= True
						Else
							.frm1.rdoAvailInvFlg2.checked	= True
						End If
						
						If "<%= UCase(lgObjRs(8)) %>" = "Y" Then
							.frm1.rdoMpsConfirmFlg2.checked = True
						ElseIf "<%= UCase(lgObjRs(8)) %>" = "N" Then
							.frm1.rdoMpsConfirmFlg3.checked = True
						Else
							.frm1.rdoMpsConfirmFlg1.checked = True
						End If
						
				    End With  
				            
				</Script>       
<%			
			End If

	End Select
					
	Call SubCloseRs(lgObjRs) 
   
End Sub    
'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
    Dim iDx
    Dim iLoopMax
    Dim iKey1, iKey2, iKey3, iKey4, iKey5
    Dim FirstDt_D, LastDt_D
    Dim FirstDt, LastDt
        
    On Error Resume Next
    Err.Clear 
    
    iKey1 = FilterVar(lgKeyStream(0),"''", "S")
    iKey2 = FilterVar(lgKeyStream(1),"''", "S")
    iKey3 = FilterVar(lgKeyStream(2),"''", "S")

    FirstDt_D = FilterVar(UNIConvYYYYMMDDToDate(gServerDateFormat,"1900","01","01"),"''", "S")
    FirstDt = UniConvDateAToB(lgKeyStream(3), gDateFormat, gServerDateFormat)
    iKey4 = FilterVar(FirstDt, FirstDt_D, "S")
    
    LastDt_D = FilterVar(UNIConvYYYYMMDDToDate(gServerDateFormat,"2999","12","31"),"''", "S")
    LastDt = UniConvDateAToB(lgKeyStream(4), gDateFormat, gServerDateFormat)
    iKey5 = FilterVar(LastDt, LastDt_D, "S")
        
    Call SubMakeSQLStatements("MI",iKey1,iKey2,iKey3,iKey4,iKey5)
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
        lgStrPrevKeyIndex = ""
        iNfFlg = True
        lgErrorStatus = "YES"
        If lgPos = "1" Then
			Call SetErrorStatus()
		End If
    Else
		iNfFlg = False
		
		lgLngMaxRow       = Request("txtMaxRows") 
		lgMaxCount        = C_SHEETMAXROWS_D
		
        Call SubSkipRs(lgObjRs,lgMaxCount * lgStrPrevKeyIndex)

        lgstrData = ""
        
        iDx       = 1
        
        Do While Not lgObjRs.EOF   
			lgstrData = lgstrData & Chr(11) & ""					
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(0))			'품목 
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs(10))			'품목명 
			lgstrData = lgstrData & Chr(11) & lgObjRs(13)						'규격 
            lgstrData = lgstrData & Chr(11) & UniDateClientFormat(lgObjRs(3))	'시작일	
            lgstrData = lgstrData & Chr(11) & UniDateClientFormat(lgObjRs(4))	'완료일 
            lgstrData = lgstrData & Chr(11) & UniConvNumberDBToCompany(lgObjRs(5),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)
            lgstrData = lgstrData & Chr(11) & lgObjRs(7)						'단위 
            lgstrData = lgstrData & Chr(11) & lgObjRs(9)						'조달구분명 
            lgstrData = lgstrData & Chr(11) & lgObjRs(11)						'status	
            lgstrData = lgstrData & Chr(11) & lgObjRs(2)						'tracking no
            lgstrData = lgstrData & Chr(11) & lgObjRs(16)						'생산담당자 
            lgstrData = lgstrData & Chr(11) & lgObjRs(15)						'구매조직 
            lgstrData = lgstrData & Chr(11) & lgObjRs(14)						'생산담당자 
            lgstrData = lgstrData & Chr(11) & lgObjRs(12)						'조달구분 
            lgstrData = lgstrData & Chr(11) & lgObjRs(17)						'seq
            lgstrData = lgstrData & Chr(11) & lgLngMaxRow + iDx
            lgstrData = lgstrData & Chr(11) & Chr(12)
            
		    lgObjRs.MoveNext

            iDx =  iDx + 1
            If iDx > lgMaxCount Then
               lgStrPrevKeyIndex = lgStrPrevKeyIndex + 1
               Exit Do
            End If   
               
        Loop 
        
    End If
    
    If iDx <= lgMaxCount Then
       lgStrPrevKeyIndex = ""
    End If   

	Call SubHandleError("MR",lgObjConn,lgObjRs,Err)
    Call SubCloseRs(lgObjRs) 
 
End Sub    

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti(lgOpModeCRUD)

    Dim arrRowVal
    Dim arrColVal
    Dim iDx

    On Error Resume Next 
    Err.Clear 
    
	arrRowVal = Split(Request("txtSpread"), gRowSep)
	
    For iDx = 1 To lgLngMaxRow
        arrColVal = Split(arrRowVal(iDx-1), gColSep)
        
        Call SubBizSaveMultiUpdate(arrColVal, lgOpModeCRUD) 
        
        If lgErrorStatus    = "YES" Then
           lgErrorPos = lgErrorPos & arrColVal(1) & gColSep
           Exit For
        End If
        
    Next

End Sub    

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pCode2,pCode3,pCode4)

    Select Case pDataType
		Case "P"
			lgStrSQL = "SELECT * " 
            lgStrSQL = lgStrSQL & " FROM  b_plant "
            lgStrSQL = lgStrSQL & " WHERE plant_cd = " & pCode
		Case "I"
        	lgStrSQL = "SELECT * " 
            lgStrSQL = lgStrSQL & " FROM  b_item a, b_item_by_plant b"
            lgStrSQL = lgStrSQL & " WHERE a.item_cd = b.item_cd "
            lgStrSQL = lgStrSQL & " AND b.plant_cd = " & pCode1
            lgStrSQL = lgStrSQL & " AND a.item_cd = " & pCode
		Case "M"
			lgStrSQL = "SELECT * " 
            lgStrSQL = lgStrSQL & " FROM  p_mrp_history_s "
            lgStrSQL = lgStrSQL & " WHERE plant_cd = " & pCode
            lgStrSQL = lgStrSQL & " AND run_no = (SELECT max(run_no) FROM p_mrp_history_s WHERE plant_cd = " & pCode & ")"            
        Case "MI" 
			lgStrSQL = "SELECT a.item_cd, a.bom_no, a.tracking_no, a.start_dt, a.due_dt, a.plan_qty, a.mps_no, "
			lgStrSQL = lgStrSQL & " c.basic_unit, a.pm_odr_flg, d.minor_nm, c.item_nm, a.status, b.procur_type, "
			lgStrSQL = lgStrSQL & " c.spec, b.prod_mgr, b.pur_org, e.minor_nm, a.key_temp "  
            lgStrSQL = lgStrSQL & " FROM  p_planned_order_temp_s a, b_item_by_plant b, b_item c , b_minor d, b_minor e"
            lgStrSQL = lgStrSQL & " WHERE a.plant_cd = b.plant_cd AND a.item_cd = b.item_cd AND a.item_cd = c.item_cd "
            lgStrSQL = lgStrSQL & " AND d.major_cd = 'P1003' AND d.minor_cd = b.procur_type "
            lgStrSQL = lgStrSQL & " AND e.major_cd = 'P1015' AND e.minor_cd =* b.prod_mgr "
            lgStrSQL = lgStrSQL & " AND a.plant_cd = " & pCode
            If pCode1 <> "''" Then 
				lgStrSQL = lgStrSQL & " AND a.item_cd >= " & pCode1
			End If
			IF pCode2 <> "''" Then
				lgStrSQL = lgStrSQL & " AND b.procur_type = " & pCode2 
			End If
			lgStrSQL = lgStrSQL & " AND a.start_dt >= " & pCode3
			lgStrSQL = lgStrSQL & " AND a.start_dt <= " & pCode4
			
            lgStrSQL = lgStrSQL & " ORDER BY a.item_cd, a.due_dt "
		
    End Select

End Sub

'============================================================================================================
' Name : SubBizSaveMultiUpdate
' Desc : Update Data from Db
'============================================================================================================
Sub SubBizSaveMultiUpdate(arrColVal, lgOpModeCRUD)

    On Error Resume Next 
    Err.Clear
    
    Dim strStartDt, strDueDt, strPlanQty
    Dim strSeq
    
    If lgOpModeCRUD = CStr(UID_M0002) Then
		strStartDt = UniConvDate(Trim(UCase(arrColVal(3))))
		strDueDt = UniConvDate(Trim(UCase(arrColVal(4))))
		strPlanQty = UniConvNum(arrColVal(5),0)
		
		If UniConvDateToYYYYMMDD(arrColVal(3),gDateFormat,"") > UniConvDateToYYYYMMDD(arrColVal(4),gDateFormat,"") Then
			Call DisplayMsgBox("970023", VBOKOnly, "시작일", "완료일", I_MKSCRIPT)
			Call SheetFocus(arrVal(1), 7, I_MKSCRIPT)
			Response.End
		End If

		lgStrSQL = "UPDATE  P_PLANNED_ORDER_TEMP_S"
		lgStrSQL = lgStrSQL & " SET " 
		lgStrSQL = lgStrSQL & " START_DT = " &  FilterVar(strStartDt,NULL,"S")   & ","
		lgStrSQL = lgStrSQL & " DUE_DT = " &  FilterVar(strDueDt,NULL,"S")   & ","
		lgStrSQL = lgStrSQL & " PLAN_QTY = " &  FilterVar(strPlanQty,"''","S")  
		lgStrSQL = lgStrSQL & " WHERE        "
		lgStrSQL = lgStrSQL & " KEY_TEMP = " &  FilterVar(Trim(UCase(arrColVal(2))),"''","S")

		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords 
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

	
	Else
		strSeq = arrColVal(0)    
		
	    '---------- Developer Coding part (Start) ---------------------------------------------------------------
		'A developer must define field to update record
		'--------------------------------------------------------------------------------------------------------
		lgStrSQL = "UPDATE  P_PLANNED_ORDER_TEMP_S"
		lgStrSQL = lgStrSQL & " SET " 
		lgStrSQL = lgStrSQL & " STATUS = " &  FilterVar("CL",NULL,"S")
		lgStrSQL = lgStrSQL & " WHERE        "
		lgStrSQL = lgStrSQL & " KEY_TEMP = " &  arrColVal(0)

		'---------- Developer Coding part (End  ) ---------------------------------------------------------------
		lgObjConn.Execute lgStrSQL,,adCmdText + adExecuteNoRecords 
		Call SubHandleError("MU",lgObjConn,lgObjRs,Err)

	End If
    			

End Sub
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizBatch()

	Dim strMsg_cd
    Dim strMsg_text
	Dim strPlantCd, strSafeFlg, strInvFlg, strDate
	Dim strIdepFlg
	
    On Error Resume Next
    Err.Clear

	strPlantCd = UCase(Trim(Request("txtPlantCd")))										'☆: Plant Code    
    
	If Request("rdoSafeInvFlg") = "Y" Then
         strSafeFlg  = "Y"
    Else
    	 strSafeFlg  = "N"
    End If

    strInvFlg  = "M"
    strIdepFlg = "S"
    strDate = UniConvDateToYYYYMMDD(GetSvrDate,gServerDateFormat,"")

	With lgObjComm
	    .CommandText = "usp_mmpc000pc"
	    .CommandType = adCmdStoredProc
	        
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("RETURN_VALUE",	adInteger,adParamReturnValue)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@orgid",	advarXchar,adParamInput,4, strPlantCd)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@emp_no", advarXchar,adParamInput,13, UCASE(gUsrId))
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@vrp_chk", advarXchar,adParamInput,1, strSafeFlg)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@option_flg_2", advarXchar,adParamInput,1,strIdepFlg)
	    lgObjComm.Parameters.Append lgObjComm.CreateParameter("@string_entdt", advarXchar,adParamInput,Len(strDate),strDate)
		    
	    lgObjComm.Execute ,, adExecuteNoRecords
	     
	End With

	If  Err.number = 0 Then
	    IntRetCD = lgObjComm.Parameters("RETURN_VALUE").Value
	        
	    If  IntRetCD <> 0 then
			Call SvrMsgBox(Err.Description, vbinformation, i_mkscript)
			Call SubHandleError("MB",lgObjComm.ActiveConnection,lgObjRs,Err)
	        IntRetCD = -1
	    Else
			Call DisplayMsgBox("183114", vbInformation, "", "", I_MKSCRIPT) 
			IntRetCD = 1
	    End if
	Else           
	    Call SvrMsgBox(Err.Description, vbinformation, i_mkscript)
	    Call SubHandleError("MB",lgObjComm.ActiveConnection,lgObjRs,Err)
	    IntRetCD = -1
	End if
	
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
        Case "MD"
        Case "MR"
        Case "MU"
        Case "MB"
    End Select
End Sub
		
%>
<Script Language="VBScript">

	Select Case <%=lgOpModeCRUD%>
		Case "<%=UID_M0001%>"
			If Trim("<%=lgErrorStatus%>") = "NO" Then
				With Parent
					If "<%=lgPos%>" = "0" Or "<%=lgPos%>" = "1" Then
					    .ggoSpread.Source     = .frm1.vspdData
					    .lgStrPrevKeyIndex    = "<%=lgStrPrevKeyIndex%>"
					    .ggoSpread.SSShowData "<%=ConvSPChars(lgstrData)%>"
					End If

					If "<%=lgPos%>" = "0" Then
						.frm1.hPlantCd.value	= "<%=lgKeyStream(0)%>"
						.frm1.hItemCd.value		= "<%=lgKeyStream(1)%>"
						.frm1.hProcType.value	= "<%=lgKeyStream(2)%>"
						.frm1.hBaseFromDt.value = "<%=lgKeyStream(3)%>"
						.frm1.hBaseToDt.value	= "<%=lgKeyStream(4)%>"

					End If
					
					.DBQueryOk(lgLngMaxRow + 1) 

				 End with
			Else
				Parent.DBQueryNotOk() 
	        End If   
		Case "<%=UID_M0002%>"
			If Trim("<%=lgErrorStatus%>") = "NO" Then
			   Parent.DBSaveOk
			Else
			   Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
			End If   	       
        Case "<%=UID_M0003%>"     
			If Trim("<%=lgErrorStatus%>") = "NO" Then
			   Parent.DBSaveOk
			Else
			   Parent.SubSetErrPos(Trim("<%=lgErrorPos%>"))
			End If   	       
	End Select    
       
</Script>	
