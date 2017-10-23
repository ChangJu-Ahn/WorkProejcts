<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<%
'**********************************************************************************************
'*  1. Module Name          : Accounting
'*  2. Function Name        : Fixed Asset Management
'*  3. Program ID           : a7109mb1(고정자산부서이동등록) - 자산부서별정보 내역부분 
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Comproxy List        : +B19029LookupNumericFormat
'                             +B25011ManagePlant
'                             +B25011ManagePlant
'                             +B25018ListPlant
'                             +B25019LookUpPlant
'*  7. Modified date(First) : 2000/3/21
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Kim Hee Jung
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
    
    Call LoadBasisGlobalInf()
	Dim  lgOpModeCRUD
	Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")
    '---------------------------------------Common-----------------------------------------------------------
'    lgErrorStatus     = "NO"
'    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

	'------ Developer Coding part (Start ) ------------------------------------------------------------------

'    lgLngMaxRow       = Request("txtMaxRows")                                        '☜: Read Operation Mode (CRUD)
'    lgMaxCount        = CInt(Request("lgMaxCount"))                                '☜: Fetch count at a time for VspdData
'    lgStrPrevKeyIndex = UNICInt(Trim(Request("lgStrPrevKeyIndex")),0)                '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
'        Case CStr(UID_M0002)                                                         '☜: Save,Update
'             Call SubBizSave()
'        Case CStr(UID_M0003)                                                         '☜: Delete
'             Call SubBizDelete()
    End Select
    'Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection
Response.End
'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()
    Const A519_I1_dept_cd = 0    
    Const A519_I2_asst_no = 0    
    Const A519_I3_org_change_id = 0    

    Const A519_E1_asst_no = 0    
    Const A519_E1_asst_nm = 1
    Const A519_E2_dept_cd = 2    

    Const A519_EG1_E1_biz_area_cd = 0    
    Const A519_EG1_E1_biz_area_nm = 1
    Const A519_EG1_E2_dept_cd = 2    
    Const A519_EG1_E2_dept_nm = 3
    Const A519_EG1_E2_org_change_id = 4    
    Const A519_EG1_E3_cost_cd = 5    
    Const A519_EG1_E3_cost_nm = 6
    Const A519_EG1_E3_cost_type = 7
    Const A519_EG1_E3_di_fg = 8
    Const A519_EG1_E4_minor_nm = 9   
    Const A519_EG1_E5_inv_qty = 10
    Const A519_EG1_E5_assn_rate = 11
    
	dim IG1_a_acct_dept 
    dim E1_a_asset_master 
    dim EG1_export_group 

    Dim iPAAG025
    Dim iStrData
    Dim exportData
    Dim exportReturn
    Dim iLngRow,iLngCol
    Dim iStrPrevKey
    Dim iStrAsstNo
    Dim iIntMaxRows
    Dim iIntQueryCount
    Dim importArray
    Dim iIntLoopCount
    Dim LngMaxRow
    
	Dim iChgQty
	Dim temp_inv_no_sum
	Dim iOptMeth 

    Const C_SHEETMAXROWS  = 100
    
    Const C_QueryConut		= 0
    Const C_MaxQueryReCord = 1
    Const C_AsstNo = 2
	
	Const C_E1_asst_nm = 0
	Const C_E1_inv_qty = 1

	iChgQty = 0
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	iStrPrevKey		= Trim(Request("lgStrPrevKey"))        
    iIntMaxRows		= Request("txtMaxRows")
    iIntQueryCount	= Request("lgPageNo")
    iOptMeth		= Request("txtOptMeth")

    If iStrPrevKey = "" Then
		iStrAsstNo	=  Trim(Request("txtAsstNo"))
	Else
		iStrAsstNo	= iStrPrevKey
    End If

    If Len(Trim(iIntQueryCount))  Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(iIntQueryCount) Then
          iIntQueryCount = CInt(iIntQueryCount)          
       End If   
    Else   
       iIntQueryCount = 0
    End If

        
    ReDim importArray(2)        
    importArray(C_QueryConut)	  = iIntQueryCount
    importArray(C_MaxQueryReCord) = C_SHEETMAXROWS
    importArray(C_AsstNo)		  = iStrAsstNo
    
    Redim exportData(1)
    
	Set iPAAG025 = Server.CreateObject("PAAG025.cAAS0068ListSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    
	
	Call iPAAG025.AS0068_LIST_SVR(gStrGloBalCollection, importArray, exportData, exportReturn)

    If CheckSYSTEMError(Err, True) = True Then					
       Set iPAAG025 = Nothing
       Exit Sub
    End If    

    Set iPAAG025 = Nothing

	iStrData = ""
	iIntLoopCount = 0	
	For iLngRow = 0 To UBound(exportReturn, 1) 		
		iIntLoopCount = iIntLoopCount + 1

	    If  iIntLoopCount < (C_SHEETMAXROWS + 1) Then
'			If Trim(exportReturn(iLngRow, A519_EG1_E5_inv_qty)) <> 0 Then '수량이  0 이 아닌것만 조회 
				temp_inv_no_sum = Trim(exportReturn(iLngRow, A519_EG1_E5_inv_qty))
'				iChgQty = iChgQty + temp_inv_no_sum ' 수량 합계 
				iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, A519_EG1_E2_dept_cd)) 
				iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, A519_EG1_E2_dept_nm)) 
				istrData = istrData & Chr(11) & UNINumClientFormat(exportReturn(iLngRow, A519_EG1_E5_assn_rate), ggExchRate.DecPoint, 0)
				iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, A519_EG1_E3_cost_nm)) 
				iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, A519_EG1_E1_biz_area_nm)) 		
				iStrData = iStrData & Chr(11) & LngMaxRow +iIntLoopCount
				iStrData = iStrData & Chr(11) & Chr(12)
'			End If
	    Else
			iStrPrevKey = exportReturn(UBound(exportReturn, 1), 0)
			iIntQueryCount = iIntQueryCount + 1
			Exit For
		End If
	Next

	iChgQty = UNINumClientFormat(iChgQty, ggQty.DecPoint, 0)
	If  iIntLoopCount < (C_SHEETMAXROWS + 1) Then
		iStrPrevKey = ""
	    iIntQueryCount = ""
	End If

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
	Response.Write "     .frm1.txtAsstNo.value = """ & ConvSPChars(iStrAsstNo) & """" & vbCr
	Response.Write "     .frm1.txtAsstNm.value = """ & ConvSPChars(	exportData(C_E1_asst_nm)) & """" & vbCr
	If iOptMeth = "R" Then				'
	Response.Write "     .frm1.txtChgQty.value = """ & ConvSPChars(	exportData(C_E1_inv_qty)) & """" & vbCr
	End If
	Response.Write " LngMaxRow = .frm1.vspdData.MaxRows	                                        " & vbCr
    Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr 			 
    Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr
'    Response.Write "	.lgPageNo = """ & iIntQueryCount		   & """" & vbCr
    Response.Write "	.lgStrPrevKey = """ & iStrPrevKey		   & """" & vbCr
	Response.Write "	If .frm1.vspdData.MaxRows < .C_SHEETMAXROWS And .lgStrPrevKey <> """" Then "	& vbCr
	Response.Write "			.DbQuery_master													 " & vbCr
	Response.Write "		Else											"	& vbCr
	Response.Write "			.frm1.txthAsstNo.value  = """ & Request("txtAsstNo")			& """" & vbCr
	Response.Write "			.DbQueryOk_master												" & vbCr			
	Response.Write "		End If					" & vbCr
    Response.Write " End With   " & vbCr
    Response.Write " </Script>  " & vbCr
    Response.End

End Sub	
%>
