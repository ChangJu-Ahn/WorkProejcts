<%@ LANGUAGE=VBSCript%>
<%Option Explicit%>
<!-- #Include file="../../inc/incSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%
'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a7105b1
'*  4. Program Name         : 고정자산 부서별배분율등록 
'*  5. Program Desc         : 고정자산 부서별배분율을 등록,수정,삭제,조회 
'*  6. Comproxy List        : +As0061ManageSvr
'                             +As0068ListSvr
'*  7. Modified date(First) : 2000/09/19
'*  8. Modified date(Last)  : 2001/05/31
'*  9. Modifier (First)     : hersheys
'* 10. Modifier (Last)      : Kim Hee Jung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

    Call HideStatusWnd                                                               '☜: Hide Processing message
    

    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")                                                        '☜: Hide Processing message
'Dim lgCurrency, lgStrPrevKey_i, lgBlnFlgChgValue, plgStrPrevKey_i

	Dim lgOpModeCRUD
'	Dim lgPageNo, lgStrPrevKey

    '---------------------------------------Common-----------------------------------------------------------
'    lgErrorStatus     = "NO"
'    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

	'------ Developer Coding part (Start ) ------------------------------------------------------------------


	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '☜: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '☜: Delete
             Call SubBizDelete()
    End Select

    'Call SubCloseDB(lgObjConn)                                                       '☜: Close DB Connection

'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'		 -- Spread Setting
'		 ggoSpread.SSSetEdit    C_DeptCd,     "관리부서",      12, 2, , 10,2 '3
'		 ggoSpread.SSSetButton  C_DeptCdPopUp								 '4
'		 ggoSpread.SSSetEdit    C_DeptNm,     "부서명",        35			 '5
'		 ggoSpread.SSSetEdit    C_CostCd,	  "코스트센타",		0			 '6
'		 ggoSpread.SSSetEdit    C_CostNm,	  "코스트센타명",  30			 '7
'		 ggoSpread.SSSetEdit    C_CostType,   "",              10			 '8
'		 ggoSpread.SSSetEdit    C_CostTypeNm, "직간접구분",    15			 '9
'		 ggoSpread.SSSetFloat   C_InvQty,     "재고수량",      10, ggQtyNo,       ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec
'		 ggoSpread.SSSetFloat   C_AssnRate,   "배분비율(%)",   21, ggAmtOfMoneyNo,ggStrIntegeralPart, ggStrDeciPointPart,gComNum1000,gComNumDec,,,"Z","0","100"
'
'		 -- QueryData 
'		 10	(주)UNIERP	10        	본사	CS04	공통부문	C 	I 	간접	1	100
'============================================================================================================
Sub SubBizQuery()

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
    
    Const C_SHEETMAXROWS_D  = 100
    Const C_QueryConut		= 0
    Const C_MaxQueryReCord = 1
    Const C_AsstNo = 2
   
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	' -- 권한관리추가 
	Const A519_I2_a_data_auth_data_BizAreaCd = 0
	Const A519_I2_a_data_auth_data_internal_cd = 1
	Const A519_I2_a_data_auth_data_sub_internal_cd = 2
	Const A519_I2_a_data_auth_data_auth_usr_id = 3

	Dim I2_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

  	Redim I2_a_data_auth(3)
	I2_a_data_auth(A519_I2_a_data_auth_data_BizAreaCd)       = Trim(Request("lgAuthBizAreaCd"))
	I2_a_data_auth(A519_I2_a_data_auth_data_internal_cd)     = Trim(Request("lgInternalCd"))
	I2_a_data_auth(A519_I2_a_data_auth_data_sub_internal_cd) = Trim(Request("lgSubInternalCd"))
	I2_a_data_auth(A519_I2_a_data_auth_data_auth_usr_id)     = Trim(Request("lgAuthUsrID"))

	iStrPrevKey		= Trim(Request("lgStrPrevKey"))        
    iIntMaxRows		= Request("txtMaxRows")
    iIntQueryCount	= Request("lgPageNo")
    If iStrPrevKey = "" Then
		iStrAsstNo	= Request("txtCondAsstNo")
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
    importArray(C_MaxQueryReCord) = C_SHEETMAXROWS_D
    importArray(C_AsstNo)		  = iStrAsstNo
    
	Set iPAAG025 = Server.CreateObject("PAAG025.cAAS0068ListSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    

	Call iPAAG025.AS0068_LIST_SVR(gStrGloBalCollection, importArray, exportData, exportReturn, I2_a_data_auth)

    If CheckSYSTEMError(Err, True) = True Then					
       Set iPAAG025 = Nothing
       Response.End
       Exit Sub
       
    End If    

    Set iPAAG025 = Nothing




	iStrData = ""
	iIntLoopCount = 0	
	For iLngRow = 0 To UBound(exportReturn, 1) 		
		iIntLoopCount = iIntLoopCount + 1

	    If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
			For iLngCol = 0 To UBound(exportReturn, 2)
				select case iLngCol
					case 0	'관리부서코드 
						'iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, iLngCol)) & Chr(11) 
					case 1	'관리부서명 
						'iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, iLngCol))
					case 2	'부서코드 
						iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, iLngCol)) & Chr(11) 
					case 3	'부서명 
						iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, iLngCol))
					case 4	'조직변경id
						iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, iLngCol))
					case 5	'코스트센터코드 
						iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, iLngCol))
					case 6	'코스트센터명 
						iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, iLngCol))
					case 7	'!@#@#%
						'iStrData = iStrData & Chr(11) & Chr(11) & ConvSPChars(exportReturn(iLngRow, iLngCol))
					case 8	'@#$^@#
						'iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, iLngCol)) & Chr(11)
					case 9	'직간접구분 
						iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, iLngCol)) & Chr(11)
					case 10	'재고수량 
						iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, iLngCol))
					case 11	'배분비율 
						iStrData = iStrData & Chr(11) & UNINumClientFormat(exportReturn(iLngRow, iLngCol), ggExchRate.DecPoint, 0)
					case else
'						iStrData = iStrData & Chr(11) & ConvSPChars(exportReturn(iLngRow, iLngCol))
				end select
			Next
				iStrData = iStrData & gColSep & iIntMaxRows + iLngRow + 1
			    iStrData = iStrData & Chr(11) & Chr(12)
	    Else
			iStrPrevKey = exportReturn(UBound(exportReturn, 1), 0)
			iIntQueryCount = iIntQueryCount + 1
			Exit For
		End If
	Next

	If  iIntLoopCount < (C_SHEETMAXROWS_D + 1) Then
		iStrPrevKey = ""
	    iIntQueryCount = ""
	End If

	Response.Write " <Script Language=vbscript>	                        " & vbCr
	Response.Write " With parent                                        " & vbCr
    Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr 			 
    Response.Write "	.ggoSpread.SSShowData """ & iStrData       & """" & vbCr
    Response.Write "	.frm1.txtCondAsstNm.value = """ & ConvSPChars(exportData)  & """" & vbCr
    Response.Write "	.lgPageNo = """ & iIntQueryCount		   & """" & vbCr
    Response.Write "	.lgStrPrevKey = """ & ConvSPChars(iStrPrevKey)		   & """" & vbCr
    Response.Write "	.DbQueryOk " & vbCr
    Response.Write " End With   " & vbCr
    Response.Write " </Script>  " & vbCr
    Response.End

End Sub	
'============================================================================================================
' Name : SubBizQuery
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()

    Dim iPAAG025
    'Dim import_String
    Dim import_Group
    Dim import_GroupString
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status

	' -- 권한관리추가 
	Const A519_I2_a_data_auth_data_BizAreaCd = 0
	Const A519_I2_a_data_auth_data_internal_cd = 1
	Const A519_I2_a_data_auth_data_sub_internal_cd = 2
	Const A519_I2_a_data_auth_data_auth_usr_id = 3

	Dim I2_a_data_auth  '--> 파라미터의 순서에 따라 네이밍 변동 

  	Redim I2_a_data_auth(3)
	I2_a_data_auth(A519_I2_a_data_auth_data_BizAreaCd)       = Trim(Request("txthAuthBizAreaCd"))
	I2_a_data_auth(A519_I2_a_data_auth_data_internal_cd)     = Trim(Request("txthInternalCd"))
	I2_a_data_auth(A519_I2_a_data_auth_data_sub_internal_cd) = Trim(Request("txthSubInternalCd"))
	I2_a_data_auth(A519_I2_a_data_auth_data_auth_usr_id)     = Trim(Request("txthAuthUsrID"))
    
    import_Group = Trim(Request("txtCondAsstNo"))
    import_GroupString = replace(Trim(Request("txtSpread")),",","")
    
    Set iPAAG025 = Server.CreateObject("PAAG025.cAMngAsDptSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    
    Call iPAAG025.AS0061_MANAGE_ASSET_DEPT_SVR(gStrGloBalCollection, import_Group, import_GroupString, I2_a_data_auth)
    'Call iPAAG025.AS0061_MANAGE_ASSET_DEPT_SVR(gStrGloBalCollection, import_GroupString)

    If CheckSYSTEMError(Err, True) = True Then					
       Set iPAAG025 = Nothing
       response.end
       Exit Sub
    End If    
    
    Set iPAAG025 = Nothing
	
	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr    

End Sub	
	    
'============================================================================================================
' Name : SubBizDelete
' Desc : Delete DB data
'============================================================================================================
Sub SubBizDelete()

    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubBizSaveCreate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleCreate()
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
	
End Sub

'============================================================================================================
' Name : SubBizSaveSingleUpdate
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizSaveSingleUpdate()
    
    On Error Resume Next                                                             '☜: Protect system from crashing
    Err.Clear                                                                        '☜: Clear Error status
    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    'A developer must define field to update record
    '--------------------------------------------------------------------------------------------------------

    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pMode)
    On Error Resume Next
    
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
    
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
	
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
    Call SetErrorStatus()
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SetErrorStatus
' Desc : This Sub set error status
'============================================================================================================
Sub SetErrorStatus()
'    lgErrorStatus     = "YES"                                                         '☜: Set error status
	'------ Developer Coding part (Start ) ------------------------------------------------------------------
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
%>

<Script Language="VBScript">
	parent.DbSaveOk																		'☜: 화면 처리 ASP 를 지칭함 
</Script>	
<%					

    Set pAS0011 = Nothing                                                   '☜: Unload Comproxy

	Response.End
%>
























