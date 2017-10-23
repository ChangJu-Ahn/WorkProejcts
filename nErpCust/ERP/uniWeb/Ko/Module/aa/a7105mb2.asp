<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

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


	Dim lgOpModeCRUD
	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 ,RS1
	DIM lgStrPrevToKey
	Const C_SHEETMAXROWS_D = 100
	
	lgStrPrevToKey=Request("lgStrPrevToKey")
	if lgStrPrevToKey="" then lgStrPrevToKey="1"


    '---------------------------------------Common-----------------------------------------------------------
'    lgErrorStatus     = "NO"
'    lgErrorPos        = ""                                                           '☜: Set to space
    lgOpModeCRUD      = Request("txtMode")                                           '☜: Read Operation Mode (CRUD)

	'------ Developer Coding part (Start ) ------------------------------------------------------------------


	'------ Developer Coding part (End   ) ------------------------------------------------------------------ 

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '☜: Query
             Call SubBizQueryMulti()
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

Sub SubBizQueryMulti()
  
	Call FixUNISQLData()
	Call QueryData()	

End Sub    

'============================================================================================================
' Set DB Agent arg
'============================================================================================================

Sub FixUNISQLData()
    Dim strVal
	Redim UNISqlId(0)                                                     
    Redim UNIValue(0,3)    
    dim txtYYYYMM,txtAssetNo
    

             
      txtYYYYMM = Request("txtYyyymm")
      txtAssetNo = trim(request("txtCondAsstNo"))  
                                                              
     UNISqlId(0) = "a7105ma2" 										

     UNIValue(0,0)=" TOP " & lgStrPrevToKey * C_SHEETMAXROWS_D &"  A1.ASST_NO,'',C.ASST_NM ,   B2.DEPT_CD,'',B2.DEPT_NM,"
     UNIValue(0,0)=UNIValue(0,0) & "B2.ORG_CHANGE_ID, B2.COST_CD,COST_NM,dbo.ufn_GetCodeName('B9013' , DI_FG ),'',  A1.INV_QTY ,A1.ASSN_RATE   "

     UNIValue(0,1) = filterVar(txtAssetNo&"%","''","S") 
     UNIValue(0,2) = filterVar(txtYYYYMM,"''","S") 
     
     UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
   
End Sub
'============================================================================================================
' Query Data
' ADO의 Record Set이용하여 Query를 하고 Record Set을 넘겨서 MakeSpreadSheetData()으로 Spreadsheet에 데이터를 
' 뿌림 
' ADO 객체를 생성할때 prjPublic.dll파일을 이용한다.(상세내용은 vb로 작성된 prjPublic.dll 소스 참조)
'============================================================================================================
Sub QueryData()
    Dim lgstrRetMsg                                             
    Dim lgADF                                                  
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)

	Set lgADF   = Nothing
	iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
		Response.end
    End If 
 
 
  If  rs0.EOF And rs0.BOF  Then
       
		Response.Write "<Script Language=VBScript>" & vbCrLF
		Response.Write "</Script>" & vbCrLF
		
		 Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
		Response.end
    ELSE
		call ListupDataGrid (rs0.getRows,"","","vspdData")
			
    End If  
End Sub
	    

'============================================================================
'ListupDataGrid
'============================================================================
 
 Sub ListupDataGrid(pArr,dataFormatCol,NFormatCol,grid)
	Dim strData
	Dim i,j,moveLine,RowCnt
	dim iStrData
	'on error resume next

	RowCnt=0
	moveLine = (lgStrPrevToKey - 1) * C_SHEETMAXROWS_D
	

		for i=moveLine to uBound(pArr,2)
			RowCnt=RowCnt+1
			for j=0 to uBound(pArr,1)
			
			if inStr(dataFormatCol,"," & j&",") > 0 then
				strData = strData & Chr(11) & UniConvDateDbToCompany(pArr(j,i),"")
			elseif inStr(NFormatCol,"," & j&",") > 0 then
				strData = strData & Chr(11) & UniConvNumDBToCompanyWithOutChange(trim(ConvSPChars(pArr(j,i))),0)
			else
				strData = strData & Chr(11) & trim(ConvSPChars(pArr(j,i)))
			end if	
			next 
			strData =  strData & Chr(11) & i &  Chr(11) & Chr(12) 
		next 
		
		Response.Write " <Script Language=vbscript>	                        " & vbCr
		Response.Write " With parent                                        " & vbCr
		Response.Write "	.ggoSpread.Source = .frm1.vspdData              " & vbCr 			 
		Response.Write "	.ggoSpread.SSShowData """ & strData       & """" & vbCr
		Response.Write "	.lgStrPrevToKey  = """ & lgStrPrevToKey + 1 & """" & vbCr 
		Response.Write "	.DbQueryOk " & vbCr
		if RowCnt<C_SHEETMAXROWS_D then
			Response.Write "    .lgStrPrevToKey= """"  "                  & vbCr 
		
		end if
		Response.Write " End With   " & vbCr
		Response.Write " </Script>  " & vbCr
		Response.End
End Sub	
	  
'============================================================================================================
' Name : SubBizSave
' Desc : Date data 
'============================================================================================================
Sub SubBizSave()

    Dim iPAAG030
    'Dim import_String
    Dim import_Group
    Dim import_GroupString
	Dim strYear, strMonth, strDay, stryyyymm, stDt
    
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
    ReDim import_Group(1)        
    import_Group(0)	= Trim(Request("txtCondAsstNo"))
    
    stryyyymm = Trim(Request("txtYyyymm"))
	stDt = UniConvYYYYMMDDToDate(gDateFormat, Mid(stryyyymm,1,4), Mid(stryyyymm,6,2), "01")
    Call ExtractDateFrom(stDt, gDateFormat, gComDateType, strYear, strMonth, strDay)
    
    import_Group(1)	= strYear & Right("0" & strMonth, 2)
    
    import_GroupString = replace(Trim(Request("txtSpread")),",","")
    
    Set iPAAG030 = Server.CreateObject("PAAG030.cAMngAsDptHistorySvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    
    Call iPAAG030.AS0061_MANAGE_ASSET_DEPT_HISTORY_SVR(gStrGloBalCollection, import_Group, import_GroupString, I2_a_data_auth, "1")
    'Call iPAAG030.AS0061_MANAGE_ASSET_DEPT_SVR(gStrGloBalCollection, import_GroupString)

    If CheckSYSTEMError(Err, True) = True Then					
       Set iPAAG030 = Nothing
       response.end
       Exit Sub
    End If    
    
    Set iPAAG030 = Nothing
	
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

	Response.End
%>
























