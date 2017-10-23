<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%'======================================================================================================
'*  1. Module Name          : Basic Architect
'*  2. Function Name        : Asset Acquisition Reference Popup
'*  3. Program ID           : A7107rb1.asp
'*  4. Program Name         : 자산변동 참조 팝업
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001/06/10
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Kim Hee Jung
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================

Response.Expires = -1                                                      '☜ : ASP가 캐쉬되지 않도록 한다.
Response.Buffer = True                                                     '☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다

On Error Resume Next
Err.Clear                                                                  '☜ : Clear Error status

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언
Dim lgPID                                                                  '☜ : ActiveX Data Factory 지정 변수선언
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언
Dim lgStrPrevKey                                                           '☜ : 이전 값
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------

Dim lgFrChgDt		
Dim lgToChgDt		
Dim lgFrAsstChgNo	
Dim lgToAsstChgNo	
Dim lgFrChgNo		
Dim lgDeptCd		
Dim lgAsstChgDesc
Dim lgGubun
Const C_SHEETMAXROWS_D = 100
Dim strCond

' 권한관리 추가
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장
Dim lgInternalCd ', lgDeptCd, lgDeptNm			' 내부부서
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)
Dim lgAuthUsrID, lgAuthUsrNm					' 개인

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "A", "NOCOOKIE", "RB")   'ggQty.DecPoint Setting...
	Call LoadBNumericFormatB("Q", "A", "NOCOOKIE", "RB") 

	lgFrChgDt		= UNIConvDate(Request("txtFrChgDt"))
	lgToChgDt		= UNIConvDate(Request("txtToChgDt"))
	lgFrAsstChgNo	= Request("txtFrAsstChgNo")
	lgToAsstChgNo	= Request("txtToAsstChgNo")
	lgFrChgNo		= Request("txtFrChgNo")
	lgDeptCd		= Request("txtDeptCd")
	lgAsstChgDesc	= Request("txtAsstChgDesc")
	lgGubun			= Request("txtGubun")

	lgPID          = UCase(Request("PID"))  
    lgStrPrevKey   = Request("lgStrPrevKey")                               '☜ : Next key flag
    lgMaxCount     = CInt(Request("lgMaxCount"))                           '☜ : 한번에 가져올수 있는 데이타 건수

	Call SubOpenDB(lgObjConn)                                              '☜ : Make a DB Connection
	Call SubBizQuery()

   ' Call QueryData()

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
'============================================================================================================
' Name : SubBizQueryMulti1    두번째 dbqueryok()에서 호출된 두번째 
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQuery()

    Dim iDx
    Dim iKey1
    Dim strWhere
    Dim YYYYMM
	Dim iIntLoopCount
	Dim iRCnt
	Dim iCnt

    On Error Resume Next                                                   '☜ : Protect system from crashing
    Err.Clear                                                              '☜ : Clear Error status
    iCnt = 0

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    Call SubMakeSQLStatements("MR",strWhere,"X",C_LIKE)                                 '☆: Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
		lgStrPrevKey = ""        
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '☜ : No data is found. 
    Else
		If Len(Trim(lgStrPrevKey)) Then                                        '☜ : Chnage Nextkey str into int value
		   If Isnumeric(lgStrPrevKey) Then
			  iCnt = CInt(lgStrPrevKey)
		   End If   
		End If   

		For iRCnt = 1 to C_SHEETMAXROWS_D * Cint(iCnt)                                  '☜ : Discard previous data
			lgObjRs.MoveNext
		Next

        lgstrData = ""
        iRCnt = 0
        iDx = 1

        Do While Not lgObjRs.EOF
			iRCnt = iRCnt + 1

            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ASST_CHG_NO"))
			lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ASST_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ASST_NM"))
            lgstrData = lgstrData & Chr(11) & UNIDateClientFormat(lgObjRs("CHG_DT"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("CHG_FG"))
            lgstrData = lgstrData & Chr(11) & ""
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("FROM_DEPT_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DEPT_NM"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("DOC_CUR"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_CD"))
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("BP_NM"))

            lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("CHG_TOT_AMT"),gCurrency,ggAmtOfMoneyNo, "X" , "X")
            lgstrData = lgstrData & Chr(11) & UNIConvNumDBToCompanyByCurrency(lgObjRs("CHG_TOT_LOC_AMT"),gCurrency,ggAmtOfMoneyNo, gLocRndPolicyNo , "X")
            lgstrData = lgstrData & Chr(11) & ConvSPChars(lgObjRs("ASST_CHG_DESC"))
			lgstrData = lgstrData & Chr(11) & C_SHEETMAXROWS_D * iCnt  + iDx
			lgstrData = lgstrData & Chr(11) & Chr(12)
 			If  iRCnt < C_SHEETMAXROWS_D Then                                            '☜: Check if next data exists
			Else
				iCnt = iCnt + 1
				lgStrPrevKey = Cint(iCnt)
				Exit Do
			End If

			iDx = iDx + 1
			lgObjRs.MoveNext
		Loop 
		If  iRCnt < C_SHEETMAXROWS_D Then                                            '☜: Check if next data exists
			lgStrPrevKey = ""                                                  '☜: 다음 데이타 없다.
		End If
   End If
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)

	' 권한관리 추가
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID		= Trim(Request("lgAuthUsrID"))

	Select Case Mid(pDataType,1,1)
		Case "S"
		Case "M"
		   Select Case Mid(pDataType,2,1)
			   Case "R"
				  select case pCode1
				  case "X"
						If UCase(Trim(lgGubun)) = "A" Then ' 매각폐기번호
							lgStrSQL = ""
							lgStrSQL = lgStrSQL & "	  SELECT distinct B.asst_chg_no,   "
							lgStrSQL = lgStrSQL & "	         '' asst_cd,   "
							lgStrSQL = lgStrSQL & "	         B.chg_dt,   "
							lgStrSQL = lgStrSQL & "	         B.chg_fg,   "
							lgStrSQL = lgStrSQL & "	         B.doc_cur,   "
							lgStrSQL = lgStrSQL & "	         B.from_dept_cd,   "
							lgStrSQL = lgStrSQL & "	         '' asst_nm,   "
							lgStrSQL = lgStrSQL & "	         B.chg_tot_amt,   "
							lgStrSQL = lgStrSQL & "	         B.chg_tot_loc_amt,   "
							lgStrSQL = lgStrSQL & "	         B.bp_cd,   "
							lgStrSQL = lgStrSQL & "	         D.bp_nm,   "
							lgStrSQL = lgStrSQL & "	         C.dept_nm,   "
							lgStrSQL = lgStrSQL & "	         B.asst_chg_desc  "
							lgStrSQL = lgStrSQL & "	    FROM a_asset_chg_master B,   "
							lgStrSQL = lgStrSQL & "	         b_acct_dept C,   "
							lgStrSQL = lgStrSQL & "	         b_biz_partner D   "
							lgStrSQL = lgStrSQL & "	   WHERE ( B.from_dept_cd  = C.dept_cd) and  "
							lgStrSQL = lgStrSQL & "	         ( B.bp_cd  *= D.bp_cd) and  "
							lgStrSQL = lgStrSQL & "	         ( B.from_org_change_id  = C.org_change_id) 	 "	
						Else
							lgStrSQL = ""
							lgStrSQL = lgStrSQL & "	  SELECT distinct B.asst_chg_no,   "
							lgStrSQL = lgStrSQL & "	         a.asst_cd,   "
							lgStrSQL = lgStrSQL & "	         B.chg_dt,   "
							lgStrSQL = lgStrSQL & "	         B.chg_fg,   "
							lgStrSQL = lgStrSQL & "	         B.doc_cur,   "
							lgStrSQL = lgStrSQL & "	         B.from_dept_cd,   "
							lgStrSQL = lgStrSQL & "	         E.asst_nm,   "
							lgStrSQL = lgStrSQL & "	         B.chg_tot_amt,   "
							lgStrSQL = lgStrSQL & "	         B.chg_tot_loc_amt,   "
							lgStrSQL = lgStrSQL & "	         B.bp_cd,   "
							lgStrSQL = lgStrSQL & "	         D.bp_nm,   "
							lgStrSQL = lgStrSQL & "	         C.dept_nm,   "
							lgStrSQL = lgStrSQL & "	         B.asst_chg_desc  "
							lgStrSQL = lgStrSQL & "	    FROM a_asset_chg A,   "
							lgStrSQL = lgStrSQL & "	         a_asset_chg_master B,   "
							lgStrSQL = lgStrSQL & "	         b_acct_dept C,   "
							lgStrSQL = lgStrSQL & "	         b_biz_partner D,   "
							lgStrSQL = lgStrSQL & "	         a_asset_master E "
							lgStrSQL = lgStrSQL & "	   WHERE ( B.asst_chg_no = A.asst_chg_no ) and  "
							lgStrSQL = lgStrSQL & "	         ( A.asst_cd *= E.asst_no ) and  "
							lgStrSQL = lgStrSQL & "	         ( B.from_dept_cd  *= C.dept_cd) and  "
							lgStrSQL = lgStrSQL & "	         ( B.bp_cd  *= D.bp_cd) and  "
							lgStrSQL = lgStrSQL & "	         ( B.from_org_change_id  *= C.org_change_id) 	 "		 
						End If

						IF Trim(Request("txtFrChgDt")) <> "" Then '조회기간
							lgStrSQL = lgStrSQL & "	And (B.CHG_DT >=  " & FilterVar(lgFrChgDt , "''", "S") & " )"
						End If

						IF Trim(Request("txtToChgDt")) <> "" Then
							lgStrSQL = lgStrSQL & "	And (B.CHG_DT <=  " & FilterVar(lgToChgDt , "''", "S") & " )"
						End If


						If Trim(lgDeptCd) <> "" Then '부서
							lgStrSQL = lgStrSQL & "		AND B.FROM_DEPT_CD = " & FilterVar(lgDeptCd, "''", "S") & "  "
						End If

						If Trim(lgAsstChgDesc) <> "" Then '적요
							lgStrSQL = lgStrSQL & "		AND B.ASST_CHG_DESC LIKE " & FilterVar("%" & lgAsstChgDesc & "%" , "''", "S")
						End If

						' 권한관리 추가
						If lgAuthBizAreaCd <> "" Then
							lgBizAreaAuthSQL		= " AND B.FROM_BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")
						End If
	
						If lgInternalCd <> "" Then
							lgInternalCdAuthSQL		= " AND B.FROM_INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")
						End If
	
						If lgSubInternalCd <> "" Then
							lgSubInternalCdAuthSQL	= " AND B.FROM_INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")
						End If
	
						If lgAuthUsrID <> "" Then
							lgAuthUsrIDAuthSQL		= " AND B.UPDT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")
						End If

						lgStrSQL = lgStrSQL & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL

						If UCase(Trim(lgGubun)) = "A" Then ' 매각폐기번호
							If(Trim(lgFrAsstChgNo) <> "" And Trim(lgToAsstChgNo)  <> "" )Then 
								lgStrSQL = lgStrSQL & "		AND (B.ASST_CHG_NO >=  " & FilterVar(lgFrAsstChgNo, "''", "S") & "  AND B.ASST_CHG_NO <= " & FilterVar(lgToAsstChgNo, "''", "S") & " ) "
							ElseIf( Trim(lgFrAsstChgNo) <> "" And Trim(lgToAsstChgNo)  = "" )Then 
								lgStrSQL = lgStrSQL & "		AND (B.ASST_CHG_NO >=  " & FilterVar(lgFrAsstChgNo, "''", "S") & " ) "
							ElseIf( Trim(lgFrAsstChgNo) = "" And Trim(lgToAsstChgNo)  <> "" )Then 
								lgStrSQL = lgStrSQL & "		AND (B.ASST_CHG_NO <= " & FilterVar(lgToAsstChgNo, "''", "S") & " ) "
							End If

							lgStrSQL = lgStrSQL & "	ORDER BY B.ASST_CHG_NO "
						Else
							If Trim(lgFrChgNo) <> "" Then 
								lgStrSQL = lgStrSQL & "		AND A.ASST_CD= " & FilterVar(lgFrChgNo, "''", "S") & " "
							End If
							lgStrSQL = lgStrSQL & "	ORDER BY A.ASST_CD "
						End If
					End Select
			End Select
	End Select
	'------ Developer Coding part (End   ) ------------------------------------------------------------------
End Sub
%>

<Script Language=vbscript>
    With parent
         .ggoSpread.Source    = .frm1.vspdData 
         .ggoSpread.SSShowData "<%=lgstrData%>"                            '☜: Display data 
         .lgStrPrevKey        =  "<%=lgStrPrevKey%>"                       '☜: set next data tag
         .DbQueryOk
	End with
</Script>

