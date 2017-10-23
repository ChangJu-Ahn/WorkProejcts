<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%'======================================================================================================
'*  1. Module Name          : Basic Architect
'*  2. Function Name        : Asset Acquisition Reference Popup
'*  3. Program ID           : A7107rb1.asp
'*  4. Program Name         : �ڻ꺯�� ���� �˾�
'*  5. Program Desc         :
'*  6. Modified date(First) : 2001/06/10
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Kim Hee Jung
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================

Response.Expires = -1                                                      '�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True                                                     '�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.
%>

<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/lgsvrvariables.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ�

On Error Resume Next
Err.Clear                                                                  '�� : Clear Error status

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� ��������
Dim lgPID                                                                  '�� : ActiveX Data Factory ���� ��������
Dim lgstrRetMsg                                                            '�� : Record Set Return Message ��������
Dim lgStrPrevKey                                                           '�� : ���� ��
'--------------- ������ coding part(��������,Start)--------------------------------------------------------

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

' ���Ѱ��� �߰�
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' �����
Dim lgInternalCd ', lgDeptCd, lgDeptNm			' ���κμ�
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ����

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL

'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
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
    lgStrPrevKey   = Request("lgStrPrevKey")                               '�� : Next key flag
    lgMaxCount     = CInt(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ�

	Call SubOpenDB(lgObjConn)                                              '�� : Make a DB Connection
	Call SubBizQuery()

   ' Call QueryData()

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
'============================================================================================================
' Name : SubBizQueryMulti1    �ι�° dbqueryok()���� ȣ��� �ι�° 
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

    On Error Resume Next                                                   '�� : Protect system from crashing
    Err.Clear                                                              '�� : Clear Error status
    iCnt = 0

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    Call SubMakeSQLStatements("MR",strWhere,"X",C_LIKE)                                 '��: Make sql statements
    
    If 	FncOpenRs("R",lgObjConn,lgObjRs,lgStrSQL,"X","X") = False Then
		lgStrPrevKey = ""        
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)      '�� : No data is found. 
    Else
		If Len(Trim(lgStrPrevKey)) Then                                        '�� : Chnage Nextkey str into int value
		   If Isnumeric(lgStrPrevKey) Then
			  iCnt = CInt(lgStrPrevKey)
		   End If   
		End If   

		For iRCnt = 1 to C_SHEETMAXROWS_D * Cint(iCnt)                                  '�� : Discard previous data
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
 			If  iRCnt < C_SHEETMAXROWS_D Then                                            '��: Check if next data exists
			Else
				iCnt = iCnt + 1
				lgStrPrevKey = Cint(iCnt)
				Exit Do
			End If

			iDx = iDx + 1
			lgObjRs.MoveNext
		Loop 
		If  iRCnt < C_SHEETMAXROWS_D Then                                            '��: Check if next data exists
			lgStrPrevKey = ""                                                  '��: ���� ����Ÿ ����.
		End If
   End If
End Sub

'============================================================================================================
' Name : SubMakeSQLStatements
' Desc : Make SQL statements
'============================================================================================================
Sub SubMakeSQLStatements(pDataType,pCode,pCode1,pComp)

	' ���Ѱ��� �߰�
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
						If UCase(Trim(lgGubun)) = "A" Then ' �Ű�����ȣ
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

						IF Trim(Request("txtFrChgDt")) <> "" Then '��ȸ�Ⱓ
							lgStrSQL = lgStrSQL & "	And (B.CHG_DT >=  " & FilterVar(lgFrChgDt , "''", "S") & " )"
						End If

						IF Trim(Request("txtToChgDt")) <> "" Then
							lgStrSQL = lgStrSQL & "	And (B.CHG_DT <=  " & FilterVar(lgToChgDt , "''", "S") & " )"
						End If


						If Trim(lgDeptCd) <> "" Then '�μ�
							lgStrSQL = lgStrSQL & "		AND B.FROM_DEPT_CD = " & FilterVar(lgDeptCd, "''", "S") & "  "
						End If

						If Trim(lgAsstChgDesc) <> "" Then '����
							lgStrSQL = lgStrSQL & "		AND B.ASST_CHG_DESC LIKE " & FilterVar("%" & lgAsstChgDesc & "%" , "''", "S")
						End If

						' ���Ѱ��� �߰�
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

						If UCase(Trim(lgGubun)) = "A" Then ' �Ű�����ȣ
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
         .ggoSpread.SSShowData "<%=lgstrData%>"                            '��: Display data 
         .lgStrPrevKey        =  "<%=lgStrPrevKey%>"                       '��: set next data tag
         .DbQueryOk
	End with
</Script>

