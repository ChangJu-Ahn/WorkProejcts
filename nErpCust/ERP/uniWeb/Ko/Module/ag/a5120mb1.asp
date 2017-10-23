<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../comasp/loadinftb19029.asp" -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

Call LoadBasisGlobalInf()
Call loadInfTB19029B("Q", "A","NOCOOKIE","QB")
Call LoadBNumericFormatB("Q", "A","NOCOOKIE","QA")

Err.Clear
On Error Resume Next

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4 ,rs5 ,rs6    '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Dim lgtxtGlNo

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim lgtxtFromGlDt
Dim lgtxtToGlDt
Dim lgtxtBizArea
Dim lgtxtBizArea1
Dim lgtxtCOST_CENTER_CD
Dim lgtxtdeptcd
Dim lgcboGlInputType
Dim lgtxtMaxRows
Dim lgtxtUsr_Id

Dim dr_loc_amt
Dim cr_loc_amt
Dim biz_area_nm
Dim biz_area_nm1
Dim cost_nm
Dim dept_nm
Dim Usr_Nm

Dim StrDesc, StrRefNo,strAmtFr, strAmtTo

Dim strSql


' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)				
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL					



'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 


    lgPageNo       = Request("lgPageNo")                               '☜ : Next key flag
'    lgMaxCount     = CInt(Request("lgMaxCount"))                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
    lgDataExist    = "No"
    
	
	lgtxtFromGlDt		= FilterVar(UNIConvDate(Request("txtFromGlDt")), "''", "S")
	lgtxtToGlDt			= FilterVar(UNIConvDate(Request("txtToGlDt")), "''", "S")
	lgtxtBizArea		= Trim(Request("txtBizArea"))
	lgtxtBizArea1		= Trim(Request("txtBizArea1"))
	lgtxtCOST_CENTER_CD	= Trim(Request("txtCOST_CENTER_CD"))
	lgtxtGlNo           = Trim(Request("txtGlNo"))
	lgtxtdeptcd			= Request("txtdeptcd")
	lgcboGlInputType	= Request("cboGlInputType")
	lgtxtMaxRows		= Request("txtMaxRows")
	lgtxtUsr_Id		= Request("txtUsr_Id")
	
	
	StrDesc				= Trim(Request("txtDesc"))
	StrRefNo			= Trim(Request("txtRefNo"))
	strAmtFr			= UNIConvNum(Request("txtAmtFr"),0)
	strAmtTo			= UNIConvNum(Request("txtAmtTo"),0)

	' 권한관리 추가 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))		
	lgInternalCd		= Trim(Request("lgInternalCd"))	
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))	
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))

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

    If Len(Trim(lgPageNo))  Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(lgPageNo) Then
          lgPageNo = CInt(lgPageNo)
          
       End If   
    Else   
       lgPageNo = 0
    End If   

   
    'rs0에 대한 결과 
    rs0.PageSize     = C_SHEETMAXROWS_D                                                'Seperate Page with page count (MA : C_SHEETMAXROWS_D )
    rs0.AbsolutePage = lgPageNo + 1

    iLoopCount = -1
    'Call ServerMsgBox("lgPageNo::" &lgPageNo , vbInformation, I_MKSCRIPT) '#
    
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

    If  iLoopCount < C_SHEETMAXROWS_D Then                                            '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
    
    
    'rs1에 대한 결과 
    IF NOT (rs1.EOF or rs1.BOF) then
		dr_loc_amt = rs1("dr_loc_amt")
		cr_loc_amt = rs1("cr_loc_amt")
    End if
    rs1.Close
    Set rs1 = Nothing 
    
    'rs2에 대한 결과 
    IF NOT (rs2.EOF or rs2.BOF) then
	    biz_area_nm = rs2("biz_area_nm")
    END IF
    rs2.Close
    Set rs2 = Nothing
    
    'rs3에 대한 결과    
    IF NOT (rs3.EOF or rs3.BOF) then
		cost_nm = rs3("cost_nm")
    END IF
    rs3.Close
    Set rs3 = Nothing
    
    'rs4에 대한 결과 
    IF NOT (rs4.EOF or rs4.BOF) then
		dept_nm = rs4("dept_nm")
    END IF
    rs4.Close
    Set rs4 = Nothing
    
    'rs5에 대한 결과 
    IF NOT (rs5.EOF or rs5.BOF) then
	    biz_area_nm1 = rs5("biz_area_nm")
    END IF
    rs5.Close
    Set rs5 = Nothing  
    
    IF NOT (rs6.EOF or rs6.BOF) then
	    Usr_Nm = rs6("Usr_Nm")
    END IF
    rs6.Close
    Set rs6 = Nothing    

End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(6)                                                    '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(6,8)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 

    'UNISqlId(0) = "a5120ma101"
    'UNISqlId(1) = "a5120ma102"
    UNISqlId(0) = "a5120ma1Q1"
	UNISqlId(1) = "a5120ma1Q2"
    UNISqlId(2) = "ABIZNM"
    UNISqlId(3) = "M6111QA104"
    UNISqlId(4) = "ADEPTNM"
    UNISqlId(5) = "ABIZNM"    
    UNISqlId(6) = "CommonQry"    
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    
    'rs0에 대한 Value값 setting    
    UNIValue(0,0) = lgSelectList  
  	UNIValue(0,1)  = lgtxtFromGlDt	'UNIConvDate(Request("txtFromGlDt") )
	UNIValue(0,2)  = lgtxtToGlDt
	
	strSql = " "

	If StrRefNo <> "" Then
		strSql  = strSql + "  AND A.REF_NO  LIKE   " & FilterVar(StrRefNo & "%", "''", "S") & " "
	End If

	If StrDesc <> "" Then
		strSql  = strSql + "  AND A.GL_DESC LIKE  " & FilterVar("%" & StrDesc & "%", "''", "S") & " "
	End If

	'-------------------------
	'금액 
	'-----------------------
	If strAmtFr <> 0 or strAmtTo <> 0 Then
		If strAmtFr > 0 and strAmtTo <= 0 Then
			strSql = strSql & " AND (a.DR_LOC_AMT >= " & strAmtFr & " AND  a.CR_LOC_AMT >= " & strAmtFr & " ) "
		ElseIf strAmtFr <= 0 and strAmtTo > 0 Then
			strSql = strSql & " AND (a.DR_LOC_AMT <= " & strAmtTo & " AND  a.CR_LOC_AMT <= " & strAmtTo & " ) "
		Else
			strSql = strSql & " AND (a.DR_LOC_AMT between " & strAmtFr & " AND " & strAmtTo & " AND a.CR_LOC_AMT between  " & strAmtFr & " AND " & strAmtTo & " ) "
		End If
	End If

	If lgtxtBizArea = "" Then
		strSql = strSql & " and A.biz_area_cd >= " & FilterVar("0", "''", "S") 
	Else		
		strSql = strSql & " and A.biz_area_cd >= " & FilterVar(lgtxtBizArea, "''", "S") 
	End If
	
	If lgtxtBizArea1 = "" Then
		strSql = strSql & " and A.biz_area_cd <= " & FilterVar("ZZZZZZZZZZZ", "''", "S") & " "
	Else		
		strSql = strSql & " and A.biz_area_cd <= " & FilterVar(lgtxtBizArea1, "''", "S") 
	End If
	
	If lgtxtGlNo <> "" Then

		strSql = strSql & " and A.gl_no = " & FilterVar(lgtxtGlNo, "''", "S") & " "
	End If	
	
	If lgtxtUsr_Id <> "" Then

		strSql = strSql & " and A.INSRT_USER_ID = " & FilterVar(lgtxtUsr_Id, "''", "S") & " "
	End If	
	


	' 권한관리 추가 
	If lgAuthBizAreaCd <> "" Then			
		lgBizAreaAuthSQL		= " AND A.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		lgInternalCdAuthSQL		= " AND A.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		lgSubInternalCdAuthSQL	= " AND A.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		lgAuthUsrIDAuthSQL		= " AND A.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	

	strSql	= strSql	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL	
	
	strSql = strSql + " "

	

	UNIValue(0,3)  = strSql 

	If lgtxtCOST_CENTER_CD = "" Then
		UNIValue(0,4)  = ""
	Else 
		UNIValue(0,4)  = " AND A.COST_CD = " & FilterVar(lgtxtCOST_CENTER_CD, "''", "S")  
	End If
	 
	If lgtxtdeptcd = "" Then
		UNIValue(0,5)  = ""
	Else 
		UNIValue(0,5)  = " AND A.DEPT_CD = " & FilterVar(lgtxtdeptcd, "''", "S")   
	End If
	 
	If lgcboGlInputType = "" Then
		UNIValue(0,6)  = ""
	Else 
		UNIValue(0,6)  = " AND A.GL_INPUT_TYPE = " & FilterVar(lgcboGlInputType, "''", "S")  
	End If
	
	If lgtxtdeptcd = "" Then
		UNIValue(0,7)  = ""
	Else 
		If Request("lgAuthorityFlag") = "Y" Then      '권한관리 추가 
			UNIValue(0,7) = " AND C.ORG_CHANGE_ID  =  " & FilterVar(GetGlobalInf("gChangeOrgId"), "''", "S") & " and EXISTS ( SELECT 1 FROM z_usr_authority_value S WHERE a.dept_cd = S.code_value and S.usr_id =  " & FilterVar(gUsrId , "''", "S") & " AND S.module_cd = " & FilterVar("A", "''", "S") & "  )  "   '권한관리 추가 
		Else
			UNIValue(0,7)  = " AND C.ORG_CHANGE_ID  =  " & FilterVar(request("OrgChangeId"), "''", "S") & " "
		End If
	End If



 
	'rs1에 대한 Value값 setting
  	
	UNIValue(1,0)  = lgtxtFromGlDt	
	UNIValue(1,1)  = lgtxtToGlDt
	
	UNIValue(1,2)  = strSql 

	If lgtxtCOST_CENTER_CD = "" Then
		UNIValue(1,3)  = ""
	Else 
		UNIValue(1,3)  = " AND A.COST_CD = " & FilterVar(lgtxtCOST_CENTER_CD,"","S") 
	End If
	 
	If lgtxtdeptcd = "" Then
		UNIValue(1,4)  = ""
	Else 
		UNIValue(1,4)  = " AND A.DEPT_CD = " & FilterVar(lgtxtdeptcd,"","S")
	End If
	 
	If lgcboGlInputType = "" Then
		UNIValue(1,5)  = ""
	Else 
		UNIValue(1,5)  = " AND A.GL_INPUT_TYPE = " & FilterVar(lgcboGlInputType,"","S")
	End If
	
	If lgtxtdeptcd = "" Then
		UNIValue(1,6)  = ""
	Else 
		UNIValue(1,6)  = " AND A.ORG_CHANGE_ID =  " & FilterVar(request("OrgChangeId"), "''", "S") & " "
	End If



    
    'rs2에 대한 Value값 setting
	UNIValue(2,0) = FilterVar(lgtxtBizArea, "''", "S")
	



	'rs3에 대한 Value값 setting
	If lgtxtCOST_CENTER_CD = "" Then
		UNIValue(3,0)  = FilterVar(lgtxtCOST_CENTER_CD,"","S")				'입력된 값이 없을때 더미값을 넘겨준다 
	Else 
		UNIValue(3,0)  = FilterVar(lgtxtCOST_CENTER_CD,"","S")
	End If




	'rs4에 대한 Value값 setting
	If lgtxtdeptcd = "" Then
		UNIValue(4,0)  = FilterVar(lgtxtdeptcd,"","S")				'입력된 값이 없을때 더미값을 넘겨준다 
		UNIValue(4,1)  = FilterVar(lgtxtdeptcd,"","S")
	Else 
		UNIValue(4,0)  = FilterVar(lgtxtdeptcd,"","S")
		UNIValue(4,1)  = FilterVar(request("OrgChangeId"), "''", "S")
	End If



	UNIValue(5,0) = FilterVar(lgtxtBizArea1,"","S")

    
    UNIValue(6,0) = " select USR_ID,USR_NM from Z_USR_MAST_REC where USR_ID = " & FilterVar(UCase(lgtxtUsr_Id),"''","S")
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

       
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"										'☜: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg																'☜ : Record Set Return Message 변수선언 
    Dim iStr
    Dim lgADF																	'☜ : ActiveX Data Factory 지정 변수선언 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4,rs5,rs6)
    Set lgADF = Nothing															'☜: ActiveX Data Factory Object Nothing
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMsgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)				'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If
    
End Sub

%>

<Script Language=vbscript>

	If "<%=lgDataExist%>" = "Yes" Then
       'Set condition data to hidden area
		With parent
			If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
				.Frm1.htxtFromGlDt.Value		= .Frm1.txtFromGlDt.text
				.Frm1.htxtToGlDt.Value			= .Frm1.txtToGlDt.text
				.Frm1.htxtBizArea.Value			= .Frm1.txtBizArea.Value
				.Frm1.htxtBizArea1.Value		= .Frm1.txtBizArea1.Value					
				.Frm1.htxtCOST_CENTER_CD.Value  = .Frm1.txtCOST_CENTER_CD.Value
				.Frm1.htxtdeptcd.Value			= .Frm1.txtdeptcd.Value
				.Frm1.hcboGlInputType.Value     = .Frm1.cboGlInputType.Value
				.Frm1.htxtDesc.Value			= .Frm1.txtDesc.Value
				.Frm1.htxtRefNo.Value			= .Frm1.txtRefNo.Value
				.Frm1.htxtAmtFr.Value			= .Frm1.txtAmtFr.Value
				.Frm1.htxtAmtTo.Value			= .Frm1.txtAmtTo.Value
				.Frm1.htxtUsr_Id.Value			= .Frm1.txtUsr_Id.Value
			End If
       
			'Show multi spreadsheet data from this line       
			.ggoSpread.Source	= .frm1.vspdData      
			.ggoSpread.SSShowData "<%=lgstrData%>"                  '☜ : Display data
			.lgPageNo			=  "<%=lgPageNo%>"               '☜ : Next next data tag
       
			.frm1.txtDrlocAmt.value			= "<%=UNINumClientFormat(dr_loc_amt, ggAmtOfMoney.DecPoint, 0)%>"
			.frm1.txtCrlocAmt.value			= "<%=UNINumClientFormat(cr_loc_amt, ggAmtOfMoney.DecPoint, 0)%>"		
			.frm1.txtBizAreaNm.value		= "<%=biz_area_nm%>"
			.frm1.txtBizAreaNm1.value		= "<%=biz_area_nm1%>"		
			.frm1.txtCOST_CENTER_NM.value	= "<%=cost_nm%>"
			.frm1.txtdeptnm.value			= "<%=dept_nm%>"
			.frm1.txtUsr_nm.value			= "<%=Usr_nm%>"
		End With
       
		Parent.DbQueryOk
    Else
		With parent
			.frm1.txtDrlocAmt.value			= "<%=UNINumClientFormat(dr_loc_amt, ggAmtOfMoney.DecPoint, 0)%>"
			.frm1.txtCrlocAmt.value			= "<%=UNINumClientFormat(cr_loc_amt, ggAmtOfMoney.DecPoint, 0)%>"
			.frm1.txtBizArea.value			= ""
			.frm1.txtBizAreaNm.value		= ""
			.frm1.txtBizArea1.value			= ""
			.frm1.txtBizAreaNm1.value		= ""		
			.frm1.txtCOST_CENTER_Cd.value	= ""
			.frm1.txtCOST_CENTER_NM.value	= ""
			.frm1.txtdeptCd.value			= ""
			.frm1.txtdeptnm.value			= ""
			.frm1.cboGlInputType.Value		= ""
			.Frm1.txtDesc.Value				= ""
			.Frm1.txtRefNo.Value			= ""
			.Frm1.txtAmtFr.Text				= ""
			.Frm1.txtAmtTo.Text				= ""
			.frm1.txtUsr_Id.value			= ""
			.frm1.txtUsr_nm.value			= ""
		End with
	End If

</Script>	

