<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<!--'**********************************************************************************************
'*  1. Module Name          : ACCOUNT
'*  2. Function Name        : 
'*  3. Program ID           : a7108mb1
'*  4. Program Name         : 감가상각 상세조회 
'*  5. Program Desc         : 고정자산별 감가상각을 조회 
'*  6. Comproxy List        : +As0069LookupSvr
'                             +B1a028ListMinorCode
'*  7. Modified date(First) : 2002/03/26
'*  8. Modified date(Last)  : 2002/03/26
'*  9. Modifier (First)     : 황은희 
'* 10. Modifier (Last)      : 황은희 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                       
'********************************************************************************************** -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next								'☜: 

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2                              '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT

Dim lgDataExist

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim txtAsstNo
DIm txtDepryyyymm
Dim DurMnthFg
Dim Asst_NM

Dim txtAcctCd
Dim txtAcctNm
Dim cboDeprMthd
Dim txtRegDt
Dim txtLocAcqAmt
Dim txtAcqQty
Dim txtInvQty
Dim txtDurMnth
'Dim txtDeprRate

Dim txtEndLTermAcqAmt			'전기말 현재 
Dim txtEndLTermCptAmt
Dim txtEndLTermDeprAmt
Dim txtEndLTermBalAmt
Dim txtEndLTermInvQty

Dim txtFMnthAcqAmt				'당월초 
Dim txtFMnthCptAmt
Dim txtFMnthDeprAmt
Dim txtFMnthBalAmt
Dim txtFMnthInvQty

Dim txtMnthCptAmt				'당월발생 
Dim txtMnthDeprAmt
Dim txtMnthDisUseQty

Dim txtLMnthAcqAmt				'당월말 
Dim txtLMnthCptAmt
Dim txtLMnthDeprAmt
Dim txtLMnthBalAmt
Dim txtLMnthInvQty


' 권한관리 추가 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)				
Dim lgAuthUsrID, lgAuthUsrNm					' 개인 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL					

'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------

	Call HideStatusWnd
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "A", "NOCOOKIE", "MB")   'ggQty.DecPoint Setting...

	lgDataExist			= "No"
	txtAsstNo			= Trim(Request("txtAsstNo"))
	txtDepryyyymm		= Trim(Request("txtDepryyyymm"))
	DurMnthFg			= Trim(Request("DurMnthFg"))

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
    
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(2)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	Dim strWhereUP
	Dim strWhereDown
	
    Redim UNIValue(2,1)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 

	If DurMnthFg= "T" then
		UNISqlId(0) = "A7112MA01KO441"	'상위 
	else
		UNISqlId(0) = "A7112MA03KO441"	'상위 
	End If
	
    UNISqlId(1) = "A7112MA02"	'하위 
    UNISqlId(2) = "commonqry"	

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    'UNIValue(0,0) = lgSelectList                                          '☜: Select list
    
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	If DurMnthFg= "T" then
		'strWhereUP	= " and e.dur_yrs = a.tax_dur_yrs"
	else
		'strWhereUP	= " and e.dur_yrs = a.cas_dur_yrs"
	End If

	strWhereUP = ""
	strWhereUP = strWhereUP & " and d.major_cd	= " & FilterVar("a2002" , "''", "S") & ""  
	strWhereUP = strWhereUP & " and asst_no		= " & FilterVar(txtAsstNo , "''", "S") 
'Call ServerMesgBox(strWhereUP , vbInformation, I_MKSCRIPT)
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

	strWhereUP	= strWhereUP	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL	


	
    strWhereDown = "asst_no = " & FilterVar(txtAsstNo , "''", "S")
    strWhereDown =  strWhereDown & " and depr_yyyymm	= " & FilterVar(txtDepryyyymm , "''", "S")
    strWhereDown =  strWhereDown & " and dur_yrs_fg		= " & FilterVar(DurMnthFg ,"''" ,"S")	'회계기준구분

	' 자산코드로 가져오므로 상각정보는 권한관리 조건 없음 

	UNIValue(0,0) = strWhereUP
	UNIValue(1,0) = strWhereDown
	UNIValue(2,0) = "Select Asst_NM from a_asset_master(NOLOCK) where asst_no=" & FilterVar(txtAsstNo , "''", "S")
	    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'--------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
    Dim iStr
    Dim lgADF
                                                                      '☜ : ActiveX Data Factory 지정 변수선언 

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)   
   
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
    
    'rs2 자산명 가져오기 
    If Not(rs2.EOF Or rs2.BOF) Then
		Asst_NM = Trim(rs2("Asst_NM"))		
		rs2.Close
		Set rs2 = Nothing   
	Else
		Asst_NM = ""	
		Call DisplayMsgBox("117400", vbOKOnly, "", "", I_MKSCRIPT)     'No Data Found!!			
 %>
	<Script Language=vbscript>
		parent.frm1.txtAsstNo.focus
	</Script>
 <%
      rs2.Close
      Set rs2 = Nothing
	  Response.End
	  Exit sub 
	End IF
	
	If NOt(rs0.EOF or rs0.BOF) Then
		Call  MakeDataUP()
    Else 		
		Call DisplayMsgBox("117500", vbOKOnly, "", "", I_MKSCRIPT)     'No Data Found!!
 %>
	<Script Language=vbscript>
		parent.frm1.txtAsstNo.focus
	</Script>
 <%
        rs0.Close
        Set rs0 = Nothing
			Response.End 
        Exit Sub
    End If
	
	'rs1	
    If Not (rs1.EOF OR rs1.BOF) Then
		Call MakeDataDown()
	Else
		Call DisplayMsgBox("117500", vbOKOnly, "", "", I_MKSCRIPT)
%>
	<Script Language=vbscript>
		parent.frm1.txtDeprYyyymm.focus
	</Script>
 <%
			rs1.Close
			Set rs1 = Nothing
			Exit Sub
		
	End IF
End Sub
'----------------------------------------------------------------------------------------------------------
' Make Data 
'--------------------------------------------------------------------------------------------------------
Sub MakeDataUP()

	txtAcctCd	= Rs0(0)
	txtAcctNm	= Rs0(1)
	cboDeprMthd	= Rs0(2)
	txtRegDt	= Rs0(3)
	txtLocAcqAmt= Rs0(4)
	txtAcqQty	= Rs0(5)
	txtInvQty	= Rs0(6)
	txtDurMnth	= Rs0(7)
	'txtDeprRate	= Rs0(8)
End Sub

Sub MakeDataDown()

	lgDataExist    = "Yes"
	
	txtEndLTermAcqAmt	= Rs1(0)			       '전기말 현재 
	txtEndLTermCptAmt	= Rs1(1)	
	txtEndLTermDeprAmt	= Rs1(2)	
	txtEndLTermBalAmt	= Rs1(3)	
	txtEndLTermInvQty	= Rs1(4)	

	txtFMnthAcqAmt	= Rs1(5)	'당월초 
	txtFMnthCptAmt	= Rs1(6)	
	txtFMnthDeprAmt	= Rs1(7)	
	txtFMnthBalAmt	= Rs1(8)	
	txtFMnthInvQty	= Rs1(9)	

	txtMnthCptAmt	= Rs1(10)	'당월발생 
	txtMnthDeprAmt	= Rs1(11)	
	txtMnthDisUseQty= Rs1(12)	

	txtLMnthAcqAmt	= Rs1(13)	'당월말 
	txtLMnthCptAmt	= Rs1(14)	
	txtLMnthDeprAmt	= Rs1(15)	
	txtLMnthBalAmt	= Rs1(16)	
	txtLMnthInvQty	= Rs1(17)	

End Sub

%>

<Script Language=vbscript>

With Parent

	If "<%=lgDataExist%>" = "Yes" Then

    	.frm1.txtAcctCd.Value				= "<%=ConvSPChars(txtAcctCd)%>"				'계정코드 
		.frm1.txtAcctNm.value				= "<%=ConvSPChars(txtAcctNm)%>"				'계정명 
		.frm1.cboDeprMthd.value				= "<%=ConvSPChars(cboDeprMthd)%>"			'상각방법 
		.frm1.txtRegDt.text					= "<%=UNIDateClientFormat(txtRegDt)%>"				'등록일자 

		.frm1.txtLocAcqAmt.value			= "<%=UNINumClientFormat(txtLocAcqAmt, ggAmtOfMoney.DecPoint, 0)%>"		'취득금액(자국)
		.frm1.txtAcqQty.value				= "<%=UNINumClientFormat(txtAcqQty, ggQty.DecPoint, 0)%>"				'취득수량 
		.frm1.txtInvQty.Value				= "<%=UNINumClientFormat(txtInvQty, ggQty.DecPoint, 0)%>"				'재고수량 
		.frm1.txtDurMnth.value				= "<%=txtDurMnth%>"														'내용월수 
		'.frm1.txtDeprRate.value				= "<%=UNINumClientFormat(txtDeprRate, ggExchRate.DecPoint, 0)%>"			'상각율 

		'''''전기말 
		.frm1.txtEndLTermAcqAmt.value		= "<%=UNINumClientFormat(txtEndLTermAcqAmt, ggAmtOfMoney.DecPoint, 0)%>"	'취득가액 
		.frm1.txtEndLTermCptAmt.value		= "<%=UNINumClientFormat(txtEndLTermCptAmt, ggAmtOfMoney.DecPoint, 0)%>"	'자본적지출 
		.frm1.txtEndLTermDeprAmt.value		= "<%=UNINumClientFormat(txtEndLTermDeprAmt, ggAmtOfMoney.DecPoint, 0)%>"	'상각액 
		.frm1.txtEndLTermBalAmt.value		= "<%=UNINumClientFormat(txtEndLTermBalAmt, ggAmtOfMoney.DecPoint, 0)%>"	'미상각액 
		.frm1.txtEndLTermInvQty.value		= "<%=UNINumClientFormat(txtEndLTermInvQty, ggQty.DecPoint, 0)%>"	'재고량 
		
		'''''당월초			
		.frm1.txtFMnthAcqAmt.value			= "<%=UNINumClientFormat(txtFMnthAcqAmt, ggAmtOfMoney.DecPoint, 0)%>"				'취득가액 
		.frm1.txtFMnthCptAmt.value			= "<%=UNINumClientFormat(txtFMnthCptAmt, ggAmtOfMoney.DecPoint, 0)%>"			'자본적지출 
		.frm1.txtFMnthDeprAmt.value			= "<%=UNINumClientFormat(txtFMnthDeprAmt, ggAmtOfMoney.DecPoint, 0)%>"			'상각액 
		.frm1.txtFMnthBalAmt.value			= "<%=UNINumClientFormat(txtFMnthBalAmt, ggAmtOfMoney.DecPoint, 0)%>"			'미상각액 
		.frm1.txtFMnthInvQty.value			= "<%=UNINumClientFormat(txtFMnthInvQty, ggQty.DecPoint, 0)%>"			'재고량 
		'''''당월발생			
		.frm1.txtMnthCptAmt.value			= "<%=UNINumClientFormat(txtMnthCptAmt, ggAmtOfMoney.DecPoint, 0)%>"				'자본적지출 
		.frm1.txtMnthDeprAmt.value			= "<%=UNINumClientFormat(txtMnthDeprAmt, ggAmtOfMoney.DecPoint, 0)%>"			'상각액 
		.frm1.txtMnthDisUseQty.value		= "<%=UNINumClientFormat(txtMnthDisUseQty, ggQty.DecPoint, 0)%>"			'매각폐기량 
		'''''당월말	
		.frm1.txtLMnthAcqAmt.value			= "<%=UNINumClientFormat(txtLMnthAcqAmt, ggAmtOfMoney.DecPoint, 0)%>"				'취득가액 
		.frm1.txtLMnthCptAmt.value			= "<%=UNINumClientFormat(txtLMnthCptAmt, ggAmtOfMoney.DecPoint, 0)%>"				'자본적지출 
		.frm1.txtLMnthDeprAmt.value			= "<%=UNINumClientFormat(txtLMnthDeprAmt, ggAmtOfMoney.DecPoint, 0)%>"			'상각액 
		.frm1.txtLMnthBalAmt.value			= "<%=UNINumClientFormat(txtLMnthBalAmt, ggAmtOfMoney.DecPoint, 0)%>"				'미상각액 
		.frm1.txtLMnthInvQty.value			= "<%=UNINumClientFormat(txtLMnthInvQty, ggQty.DecPoint, 0)%>"				'재고량 

		.DbQueryOk
    End If
    .frm1.txtAsstNm.value = "<%=ConvSPChars(Asst_NM)%>"
End With
</Script>
