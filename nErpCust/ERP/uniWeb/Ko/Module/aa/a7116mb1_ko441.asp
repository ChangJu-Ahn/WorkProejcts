<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% Option Explicit%>
<% session.CodePage=949 %>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% 


On Error Resume Next

Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3 , rs4, rs5, rs6, rs7   '☜ : DBAgent Parameter 선언 
Dim lgstrData																		'☜ : data for spreadsheet data
Dim lgStrPrevKey																	'☜ : 이전 값 
Dim lgMaxCount																		'☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList																		'☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

Dim txtFryymm
Dim txtToyymm
Dim txtDurYrsFg
Dim txtDeptCd
Dim txtDeptNm
Dim txtAcctCd
Dim txtAcctNm
Dim txtCondAsstNo
Dim txtCondAsstNm
Dim strBizAreaCd															'⊙ : 시작사업장 
Dim strBizAreaNm
Dim strBizAreaCd1															'⊙ : 종료사업장 
Dim strBizAreaNm1
Dim txtBizUnitCd
Dim txtBizUnitNm
Dim txtInternalCd 
Dim txtRFryymmdd
Dim txtRToyymmdd

Dim strMsgCd, strMsg1, strMsg2

Dim iPrevEndRow
Dim iEndRow

' 권한관리 추가
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' 사업장
Dim lgInternalCd, lgDeptCd, lgDeptNm			' 내부부서		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' 내부부서(하위포함)				
Dim lgAuthUsrID, lgAuthUsrNm					' 개인

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL					


    Call LoadBasisGlobalInf()    

    Call LoadInfTB19029B("Q","A","NOCOOKIE","QB")   
    Call LoadBNumericFormatB("Q", "A","NOCOOKIE","QB") 

    Call HideStatusWnd 


    lgPageNo		= UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount		= CInt(Request("lgMaxCount"))                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList	= Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT	= Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList		= Request("lgTailList")                                 '☜ : Orderby value
    lgDataExist		= "No"

    txtFryymm		= Trim(Request("txtFryymm"))
    txtToyymm		= Trim(Request("txtToyymm"))
    txtDurYrsFg		= Trim(Request("DurYrsFg"))

    txtRFryymmdd        = Trim(Request("txtRFryymmdd"))
    txtRToyymmdd        = Trim(Request("txtRToyymmdd"))

	txtDeptCd		= Trim(Request("txtDeptCd"))
	txtAcctCd		= Trim(Request("txtAcctCd"))
	txtCondAsstNo	= Trim(Request("txtCondAsstNo"))
	txtBizUnitCd	= Trim(Request("txtBizUnitCd"))
	
	strBizAreaCd	= Trim(Ucase(Request("txtBizAreaCd")))					'사업장From
	strBizAreaCd1	= Trim(Ucase(Request("txtBizAreaCd1")))					'사업장To
	txtInternalCd= Trim(Ucase(Request("txtinternalCd")))	
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
    
    lgDataExist    = "Yes"
    lgstrData      = ""

    iPrevEndRow = 0

    If CDbl(lgPageNo) > 0 Then
		iPrevEndRow = CDbl(lgMaxCount) * CDbl(lgPageNo)    
		rs0.Move= iPrevEndRow                 
    End If

    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
 
        If  iLoopCount < lgMaxCount Then
            lgstrData		=	lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop
    If  iLoopCount < lgMaxCount Then                                            '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
        iEndRow = iPrevEndRow + iLoopCount + 1
    Else
        iEndRow = iPrevEndRow + iLoopCount
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(7)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Dim strWhere, strWhere2

    Redim UNIValue(7,8)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 

    UNISqlId(0) = "a7116ma1_ko441"
    UNISQLID(1) = "commonqry"
    UNISQLID(2) = "commonqry"
    UNISQLID(3) = "commonqry"
	UNISqlId(4) = "a7116ma2_ko441"
	UNISqlId(5) = "A_GETBIZ"
	UNISqlId(6) = "A_GETBIZ"
	UNISQLID(7) = "commonqry"

    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    
    strWhere = ""
    strWhere2 = ""

    strWhere2 = " S.DEPR_YYYYMM BETWEEN " & FilterVar(txtFryymm ,"''"	,"S") & " AND " & FilterVar(txtToyymm ,"''"	,"S")

	'strWhere = strWhere & " A.ASST_NO >= " & FilterVar(txtCondAsstNo ,"''"	,"S")		'자산번호 
	
	If txtCondAsstNo <>"" Then
		strWhere = strWhere & " A.ASST_NO = " & FilterVar(txtCondAsstNo ,"''"	,"S")
	Else
		strWhere = strWhere & " A.ASST_NO >= " & FilterVar(txtCondAsstNo ,"''"	,"S")	
	End If
	
		
    'If txtDeptCd <> "" Then
	'	strWhere = strWhere & " AND ISNULL(ISNULL(G.TO_DEPT_CD,G.FROM_DEPT_CD),A.DEPT_CD) = " & FilterVar(txtDeptCd ,"''"	,"S")		'관리부서 
	'End If
    '2008.07.25 by Lws
	if txtInternalCd <> "" then
		strWhere = strWhere & " AND ISNULL(ISNULL(G.TO_INTERNAL_CD,G.FROM_INTERNAL_CD),A.INTERNAL_CD) = " & FilterVar(txtInternalCd ,"''"	,"S")
	end if
	
	If txtAcctCd <> "" Then
		strWhere = strWhere & " AND A.ACCT_CD = " & FilterVar(txtAcctCd ,"''"	,"S")		'계정명 
	ENd If
	
	If strBizAreaCd <> "" Then
		strWhere = strWhere & " AND ISNULL(ISNULL(G.TO_BIZ_AREA_CD,G.FROM_BIZ_AREA_CD),A.BIZ_AREA_CD) >= "	& FilterVar(strBizAreaCd , "''", "S")	'사업장
	Else
		strWhere = strWhere & " AND ISNULL(ISNULL(G.TO_BIZ_AREA_CD,G.FROM_BIZ_AREA_CD),A.BIZ_AREA_CD) >= " & FilterVar("0", "''", "S") & " "
	End If
	
	If strBizAreaCd1 <> "" Then
		strWhere = strWhere & " AND ISNULL(ISNULL(G.TO_BIZ_AREA_CD,G.FROM_BIZ_AREA_CD),A.BIZ_AREA_CD) <= "	& FilterVar(strBizAreaCd1 , "''", "S") 
	Else
		strWhere = strWhere & " AND ISNULL(ISNULL(G.TO_BIZ_AREA_CD,G.FROM_BIZ_AREA_CD),A.BIZ_AREA_CD) <= " & FilterVar("ZZZZZZZZZZ", "''", "S") & " "
	End If
	
	If txtBizUnitCd <> "" Then
		strWhere = strWhere & " AND J.biz_unit_cd = " & FilterVar(txtBizUnitCd ,"''"	,"S")		'사업부
	ENd If
	
	'if strBizAreaCd <> "" then
	'	strWhere = strWhere &" AND ISNULL(G.TO_BIZ_AREA_CD,G.FROM_BIZ_AREA_CD) = "	& FilterVar(strBizAreaCd ,"''","S") 
	'end if
	
        '조회조건 추가(취득일자 From ~ To)
                strWhere = strWhere &" AND A.REG_DT BETWEEN " & FilterVar(txtRFryymmdd ,"''"	,"S") & " AND " & FilterVar(txtRToyymmdd ,"''"	,"S")               

	' 권한관리 추가
	If lgAuthBizAreaCd <> "" Then			
		lgBizAreaAuthSQL		= " AND ISNULL(G.TO_BIZ_AREA_CD,G.FROM_BIZ_AREA_CD) = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		lgInternalCdAuthSQL		= " AND  (case when ISNULL(C.TO_INTERNAL_CD,'') <> '' then C.TO_INTERNAL_CD else C.FROM_INTERNAL_CD end) = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		lgSubInternalCdAuthSQL	= " AND  (case when ISNULL(C.TO_INTERNAL_CD,'') <> '' then C.TO_INTERNAL_CD else C.FROM_INTERNAL_CD end) LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		lgAuthUsrIDAuthSQL		= " AND A.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  

	End If	

      
 
	strWhere	= strWhere	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL	



	UNIValue(0,1)  = FilterVar(txtDurYrsFg ,"''"	,"S")
	UNIValue(0,2)  = strWhere2
	UNIValue(0,3)  = FilterVar(txtDurYrsFg ,"''"	,"S")
	UNIValue(0,4)  = strWhere2
	UNIValue(0,5)  = FilterVar(txtDurYrsFg ,"''"	,"S")
	UNIValue(0,6)  = strWhere2
	UNIValue(0,7)  = strWhere

	UNIValue(0,8)  = lgTailList
	

 	'UNIValue(4,0) = "B.DUR_YRS_FG =" & FilterVar(txtDurYrsFg ,"''","S") & " AND B.DEPR_YYYYMM BETWEEN " & FilterVar(txtFryymm ,"''"	,"S") & " AND " & FilterVar(txtToyymm ,"''"	,"S") & " AND " & strWhere
 	'UNIValue(4,0)  = "SUM(isnull(F.당월감가상각액,0))	"
 	UNIValue(4,0)  = "SUM(CASE WHEN C.BAL_FG ='CR' THEN  F.당월감가상각액*-1 ELSE F.당월감가상각액 END)      "

 
 	UNIValue(4,1)  = FilterVar(txtDurYrsFg ,"''"	,"S")
	UNIValue(4,2)  = strWhere2
	UNIValue(4,3)  = FilterVar(txtDurYrsFg ,"''"	,"S")
	UNIValue(4,4)  = strWhere2
	UNIValue(4,5)  = FilterVar(txtDurYrsFg ,"''"	,"S")
	UNIValue(4,6)  = strWhere2
	UNIValue(4,7)  = strWhere
 	
 	
 	
 	

	UNIValue(1,0) = "select DEPT_NM from B_ACCT_DEPT Where dept_cd= " & FilterVar(txtDeptCd ,"''"	,"S")
	UNIValue(2,0) = "select acct_nm from A_ACCT Where acct_cd = " & FilterVar(txtAcctCd ,"''"	,"S")
	UNIValue(3,0) = "select asst_nm from A_ASSET_MASTER Where asst_no = " & FilterVar(txtCondAsstNo ,"''"	,"S")
	UNIValue(5,0)  = FilterVar(strBizAreaCd,"''","S")
	UNIValue(6,0)  = FilterVar(strBizAreaCd1,"''","S")
	UNIValue(7,0) = "select cost_nm from B_cost_center Where biz_unit_cd = " & FilterVar(txtBizUnitCd ,"''"	,"S")

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode

End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
    Dim iStr
    Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 



    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
  
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6, rs7)
    
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    






'rs1 관리부서
    If txtDeptCd <> "" Then
		If Not (rs1.EOF OR rs1.BOF) Then
			txtDeptNm = Trim(rs1("Dept_Nm"))

		Else
	'	Response.Write "txtDeptCd no" & "<br>"
			txtDeptNm = ""
			Call DisplayMsgBox("127800", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
		    rs1.Close
		    Set rs1 = Nothing 
		    Exit sub
		End IF
		rs1.Close
		Set rs1 = Nothing
	End If


'rs0
    If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("117400", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
        rs0.Close
        Set rs0 = Nothing
        Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If
    
    
'rs2 계정코드
    If txtAcctCd <> "" Then
		If Not (rs2.EOF OR rs2.BOF) Then
			txtAcctNm = Trim(rs2("acct_nm"))
		Else
			txtAcctNm = ""
			Call DisplayMsgBox("110100", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
		    rs2.Close
		    Set rs2 = Nothing 
		    Exit sub
		End IF
		rs2.Close
		Set rs2 = Nothing
	End If
    
'rs3 자산번호
    If txtCondAsstNo <> "" Then
		If Not (rs3.EOF OR rs3.BOF) Then
			txtCondAsstNm = Trim(rs3("asst_nm"))
		Else
			txtCondAsstNm = ""
		End IF
		rs3.Close
		Set rs3 = Nothing
    End If

'rs5 시작사업장
	If (rs5.EOF And rs5.BOF) Then
		If strMsgCd = "" and strBizAreaCd <> ""  Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBizAreaCd_ALT")
		End If
    Else
%>
	<Script Language=vbScript>
	With parent
		.frm1.txtBizAreaCd.value = "<%=Trim(rs5(0))%>"
		.frm1.txtBizAreaNm.value = "<%=Trim(rs5(1))%>"					
	End With
	</Script>
<%
    End If
	
	rs5.Close
	Set rs5 = Nothing   
    
    
	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
		Response.End 
	End If
	
'rs6 종료사업장
	If (rs6.EOF And rs6.BOF) Then
		If strMsgCd = "" and strBizAreaCd1 <> ""  Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBizAreaCd1_ALT")
		End If
    Else
%>
	<Script Language=vbScript>
	With parent
		.frm1.txtBizAreaCd1.value = "<%=Trim(rs6(0))%>"
		.frm1.txtBizAreaNm1.value = "<%=Trim(rs6(1))%>"					
	End With
	</Script>
<%
    End If
	
	rs6.Close
	Set rs6 = Nothing   
    
    
	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
		Response.End 
	End If

'rs7 사업부
    If txtBizUnitCd <> "" Then
		If Not (rs7.EOF OR rs7.BOF) Then
			txtBizUnitNm = Trim(rs7("cost_nm"))
		Else
			txtBizUnitNm = ""
		End IF
		rs7.Close
		Set rs7 = Nothing
	End If
	
'rs4 당기상각합계
    If Not (rs4.EOF OR rs4.BOF) Then
%>
		<Script Language=vbScript>
		With parent
			
			.frm1.txtSum1.Text  = "<%=UNINumClientFormat(Trim(rs4(0)), ggAmtOfMoney.DecPoint, 0)%>"
		End With
		</Script>
<%	Else %>
		<Script Language=vbScript>
		With parent
			.frm1.txtSum1.Text = "<%=UNINumClientFormat(0, ggAmtOfMoney.DecPoint, 0)%>"
		End With
		</Script>
<%	
	End IF

	rs4.Close
	Set rs4 = Nothing    
    
End Sub

%>


<Script Language=vbscript>
With Parent
	If "<%=lgDataExist%>" = "Yes" Then

       'Set condition data to hidden area
		If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			.Frm1.htxtFr_dt.Value   	= .Frm1.txtFr_dt.text
			.Frm1.htxtTo_dt.Value   	= .Frm1.txtTo_dt.text
			.Frm1.htxtDeptCd.Value		= .Frm1.txtDeptCd.Value             
			.Frm1.htxtAcctCd.Value		= .Frm1.txtAcctCd.Value
			.Frm1.htxtCondAsstNo.Value	= .Frm1.txtCondAsstNo.Value
			.frm1.htxtBizAreaCd.value	= .frm1.txtBizAreaCd.value
			.frm1.htxtBizAreaCd1.value	= .frm1.txtBizAreaCd1.value
			.frm1.htxtBizUnitCd.value	= .frm1.txtBizUnitCd.value
		        if .frm1.rdoDurYrsFg(0).checked then 
  			  .frm1.hDurYrsFg.value	        = "C"
                        else
  			  .frm1.hDurYrsFg.value	        = "T"
		        End If
			
		End If

		Parent.ggoSpread.Source  = Parent.frm1.vspdData
		Parent.frm1.vspdData.Redraw = False
		Parent.ggoSpread.SSShowData "<%=lgstrData%>", "F"                    '☜ : Display data

		Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,<%=iPrevEndRow+1%>,<%=iEndRow%>,Parent.GetKeyPos("A",2),Parent.GetKeyPos("A",3),"A", "Q" ,"X","X")
		Parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
		Parent.DbQueryOk
		Parent.frm1.vspdData.Redraw = True
    End If


'	.frm1.txtDeptNm.value = "<%=ConvSPChars(txtDeptNm)%>"			

	.frm1.txtAcctNm.value = "<%=ConvSPChars(txtAcctNm)%>"			'rs2 값 받기 팝업으로 안하고 그냥 입력했을때 값넣어주기 
	.frm1.txtCondAsstNm.value = "<%=ConvSPChars(txtCondAsstNm)%>"	'rs3 값 받기 팝업으로 안하고 그냥 입력했을때 값넣어주기 	
	.frm1.txtBizUnitNm.value = "<%=ConvSPChars(txtBizUnitNm)%>"		'rs7 값 받기 팝업으로 안하고 그냥 입력했을때 값넣어주기
	.frm1.txtDeptCd.focus
End With
</Script>