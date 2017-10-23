<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% 
Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("Q", "A","NOCOOKIE","MB")
'Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4          '☜ : DBAgent Parameter 선언 
Const C_SHEETMAXROWS_D = 100

Call HideStatusWnd 
Call FixUNISQLData()
Call QueryData()

Dim strSql 

Sub FixUNISQLData()
On Error Resume Next
	Dim intI
    Redim UNISqlId(4)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    UNISqlId(0) = "a6101mb101":	UNISqlId(1) = "a6101mb102":	UNISqlId(2) = "a6101mb103"
    UNISqlId(3) = "ATAXBIZNM" : UNISqlId(4) = "ABPNM"

    Redim UNIValue(4,5)
	
	strSql = "Select	dbo.ufn_A_GetBpRgStNo(c.bp_rgst_no) own_rgst_no," & vbCr
	strSql = strSql & "	dbo.ufn_bpname_max(c.bp_rgst_no, " & FilterVar(UNIConvDate(Request("txtIssueDT2")), "''", "S") & ") bp_full_nm," & vbCr
	strSql = strSql & " dbo.ufn_indType_name_max(c.bp_rgst_no, " & FilterVar(UNIConvDate(Request("txtIssueDT2")), "''", "S") & ") ind_type_nm," & vbCr
	strSql = strSql & " dbo.ufn_indClass_name_max(c.bp_rgst_no, " & FilterVar(UNIConvDate(Request("txtIssueDT2")), "''", "S") & ") ind_class_nm," & vbCr
	strSql = strSql & " COUNT(A.vat_no) cnt, SUM(A.net_loc_amt) net_loc_amt,SUM(A.vat_loc_amt) vat_loc_amt " & vbCr
	strSql = strSql & " FROM  	a_vat A" & vbCr
	strSql = strSql & " JOIN 	b_configuration B " & vbCr
	strSql = strSql & " ON  	B.minor_cd = A.vat_type " & vbCr
	strSql = strSql & " AND 	B.major_cd = " & FilterVar("B9001", "''", "S") & " " & vbCr
	strSql = strSql & " AND 	B.seq_no = 3 AND B.reference = " & FilterVar("Y", "''", "S") & " "   & vbCr
	strSql = strSql & " JOIN 	B_BIZ_PARTNER_HISTORY C" & vbCr
	strSql = strSql & " ON 		A.bp_cd = C.bp_cd" & vbCr
	strSql = strSql & " LEFT JOIN b_minor D ON D.major_cd = " & FilterVar("B9002", "''", "S") & " " & vbCr
	strSql = strSql & " AND 	D.minor_cd = C.ind_type"   & vbCr
	strSql = strSql & " LEFT JOIN b_minor E ON E.major_cd = " & FilterVar("B9003", "''", "S") & " "   & vbCr
	strSql = strSql & " AND 	E.minor_cd = C.ind_class"	 & vbCr
	
	If Request("cboIOFlag") = "O" Then
		strSql = strSql & " JOIN 	b_configuration G " & vbCr
		strSql = strSql & " ON  	G.minor_cd = A.vat_type " & vbCr
	End If
	
	strSql = strSql & " Where	A.conf_fg = " & FilterVar("C", "''", "S") & " " & vbCr
	
	If Request("cboIOFlag") = "O" Then
		strSql = strSql & " AND		G.Major_Cd = " & FilterVar("B9001", "''", "S") & " " & vbCr
		strSql = strSql & " AND		G.Seq_No = 5 " & vbCr
		strSql = strSql & " AND		G.Reference =  " & FilterVar("N", "''", "S") & " " & vbCr
		strSql = strSql & " AND		B.MAJOR_CD = G.MAJOR_CD " & vbCr
		strSql = strSql & " AND		B.MINOR_CD = G.MINOR_CD " & vbCr
		strSql = strSql & " AND		A.Vat_Type <>   " & FilterVar("D", "''", "S") & "  " & vbCr
	End IF
	
	strSql = strSql & " AND 	C.valid_from_dt = (SELECT MAX(F.valid_from_dt)  FROM B_BIZ_PARTNER_HISTORY F"   & vbCr
	strSql = strSql & " 			WHERE   F.BP_CD = A.BP_CD AND F.valid_from_dt <= A.issued_dt)" & vbCr
	strSql = strSql & " AND  	A.issued_dt >=  " & FilterVar(UNIConvDate(Request("txtIssueDT1")), "''", "S") & ""   & vbCr
	strSql = strSql & " AND		A.issued_dt <=  " & FilterVar(UNIConvDate(Request("txtIssueDT2")), "''", "S") & ""   & vbCr
	strSql = strSql & " AND 	A.io_fg = " & FilterVar(Request("cboIOFlag"), "''", "S")  & vbCr

	If "" & UCase(Trim(Request("txtBizAreaCD"))) <> "" Then
		strSql = strSql & " AND		A.report_biz_area_cd = " & FilterVar(UCase(Request("txtBizAreaCD")), "''", "S") & vbCr 'jsk 20030825
	End IF

	If "" & UCase(Trim(Request("txtBPCd"))) <> "" Then
		strSql = strSql & " AND		A.bp_cd = " & FilterVar(UCase(Request("txtBPCd")), "''", "S")  & vbCr
	End IF

	strSql = strSql & " Group By" & vbCr
	strSql = strSql & " dbo.ufn_A_GetBpRgStNo(c.bp_rgst_no)," & vbCr
	strSql = strSql & "	dbo.ufn_bpname_max(c.bp_rgst_no, " & FilterVar(UNIConvDate(Request("txtIssueDT2")), "''", "S") & ")," & vbCr
	strSql = strSql & " dbo.ufn_indType_name_max(c.bp_rgst_no, " & FilterVar(UNIConvDate(Request("txtIssueDT2")), "''", "S") & ")," & vbCr
	strSql = strSql & " dbo.ufn_indClass_name_max(c.bp_rgst_no, " & FilterVar(UNIConvDate(Request("txtIssueDT2")), "''", "S") & ") " & vbCr
	
	
	UNIValue(0,0) = strSql

	
	For intI = 1 to 2
		UNIValue(intI,0) = "" & FilterVar(UNIConvDate(Request("txtIssueDT2")), "''", "S") 
		UNIValue(intI,1) = "" & FilterVar(UNIConvDate(Request("txtIssueDT1")), "''", "S") 
		UNIValue(intI,2) = "" & FilterVar(UNIConvDate(Request("txtIssueDT2")), "''", "S") 
		UNIValue(intI,3) = "" & FilterVar(Request("cboIOFlag"), "''", "S") 


		If "" & UCase(Trim(Request("txtBizAreaCD"))) = "" Then
			UNIValue(intI,4) = "|"
		Else
			UNIValue(intI,4) = "" & FilterVar(UCase(Request("txtBizAreaCD")), "''", "S")
		End If
		If "" & UCase(Trim(Request("txtBPCd"))) = "" Then
			UNIValue(intI,5) = "|"
		Else
			UNIValue(intI,5) = "" & FilterVar(UCase(Request("txtBPCd")), "''", "S")
		End If
	Next

	UNIValue(3,0) = FilterVar(UCase(Request("txtBizAreaCD")), "''", "S") 

	UNIValue(4,0) = FilterVar(UCase(Request("txtBPCd")), "''", "S") 
	
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub

Sub QueryData()
	On Error Resume Next
    Dim iStr
    
    If UNIValue(0,0) = "" Then
		Exit Sub
    End If
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    'lgstrRetMsg = lgADF.QryRs("3", UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2)
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)
    iStr = Split(lgstrRetMsg,gColSep)
  
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If

	%>
	<Script Language=vbscript>
Option Explicit
	
			With parent									'☜: 화면 처리 ASP 를 지칭함 
    			.frm1.txtCntPer.text = "<%=UNINumClientFormat(rs2("cnt"), ggAmtOfMoney.DecPoint, 0)%>"
    			.frm1.txtAmtPer.text = "<%=UNINumClientFormat(rs2("net_loc_amt"), ggAmtOfMoney.DecPoint, 0)%>"
				.frm1.txtTaxPer.text = "<%=UNINumClientFormat(rs2("vat_loc_amt"), ggAmtOfMoney.DecPoint, 0)%>"
				
				.frm1.txtCntSum.text = "<%=UNINumClientFormat(rs1("cnt"), ggAmtOfMoney.DecPoint, 0)%>"
				.frm1.txtAmtSum.text = "<%=UNINumClientFormat(rs1("net_loc_amt"), ggAmtOfMoney.DecPoint, 0)%>"
				.frm1.txtTaxSum.text = "<%=UNINumClientFormat(rs1("vat_loc_amt"), ggAmtOfMoney.DecPoint, 0)%>"
			End With			
	</Script>	
	<%
	
	If rs3.EOF And rs3.BOF Then
		If "" & UCase(Trim(Request("txtBizAreaCD"))) <> "" Then
			Call DisplayMsgBox("124200", vbOKOnly, "", "", I_MKSCRIPT)
			Call ReleaseObj()
			Response.End
		End If
	Else
	%>
	<Script Language=vbscript>
		parent.frm1.txtBizAreaCD.value = "<%=ConvSPChars(rs3("TAX_BIZ_AREA_CD"))%>"
		parent.frm1.txtBizAreaNM.value = "<%=ConvSPChars(rs3("TAX_BIZ_AREA_NM"))%>"
	</Script>
	<%
	End If
	
	If rs4.EOF And rs4.BOF Then
		If "" & UCase(Trim(Request("txtBPCd"))) <> "" Then
			Call DisplayMsgBox("126100", vbOKOnly, "", "", I_MKSCRIPT)
			Call ReleaseObj()
			Response.End
		End If
	Else
	%>
	<Script Language=vbscript>
		parent.frm1.txtBPCd.value = "<%=ConvSPChars(rs4("BP_CD"))%>"
		parent.frm1.txtBPNm.value = "<%=ConvSPChars(rs4("BP_NM"))%>"
	</Script>
	<%
	End If
				
    If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
	Else    
        Call  MakeSpreadSheetData()
    End If
    
    Call ReleaseObj()
End Sub

Sub ReleaseObj()
	rs0.Close:	Set rs0 = Nothing
	rs1.Close:	Set rs1 = Nothing
	rs2.Close:	Set rs2 = Nothing
	rs3.Close:	Set rs3 = Nothing
	rs4.Close:	Set rs4 = Nothing
	Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
End Sub

Sub MakeSpreadSheetData()
	Dim LngRow
	%>
	<Script Language=vbscript>
Option Explicit

	Dim LngMaxRow       
	Dim strData

		LngMaxRow = parent.frm1.vspdData.MaxRows										'Save previous Maxrow                                                

	<%
		For LngRow = 1 To rs0.recordcount
	%>
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("own_rgst_no"))%>" 	'사업자등록번호			
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("bp_full_nm"))%>"   '거래처명(상호)
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ind_type_nm"))%>"			
			strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ind_class_nm"))%>"
			strData = strData & Chr(11) & "<%=UNINumClientFormat(rs0("cnt"), ggAmtOfMoney.DecPoint, 0)%>"
			strData = strData & Chr(11) & "<%=UNINumClientFormat(rs0("net_loc_amt"), ggAmtOfMoney.DecPoint, 0)%>"
			strData = strData & Chr(11) & "<%=UNINumClientFormat(rs0("vat_loc_amt"), ggAmtOfMoney.DecPoint, 0)%>"
						
			strData = strData & Chr(11) & LngMaxRow + <%=LngRow%>
			strData = strData & Chr(11) & Chr(12)
	<%		
			rs0.MoveNext
		Next			
	%>
			parent.ggoSpread.Source = parent.frm1.vspdData 
			parent.ggoSpread.SSShowData strData
			Call parent.DbQueryOk()
	</script>
	<%      
End Sub
	
%>

