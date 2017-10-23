<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/incSvrDate.inc"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% 
Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("Q", "A","NOCOOKIE","QB")

On Error Resume Next

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0,rs1,rs2,rs3,rs4                              '☜ : DBAgent Parameter 선언Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgstrData 
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strIssueDT1, strIssueDT2, strBizAreaCd, strBPCd
Dim AmtSumI,VatSumI,CntSumI,AmtSumO,VatSumO,CntSumO
Dim strMsgCd, strMsg1, strMsg2 
Dim strWhere0, strWhere1, strWhere2
Dim strReportFg
Dim	strIOFlag 
Dim	strVatType 
		
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
	Const C_SHEETMAXROWS_D = 40
  
    Call HideStatusWnd 

    lgStrPrevKey   = Request("lgStrPrevKey")                               '☜ : Next key flag
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value

    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iCnt
    Dim  iRCnt
    Dim  iStr

    iCnt = 0

    If Len(Trim(lgStrPrevKey)) Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(lgStrPrevKey) Then
          iCnt = CInt(lgStrPrevKey)
       End If   
    End If   

    For iRCnt = 1 to iCnt  *  C_SHEETMAXROWS_D                                   '☜ : Discard previous data
        rs0.MoveNext
    Next

    iRCnt = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
        iStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
             iStr = iStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iRCnt < C_SHEETMAXROWS_D Then
            lgstrData      = lgstrData      & iStr & Chr(11) & Chr(12)
        Else
            iCnt = iCnt + 1
            lgStrPrevKey = CStr(iCnt)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iRCnt < C_SHEETMAXROWS_D Then                                            '☜: Check if next data exists
        lgStrPrevKey = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
'	rs0.Close
'	Set rs0 = Nothing 
'	Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(5)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    UNISqlId(0) = "A6113MA101"	
    UNISqlId(1) = "A6113MA102"	
    UNISqlId(2) = "A6113MA102"	
	UNISqlId(3) = "ATAXBIZNM"
	
	UNISqlId(4) = "ABPNM"
	
	Redim UNIValue(4,2)
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    UNIValue(0,1) = UCase(Trim(strWhere0))

	UNIValue(1,0) = UCase(Trim(strWhere1))
	UNIValue(2,0) = UCase(Trim(strWhere2))
	UNIValue(3,0) = FilterVar(UCase(Request("txtBizAreaCD")), "''", "S")  
    UNIValue(4,0) = FilterVar(UCase(Request("txtBPCd")), "''", "S") 
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,2) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2,rs3,rs4)
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If   
   
		 
      	
	AmtSumI = 0
	VatSumI = 0
	CntSumI = 0
	AmtSumO = 0
	VatSumO = 0
	CntSumO = 0
	
	If not(rs1.EOF And rs1.BOF) Then
		If IsNull(rs1(0)) = False then AmtSumI   = rs1(0)
		If IsNull(rs1(1)) = False then VatSumI   = rs1(1)
		If IsNull(rs1(2)) = False then CntSumI   = rs1(2)
	End If

		
	rs1.close
	Set rs1 = Nothing
	
	If not(rs2.EOF And rs2.BOF) Then
		If IsNull(rs2(0)) = False then AmtSumO   = rs2(0)
		If IsNull(rs2(1)) = False then VatSumO   = rs2(1)
		If IsNull(rs2(2)) = False then CntSumO   = rs2(2)
	End If
	rs2.close
	Set rs2 = Nothing
	
	
		
%>
		<Script Language=vbscript>
		With parent.frm1
		.txtAmtSumI.value		= "<%=UNINumClientFormat(AmtSumI, ggAmtOfMoney.DecPoint, 0)%>"
   		.txtVatSumI.value		= "<%=UNINumClientFormat(VatSumI, ggAmtOfMoney.DecPoint, 0)%>"
   		.txtCntSumI.value		= "<%=UNINumClientFormat(CntSumI, ggAmtOfMoney.DecPoint, 0)%>"
   		.txtAmtSumO.value		= "<%=UNINumClientFormat(AmtSumO, ggAmtOfMoney.DecPoint, 0)%>"
   		.txtVatSumO.value		= "<%=UNINumClientFormat(VatSumO, ggAmtOfMoney.DecPoint, 0)%>"
   		.txtCntSumO.value		= "<%=UNINumClientFormat(CntSumO, ggAmtOfMoney.DecPoint, 0)%>"
		
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
		If strMsgCd = "" Then strMsgCd = "900014"
'		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
'		rs0.Close
'		Set rs0 = Nothing
'		Response.End													'☜: 비지니스 로직 처리를 종료함 
    Else    
        Call  MakeSpreadSheetData()
    End If

        	
    %>
    
    <%
	
	rs0.Close
	Set rs0 = Nothing 
	Set lgADF = Nothing  	
	                                                  '☜: ActiveX Data Factory Object Nothing
	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
		'Response.End 
	
	End If
End Sub

Sub ReleaseObj()
	rs0.Close:	Set rs0 = Nothing
	rs3.Close:	Set rs3 = Nothing
	rs4.Close:	Set rs4 = Nothing
	Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()  
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	strIssueDT1 = UNIConvDate(Request("txtIssueDT1"))
	strIssueDT2 = UNIConvDate(Request("txtIssueDT2"))
	strBizAreaCd = UCase(Request("txtBizAreaCd"))
	strBPCd = UCase(Request("txtBPCd"))
	
	strReportFg = UCase(Request("cboReportFg"))
	strIOFlag = UCase(Request("cboIOFlag"))
	strVatType = UCase(Request("cboVatType"))
	
	strWhere0 = ""
	strWhere0 = strWhere0 & "  a.issued_dt between  " & FilterVar(strIssueDT1, "''", "S") & " and  " & FilterVar(strIssueDT2, "''", "S") & " "

	IF Trim(strBizAreaCd) <> "" Then
		strWhere0 = strWhere0 & " and A.REPORT_BIZ_AREA_CD = " & FilterVar(strBizAreaCd, "''", "S") 
	End If
	strWhere0 = strWhere0 & " AND	A.conf_fg = " & FilterVar("C", "''", "S") & "  AND	B.major_cd = " & FilterVar("a1003", "''", "S") & "  AND B.minor_cd = A.io_fg "
	strWhere0 = strWhere0 & " AND	C.major_cd = " & FilterVar("b9001", "''", "S") & "  AND C.minor_cd = A.vat_type AND	D.bp_cd = A.bp_cd"
	strWhere0 = strWhere0 & " AND	D.valid_from_dt = (SELECT MAX(valid_from_dt) FROM b_biz_partner_history BP"
	strWhere0 = strWhere0 & " WHERE BP.bp_cd = A.bp_cd AND BP.valid_from_dt <= A.issued_dt) "
 
			 
			
	If strReportFg <> "" Then
		strWhere0 = strWhere0 & " and A.MADE_VAT_FG = " & FilterVar(strReportFg, "''", "S")
	End If
	If strBPCd <> "" Then
		strWhere0 = strWhere0 & " and A.BP_CD = " & FilterVar(strBPCd, "''", "S") 

	End If
	If strIOFlag <> "" Then
		strWhere0 = strWhere0 & " and A.IO_FG = " & FilterVar(strIOFlag, "''", "S")  
	End If
	If strVatType <> "" Then
		strWhere0 = strWhere0 & " and A.VAT_TYPE = " & FilterVar(strVatType, "''", "S")  
	End If	

	strWhere1 = ""
	strWhere1 = strWhere0 & " and A.IO_FG = " & FilterVar("I", "''", "S") & " "

	
	strWhere2 = ""
	strWhere2 = strWhere0 & " and A.IO_FG = " & FilterVar("O", "''", "S") & " "

End Sub    

%>



<Script Language=vbscript>
	With parent
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowData "<%=lgstrData%>"                            '☜: Display data 
		.lgStrPrevKey =  "<%=ConvSPChars(lgStrPrevKey)%>"                       '☜: set next data tag
		.DbQueryOk
	End with
	
</Script>	

<%
	Response.End 
%>

