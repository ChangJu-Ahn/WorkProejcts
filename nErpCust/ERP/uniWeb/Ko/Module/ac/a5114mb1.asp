<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->


<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next
Err.Clear

Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("Q", "A", "NOCOOKIE", "QB")

Dim lgADF																	'�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg																'�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2						'�� : DBAgent Parameter ���� 
Dim lgPageNo																'�� : ���� �� 
Dim lgTailList																'�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgstrData 
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strBizAreaCd, strBizAreaNm,strBizAreaCd1, strBizAreaNm1
Dim strMsgCd, strMsg1, strMsg2 
Dim strMode	
Dim strWhere0, strWhere1, strWhere2											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

Dim StrNextseq		' ���� �� 

DIM strfiscdt, strfiscyymm, strfiscdd, strcond

Dim sFromGlDt, sToGlDt

	Call HideStatusWnd
	Const C_SHEETMAXROWS_D  = 100														'��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 

	lgPageNo		= Request("lgPageNo")											'�� : Next key flag
	lgSelectList	= Request("lgSelectList")										'�� : select ����� 
	lgSelectListDT	= Split(Request("lgSelectListDT"), gColSep)						'�� : �� �ʵ��� ����Ÿ Ÿ�� 
	lgTailList		= Request("lgTailList")											'�� : Orderby value

	sFromGlDt = Request("txtFromGlDt")
	sToGlDt = Request("txtToGlDt")
	strMode = Request("txtMode")														'�� : ���� ���¸� ���� 
	strBizAreaCd= Request("txtBizAreaCd")
	strBizAreaCd1= Request("txtBizAreaCd1")

	Select Case strMode

	Case CStr(UID_M0001)																'��: ���� ��ȸ/Prev/Next ��û�� ���� 
	    
	    Call QUERYIWOL()
		Call FixUNISQLDATA()
		Call QueryData()

	Sub FixUNISQLData()
	    Dim intI
	    Redim UNISqlId(0)																'��: SQL ID ������ ���� ����Ȯ�� 

	    UNISqlId(0) = "a5114MA01"	  '�����ⳳ����ȸ	
	    Redim UNIValue(0,1)

		UNIValue(0,0) = lgSelectList 
		UNIValue(0,1) = Trim(strWhere0)

		UNILock = DISCONNREAD :	UNIFlag = "1"											'��: set ADO read mode
	End Sub

Sub QueryData()
    Dim iStr
		    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
		    
    iStr = Split(lgstrRetMsg,gColSep)
		  
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
        Response.End
    End If

    If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close:		Set rs0 = Nothing
		Set lgADF = Nothing
		Response.End																'��: �����Ͻ� ���� ó���� ������ 
	Else
		Call  MakeSpreadSheetData()
    End If				
		    
    Call ReleaseObj()
End Sub

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

    If Len(Trim(lgPageNo)) Then														'�� : Chnage Nextkey str into int value
       If Isnumeric(lgPageNo) Then
          iCnt = CInt(lgPageNo)
       End If   
    End If   

    For iRCnt = 1 to iCnt  *  C_SHEETMAXROWS_D										'�� : Discard previous data
        rs0.MoveNext
    Next

    iRCnt = -1
    
    Do While Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
        iStr = ""

		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
			iStr = iStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next

        If iRCnt < C_SHEETMAXROWS_D Then
            lgstrData      = lgstrData      & iStr & Chr(11) & Chr(12)
        Else
            iCnt = iCnt + 1
            lgPageNo = CStr(iCnt)
            Exit Do
        End If

        rs0.MoveNext
	Loop

    If iRCnt < C_SHEETMAXROWS_D Then                                            '��: Check if next data exists
        lgPageNo = ""															'��: ���� ����Ÿ ����.
    End If
%>    
    <Script Language=vbscript>
	With parent
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowData "<%=lgstrData%>"									'��: Display data 
		.lgPageNo =  "<%=ConvSPChars(lgPageNo)%>"								'��: set next data tag
		'.hBIZ_AREA_CD.value =  "<%=ConvSPChars(strBizAreaCd)%>"                '��: set next data tag
		'.hFromGlDt.value =  "<%=ConvSPChars(sFromGlDt)%>"                      '��: set next data tag
		'.hToGlDt.value =  "<%=ConvSPChars(sToGlDt)%>"							'��: set next data tag
		.DbQueryOk
	End with
	</Script>	
<%    
End Sub

Sub ReleaseObj()
	Set rs0 = Nothing
	Set rs1 = Nothing
	Set rs2 = Nothing
	Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
End Sub	
	
'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub QUERYIWOL()
    Dim Fiscyyyy,Fiscmm,Fiscdd,VarBaseFiscDt,DateFryyyy,DateFrmm,DateFrdd
    Dim FiscEyyyy,FiscEmm,FiscEdd
    Dim startdate
    Dim strCond1

	gFiscStart = GetGlobalInf("gFiscStart")
	gFiscEnd = GetGlobalInf("gFiscEnd")
	
    Call ExtractDateFrom(gFiscStart,gAPDateFormat,gAPDateSeperator,Fiscyyyy,Fiscmm,Fiscdd)
    Call ExtractDateFrom(gFiscEnd ,gAPDateFormat,gAPDateSeperator,FiscEyyyy,FiscEmm,FiscEdd)

    Call ExtractDateFrom(sFromGlDt,gDateFormat,gComDateType,DateFryyyy,DateFrmm,DateFrdd)
    
    strBizAreaCd  = Request("txtBizAreaCd")
    strBizAreaCd1 = Request("txtBizAreaCd1")

	strWhere0 = ""
	strWhere0 = strWhere0 & " GL_DT >=  " & FilterVar(UniConvDateAToB(sFromGlDt,gDateFormat,gServerDateFormat), "''", "S") & " "
	strWhere0 = strWhere0 & " and GL_DT <=  " & FilterVar(UniConvDateAToB(Request("txtToGlDt"),gDateFormat,gServerDateFormat), "''", "S") & " "
    
    If strBizAreaCd <> "" Then
		Call fnGetBizAreaCd
		If strBizAreaCd1 = "" Then
			strWhere0 = strWhere0 & " and BIZ_AREA_CD >=  " & FilterVar(strBizAreaCd , "''", "S") & ""
		Else
			Call fnGetBizAreaCd1
			strWhere0 = strWhere0 & " and BIZ_AREA_CD >=  " & FilterVar(strBizAreaCd , "''", "S") & " and biz_area_cd <=  " & FilterVar(strBizAreaCd1 , "''", "S") & ""		
		End If
	Else		
		If strBizAreaCd1 <> "" Then
			Call fnGetBizAreaCd1
			strWhere0 = strWhere0 & " and biz_area_cd <=  " & FilterVar(strBizAreaCd1 , "''", "S") & ""		
		End If
	End If		
	
	Fiscyyyy =  DateFryyyy
	If Fiscmm > DateFrmm  then                         ' ��ȸ���ۿ��� ��� ���ۿ��������� ��� ���� ���ڰ�� 
	   Fiscyyyy	= cstr(cint(DateFryyyy) - 1)
	End If		   
	startdate = DateFryyyy & DateFrmm & DateFrdd	
	
	VarBaseFiscDt = Fiscyyyy & Fiscmm & "00"
	
	strcond = "  ( fisc_yr+fisc_mnth +fisc_dt = (select isnull(max(fisc_yr+fisc_mnth)," & FilterVar("190001", "''", "S") & " ) + " & FilterVar("00", "''", "S") & "  from a_gl_sum where fisc_yr + fisc_mnth + fisc_dt <= substring(convert(char(8)," & FilterVar(startdate , "''", "S") & " ,112),1,6) and fisc_dt = " & FilterVar("00", "''", "S") & " ) "
	strCond =  strCond & "  or  ( fisc_yr+fisc_mnth +fisc_dt >= (select isnull(max(fisc_yr+fisc_mnth)," & FilterVar("190001", "''", "S") & " ) + " & FilterVar("01", "''", "S") & "  from a_gl_sum where fisc_yr + fisc_mnth + fisc_dt <= substring(convert(char(8)," & FilterVar(startdate , "''", "S") & " ,112),1,6) and fisc_dt = " & FilterVar("00", "''", "S") & " ) "
	strcond =  strCond & " and   fisc_yr+fisc_mnth +fisc_dt <  " & FilterVar(startdate, "''", "S") & ""
	strcond =  strCond & "and fisc_dt not in (" & FilterVar("00", "''", "S") & " ," & FilterVar("99", "''", "S") & " ))) "  
	
	
	'strcond = "  ( fisc_yr+fisc_mnth +fisc_dt =  " & FilterVar(VarBaseFiscDt , "''", "S") & ""
	'strCond =  strCond & "  or  ( fisc_yr+fisc_mnth +fisc_dt <  " & FilterVar(startdate, "''", "S") & ""
	'strcond =  strCond & " and   fisc_yr+fisc_mnth +fisc_dt >  " & FilterVar(VarBaseFiscDt, "''", "S") & ""
	'strcond =  strCond & "and fisc_dt not in (" & FilterVar("00", "''", "S") & " ," & FilterVar("99", "''", "S") & " ))) "  

    If strBizAreaCd <> "" Then
		If strBizAreaCd1 = "" Then
			strcond = strcond & " and a.BIZ_AREA_CD >=  " & FilterVar(strBizAreaCd , "''", "S") & ""
		Else
			strcond = strcond & " and a.BIZ_AREA_CD >=  " & FilterVar(strBizAreaCd , "''", "S") & " and a.biz_area_cd <=  " & FilterVar(strBizAreaCd1 , "''", "S") & ""		
		End If
	Else		
		If strBizAreaCd1 <> "" Then
			strcond = strcond & " and a.biz_area_cd <=  " & FilterVar(strBizAreaCd1 , "''", "S") & ""		
		End If
	End If
		
	'=====================================================
	'Gl No, Ref No�� Gl Header �б� 
	'a5114ma02 : ���ݾ� (statements) - rs2
	'a5114ma03 : �����̿��ݾ� (Condition �� ASP���� ó��) - rs1
	'=====================================================
    Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(1,0)

    UNISqlId(0)   = "a5114MA03"  '�̿��ݾ�                            
    UNIValue(0,0) = strCond

    UNISqlId(1)   = "a5114MA02"   '�߻��ݾ� 

	strCond1 = ""
	strCond1 = strCond1 & " and gl_dt >=  " & FilterVar(UniConvDateAToB(sFromGlDt,gDateFormat,gServerDateFormat), "''", "S") & ""
	strCond1 = strCond1 & " and gl_dt <=  " & FilterVar(UniConvDateAToB(Request("txtToGlDt"),gDateFormat,gServerDateFormat), "''", "S") & ""

	If strBizAreaCd <> "" Then
		If strBizAreaCd1 = "" Then
			strcond1 = strcond1 & " and BIZ_AREA_CD >=  " & FilterVar(strBizAreaCd , "''", "S") & ""
		Else
			strcond1 = strcond1 & " and BIZ_AREA_CD >=  " & FilterVar(strBizAreaCd , "''", "S") & " and biz_area_cd <=  " & FilterVar(strBizAreaCd1 , "''", "S") & ""		
		End If
	Else		
		If strBizAreaCd1 <> "" Then
			strcond1 = strcond1 & " and biz_area_cd <=  " & FilterVar(strBizAreaCd1 , "''", "S") & ""
		End If
	End If

    UNIValue(1,0) = strCond1

'	If UCase(FilterVar(Trim(Request("txtBizAreaCd")),"","S")) = "" Then
'		UNIValue(1,2) = "|"	
'	Else
'		UNIValue(1,2) = "" & FilterVar(Request("txtBizAreaCd"),"","S")
'	End If				
    
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode

    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs1, rs2 )

    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    

    If rs1.EOF And rs1.BOF Then
		rs1.Close
		Set rs1 = Nothing
		Set lgADF = Nothing

		strMsgCd = "970000"
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
		Response.End																					'��: �����Ͻ� ���� ó���� ������ 
    Else
%>
		<Script Language=vbScript>
			With parent
				.frm1.txtTDrAmt.text   = "<%=UNINumClientFormat(rs1("drsum"), ggAmtOfMoney.DecPoint, 0)%>"
				.frm1.txtTCrAmt.text   = "<%=UNINumClientFormat(rs1("crsum"), ggAmtOfMoney.DecPoint, 0)%>"
				.frm1.txtTSumAmt.text  = "<%=UNINumClientFormat(rs1("balamt"), ggAmtOfMoney.DecPoint, 0)%>"
			End With
		</Script>
<%
    End If

    If  rs2.EOF And rs2.BOF Then
		rs2.Close
		Set rs2 = Nothing
		Set lgADF = Nothing

		strMsgCd = "970000"
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
		Response.End																						'��: �����Ͻ� ���� ó���� ������ 
    Else
%>
		<Script Language=vbScript>
			With parent
				.frm1.txtNDrAmt.text   = "<%=UNINumClientFormat(rs2("drsum"), ggAmtOfMoney.DecPoint, 0)%>"
				.frm1.txtNCrAmt.text	= "<%=UNINumClientFormat(rs2("crsum"), ggAmtOfMoney.DecPoint, 0)%>"
				.frm1.txtNSumAmt.text  = "<%=UNINumClientFormat(rs2("balamt"), ggAmtOfMoney.DecPoint, 0)%>"
				.frm1.txtSDrAmt.text   = "<%=UNINumClientFormat(CDbl(rs1("drsum")) + CDbl(rs2("drsum")), ggAmtOfMoney.DecPoint, 0)  %>"
				.frm1.txtSCrAmt.text	= "<%=UNINumClientFormat(CDbl(rs1("crsum")) + CDbl(rs2("crsum")), ggAmtOfMoney.DecPoint, 0)%>"
				.frm1.txtSSumAmt.text  = "<%=UNINumClientFormat(CDbl(rs1("balamt")) + CDbl(rs2("balamt")), ggAmtOfMoney.DecPoint, 0)%>"
           End With 
		</Script>
<%
    End If

	rs2.Close
	Set rs2 = Nothing 
    rs2.Close
	Set rs2 = Nothing
End Sub
    
End Select

'--------------------------------------------
'������ 
'--------------------------------------------
Sub fnGetBizAreaCd()
    Dim iStr

    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(0,1)

	On Error Resume Next
	Err.Clear
 
    UNISqlId(0) = "ABIZNM"	'������ڵ� 
	
	UNIValue(0,0) = FilterVar(strBizAreaCd,"","S")

    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    
    If rs0.EOF And rs0.BOF Then
        rs0.Close
        Set rs0 = Nothing

		'strMsgCd = "970000"												'Not Found	
		'strMsg1 = Request("txtBizAreaCd_Alt")
		'Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
		'Response.End													'��: �����Ͻ� ���� ó���� ������ 
    Else    
    	strBizAreaCd = Trim(rs0(0))
		strBizAreaNm = Trim(rs0(1))
%>
		<Script Language=vbscript>
			With parent.frm1
			.txtBizAreaCd.value = "<%=ConvSPChars(strBizAreaCd)%>"
			.txtBizAreaNm.value = "<%=ConvSPChars(strBizAreaNm)%>"
			End With
		</Script>
<%			
    End If    

	rs0.Close
	Set rs0 = Nothing
End Sub 

'--------------------------------------------
'������1
'--------------------------------------------
Sub fnGetBizAreaCd1()
    Dim iStr

    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(0,1)

	On Error Resume Next
	Err.Clear
 
    UNISqlId(0) = "ABIZNM"	'������ڵ� 
	
	UNIValue(0,0) =  FilterVar(strBizAreaCd1,"","S")	

    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    
    If rs0.EOF And rs0.BOF Then
        rs0.Close
        Set rs0 = Nothing

		'strMsgCd = "970000"												'Not Found	
		'strMsg1 = Request("txtBizAreaCd1_Alt")
		'Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
		'Response.End													'��: �����Ͻ� ���� ó���� ������ 
    Else    
    	strBizAreaCd1 = Trim(rs0(0))
		strBizAreaNm1 = Trim(rs0(1))
%>
		<Script Language=vbscript>
			With parent.frm1
			.txtBizAreaCd1.value = "<%=ConvSPChars(strBizAreaCd1)%>"
			.txtBizAreaNm1.value = "<%=ConvSPChars(strBizAreaNm1)%>"
			End With
		</Script>
<%			
    End If    

	rs0.Close
	Set rs0 = Nothing
End Sub 

%>

<%
	Response.End 
%>
