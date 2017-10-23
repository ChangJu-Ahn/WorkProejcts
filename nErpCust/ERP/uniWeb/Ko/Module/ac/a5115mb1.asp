<%Option Explicit%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->


<% 

On Error Resume Next

Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("I","*", "NOCOOKIE", "MB")

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim  UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5

Dim lgPageNo                                                           '�� : ���� �� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgstrData 
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strFromGlDt, strToGLDt, strBizAreaCd, strBizAreaCd1, strFrAcctCd, strToAcctCd
Dim strGlDt

Dim TDrLocAmt,TCrLocAmt,NDrLocAmt,NCrLocAmt
Dim TTotSumAmt,NTotSumAmt,STotSumAmt,SDrAmt, SCrAmt
Dim strMsgCd, strMsg1, strMsg2 
Dim strWhere0, strWhere1
Dim strBizAreaNm
Dim strBizAreaNm1
Dim strCompYr,strCompMnth,strCompDt, strGlDtYr, strGlDtMnth, strGlDtDt
Dim strCompFiscStartDt

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ�		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)				
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL					

'--------------- ������ coding part(��������,End)----------------------------------------------------------
Const C_SHEETMAXROWS_D  = 100
    Call HideStatusWnd 

    lgPageNo   = Request("lgPageNo")                               '�� : Next key flag
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value

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

    If Len(Trim(lgPageNo)) Then                                        '�� : Chnage Nextkey str into int value
       If Isnumeric(lgPageNo) Then
          iCnt = CInt(lgPageNo)
       End If   
    End If   

    For iRCnt = 1 to iCnt  *  C_SHEETMAXROWS_D                                   '�� : Discard previous data
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
            lgPageNo = CStr(iCnt)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iRCnt < C_SHEETMAXROWS_D Then                                            '��: Check if next data exists
        lgPageNo = ""                                                  '��: ���� ����Ÿ ����.
    End If
  	
'	rs0.Close
'	Set rs0 = Nothing 
'	Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(5)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    UNISqlId(0) = "A5115MA101"	'�Ѱ�����ȸ 
    UNISqlId(1) = "A5115MA102"	'�̿��ݾ� 
	UNISqlId(2) = "A5115MA103"	'�߻��ݾ�		
	UNISqlId(3) = "AACCTNM"	'�����ڵ� 
    UNISqlId(4) = "AACCTNM"	'�����ڵ�    	
	
	Redim UNIValue(5,2)
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    UNIValue(0,1) = UCase(Trim(strWhere0))

	UNIValue(1,0) = FilterVar(Trim(strGlDt),"''","S")
	UNIValue(1,1) = Trim(strWhere1)	
	
	UNIValue(2,0) = Trim(strWhere0)
	
	UNIValue(3,0) = Filtervar(Trim(strFrAcctCd),"''","S")

	UNIValue(4,0) = FilterVar(Trim(strToAcctCd),"''","S")	
	
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    'UNIValue(0,2) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    'lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5)
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)    

    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If        	

	TDrLocAmt = 0
	TCrLocAmt = 0
	NDrLocAmt = 0
	NCrLocAmt = 0 
	
	If Not(rs1.EOF And rs1.BOF) Then
		If IsNull(rs1(0)) = False Then TDrLocAmt    = rs1(0)
		If IsNull(rs1(1)) = False Then TCrLocAmt    = rs1(1)
		If IsNull(rs1(2)) = False Then TTotSumAmt    = rs1(2)
	End If
	
	rs1.Close
	Set rs1 = Nothing
	
	If Not(rs2.EOF And rs2.BOF) Then
		If IsNull(rs2(0)) = False Then NDrLocAmt    = rs2(0)
		If IsNull(rs2(1)) = False Then NCrLocAmt    = rs2(1)
		If IsNull(rs2(2)) = False Then NTotSumAmt    = rs2(2)
	End If
	
	rs2.Close
	Set rs2 = Nothing	
	
	If rs0.EOF And rs0.BOF Then
		If strMsgCd = "" Then strMsgCd = "900014"
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
'		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
'		rs0.Close
'		Set rs0 = Nothing
'		Response.End													'��: �����Ͻ� ���� ó���� ������ 
    Else    
        Call  MakeSpreadSheetData()
    End If

    STotSumAmt = cdbl(TTotSumAmt) + cdbl(NTotSumAmt) 
    SDrAmt  = cdbl(TDrLocAmt) + cdbl(NDrLocAmt)
    SCrAmt  = cdbl(TCrLocAmt) + cdbl(NCrLocAmt)    
        	
    %>
    
    <Script Language=vbscript>
		With parent
    	.frm1.txtTDrAmt.text		= "<%=UNINumClientFormat(TDrLocAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtTCrAmt.text		= "<%=UNINumClientFormat(TCrLocAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtTSumAmt.text		= "<%=UNINumClientFormat(TTotSumAmt, ggAmtOfMoney.DecPoint, 0)%>"
				
		.frm1.txtNDrAmt.text		= "<%=UNINumClientFormat(NDrLocAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtNCrAmt.text		= "<%=UNINumClientFormat(NCrLocAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtNSumAmt.text		= "<%=UNINumClientFormat(NTotSumAmt, ggAmtOfMoney.DecPoint, 0)%>"		
		
		.frm1.txtSSumAmt.text		= "<%=UNINumClientFormat(STotSumAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtSDrAmt.text		= "<%=UNINumClientFormat(SDrAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtSCrAmt.text		= "<%=UNINumClientFormat(SCrAmt, ggAmtOfMoney.DecPoint, 0)%>"

		End With
	</script>
	<%
	
	If (rs3.EOF And rs3.BOF) Then
		If strMsgCd = "" And strFrAcctCd <> "" Then 
			strMsgCd = "970000"												'Not Found	
			strMsg1 = Request("txtFrAcctCd_Alt")
			Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
			Response.End
		End If
	Else
%>
		<Script Language=vbscript>
			With parent.frm1
			.txtFrAcctCd.value = "<%=ConvSPChars(strFrAcctCd)%>"
			.txtFrAcctNm.value = "<%=ConvSPChars(Trim(rs3(0)))%>"
			End With
		</Script>
<%			
	End If

	rs3.Close
	Set rs3 = Nothing

	If (rs4.EOF And rs4.BOF) Then
		If strMsgCd = "" And strToAcctCd <> "" Then 
			strMsgCd = "970000"												'Not Found	
			strMsg1 = Request("txtToAcctCd_Alt")
			Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
			Response.End
		End If
	Else
%>
		<Script Language=vbscript>
			With parent.frm1
			.txtToAcctCd.value = "<%=ConvSPChars(strToAcctCd)%>"
			.txtToAcctNm.value = "<%=ConvSPChars(Trim(rs4(0)))%>"
			End With
		</Script>
<%			
	End If
	
	rs4.Close
	Set rs4 = Nothing	
	
	rs0.Close
	Set rs0 = Nothing 	
	Set lgADF = Nothing  	
	                                                  '��: ActiveX Data Factory Object Nothing
	
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()  
	
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
	strFromGlDt   = uniconvdate(Request("txtFromGlDt"))
	strToGLDt     = uniconvdate(Request("txtToGlDt"))
	strBizAreaCd  = Request("txtBizAreaCd")
	strBizAreaCd1 = Request("txtBizAreaCd1")
	strFrAcctCd = UCase(Request("txtFrAcctCd"))
	strToAcctCd = UCase(Request("txtToAcctCd"))
	
	' ���Ѱ��� �߰� 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))
	
	Call	fnGetCompStDt	
	
	If strBizAreaCd <> "" Then 
		Call fnGetBizAreaCd
	Else
		If lgAuthBizAreaCd <> "" Then 
			strBizAreaCd = lgAuthBizAreaCd
		End IF		
	End IF		
	
	IF strBizAreaCd1 <> "" Then 
		Call fnGetBizAreaCd1
	Else
		If lgAuthBizAreaCd <> "" Then 
			strBizAreaCd1 = lgAuthBizAreaCd
		End IF		
	End IF
	
	Call ExtractDateFrom(strCompFiscStartDt,gAPDateFormat,gApDateSeperator,strCompYr,strCompMnth,strCompDt)
	Call ExtractDateFrom(strFromGlDt,gAPDateFormat,gApDateSeperator,strGlDtYr,strGlDtMnth,strGlDtDt)
	
	strGlDt = 	strGlDtYr +  strGlDtMnth + strGlDtDt
	strWhere0 = ""
	strWhere0 = strWhere0 & " X.Acct_cd between  " & FilterVar(strFrAcctCd, "''", "S") & " and  " & FilterVar(strToAcctCd, "''", "S") & " "
	strWhere0 = strWhere0 & " and convert(datetime,X.gl_dt) between  " & FilterVar(strFromGlDt, "''", "S") & " and  " & FilterVar(strToGLDt, "''", "S") & " "
	
	If strBizAreaCd <> "" Then		
		strWhere0 = strWhere0 & " and X.biz_area_cd >=  " & FilterVar(strBizAreaCd , "''", "S") & " "
	else
		strWhere0 = strWhere0 & " and X.biz_area_cd >= " & FilterVar("0", "''", "S") & "  " 		
	End If
	
	If strBizAreaCd1 <> "" Then		
		strWhere0 = strWhere0 & " and X.biz_area_cd <=  " & FilterVar(strBizAreaCd1 , "''", "S") & " "
	else
		strWhere0 = strWhere0 & " and X.biz_area_cd <= " & FilterVar("ZZZZZZZZZZ", "''", "S") & "  " 		
	End If
	
	strWhere1 = ""
	strWhere1 = strWhere1 & " a.Acct_cd between  " & FilterVar(strFrAcctCd, "''", "S") & " and  " & FilterVar(strToAcctCd, "''", "S") & " "
	
	If strBizAreaCd <> "" Then		
		strWhere1 = strWhere1 & " and a.biz_area_cd >=  " & FilterVar(strBizAreaCd , "''", "S") & " "
	else
		strWhere1 = strWhere1 & " and a.biz_area_cd >= " & FilterVar("0", "''", "S") & "  "
	End If
	
	If strBizAreaCd1 <> "" Then		
		strWhere1 = strWhere1 & " and a.biz_area_cd <=  " & FilterVar(strBizAreaCd1 , "''", "S") & " "
	else
		strWhere1 = strWhere1 & " and a.biz_area_cd <= " & FilterVar("ZZZZZZZZZZ", "''", "S") & "  "
	End If
	
End Sub
'--------------------------------------------
'Company(start_Dt)/ �̿��ݾ� 
'--------------------------------------------
Sub fnGetCompStDt()
    Dim iStr

    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(0,1)

    ON error resume next
 
    UNISqlId(0) = "A5124MA108"

    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    iStr = Split(lgstrRetMsg,gColSep)


    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    
    If  rs0.EOF And rs0.BOF Then
        rs0.Close
        Set rs0 = Nothing

		strMsgCd = "970000"
		strMsg1 = ""
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)

        strCompFiscStartDt = "1900-01-01"

    Else    
        strCompFiscStartDt   = Trim(rs0(0))

    End If
End Sub 
'--------------------------------------------
'������ 
'--------------------------------------------
Sub fnGetBizAreaCd()
    Dim iStr
	Dim strBizAreaCd2
    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(0,1)

    ON error resume next
 
    UNISqlId(0) = "ABIZNM"	'������ڵ� 
	
	UNIValue(0,0) =  " " & FilterVar(strBizAreaCd, "''", "S") & ""
	'Response.write UNIValue(0,0)
	'Response.End 

	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" AND strBizAreaCd < lgAuthBizAreaCd Then			
		UNIValue(0,0)  = UNIValue(0,0) & " AND BIZ_AREA_CD LIKE " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			
    
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    
    If  rs0.EOF And rs0.BOF Then
        rs0.Close
        Set rs0 = Nothing

		strMsgCd = "970000"												'Not Found	
		strMsg1 = Request("txtBizAreaCd_Alt")
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
		Response.End
    Else    
    	strBizAreaCd2 = Trim(rs0(0))
		strBizAreaNm  = Trim(rs0(1))
%>
		<Script Language=vbscript>
			With parent.frm1
			.txtBizAreaCd.value = "<%=ConvSPChars(strBizAreaCd2)%>"
			.txtBizAreaNm.value = "<%=ConvSPChars(strBizAreaNm)%>"
			End With
		</Script>
<%			
    End If    
    
	rs0.Close
	Set rs0 = Nothing	
End Sub 

'--------------------------------------------
'������ 
'--------------------------------------------
Sub fnGetBizAreaCd1()
    Dim iStr
	Dim strBizAreaCd3
    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(0,1)

    ON error resume next
 
    UNISqlId(0) = "ABIZNM"	'������ڵ� 
	
	UNIValue(0,0) =  " " & FilterVar(strBizAreaCd1, "''", "S") & " "
	
	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" AND strBizAreaCd1 > lgAuthBizAreaCd Then			
		UNIValue(0,0)  = UNIValue(0,0) & " AND BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			
	
	'Response.write UNIValue(0,0)
	'Response.End 
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    iStr = Split(lgstrRetMsg,gColSep)

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    
    If  rs0.EOF And rs0.BOF Then
        rs0.Close
        Set rs0 = Nothing

		strMsgCd = "970000"												'Not Found	
		strMsg1 = Request("txtBizAreaCd_Alt1")
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
		Response.End
    Else    
    	strBizAreaCd3 = Trim(rs0(0))
		strBizAreaNm1 = Trim(rs0(1))
%>
		<Script Language=vbscript>
			With parent.frm1
			.txtBizAreaCd1.value = "<%=ConvSPChars(strBizAreaCd3)%>"
			.txtBizAreaNm1.value = "<%=ConvSPChars(strBizAreaNm1)%>"
			End With
		</Script>
<%			
    End If    
    
	rs0.Close
	Set rs0 = Nothing	
End Sub 

%>


<Script Language=vbscript>
	With parent
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowData "<%=lgstrData%>"                            '��: Display data 
		.lgPageNo =  "<%=ConvSPChars(lgPageNo)%>"                       '��: set next data tag
		.DbQueryOk
	End with
	
</Script>	

<%
	Response.End 
%>
