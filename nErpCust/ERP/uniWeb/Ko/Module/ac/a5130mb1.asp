<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Deposit
'*  3. Program ID           : a5130mb1
'*  4. Program Name         : ������(��)��ǥ ��ȸ 
'*  5. Program Desc         : 
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2002.02.01
'*  8. Modified date(Last)  : 2002.02.01
'*  9. Modifier (First)     : AHJ
'* 10. Modifier (Last)      : AHJ
'* 11. Comment              :
'=======================================================================================================
Response.Expires = -1                                                       '�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True                                                     '�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.

%>

<!-- #Include file="../../inc/IncServer.asp"  -->

<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
	                                                                   '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
'On Error Resume Next

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim  UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT

'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strDt, strBizAreaCd, strClassType, strToAcctCd							'��������� ���� 
Dim strFrDt, strToDt														'��������� ���� 

Dim TDrLocAmt,TCrLocAmt,NDrLocAmt,NCrLocAmt
Dim TTotSumAmt,NTotSumAmt,STotSumAmt,SDrAmt, SCrAmt
Dim strMsgCd, strMsg1, strMsg2 												'��������� ���� 
Dim strWhere0, strWhere1													'��������� ���� 
Dim strBizAreaNm															'��������� ���� 
Dim strCompYr,strCompMnth,strCompDt											'��������� ���� 
Dim strDtYr, strDtMnth, strDtDay											'��������� ���� 
Dim strCompFiscStartDt														'��������� ���� 
Dim strToGlDts																'��������� ���� 

'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd 

    lgStrPrevKey   = Request("lgStrPrevKey")                               '�� : Next key flag
    lgMaxCount     = CInt(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
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

    If Len(Trim(lgStrPrevKey)) Then                                        '�� : Chnage Nextkey str into int value
       If Isnumeric(lgStrPrevKey) Then
          iCnt = CInt(lgStrPrevKey)
       End If   
    End If   

    For iRCnt = 1 to iCnt  *  lgMaxCount                                   '�� : Discard previous data
        rs0.MoveNext
    Next

    iRCnt = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
        iStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            Select Case  lgSelectListDT(ColCnt)
               Case "DD"   '��¥ 
                           iStr = iStr & Chr(11) & UNIDateClientFormat(rs0(ColCnt))
               Case "F2"  ' �ݾ� 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggAmtOfMoney.DecPoint, 0)
               Case "F3"  '����7
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggQty.DecPoint       , 0)
               Case "F4"  '�ܰ� 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggUnitCost.DecPoint  , 0)
               Case "F5"   'ȯ�� 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggExchRate.DecPoint  , 0)
               Case Else
                    iStr = iStr & Chr(11) & ConvSPChars(Trim("" & rs0(ColCnt)))
            End Select
		Next
 
        If  iRCnt < lgMaxCount Then
            lgstrData      = lgstrData      & iStr & Chr(11) & Chr(12)
        Else
            iCnt = iCnt + 1
            lgStrPrevKey = CStr(iCnt)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iRCnt < lgMaxCount Then                                            '��: Check if next data exists
        lgStrPrevKey = ""                                                  '��: ���� ����Ÿ ����.
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

    UNISqlId(0) = "A5130MA101"	'�ϰ�ǥ��ȸ 
	UNISqlId(1) = "A5130MA102"	'�����ܾ� 
	UNISqlId(2) = "ACLASSNM"	'�����ڵ�    
	UNISqlId(3) = "ABIZNM"		'�����ڵ�    
	
	
	Redim UNIValue(4,4)
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    UNIValue(0,1) = UCase(Trim(strWhere0))
		
	UNIValue(1,0) = Trim(strToGlDts)	
	UNIValue(1,1) = Trim(strFrDt)	
	UNIValue(1,2) = Trim(strToDt)	
	UNIValue(1,3) = Trim(strWhere1)		
	
	UNIValue(2,0) = Trim(strClassType) 	
	
	UNIValue(3,0) = FilterVar(strBizAreaCd, "''", "S") 
	
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
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)    
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If        	   

	If Not(rs1.EOF And rs1.BOF) Then
		If IsNull(rs1(0)) = False Then NTotSumAmt    = rs1(0)
	End If
	
	rs1.Close
	Set rs1 = Nothing	
	
	If rs0.EOF And rs0.BOF Then
		If strMsgCd = "" Then strMsgCd = "900014"
'		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
'		rs0.Close
'		Set rs0 = Nothing
'		Response.End													'��: �����Ͻ� ���� ó���� ������ 
    Else    
        Call  MakeSpreadSheetData()
    End If    
        	
    %>
    
    <Script Language=vbscript>
		With parent
    '	.frm1.txtYAmt.value		= "<%=UNINumClientFormat(TTotSumAmt, ggAmtOfMoney.DecPoint, 0)%>"
		.frm1.txtTAmt.value		= "<%=UNINumClientFormat(NTotSumAmt, ggAmtOfMoney.DecPoint, 0)%>"
	
		End With
	</script>
	<%
	
	If (rs2.EOF And rs2.BOF) Then
		If strMsgCd = "" And strClassType <> "" Then 
			strMsgCd = "970000"												'Not Found	
			strMsg1 = Request("txtClassType_Alt")
		End If
	Else
%>
		<Script Language=vbscript>
			With parent.frm1
			.txtClassType.value = "<%=ConvSPChars(strClassType)%>"
			.txtClassTypeNm.value = "<%=ConvSPChars(Trim(rs2(1)))%>"
			End With
		</Script>
<%			
	End If
	
	rs2.Close
	Set rs2 = Nothing
	
	If  rs3.EOF And rs3.BOF Then        
		If strMsgCd = "" And strBizAreaCd <> "" Then 
			strMsgCd = "970000"												'Not Found	
			strMsg1 = Request("txtBizAreaCd_Atl")
		end if
    Else		
%>
		<Script Language=vbscript>
			With parent.frm1
			.txtBizAreaCd.value = "<%=ConvSPChars(strBizAreaCd)%>"
			.txtBizAreaNm.value = "<%=ConvSPChars(Trim(rs3(1)))%>"
			End With
		</Script>
<%			
    End If  
    
    rs3.Close
    Set rs3 = Nothing
    
	rs0.Close
	Set rs0 = Nothing 	
	Set lgADF = Nothing  	
	                                                  '��: ActiveX Data Factory Object Nothing
	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
		'Response.End 
	
	End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()  

    '--------------- ������ coding part(�������,Start)----------------------------------------------------
	strDt = Request("txtDateYr")
	strBizAreaCd = UCase(Request("txtBizAreaCd"))
	strClassType = UCase(FilterVar(Request("txtClassType"), "''", "S"))	
		
	Call fnGetCompStDt
	Call ExtractDateFrom(strCompFiscStartDt,gAPDateFormat,gApDateSeperator,strCompYr,strCompMnth,strCompDt)
	
	strFrDt = strDt +  strCompMnth + strCompDt
	strToDt = UniDateAdd("D", +364, strFrDt, gServerDateFormat)
	
	strWhere0 = ""
	strWhere0 = strWhere0 & " C.CLASS_TYPE = " & FilterVar(strClassType, "''", "S") 
	strWhere0 = strWhere0 & " AND D.gl_dt between  " & FilterVar(strFrDt, "''", "S") & " and  " & FilterVar(strToDt, "''", "S") & " "
		
	If strBizAreaCd <> "" Then		
		strWhere0 = strWhere0 & " AND D.biz_area_cd = " & FilterVar(strBizAreaCd, "''", "S")
	End If
	
	strToGlDts = strDt +  strCompMnth

	strWhere1 = ""
	strWhere1 = strWhere1 & " A.Acct_cd = (select acct_cd from a_acct where acct_type = " & FilterVar("A0", "''", "S") & "  ) " 	
	

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

%>

<Script Language=vbscript>
	With parent
		.ggoSpread.Source = .frm1.vspdData 
		.ggoSpread.SSShowData "<%=lgstrData%>"                            '��: Display data 
		.lgStrPrevKey =  "<%=ConvSPChars(lgStrPrevKey)%>"                       '��: set next data tag
		.DbQueryOk
	End with
	
</Script>	

<%
	Response.End 
%>