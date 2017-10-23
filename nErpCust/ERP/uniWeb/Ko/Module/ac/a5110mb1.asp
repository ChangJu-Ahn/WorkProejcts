<% Option Explicit %>
<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Deposit
'*  3. Program ID           : A5115mb1
'*  4. Program Name         : �Ѱ���������ȸ 
'*  5. Program Desc         : 
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2001.12.26
'*  8. Modified date(Last)  : 2001.12.26
'*  9. Modifier (First)     : Chang, Sung Hee
'* 10. Modifier (Last)      : Chang, Sung Hee
'* 11. Comment              :
'=======================================================================================================
Response.Expires = -1                                                       '�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True                                                     '�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.

%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrNumber.inc"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<!-- #Include file="../../inc/incSvrDBAgent.inc"  -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc"  -->

<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
	                                                                   '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
On Error Resume Next


Call LoadBasisGlobalInf() 
Call LoadInfTB19029B("Q","A", "COOKIE", "QB")

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim  UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4                              '�� : DBAgent Parameter ����Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
Dim lgstrData 
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strFromGlDt, strToGLDt, strBizAreaCd, strClassType, strToAcctCd
Dim strFrGlDts
Dim strToGlDts

Dim TDrLocAmt,TCrLocAmt,NDrLocAmt,NCrLocAmt
Dim TTotSumAmt,NTotSumAmt,STotSumAmt,SDrAmt, SCrAmt
Dim strMsgCd, strMsg1, strMsg2 
Dim strWhere0, strWhere1
Dim strBizAreaNm
Dim strCompYr,strCompMnth,strCompDt, strGlDtYr, strGlDtMnth, strGlDtDt
Dim strCompFiscStartDt
Dim strcond
Dim sFromGlDt
Dim strSPID
Dim strOUT
Dim strZeroFg

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ�		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)				
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL					


Dim Fiscyyyy,Fiscmm,Fiscdd,VarBaseFiscDt,DateFryyyy,DateFrmm,DateFrdd
DIM startdate

sFromGlDt = Request("txtFromGlDt")
'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
    Call HideStatusWnd 

    lgStrPrevKey   = Request("lgStrPrevKey")                               '�� : Next key flag
    lgMaxCount     = CInt(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
    lgSelectList   = Request("lgSelectList")                               '�� : select ����� 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
    lgTailList     = Request("lgTailList")                                 '�� : Orderby value

    Call TrimData()
    Call FixUNISQLData()
'    Call QueryData()

    IF strOUT = "2" THEN 
       Call DeleteData()
    ELSE
       Call QueryData()
    END IF
    
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
            iStr = iStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
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

    Redim UNISqlId(4)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    UNISqlId(0) = "A5110MA101"	'�ϰ�ǥ��ȸ 
	UNISqlId(3) = "ACLASSNM"	'�����ڵ�    
	UNISqlId(4) = "ABIZNM"	'�����ڵ�    
	
	
	Redim UNIValue(5,3)
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    UNIValue(0,1) = UCase(Trim(strWhere0))
   
	UNIValue(3,0) = FilterVar(strClassType, "''", "S") 	
	
	UNIValue(4,0) = " " & FilterVar(strBizAreaCd, "''", "S") & " "

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    'UNIValue(0,2) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Delete Data
'----------------------------------------------------------------------------------------------------------
Sub DeleteData()

    Dim iStr
    Redim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
	Redim UNIValue(0,0)

    '--------------- ������ coding part(�������,End)------------------------------------------------------
	UNISqlId(0) = "A5110MA105"	'DELETE    
    UNIValue(0,0) = " " & FilterVar(Request("txtSPID"), "''", "S") & ""
    '--------------- ������ coding part(�������,End)------------------------------------------------------

    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
   
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)    
    
    iStr = Split(lgstrRetMsg,gColSep)
    
	rs0.Close
	Set rs0 = Nothing 	
	Set lgADF = Nothing  	
	
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")

    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs3, rs4)    
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If        	   

	If rs0.EOF And rs0.BOF Then
		If strMsgCd = "" Then strMsgCd = "900014"
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
		Response.End													'��: �����Ͻ� ���� ó���� ������ 
    Else    
        Call  MakeSpreadSheetData()
    End If    

	If (rs3.EOF And rs3.BOF) Then
		If strMsgCd = "" And strClassType <> "" Then 
			strMsgCd = "970000"												'Not Found	
			strMsg1 = Request("txtClassType_Alt")
		End If
	Else
%>
		<Script Language=vbscript>
			With parent.frm1
			.txtClassType.value = "<%=ConvSPChars(strClassType)%>"
			.txtClassTypeNm.value = "<%=ConvSPChars(Trim(rs3(1)))%>"
			End With
		</Script>
<%			
	End If
	
	rs3.Close
	Set rs3 = Nothing
	
	If  rs4.EOF And rs4.BOF Then        
		If strMsgCd = "" And strBizAreaCd <> "" Then 
			strMsgCd = "970000"												'Not Found	
			strMsg1 = Request("txtBizAreaCd_Alt")
		end if
    Else		
%>
		<Script Language=vbscript>
			With parent.frm1
			.txtBizAreaCd.value = "<%=ConvSPChars(strBizAreaCd)%>"
			.txtBizAreaNm.value = "<%=ConvSPChars(Trim(rs4(1)))%>"
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
	strFromGlDt	 = uniconvdate(Request("txtDateFr"))
	strToGLDt	 = uniconvdate(Request("txtDateTo"))
	strBizAreaCd = UCase(Request("txtBizAreaCd"))
	strClassType = UCase(Request("txtClassType"))	
	strSPID		 = Request("txtSPID")
	strOUT		 = Request("txtOUT")
	
	' ���Ѱ��� �߰� 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))		
	lgInternalCd		= Trim(Request("lgInternalCd"))	
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))	
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))
	
	'--------------- ������ coding part(�������,Start)----------------------------------------------------

	IF strBizAreaCd <> "" OR lgAuthBizAreaCd <> "" Then 
		Call fnGetBizAreaCd
	End IF	

	If strClassType <> "" Then
	    Call fnGetClassType
	End If    	
	
	strWhere0 = ""
	strWhere0 = strWhere0 & " SPID =  " & FilterVar(strSPID , "''", "S") & " "


'''NKH


'	strWhere0 = strWhere0 & " AND D.gl_dt between '" & strFromGlDt & "' and '" & strToGLDt & "' "
	
'	If strBizAreaCd <> "" Then		
'		strWhere0 = strWhere0 & " AND D.biz_area_cd = '" & strBizAreaCd & "' " 		
'	End If
	
'	strWhere1 = ""
'	strWhere1 = strWhere1 & " A.Acct_cd in (select acct_cd from a_acct where acct_type = 'A0' ) " 	

	

End Sub
'--------------------------------------------
'Company(start_Dt)/ �̿��ݾ� 
'--------------------------------------------
'''NKH
'Sub fnGetCompStDt()
'End Sub 

'--------------------------------------------
'������ 
'--------------------------------------------
Sub fnGetBizAreaCd()
    Dim iStr
	Dim strBizAreaCd1
    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(0,1)

    ON error resume next
 
    UNISqlId(0) = "A_GetBiz"	'������ڵ� 
    
    If strBizAreaCd = "" Then
	 	UNIValue(0,0) = FilterVar("", "'%'", "S")
	Else
		UNIValue(0,0) = FilterVar(strBizAreaCd, "''", "S")
	End If

	'Response.write UNIValue(0,0)
	'Response.End 

	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then			
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
    	strBizAreaCd1 = Trim(rs0(0))
		strBizAreaNm  = Trim(rs0(1))
%>
		<Script Language=vbscript>
			With parent.frm1
			.txtBizAreaCd.value = "<%=ConvSPChars(strBizAreaCd1)%>"
			.txtBizAreaNm.value = "<%=ConvSPChars(strBizAreaNm)%>"
			End With
		</Script>
<%			
    End If    
    
	rs0.Close
	Set rs0 = Nothing	
End Sub 

'--------------------------------------------
'�ϰ�ǥ���� 
'--------------------------------------------
Sub fnGetClassType()
    Dim iStr
	Dim strClassType1
	Dim strClassTypeNm
    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(0,1)

    ON error resume next
 
    UNISqlId(0) = "ACLASSNM"	'�ϰ�ǥ���� 
	
	UNIValue(0,0) =  FilterVar(strClassType, "''", "S") 	
'	Response.write UNIValue(0,0)
'	Response.End 
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
		strMsg1 = Request("txtClassType_Alt")
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
		Response.End
    Else    
    	strClassType1 = Trim(rs0(0))
		strClassTypeNm  = Trim(rs0(1))
%>
		<Script Language=vbscript>
			With parent.frm1
			.txtClassType.value = "<%=ConvSPChars(strClassType1)%>"
			.txtClassTypeNm.value = "<%=ConvSPChars(strClassTypeNm)%>"
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
		.lgStrPrevKey =  "<%=ConvSPChars(lgStrPrevKey)%>"                       '��: set next data tag
		'msgbox "<%=Fiscyyyy%>" & "/" & "<%=fiscmm%>" & "/" & "<%=fiscdd%>"		 
		.DbQuery2Ok
	End with
	
</Script>	

<%
	Response.End 
%>
