<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Fixed Asset
'*  3. Program ID           : a7115ma1
'*  4. Program Name         : �����󰢰�������ȸ 
'*  5. Program Desc         : List Depreciation by Account
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000.11.20
'*  8. Modified date(Last)  : 2001.03.05
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : Song, Mun Gil
'* 11. Comment              :
'=======================================================================================================

Response.Expires = -1														'�� : ASP�� ĳ������ �ʵ��� �Ѵ�.
Response.Buffer = True														'�� : ASP�� ���ۿ� ������� �ʰ� �ٷ� Client�� ��������.
%>

<!-- #Include file="../../inc/IncServer.asp"  -->
<%																			'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next

Dim lgADF																	'�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg																'�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3				'�� : DBAgent Parameter ���� 
Dim lgstrData																'�� : data for spreadsheet data
Dim lgStrPrevKey															'�� : ���� �� 

Dim lgTailList																'�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strDeprYYYYMM, strDurYrsFg
Dim strWhere, strWhere1		'Where ���� 
Dim strMsgCd, strMsg1, strMsg2
Dim strBizAreaCd															'�� : ���ۻ���� 
Dim strBizAreaNm	
Dim strBizAreaCd1															'�� : �������� 
Dim strBizAreaNm1

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ�		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)				
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL					


'--------------- ������ coding part(��������,End)----------------------------------------------------------
  
	Call HideStatusWnd 

	lgStrPrevKey		= Request("lgStrPrevKey")								'�� : Next key flag
	lgSelectList		= Request("lgSelectList")								'�� : select ����� 
	lgSelectListDT		= Split(Request("lgSelectListDT"), gColSep)				'�� : �� �ʵ��� ����Ÿ Ÿ�� 
	lgTailList			= Request("lgTailList")									'�� : Orderby value

	' ���Ѱ��� �߰� 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))

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

	Const C_SHEETMAXROWS_D  = 100											'��: Server���� �ѹ��� fetch�� �ִ� ����Ÿ �Ǽ� 
	
    iCnt = 0
    lgstrData = ""

    If Len(Trim(lgStrPrevKey)) Then											'�� : Chnage Nextkey str into int value
		If Isnumeric(lgStrPrevKey) Then
			iCnt = CInt(lgStrPrevKey)
		End If   
    End If   

    For iRCnt = 1 To iCnt * C_SHEETMAXROWS_D								'�� : Discard previous data
        rs0.MoveNext
    Next

    iRCnt = -1
    
    Do While Not (rs0.EOF Or rs0.BOF)
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

    If  iRCnt < C_SHEETMAXROWS_D Then										'��: Check if next data exists
        lgStrPrevKey = ""													'��: ���� ����Ÿ ����.
    End If
  	
'	rs0.Close
'	Set rs0 = Nothing 
'	Set lgADF = Nothing														'��: ActiveX Data Factory Object Nothing
End Sub
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(3)														'��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    UNISqlId(0) = "A7115_KO441"
    UNISqlId(1) = "A7115S_KO441"	'�հ� 
    UNISqlId(2) = "A_GETBIZ"
    UNISqlId(3) = "A_GETBIZ"



    Redim UNIValue(3,4)

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList											'��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
	UNIValue(0,1) = FilterVar(strDeprYYYYMM, "''", "S") 
	UNIValue(0,2) = FilterVar(strDurYrsFg, "''", "S") 
	UNIValue(0,3) = strWhere

	
	UNIValue(1,0) = FilterVar(strDeprYYYYMM, "''", "S") 
	UNIValue(1,1) = FilterVar(strDurYrsFg, "''", "S") 
	UNIValue(1,2) = strWhere1


	UNIValue(2,0)  = FilterVar(strBizAreaCd, "''", "S")
	UNIValue(3,0)  = FilterVar(strBizAreaCd1, "''", "S")
		
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"									'��: set ADO read mode
 
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
    
	If (rs2.EOF And rs2.BOF) Then
		If strMsgCd = "" and strBizAreaCd <> ""  Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBizAreaCd_ALT")
		End If
    Else
%>
	<Script Language=vbScript>
	With parent
		.frm1.txtBizAreaCd.value = "<%=Trim(rs2(0))%>"
		.frm1.txtBizAreaNm.value = "<%=Trim(rs2(1))%>"					
	End With
	</Script>
<%
    End If
	
	rs2.Close
	Set rs2 = Nothing   
    
    
If (rs3.EOF And rs3.BOF) Then
		If strMsgCd = "" and strBizAreaCd1 <> ""  Then
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBizAreaCd1_ALT")
		End If
    Else
%>
	<Script Language=vbScript>
	With parent
		.frm1.txtBizAreaCd1.value = "<%=Trim(rs3(0))%>"
		.frm1.txtBizAreaNm1.value = "<%=Trim(rs3(1))%>"					
	End With
	</Script>
<%
    End If
    rs3.Close
	Set rs3 = Nothing 
    
	
	If Not (rs1.EOF And rs1.BOF) Then
%>
		<Script Language=vbScript>
			With Parent
				.frm1.txtAmtSum1.value = "<%=UNINumClientFormat(rs1(0), ggAmtOfMoney.DecPoint, 0)%>"
				.frm1.txtAmtSum2.value = "<%=UNINumClientFormat(rs1(1), ggAmtOfMoney.DecPoint, 0)%>"
				.frm1.txtAmtSum3.value = "<%=UNINumClientFormat(rs1(2), ggAmtOfMoney.DecPoint, 0)%>"
				.frm1.txtAmtSum4.value = "<%=UNINumClientFormat(rs1(3), ggAmtOfMoney.DecPoint, 0)%>"
				.frm1.txtAmtSum5.value = "<%=UNINumClientFormat(rs1(4), ggAmtOfMoney.DecPoint, 0)%>"
				.frm1.txtAmtSum6.value = "<%=UNINumClientFormat(rs1(5), ggAmtOfMoney.DecPoint, 0)%>"
				.frm1.txtAmtSum7.value = "<%=UNINumClientFormat(rs1(6), ggAmtOfMoney.DecPoint, 0)%>"
			End With
		</Script>
<%
	End If

	rs1.Close
	Set rs1 = Nothing
	
    If rs0.EOF And rs0.BOF Then
		If strMsgCd = "" Then strMsgCd = "900014"
'		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
'		rs0.Close
'		Set rs0 = Nothing
'		Set lgADF = Nothing
'		Response.End													'��: �����Ͻ� ���� ó���� ������ 
    Else    
        Call  MakeSpreadSheetData()
    End If

	rs0.Close
	Set rs0 = Nothing 
	Set lgADF = Nothing                                                    '��: ActiveX Data Factory Object Nothing
	
	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, strMsg2, I_MKSCRIPT)
		Response.End 
	End If

End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()

    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    strDeprYYYYMM	= UCase(Trim(Request("txtDeprYYYYMM")))
    strDurYrsFg		= UCase(Trim(Request("txtDurYrsFg")))
    strBizAreaCd	= Trim(UCase(Request("txtBizAreaCd")))					'�����From
	strBizAreaCd1	= Trim(UCase(Request("txtBizAreaCd1")))					'�����To
	
	if strBizAreaCd <> "" then
		strWhere = strWhere & " AND ISNULL(D.TO_BIZ_AREA_CD,D.FROM_BIZ_AREA_CD) >= "	& FilterVar(strBizAreaCd , "''", "S") 
	else
		strWhere = strWhere & " AND ISNULL(D.TO_BIZ_AREA_CD,D.FROM_BIZ_AREA_CD) >= "	& FilterVar("0", "''", "S") & " "
	end if
	
	if strBizAreaCd1 <> "" then
		strWhere = strWhere & " AND ISNULL(D.TO_BIZ_AREA_CD,D.FROM_BIZ_AREA_CD) <= "	& FilterVar(strBizAreaCd1 , "''", "S") 
	else
		strWhere = strWhere & " AND ISNULL(D.TO_BIZ_AREA_CD,D.FROM_BIZ_AREA_CD) <= "	& FilterVar("ZZZZZZZZZZ", "''", "S") & " "
	End if


	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then			
		strWhere	= strWhere &	" AND ISNULL(D.TO_BIZ_AREA_CD,D.FROM_BIZ_AREA_CD) = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		strWhere	= strWhere &	" AND d.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		strWhere	= strWhere &	" AND d.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		strWhere	= strWhere &	" AND d.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	



	
	if strBizAreaCd <> "" then
		strWhere1 = strWhere1 & " AND ISNULL(C.TO_BIZ_AREA_CD,C.FROM_BIZ_AREA_CD) >= " & FilterVar(strBizAreaCd , "''", "S") 
	else
		strWhere1 = strWhere1 & " AND ISNULL(C.TO_BIZ_AREA_CD,C.FROM_BIZ_AREA_CD) >= " & FilterVar("0", "''", "S") & " "
	end if
	
	if strBizAreaCd1 <> "" then
		strWhere1 = strWhere1 & " AND ISNULL(C.TO_BIZ_AREA_CD,C.FROM_BIZ_AREA_CD) <= " & FilterVar(strBizAreaCd1 , "''", "S") 
	else
		strWhere1 = strWhere1 & " AND ISNULL(C.TO_BIZ_AREA_CD,C.FROM_BIZ_AREA_CD) <= " & FilterVar("ZZZZZZZZZZ", "''", "S") & " "
	End if	


	' ���Ѱ��� �߰� 
	If lgAuthBizAreaCd <> "" Then			
		strWhere1	= strWhere1 &	" AND ISNULL(C.TO_BIZ_AREA_CD,C.FROM_BIZ_AREA_CD) = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		strWhere1	= strWhere1 &	" AND (case when ISNULL(C.TO_INTERNAL_CD,'') <> '' then C.TO_INTERNAL_CD else C.FROM_INTERNAL_CD end) = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		strWhere1	= strWhere1 &	" AND (case when ISNULL(C.TO_INTERNAL_CD,'') <> '' then C.TO_INTERNAL_CD else C.FROM_INTERNAL_CD end) LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		strWhere1	= strWhere1 &	" AND C.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	


    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
End Sub

%>

<Script Language=vbscript>
    With parent
         .ggoSpread.Source    = .frm1.vspdData 
         .ggoSpread.SSShowData "<%=lgstrData%>"                            '��: Display data 
         .lgStrPrevKey        = "<%=lgStrPrevKey%>"                       '��: set next data tag
         .DbQueryOk
	End with
	
</Script>	

<%
Response.End
%>
