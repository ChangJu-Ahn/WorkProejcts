<%'======================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Budget
'*  3. Program ID           : f2108rb1
'*  4. Program Name         : ���������˾� 
'*  5. Program Desc         : Popup of Budget
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2001.03.31
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'=======================================================================================================


%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next
Err.Clear                                                                        '��: Clear Error status

Dim lgADF                                                                  '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg                                                            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '�� : DBAgent Parameter ���� 
Dim lgstrData                                                              '�� : data for spreadsheet data
Dim lgStrPrevKey                                                           '�� : ���� �� 
Dim lgMaxCount                                                             '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
Dim lgTailList                                                             '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT
'--------------- ������ coding part(��������,Start)--------------------------------------------------------
Dim strCond
Dim strBdgYymmFr, strBdgYymmTo, strDeptCd, strBdgCdFr, strBdgCdTo, strChgFg
Dim strColYymm, strDateType
Dim strMsgCd, strMsg1, strMsg2

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL

'--------------- ������ coding part(��������,End)----------------------------------------------------------

	Call LoadBasisGlobalInf()
	Call loadInfTB19029B("Q", "A","NOCOOKIE","RB")
	Call HideStatusWnd 


	lgStrPrevKey		= Request("lgStrPrevKey")                               '�� : Next key flag
	lgMaxCount			= CInt(Request("lgMaxCount"))                           '�� : �ѹ��� �����ü� �ִ� ����Ÿ �Ǽ� 
	lgSelectList		= Request("lgSelectList")                               '�� : select ����� 
	lgSelectListDT		= Split(Request("lgSelectListDT"), gColSep)             '�� : �� �ʵ��� ����Ÿ Ÿ�� 
	lgTailList			= Request("lgTailList")                                 '�� : Orderby value

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
	On Error Resume Next
	Err.Clear                                                                        '��: Clear Error status

    iCnt = 0
    lgstrData = ""

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
	On Error Resume Next
	Err.Clear                                                                        '��: Clear Error status
    Redim UNISqlId(3)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    Redim UNIValue(3,2)

    UNISqlId(0) = "F2108RA101"
    UNISqlId(1) = "F2108RA102"	'�μ��ڵ� 
    UNISqlId(2) = "F2108RA103"	'���ۿ����ڵ� 
    UNISqlId(3) = "F2108RA103"	'���Ό���ڵ� 

    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '��: Select list
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    UNIValue(0,1) = strCond
    UNIValue(1,0) = FilterVar(strDeptCd, "''", "S")
    UNIValue(1,1) = GetGlobalInf("gChangeOrgId")
    UNIValue(2,0) = FilterVar(strBdgCdFr, "''", "S")
    UNIValue(3,0) = FilterVar(strBdgCdTo, "''", "S")
    
    '--------------- ������ coding part(�������,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
	On Error Resume Next
	Err.Clear                                                                        '��: Clear Error status
    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
	
	If rs1.EOF And rs1.BOF Then
		If strMsgCd = "" And strDeptCd <> "" Then 
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtDeptCd_Alt")
		End If
	Else
%>
		<Script Language=vbScript>
			With parent.frm1
				.txtDeptCd.value = "<%=ConvSPChars(Trim(rs1(0)))%>"
				.txtDeptNm.value = "<%=ConvSPChars(Trim(rs1(1)))%>"
			End With
		</Script>
<%	
	End If
	
	rs1.Close
	Set rs1 = Nothing

	If rs2.EOF And rs2.BOF Then
		If strMsgCd = "" And strBdgCdFr <> "" Then 
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBdgCdFr_Alt")
		End If
	Else
%>
		<Script Language=vbScript>
			With parent.frm1
				.txtBdgCdFr.value = "<%=ConvSPChars(Trim(rs2(0)))%>"
				.txtBdgNmFr.value = "<%=ConvSPChars(Trim(rs2(1)))%>"
			End With
		</Script>
<%	
	End If
	
	rs2.Close
	Set rs2 = Nothing

	If rs3.EOF And rs3.BOF Then
		If strMsgCd = "" And strBdgCdTo <> "" Then 
			strMsgCd = "970000"		'Not Found
			strMsg1 = Request("txtBdgCdTo_Alt")
		End If
	Else
%>
		<Script Language=vbScript>
			With parent.frm1
				.txtBdgCdTo.value = "<%=ConvSPChars(Trim(rs3(0)))%>"
				.txtBdgNmTo.value = "<%=ConvSPChars(Trim(rs3(1)))%>"
			End With
		</Script>
<%	
	End If
	
	rs3.Close
	Set rs3 = Nothing
	
    If  rs0.EOF And rs0.BOF Then
		If strMsgCd = "" Then strMsgCd = "900014"
'		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
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
		Call DisplayMsgBox(strMsgCd, vbInformation, strMsg1, strMsg2, I_MKSCRIPT)
		Response.End 
	End If
	
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
	On Error Resume Next
	Err.Clear                                                                        '��: Clear Error status
	Dim strInternalCd
	
    '--------------- ������ coding part(�������,Start)----------------------------------------------------
    strBdgYymmFr = Request("txtBdgYymmFr")
    strBdgYymmTo = Request("txtBdgYymmTo")
    strDeptCd    = UCase(Request("txtDeptCd"))
    strBdgCdFr   = UCase(Request("txtBdgCdFr"))
    strBdgCdTo   = UCase(Request("txtBdgCdTo"))
	strColYymm   = Request("txtColYymm")
	strDateType  = Request("txtDateType")
	strChgFg     = Request("txtChgFg")
    
	strCond = ""
	
	If strBdgYymmFr <> "" Then strCond = strCond & " and A.bdg_yyyymm >=  " & FilterVar(strBdgYymmFr , "''", "S") & " "	
	If strBdgYymmTo <> "" Then strCond = strCond & " and A.bdg_yyyymm <=  " & FilterVar(strBdgYymmTo , "''", "S") & " "
	If strDeptCd <> "" Then
		strInternalCd = fnGetInternalCd
		strCond = strCond & " and A.internal_cd =  " & FilterVar(strInternalCd , "''", "S") & " "
	End If
	If strBdgCdFr <> "" Then strCond = strCond & " and A.bdg_cd >=  " & FilterVar(strBdgCdFr , "''", "S") & " "
	If strBdgCdTo <> "" Then strCond = strCond & " and A.bdg_cd <=  " & FilterVar(strBdgCdTo , "''", "S") & " "
	
	Select Case strChgFg
		Case "A"
			strCond = strCond & " and B.add_fg = " & FilterVar("1", "''", "S") & "  "
		Case "D"
			strCond = strCond & " and B.divert_fg = " & FilterVar("1", "''", "S") & "  "
		Case "T"
			strCond = strCond & " and B.trans_fg = " & FilterVar("1", "''", "S") & "  "
		Case Else
	End Select

	' ���Ѱ��� �߰� 
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

	strCond = strCond & lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL




    '--------------- ������ coding part(�������,End)------------------------------------------------------

End Sub

'���κμ��ڵ� select
Function fnGetInternalCd()
	On Error Resume Next
	Err.Clear                                                                        '��: Clear Error status
    Redim UNISqlId(0)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
    Redim UNIValue(0,1)

    UNISqlId(0) = "F2108RA102"

    UNIValue(0,0) = FilterVar(strDeptCd, "''", "S")
    UNIValue(0,1) = GetGlobalInf("gChangeOrgId")
    
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
	
    Dim iStr

    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF Then
        fnGetInternalCd = ""
        rs0.Close
        Set rs0 = Nothing
    Else    
        fnGetInternalCd = rs0(2)
    End If
End Function

'----------------------------------------------------------------------------------------------------------
' Trim string and set string to space if string length is zero
'----------------------------------------------------------------------------------------------------------
'2004.8.19 comment ó�� 
'Function FilterVar(Byval str,Byval strALT)
'     Dim strL
'     strL = UCase(Trim(str))
'     If Len(strL) Then
'        FilterVar = " " & FilterVar(strL , "''", "S") & ""
'     Else
'        FilterVar = strALT   
'     End If
'End Function

%>

<Script Language=vbscript>
    With parent
         .ggoSpread.Source    = .frm1.vspdData 
         .ggoSpread.SSShowData "<%=lgstrData%>"                            '��: Display data 
         .lgStrPrevKey        = "<%=ConvSPChars(lgStrPrevKey)%>"                       '��: set next data tag
         
         With .frm1
			.hBdgYymmFr.value = strBdgYymmFr
			.hBdgYymmTo.value = strBdgYymmTo
			.hDeptCd.value    = strDeptCd
			.hBdgCdFr.value   = strBdgCdFr
			.hBdgCdTo.value   = strBdgCdTo
         End With
         
         Call .DbQueryOk
	End with
</Script>	

<%
	Response.End 
%>
