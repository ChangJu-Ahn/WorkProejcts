<%
'**********************************************************************************************
'*  1. Module��          : ȸ�� 
'*  2. Function��        : �������� 
'*  3. Program ID        : f5106ma
'*  4. Program �̸�      : �����Ϻ� ������ȸ 
'*  5. Program ����      : �����Ϻ��� ���������� ��ȸ 
'*  6. Comproxy ����Ʈ   : FN0018_List_Note_By_Due_Dt
'*  7. ���� �ۼ������   : 2000/10/14
'*  8. ���� ���������   : 2002/03/05
'*  9. ���� �ۼ���       : Hersheys
'* 10. ���� �ۼ���       : Heo Chung ku
'* 11. ��ü comment      :
'**********************************************************************************************


%>

<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServerAdoDb.asp" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next								'��: 
err.Clear

Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q","A","NOCOOKIE","QB")

Dim lgADF											                '�� : ActiveX Data Factory ���� �������� 
Dim lgstrRetMsg											            '�� : Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3		'�� : DBAgent Parameter ���� 
Dim lgstrData                                                       '�� : data for spreadsheet data
Dim lgStrPrevKey											        '�� : ���� ��(key flag)
Dim lgTailList													    '�� : Orderby���� ���� field ����Ʈ 
Dim lgSelectList
Dim lgSelectListDT

Dim lgPageNo

'--------------------����� ���� ����start------------------------

Dim strMode															'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Dim StrNextKey
Dim LngMaxRow														' ���� �׸����� �ִ�Row
Dim NOSumAmt
Dim LngRow
Dim strMsgCd

Dim strFromDt , strToDt
Dim strNoteFg
Dim strBankCd, strBpCd
Dim strcboSts
Dim strStsCd
Dim strBizAreaCd
Dim strBizAreaNm
Dim strBizAreaCd1
Dim strBizAreaNm1

Dim strWhere0
Dim strWhere1

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ� 
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 

Dim lgBizAreaAuthSQL, lgInternalCdAuthSQL, lgSubInternalCdAuthSQL, lgAuthUsrIDAuthSQL


'--------------------����� ���� ���� end------------------------

	Call HideStatusWnd 

	lgPageNo			= Request("lgPageNo")										'�� : Next key flag
	lgSelectList		= Request("lgSelectList")								'�� : select ����� 
	lgSelectListDT		= Split(Request("lgSelectListDT"), gColSep)				'�� : �� �ʵ��� ����Ÿ Ÿ�� 
	lgTailList			= Request("lgTailList")									'�� : Order by value
	
	' ���Ѱ��� �߰� 
	lgAuthBizAreaCd		= Trim(Request("lgAuthBizAreaCd"))
	lgInternalCd		= Trim(Request("lgInternalCd"))
	lgSubInternalCd		= Trim(Request("lgSubInternalCd"))
	lgAuthUsrID			= Trim(Request("lgAuthUsrID"))

	Call TrimData()
	Call FixUNISQLData()
	Call QueryData()

'---------------------------------------------------------------------------------------------
' Query Data
'---------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

Dim  ColCnt
Dim  iCnt
Dim  iRCnt
Dim  iStr

    iCnt = 0
    lgstrData = ""

	Const C_SHEETMAXROWS_D = 100
    
    If Len(Trim(lgPageNo)) Then												'�� : Chnage Nextkey str into int value
       If Isnumeric(lgPageNo) Then
          iCnt = CInt(lgPageNo)
       End If   
    End If   

    For iRCnt = 1 to iCnt  *  C_SHEETMAXROWS_D									'�� : Discard previous data
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


    If  iRCnt < C_SHEETMAXROWS_D Then												'��: Check if next data exists
        lgPageNo = ""														'��: ���� ����Ÿ ����.
    End If
  	
End Sub


'----------------------------------------------------------------------------------------------------------
' Set DBAgent arguments
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Redim UNISqlId(3)														'��: SQL ID ������ ���� ����Ȯ�� 
    '--------------- ������ coding part(�������,Start)----------------------------------------------------

    UNISqlId(0) = "F5106MA101"	'����(f_note)��ȸ 
	UNISqlId(1) = "F5106MA102"	'�����ݾ� �հ� 
	UNISqlId(2) = "A_GETBIZ"
    UNISqlId(3) = "A_GETBIZ"
	
	Redim UNIValue(3,2)
    UNIValue(0,0) = Trim(lgSelectList)										'��: Select list
    UNIValue(0,1) = UCase(Trim(strWhere0))									'where0������ list

	UNIValue(1,0) = UCase(Trim(strWhere1))									'where1������ list
	
	UNIValue(2,0)  = FilterVar(strBizAreaCd, "''", "S")
	UNIValue(3,0)  = FilterVar(strBizAreaCd1, "''", "S")
	
    '--------------- ������ coding part(�������,End)------------------------------------------------------
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
	

	If Not(rs1.EOF And rs1.BOF) Then
		If IsNull(rs1(0)) = False Then NOSumAmt = rs1(0)
	End If
%>
<Script Language=vbscript>
		With parent.frm1
			.txtNoteAmtSum.Text = "<%=UNINumClientFormat(NOSumAmt, ggAmtOfMoney.DecPoint, 0)%>"	'�����ݾ��հ� 
		End With 
</script>
<%
	rs1.Close
	Set rs1 = Nothing	

	If strMsgCd <> "" Then
		Call DisplayMsgBox(strMsgCd, vbOKOnly, strMsg1, "", I_MKSCRIPT)
		rs0.Close
        Set rs0 = Nothing
        Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If
	   
    If  rs0.EOF And rs0.BOF Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        Response.End														'��: �����Ͻ� ���� ó���� ������ 
    Else    
        Call  MakeSpreadSheetData()
    End If
    
	rs0.close
	Set rs0 = nothing
	Set lgADF = Nothing

End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()  

    '--------------- ������ coding part(�������,Start)----------------------------------------------------
	strFromDt	= uniconvdate(Request("txtFrDt"))
	strToDt		= uniconvdate(Request("txtToDt"))
	strNoteFg	= request("cboNoteFg")	'�������� 
	strBankCd	= request("txtBankCd")	'�����ڵ� 
	strBpCd		= request("txtBpCd")	'�ŷ�ó�ڵ� 
	strStsCd	= request("txtStsCd")	'�������� 
	strBizAreaCd  = Trim(UCase(Request("txtBizAreaCd")))    '�����From
	strBizAreaCd1 = Trim(UCase(Request("txtBizAreaCd1")))   '�����To
	

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

	
	strWhere0 = ""
	strWhere0 = strWhere0 & " a.bank_cd = c.bank_cd "
	strWhere0 = strWhere0 & " and a.bp_cd = b.bp_cd"
	strWhere0 = strWhere0 & " and a.due_dt between  " & FilterVar(strFromDt, "''", "S") & " and  " & FilterVar(strToDt, "''", "S") & " "
	strWhere0 = strWhere0 & " and a.note_fg =  " & FilterVar(strNoteFg , "''", "S") & ""
	strWhere0 = strWhere0 & " and d.major_cd = " & FilterVar("F1008", "''", "S") & " "
	strWhere0 = strWhere0 & " and d.minor_cd = a.note_sts "
	strWhere0 = strWhere0 & " and a.dept_cd = e.dept_cd "
	strWhere0 = strWhere0 & " and a.org_change_id = e.org_change_id "

	If strBankCd <> "" Then
	strWhere0 = strWhere0 & " and a.bank_cd =  " & FilterVar(strBankCd , "''", "S") & " "
	End If 
	
	If strBpCd <> "" Then
	strWhere0 = strWhere0 & " and a.bp_cd =  " & FilterVar(strBpCd , "''", "S") & " "
	End If

	If strStsCd <> "" Then
		strStsCdf = UCase(mid(strStsCd,1,1))
		if strStsCdf = "S" then 
		strWhere0 = strWhere0 & " and d.minor_cd LIKE " & FilterVar("S%", "''", "S") & " "							'���� : ����, �Ϻΰ���, �輭���� 
		else 
		strWhere0 = strWhere0 & " and d.minor_cd in (" & FilterVar("DC", "''", "S") & " , " & FilterVar("ED", "''", "S") & "  ," & FilterVar("OC", "''", "S") & " ) "			'�̰��� : �߻�, ����, �輭 
		End if 
	End If
	
	if strBizAreaCd <> "" then
		strWhere0 = strWhere0 & " AND a.BIZ_AREA_CD >= "	& FilterVar(strBizAreaCd , "''", "S") 
	else
		strWhere0 = strWhere0 & " AND a.BIZ_AREA_CD >= " & FilterVar(" ", "''", "S") & " "
	end if
	
	if strBizAreaCd1 <> "" then
		strWhere0 = strWhere0 & " AND a.BIZ_AREA_CD <= "	& FilterVar(strBizAreaCd1 , "''", "S") 
	else
		strWhere0 = strWhere0 & " AND a.BIZ_AREA_CD <= " & FilterVar("ZZZZZZZZZZ", "''", "S") & " "
	end if

	' ���Ѱ��� �߰� 
	strWhere0	= strWhere0	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL

	strWhere0	= strWhere0	& " order by a.note_no"


	strWhere1 = ""
	strWhere1 = strWhere1 & " a.bank_cd = c.bank_cd "
	strWhere1 = strWhere1 & " and a.bp_cd = b.bp_cd"
	strWhere1 = strWhere1 & " and a.due_dt between  " & FilterVar(strFromDt, "''", "S") & " and  " & FilterVar(strToDt, "''", "S") & " "
	strWhere1 = strWhere1 & " and a.note_fg =  " & FilterVar(strNoteFg , "''", "S") & ""

	If strBankCd <> "" Then
	strWhere1 = strWhere1 & " and a.bank_cd =  " & FilterVar(strBankCd , "''", "S") & " "
	End If 
	
	If strBpCd <> "" Then
	strWhere1 = strWhere1 & " and a.bp_cd =  " & FilterVar(strBpCd , "''", "S") & " "
	End If
	
	If strStsCd <> "" Then
		strStsCdf = UCase(mid(strStsCd,1,1))
		if strStsCdf = "S" then 
		strWhere1 = strWhere1 & " and a.note_sts LIKE " & FilterVar("S%", "''", "S") & "  "						'���� : ����, �Ϻΰ���, �輭���� 
		else 
		strWhere1 = strWhere1 & " and a.note_sts in (" & FilterVar("DC", "''", "S") & " , " & FilterVar("ED", "''", "S") & "  ," & FilterVar("OC", "''", "S") & " ) "			'�̰��� : ����, �߻�, �輭 
		End if 
	End If
	
	if strBizAreaCd <> "" then
		strWhere1 = strWhere1 & " AND a.BIZ_AREA_CD >= "	& FilterVar(strBizAreaCd , "''", "S") 
	else
		strWhere1 = strWhere1 & " AND a.BIZ_AREA_CD >= " & FilterVar(" ", "''", "S") & " "
	end if
	
	if strBizAreaCd1 <> "" then
		strWhere1 = strWhere1 & " AND a.BIZ_AREA_CD <= "	& FilterVar(strBizAreaCd1 , "''", "S") 
	else
		strWhere1 = strWhere1 & " AND a.BIZ_AREA_CD <= " & FilterVar("ZZZZZZZZZZ", "''", "S") & " "
	end if


	' ���Ѱ��� �߰� 
	strWhere1	= strWhere1	& lgBizAreaAuthSQL & lgInternalCdAuthSQL & lgSubInternalCdAuthSQL & lgAuthUsrIDAuthSQL

End Sub

%>

<Script Language=vbscript>
	With parent
		.ggoSpread.Source				= .frm1.vspdData 
		.ggoSpread.SSShowData			"<%=lgstrData%>"							'��: Display data 
		.lgPageNo						=  "<%=lgPageNo%>"
		.frm1.hFrDt.value				= "<%=strFromDt%>"
		.frm1.hToDt.value				= "<%=strToDt%>"  		
		.frm1.hcboNoteFg.value			= "<%=strNoteFg%>"
		.frm1.htxtBankCd.value			= "<%=strBankCd%>"
		.frm1.htxtBpCd.value			= "<%=strBpCd%>"
		.frm1.htxtStsCd.value			= "<%=strStsCd%>"
		.frm1.htxtBizAreaCd.value		="<%=strBizAreaCd%>"
		.frm1.htxtBizAreaCd1.value		="<%=strBizAreaCd1%>"		

		.DbQueryOk
	End with
	
</Script>	

