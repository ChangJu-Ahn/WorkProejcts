<%@ LANGUAGE="VBSCRIPT" %>
<% Option Explicit%>
<!--
'**********************************************************************************************
'*  1. Module Name		  : Procurement
'*  2. Function Name		:
'*  3. Program ID		   : M2111QB4
'*  4. Program Name		 : ���ſ�û����ȸ 
'*  5. Program Desc		 : ���ſ�û����ȸ 
'*  6. Component List	   :
'*  7. Modified date(First) : 2000/12/12
'*  8. Modified date(Last)  : 2003/05/21
'*  9. Modifier (First)	 : ByunJiHyun
'* 10. Modifier (Last)	  : KANG SU HWAN
'* 11. Comment			  :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*							this mark(��) Means that "may  change"
'*							this mark(��) Means that "must change"
'* 13. History			  :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%														  '�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

	On Error Resume Next

	Dim lgADF												   '�� : ActiveX Data Factory ���� �������� 
	Dim lgstrRetMsg											 '�� : Record Set Return Message �������� 
	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'�� : DBAgent Parameter ���� 
	Dim rs1, rs2, rs3, rs4, rs5, rs6							'�� : DBAgent Parameter ���� 
	Dim lgStrData											   '�� : Spread sheet�� ������ ����Ÿ�� ���� ���� 
	Dim lgStrPrevKey											'�� : ���� �� 
	Dim lgTailList
	Dim lgSelectList
	Dim lgSelectListDT

	'--------------- ������ coding part(��������,Start)----------------------------------------------------
	Dim ICount  												'   Count for column index
	Dim strPlantCd											  '   ���� 
	Dim strPlantCdFrom
	Dim strItemCd												'   ǰ�� 
	Dim strItemCdFrom
	Dim strPrFrDt											   '   ���ſ�û�� 
	Dim strPrToDt
	Dim strPdFrDt											   '   �ʿ䳳���� 
	Dim strPdToDt
	Dim strPrStsCd												'   ��û������� 
	Dim strPrStsCdFrom
	Dim StrRqDeptCd												'	��û�μ� 
	Dim StrRqDeptCdFrom
	Dim StrPrTypeCd												'	���ſ�û���� 
	Dim StrPrTypeCdFrom
	Dim lgPageNo
	Dim lgDataExist
	Dim strTrackNo

	Dim arrRsVal(11)											'* : ȭ�鿡 ��ȸ�ؿ� Name�� ��Ƴ��� ���� ���� Array
	'--------------- ������ coding part(��������,End)------------------------------------------------------


	Call HideStatusWnd
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "PB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "PB")

	lgPageNo		 = UNICInt(Trim(Request("lgPageNo")),0)			  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgSelectList	 = Request("lgSelectList")
	lgTailList	   = Request("lgTailList")
	lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)		 '�� : �� �ʵ��� ����Ÿ Ÿ�� 
	Call  TrimData()													 '�� : Parent�� ������ ����Ÿ ���� 
	Call  FixUNISQLData()												'�� : DB-Agent�� ���� parameter ����Ÿ set
	call  QueryData()													'�� : DB-Agent�� ���� ADO query


'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
 Sub MakeSpreadSheetData()
	Const C_SHEETMAXROWS_D  = 100

	Dim iLoopCount
	Dim iRowStr
	Dim ColCnt
	Dim PvArr

	lgDataExist	= "Yes"
	lgstrData	  = ""

	If CLng(lgPageNo) > 0 Then
	   rs0.Move	 = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)				  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
	End If

	iLoopCount = -1
	ReDim PvArr(C_SHEETMAXROWS_D - 1)

   Do while Not (rs0.EOF Or rs0.BOF)

		iLoopCount =  iLoopCount + 1
		iRowStr = ""

		For ColCnt = 0 To UBound(lgSelectListDT) - 1
			iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next

		If iLoopCount < C_SHEETMAXROWS_D Then
		   lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
		   PvArr(iLoopCount) = lgstrData
			lgstrData = ""
		Else
		   lgPageNo = lgPageNo + 1
		   Exit Do
		End If

		rs0.MoveNext
	Loop

	lgstrData  = Join(PvArr, "")

	If iLoopCount < C_SHEETMAXROWS_D Then									  '��: Check if next data exists
	   lgPageNo = ""
	End If
	rs0.Close													   '��: Close recordset object
	Set rs0 = Nothing												'��: Release ADF

End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

	Redim UNISqlId(6)													 '��: SQL ID ������ ���� ����Ȯ�� 
	'--------------- ������ coding part(�������,Start)----------------------------------------------------
	Redim UNIValue(6,17)												  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
															  '	parameter�� ���� ���� ������ 
	UNISqlId(0) = "M2111QA401"
																  '* : ������ ��ȸ���Ǻθ��� Name �� �������� SQL ���� ���� 
	UNISqlId(1) = "M2111QA302"											  '����� 
	UNISqlId(2) = "M2111QA303"											  'ǰ��� 
	UNISqlId(3) = "M2111QA304"											  '��û������¸� 
	UNISqlId(4) = "M2111QA305"											  '�μ��� 
	UNISqlId(5) = "M2111QA306"											  '���ſ�û���и� 
' 	UNISqlId(6) = "s0000qa017"										  'Ʈ��ŷ�ѹ� �˻� 
															  'Reusage is Strongly Recommended.
	'--------------- ������ coding part(�������,End)------------------------------------------------------

	UNIValue(0,0) = Trim(lgSelectList)									  '��: Select ������ Summary	�ʵ� 

	'--------------- ������ coding part(�������,Start)----------------------------------------------------

	UNIValue(0,1)  = " " & FilterVar(Trim(UCase(Request("txtchangorgid"))), " " , "S") & " "
	UNIValue(0,2)  = UCase(Trim(strPlantCdFrom))		'---���� 
	UNIValue(0,3)  = UCase(Trim(strPlantCd))
	UNIValue(0,4)  = UCase(Trim(strItemCdFrom))		'---ǰ�� 
	UNIValue(0,5)  = UCase(Trim(strItemCd))
	UNIValue(0,6)  = UCase(Trim(strPrFrDt))			'---���ſ�û�� 
	UNIValue(0,7)  = UCase(Trim(strPrToDt))
	UNIValue(0,8)  = UCase(Trim(strPdFrDt))			'---�ʿ䳳���� 
	UNIValue(0,9)  = UCase(Trim(strPdToDt))
	UNIValue(0,10)  = UCase(Trim(strPrStsCdFrom))		'---��û������� 
	UNIValue(0,11) = UCase(Trim(strPrStsCd))
	UNIValue(0,12) = UCase(Trim(strRqDeptCdFrom))		'---��û�μ� 
	UNIValue(0,13) = UCase(Trim(strRqDeptCd))
	UNIValue(0,14) = UCase(Trim(strPrTypeCdFrom))		'---���ſ�û���� 
	UNIValue(0,15) = UCase(Trim(strPrTypeCd))
	UNIValue(0,16) = UCase(Trim(strTrackNo))

	UNIValue(1,0) = UCase(Trim(strPlantCd))
	UNIValue(2,0) = UCase(Trim(strPlantCd))
	UNIValue(2,1) = UCase(Trim(strItemCd))
	UNIValue(3,0) = UCase(Trim(strPrStsCd))
	UNIValue(4,0) = UCase(Trim(strRqDeptCd))
	UNIValue(4,1) = " " & FilterVar(Trim(UCase(Request("txtchangorgid"))), " " , "S") & " "
	UNIValue(5,0) = UCase(Trim(strPrTypeCd))
'	UNIValue(6,0) = UCase(Trim(strTrackNo))


	'--------------- ������ coding part(�������,End)----------------------------------------------------

	UNIValue(0,UBound(UNIValue,2)) = Trim(lgTailList)	'---Order By ���� 

	UNILock = DISCONNREAD :	UNIFlag = "1"								 '��: set ADO read mode

End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
 Sub QueryData()
	Dim iStr
	Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
	lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5, rs6)

	iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
		Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	End If

	Dim FalsechkFlg

	FalsechkFlg = False


	'============================= �߰��� �κ� =====================================================================
	If  rs1.EOF And rs1.BOF Then
		rs1.Close
		Set rs1 = Nothing

		If Len(Request("txtPlantCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "����", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		   FalsechkFlg = True
		End If
	Else
		arrRsVal(0) = rs1(0)
		arrRsVal(1) = rs1(1)
		rs1.Close
		Set rs1 = Nothing
	End If

	If  rs2.EOF And rs2.BOF Then
		rs2.Close
		Set rs2 = Nothing
		If Len(Request("txtItemCd")) And FalsechkFlg = False Then
			Call DisplayMsgBox("122700", vbInformation, "", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
			FalsechkFlg = True
			rs0.Close
			Set rs0 = Nothing
			Exit Sub		'20030124 - leejt
		End If
	Else
		arrRsVal(2) = rs2(0)
		arrRsVal(3) = rs2(1)
		rs2.Close
		Set rs2 = Nothing
	End If

	If  rs3.EOF And rs3.BOF Then
		rs3.Close
		Set rs3 = Nothing
		If Len(Request("txtPrStsCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "��û�������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		   FalsechkFlg = True
		End If
	Else
		arrRsVal(4) = rs3(0)
		arrRsVal(5) = rs3(1)
		rs3.Close
		Set rs3 = Nothing
	End If

	If  rs4.EOF And rs4.BOF Then
		rs4.Close
		Set rs4 = Nothing
		If Len(Request("txtRqDeptCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "��û�μ�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		   FalsechkFlg = True
		End If
	Else
		arrRsVal(6) = rs4(0)
		arrRsVal(7) = rs4(1)
		rs4.Close
		Set rs4 = Nothing
	End If

	If  rs5.EOF And rs5.BOF Then
		rs5.Close
		Set rs5 = Nothing
		If Len(Request("txtPrTypeCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "���ſ�û����", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		   FalsechkFlg = True
		End If
	Else
		arrRsVal(8) = rs5(0)
		arrRsVal(9) = rs5(1)
		rs5.Close
		Set rs5 = Nothing
	End If

' 	If  rs6.EOF And rs6.BOF Then
' 		rs6.Close
' 		Set rs6 = Nothing
' 		If Len(Request("txtTrackNo")) And FalsechkFlg = False Then
' 		   Call DisplayMsgBox("970000", vbInformation, "Tracking No", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
' 		   FalsechkFlg = True
' 		End If
' 	Else
' 		rs6.Close
' 		Set rs6 = Nothing
' 	End If

	If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
		rs0.Close
		Set rs0 = Nothing
	Else
		Call  MakeSpreadSheetData()
	End If

End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
 Sub TrimData()

'--------------- ������ coding part(�������,Start)----------------------------------------------------
	'---���� 
	If Len(Trim(Request("txtPlantCd"))) Then
		strPlantCd	= " " & FilterVar(Trim(UCase(Request("txtPlantCd"))), " " , "S") & " "
		strPlantCdFrom = strPlantCd
	Else
		strPlantCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
		strPlantCdFrom = "''"
	End If

	'---ǰ�� 
	If Len(Trim(Request("txtItemCd"))) Then
		strItemCd	= " " & FilterVar(Trim(UCase(Request("txtItemCd"))), " " , "S") & " "
		strItemCdFrom = strItemCd
	Else
		strItemCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
		strItemCdFrom = "''"
	End If

	'---���ſ�û�� 
	If Len(Trim(Request("txtPrFrDt"))) Then
		strPrFrDt 	= " " & FilterVar(uniConvDate(Request("txtPrFrDt")), "''", "S") & ""
	Else
		strPrFrDt	= "" & FilterVar("1900/01/01", "''", "S") & ""
	End If

	If Len(Trim(Request("txtPrToDt"))) Then
		strPrToDt 	= " " & FilterVar(uniConvDate(Request("txtPrToDt")), "''", "S") & ""
	Else
		strPrToDt	= "" & FilterVar("2999/12/30", "''", "S") & ""
	End If

	'---�ʿ䳳���� 
	If Len(Trim(Request("txtPdFrDt"))) Then
		strPdFrDt 	= " " & FilterVar(uniConvDate(Request("txtPdFrDt")), "''", "S") & ""
	Else
		strPdFrDt	= "" & FilterVar("1900/01/01", "''", "S") & ""
	End If

	If Len(Trim(Request("txtPdToDt"))) Then
		strPdToDt 	= " " & FilterVar(uniConvDate(Request("txtPdToDt")), "''", "S") & ""
	Else
		strPdToDt	= "" & FilterVar("2999/12/30", "''", "S") & ""
	End If

	'---��û������� 
	If Len(Trim(Request("txtPrStsCd"))) Then
		strPrStsCd	= " " & FilterVar(Trim(UCase(Request("txtPrStsCd"))), " " , "S") & " "
		strPrStsCdFrom = strPrStsCd
	Else
		strPrStsCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
		strPrStsCdFrom = "''"
	End If

	'---��û�μ� 
	If Len(Trim(Request("txtRqDeptCd"))) Then
		strRqDeptCd	= " " & FilterVar(Trim(UCase(Request("txtRqDeptCd"))), " " , "S") & " "
		strRqDeptCdFrom = strRqDeptCd
	Else
		strRqDeptCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
		strRqDeptCdFrom = "''"
	End If

	'---���ſ�û���� 
	If Len(Trim(Request("txtPrTypeCd"))) Then
		strPrTypeCd	= " " & FilterVar(Trim(UCase(Request("txtPrTypeCd"))), " " , "S") & " "
		strPrTypeCdFrom = strPrTypeCd
	Else
		strPrTypeCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
		strPrTypeCdFrom = "''"
	End If


	If Len(Trim(Request("txtTrackNo"))) Then
		strTrackNo 	= " " & FilterVar(Trim(Request("txtTrackNo")), "''", "S") & ""
	Else
		strTrackNo	= " A.TRACKING_NO "
	End If


'--------------- ������ coding part(�������,End)------------------------------------------------------

End Sub


%>

<Script Language=vbscript>

	With Parent
		.ggoSpread.Source  = .frm1.vspdData
		.ggoSpread.SSShowData "<%=lgstrData%>"				  '�� : Display data
		.lgPageNo			=  "<%=lgPageNo%>"			   '�� : Next next data tag
		
		.frm1.hdnPlantCd.value	  = "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hdnItemCd.value	  = "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.hdnPrFrDt.value	  = "<%=ConvSPChars(Request("txtPrFrDt"))%>"
		.frm1.hdnPrToDt.value	  = "<%=ConvSPChars(Request("txtPrToDt"))%>"
		.frm1.hdnPdFrDt.value	  = "<%=ConvSPChars(Request("txtPdFrDt"))%>"
		.frm1.hdnPdToDt.value	  = "<%=ConvSPChars(Request("txtPdToDt"))%>"
		.frm1.hdnPrStsCd.value	  = "<%=ConvSPChars(Request("txtPrStsCd"))%>"
		.frm1.hdnRqDeptCd.value   = "<%=ConvSPChars(Request("txtRqDeptCd"))%>"
		.frm1.hdnPrTypeCd.value   = "<%=ConvSPChars(Request("txtPrTypeCd"))%>"
		.frm1.hdnTrackNo.value  = "<%=ConvSPChars(Request("txtTrackNo"))%>"
		
		.frm1.txtPlantNm.value		=  "<%=ConvSPChars(arrRsVal(1))%>"
		.frm1.txtItemNm.value		=  "<%=ConvSPChars(arrRsVal(3))%>"
		.frm1.txtPrStsNm.value		=  "<%=ConvSPChars(arrRsVal(5))%>"
		.frm1.txtRqDeptNm.value		=  "<%=ConvSPChars(arrRsVal(7))%>"
		.frm1.txtPrTypeNm.value		=  "<%=ConvSPChars(arrRsVal(9))%>"
	.DbQueryOk
	End with
</Script>
<%
	Response.End												'��: �����Ͻ� ���� ó���� ������ 
%>

