<%'======================================================================================================
'*  1. Module Name		  : Procurement
'*  2. Function Name		:
'*  3. Program ID		   :
'*  4. Program Name		 :
'*  5. Program Desc		 :
'*  6. Modified date(First) : 2000/12/12
'*  7. Modified date(Last)  : 2001/01/09
'*  8. Modifier (First)	 : ByunJiHyun
'*  9. Modifier (Last)	  : Min Hak-jun
'* 10. Comment			  :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================
Option Explicit
%>

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
Dim iTotstrData
Dim lgTailList
Dim lgSelectList
Dim lgSelectListDT

Dim lgDataExist
Dim lgPageNo
Dim iPrevEndRow
Dim iEndRow

'--------------- ������ coding part(��������,Start)----------------------------------------------------
Dim ICount  												'   Count for column index
Dim strPlantCd												'	���� 
Dim strPlantCdFrom
Dim strItemCd											   '   ǰ�� 
Dim strItemCdFrom
Dim strBpCd													'   �ŷ�ó 
Dim strBpCdFrom
Dim strDvFrDt											   '   ������ 
Dim strDvToDt
Dim strPurGrpCd												'	���ű׷� 
Dim strPurGrpCdFrom
Dim strPoType											   '   �������� 
Dim strPoTypeFrom
Dim strTrackNo

Dim arrRsVal(4)											'* : ȭ�鿡 ��ȸ�ؿ� Name�� ��Ƴ��� ���� ���� Array

	Call HideStatusWnd
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "M", "NOCOOKIE", "QB")
	Call LoadBNumericFormatB("Q", "M", "NOCOOKIE", "QB")

	lgPageNo		 = UNICInt(Trim(Request("lgPageNo")),0)			  '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgSelectList	 = Request("lgSelectList")
	lgTailList	   = Request("lgTailList")
	lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)		 '�� : �� �ʵ��� ����Ÿ Ÿ�� 
	iPrevEndRow = 0
	iEndRow = 0

	Call  TrimData()													 '�� : Parent�� ������ ����Ÿ ���� 
	Call  FixUNISQLData()												'�� : DB-Agent�� ���� parameter ����Ÿ set
	call  QueryData()													'�� : DB-Agent�� ���� ADO query


'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
 Sub MakeSpreadSheetData()
	Dim  RecordCnt
	Dim  ColCnt
	Dim  iLoopCount
	Dim  iRowStr
	Dim PvArr

	Const C_SHEETMAXROWS_D = 100

	lgDataExist	= "Yes"
	lgstrData	  = ""
	iPrevEndRow = 0

	If CInt(lgPageNo) > 0 Then
	   iPrevEndRow = C_SHEETMAXROWS_D * CInt(lgPageNo)
	   rs0.Move  = iPrevEndRow
	End If

	iLoopCount = -1
	ReDim PvArr(C_SHEETMAXROWS_D - 1)

	Do while Not (rs0.EOF Or rs0.BOF)
		iLoopCount =  iLoopCount + 1
		iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
			iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next

		If  iLoopCount < C_SHEETMAXROWS_D Then
			lgstrData	  = lgstrData	  & iRowStr & Chr(11) & Chr(12)
			PvArr(iLoopCount) = lgstrData
			lgstrData = ""
		Else
			lgPageNo = lgPageNo + 1
			Exit Do
		End If
		rs0.MoveNext
	Loop

	iTotstrData = Join(PvArr, "")

	If  iLoopCount < C_SHEETMAXROWS_D Then											'��: Check if next data exists
		iEndRow = iPrevEndRow + iLoopCount + 1
		lgPageNo = ""												  '��: ���� ����Ÿ ����.
	Else
		iEndRow = iPrevEndRow + iLoopCount
	End If

	rs0.Close
	Set rs0 = Nothing
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
 Sub FixUNISQLData()

	Redim UNISqlId(6)													 '��: SQL ID ������ ���� ����Ȯ�� 
	Redim UNIValue(6,14)												  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
																		  '	parameter�� ���� ���� ������ 
	 UNISqlId(0) = "M3112QA401"

	 UNISqlId(1) = "M2111QA302"											  '����� 
	 UNISqlId(2) = "M2111QA303"											  'ǰ��� 
	 UNISqlId(3) = "M3111QA102"											  '�ŷ�ó�� 
	 UNISqlId(4) = "M3111QA104"											  '���ű׷�� 
	 UNISqlId(5) = "M3111QA103"											  '�������¸� 
																		  'Reusage is Recommended
	 UNIValue(0,0) = Trim(lgSelectList)									  '��: Select ������ Summary	�ʵ� 
	 UNIValue(0,1)  = UCase(Trim(strPlantCdFrom))		'---���� 
	 UNIValue(0,2)  = UCase(Trim(strPlantCd))
	 UNIValue(0,3)  = UCase(Trim(strItemCdFrom))		'---ǰ�� 
	 UNIValue(0,4)  = UCase(Trim(strItemCd))
	 UNIValue(0,5)  = UCase(Trim(strBpCdFrom))			'---�ŷ�ó 
	 UNIValue(0,6)  = UCase(Trim(strBpCd))
	 UNIValue(0,7)  = UCase(Trim(strDvFrDt))			'---������ 
	 UNIValue(0,8)  = UCase(Trim(strDvToDt))
	 UNIValue(0,9)  = UCase(Trim(strPurGrpCdFrom))		'---���ű׷� 
	 UNIValue(0,10)  = UCase(Trim(strPurGrpCd))
	 UNIValue(0,11)  = UCase(Trim(strPoTypeFrom))		'---�������� 
	 UNIValue(0,12)  = UCase(Trim(strPoType))
	 UNIValue(0,13) = UCase(Trim(strTrackNo))

	 UNIValue(1,0)  = UCase(Trim(strPlantCd))
	 UNIValue(2,0)  = UCase(Trim(strPlantCd))
	 UNIValue(2,1)  = UCase(Trim(strItemCd))
	 UNIValue(3,0)  = UCase(Trim(strBpCd))
	 UNIValue(4,0)  = UCase(Trim(strPurGrpCd))
	 UNIValue(5,0)  = UCase(Trim(strPoType))

	 UNIValue(0,UBound(UNIValue,2)	) = Trim(lgTailList)	'---Order By ���� 
	 UNILock = DISCONNREAD :	UNIFlag = "1"								 '��: set ADO read mode

End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
 Sub QueryData()
	Dim iStr
	Dim FalsechkFlg

	FalsechkFlg = False

	Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
	lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5)

	iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
		Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	End If

	If  rs1.EOF And rs1.BOF Then
		rs1.Close
		Set rs1 = Nothing
		If Len(Request("txtPlantCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "����", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		   FalsechkFlg = True
		End If
	Else
		arrRsVal(0) = rs1(1)
		rs1.Close
		Set rs1 = Nothing
	End If

	If  rs2.EOF And rs2.BOF Then
		rs2.Close
		Set rs2 = Nothing
		If Len(Request("txtItemCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("122700", vbInformation, "", "", I_MKSCRIPT)
		   Set rs2 = Nothing
		   FalsechkFlg = True
		   Exit Sub

		End If
	Else
		arrRsVal(1) = rs2(1)
		rs2.Close
		Set rs2 = Nothing
	End If

	If  rs3.EOF And rs3.BOF Then
		rs3.Close
		Set rs3 = Nothing
		If Len(Request("txtBpCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "�ŷ�ó", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		   FalsechkFlg = True
		End If
	Else
		arrRsVal(2) = rs3(1)
		rs3.Close
		Set rs3 = Nothing
	End If

	If  rs4.EOF And rs4.BOF Then
		rs4.Close
		Set rs4 = Nothing
		If Len(Request("txtPurGrpCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "���ű׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		   FalsechkFlg = True
		End If
	Else
		arrRsVal(3) = rs4(1)
		rs4.Close
		Set rs4 = Nothing
	End If

	If  rs5.EOF And rs5.BOF Then
		rs5.Close
		Set rs5 = Nothing
		If Len(Request("txtPoType")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "��������", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		   FalsechkFlg = True
		End If
	Else
		arrRsVal(4) = rs5(1)
		rs5.Close
		Set rs5 = Nothing
	End If

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
	'---�ŷ�ó 
	If Len(Trim(Request("txtBpCd"))) Then
		strBpCd	= " " & FilterVar(Trim(UCase(Request("txtBpCd"))), " " , "S") & " "
		strBpCdFrom = strBpCd
	Else
		strBpCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
		strBpCdFrom = "''"
	End If
	 '---������

	If Trim(Request("txtPlantCd")) = "P04" Then

		strDvFrDt	= "" & FilterVar("2008-09-01", "''", "S") & ""
	Else

		If Len(Trim(Request("txtDvFrDt"))) Then
			strDvFrDt 	= " " & FilterVar(uniConvDate(Request("txtDvFrDt")), "''", "S") & ""
		Else
			strDvFrDt	= "" & FilterVar("1900-01-01", "''", "S") & ""
		End If
	End IF
	
	If Len(Trim(Request("txtDvToDt"))) Then
			strDvToDt 	= " " & FilterVar(uniConvDate(Request("txtDvToDt")), "''", "S") & ""
		Else
			strDvToDt	= "" & FilterVar("2999-12-30", "''", "S") & ""
	End If

	
	 '---���ű׷� 
	If Len(Trim(Request("txtPurGrpCd"))) Then
		strPurGrpCd	= " " & FilterVar(Trim(UCase(Request("txtPurGrpCd"))), " " , "S") & " "
		strPurGrpCdFrom = strPurGrpCd
	Else
		strPurGrpCd	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
		strPurGrpCdFrom = "''"
	End If
	 '---�������� 
	If Len(Trim(Request("txtPoType"))) Then
		strPoType	= " " & FilterVar(Trim(UCase(Request("txtPoType"))), " " , "S") & " "
		strPoTypeFrom = strPoType
	Else
		strPoType	= "" & FilterVar("zzzzzzzzz", "''", "S") & ""
		strPoTypeFrom = "''"
	End If

	If Len(Trim(Request("txtTrackNo"))) Then
		strTrackNo 	= " " & FilterVar(Trim(Request("txtTrackNo")), "''", "S") & ""
	Else
		strTrackNo	= " '' "
	End If

'--------------- ������ coding part(�������,End)------------------------------------------------------

End Sub
'----------------------------------------------------------------------------------------------------------
' Trim string and set string to space if string length is zero
'----------------------------------------------------------------------------------------------------------

%>

<Script Language=vbscript>

	With Parent
		.ggoSpread.Source  = .frm1.vspdData

		.frm1.vspdData.Redraw = False
		.ggoSpread.SSShowData "<%=iTotstrData%>"				  '�� : Display data
		.lgPageNo			=  "<%=lgPageNo%>"			   '�� : Next next data tag
		.frm1.hdnPlantCd.value	= "<%=ConvSPChars(Request("txtPlantCd"))%>"
		.frm1.hdnItemCd.value	 = "<%=ConvSPChars(Request("txtItemCd"))%>"
		.frm1.hdnBpCd.value	   = "<%=ConvSPChars(Request("txtBpCd"))%>"
		.frm1.hdnDvFrDt.value	 = "<%=ConvSPChars(Request("txtDvFrDt"))%>"
		.frm1.hdnDvToDt.value	 = "<%=ConvSPChars(Request("txtDvToDt"))%>"
		.frm1.hdnPurGrpCd.value   = "<%=ConvSPChars(Request("txtPurGrpCd"))%>"
		.frm1.hdnPoType.value	 = "<%=ConvSPChars(Request("txtPoType"))%>"
		.frm1.hdnTrackNo.value  = "<%=ConvSPChars(Request("txtTrackNo"))%>"

		.frm1.txtPlantNm.value			=  "<%=ConvSPChars(arrRsVal(0))%>"
		.frm1.txtItemNm.value			=  "<%=ConvSPChars(arrRsVal(1))%>"
		.frm1.txtBpNm.value			=  "<%=ConvSPChars(arrRsVal(2))%>"
		.frm1.txtPurGrpNm.value		=  "<%=ConvSPChars(arrRsVal(3))%>"
		.frm1.txtPoTypeNm.value		=  "<%=ConvSPChars(arrRsVal(4))%>"
		.DbQueryOk
		.frm1.vspdData.Redraw = True
	End with
</Script>

<%
	Response.End												'��: �����Ͻ� ���� ó���� ������ 
%>
