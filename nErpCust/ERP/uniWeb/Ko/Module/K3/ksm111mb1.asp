<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : MC
'*  2. Function Name        :
'*  3. Program ID           : sm111mb1
'*  4. Program Name         : ��Ƽ���۴ϼ��ֵ�� 
'*  5. Program Desc         : ��Ƽ���۴ϼ��ֵ��-��Ƽ 
'*  6. Component List       :
'*  7. Modified date(First) : 2005/01/24
'*  8. Modified date(Last)  :
'*  9. Modifier (First)     : Sim Hae Young
'* 10. Modifier (Last)      :
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/adovbs.inc" -->
<!-- #Include file="../../inc/incServeradodb.asp" -->
<%

call LoadBasisGlobalInf()
call LoadInfTB19029B("I", "*","NOCOOKIE","MB")
call LoadBNumericFormatB("I","*","NOCOOKIE","MB")

'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                 '�� : DBAgent Parameter ���� 
Dim rs1, rs2, rs3, rs4,rs5
Dim istrData
Dim iStrPoNo
Dim StrNextKey		' ���� �� 
Dim lgStrPrevKey	' ���� �� 
Dim iLngMaxRow		' ���� �׸����� �ִ�Row
Dim iLngRow
Dim GroupCount
Dim lgCurrency
Dim index,Count     ' ���� �� Return ���� ���� ������ ���� ���� 
Dim lgDataExist
Dim lgPageNo


Dim SupplierNM			'�� : ���ֹ��� 

Dim lgOpModeCRUD
Dim Inti
Dim intARows
Dim intTRows

Dim istr	'������ ���� 

intARows=0
intTRows=0
On Error Resume Next                                                             '��: Protect system from crashing
Err.Clear                                                                        '��: Clear Error status

Dim strSpread																'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 



Call HideStatusWnd                                                               '��: Hide Processing message

lgOpModeCRUD  = Request("txtMode")


'response.write "dfdsfd" & Request("txtSpread") &"<br>"
'response.write UID_M0002 & UID_M0002 &"<br>"

'response.end


Select Case lgOpModeCRUD
	Case CStr(UID_M0001)                                                         '��: Query
		Call  SubBizQueryMulti()
	Case CStr(UID_M0002)
		Call SubBizSaveMulti()
End Select

Sub SubBizQueryMulti()
	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '��: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgDataExist      = "No"
	iLngMaxRow = CLng(Request("txtMaxRows"))
	lgStrPrevKey = Request("lgStrPrevKey")

	Call FixUNISQLData()		'�� : DB-Agent�� ���� parameter ����Ÿ set

	Call QueryData()			'�� : DB-Agent�� ���� ADO query

	'-----------------------
	'Result data display area
	'-----------------------
%>
	<Script Language=vbscript>
		With parent
			.frm1.txtSupplierNM.Value	= "<%=SupplierNM%>"

			.frm1.hdnItem.value = "<%=ConvSPChars(Request("txtitemcd"))%>"

			.frm1.txtPlantCd.focus
			Set .gActiveElement = .document.activeElement

			If "<%=lgDataExist%>" = "Yes" Then

				'Show multi spreadsheet data from this line
				.ggoSpread.Source    = .frm1.vspdData
				.ggoSpread.SSShowData "<%=istrData%>"                  '��: Display data

				.lgPageNo			 =  "<%=lgPageNo%>"				    '��: Next next data tag

				.DbQueryOk <%=intARows%>,<%=intTRows%>

			End If
		End with
	</Script>
<%
End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data
'============================================================================================================
Sub SubBizSaveMulti()
On Error Resume Next                                                             '��: Protect system from crashing
Err.Clear
	Dim iErrorPosition
	Dim LngMaxRow
	Dim arrTemp
	Dim arrVal
	Dim lGrpCnt
	Dim LngRow
	Dim iRow_cnt

	Dim iNumRow
	Dim iCFM_YN
	Dim iPO_COMPANY
	Dim iSO_COMPANY
	Dim iPO_NO
	Dim itxtSo_Type
	Dim itxtSo_dt
	Dim itxtDeal_Type
	Dim itxtSales_Grp
	Dim itxtPlantCd
	Dim itxtvalid_dt

	Dim ObjPSMG111



	LngMaxRow = CInt(Request("txtMaxRows"))								'��: �ִ� ������Ʈ�� ���� 
	arrTemp = Split(Request("txtSpread"), gRowSep)

	'Response.write 	"txtSpread:" & Request("txtSpread")							'��: Spread Sheet ������ ��� �ִ� Element�� 
	'Response.end

	lGrpCnt = 0

	Set ObjPSMG111 = Server.CreateObject ("PSMG111.CMaintMcCustPoSoSvr")

	If CheckSYSTEMError(Err,True) = true then
		Set ObjPSMG111 = Nothing
		Exit Sub
	End If

	'//Response.Write "arrTemp(0):" & arrTemp(0) & "<br>"
	'//Response.Write "arrTemp(1):" & arrTemp(1) & "<br>"

	For LngRow = 1 To LngMaxRow
			Err.Clear


			arrVal = Split(arrTemp(LngRow-1), gColSep)


			iNumRow		= arrVal(1)
			iCFM_YN		= arrVal(2)
			iPO_COMPANY	= arrVal(3)
			iSO_COMPANY	= arrVal(4)
			iPO_NO 		= arrVal(5)

			'�������� 	txtSo_Type
			itxtSo_Type	= arrVal(6)
			'������		txtSo_dt
			itxtSo_dt 	= arrVal(7)
			'�Ǹ�����	txtDeal_Type
			itxtDeal_Type	= arrVal(8)
			'�����׷�	txtSales_Grp
			itxtSales_Grp 	= arrVal(9)
			'����		txtPlantCd
			itxtPlantCd 	= arrVal(10)
			'��ȿ�� 
			itxtvalid_dt	= arrVal(11)

			'Response.write "--------------------------" &"<br>"
			'Response.write "iCFM_YN:" & iCFM_YN &"<br>"
			'Response.write "iPO_COMPANY:" & iPO_COMPANY &"<br>"
			'Response.write "iSO_COMPANY:" & iSO_COMPANY &"<br>"
			'Response.write "iPO_NO:" & iPO_NO &"<br>"
			'Response.write "itxtSo_Type:" & itxtSo_Type &"<br>"
			'Response.write "itxtDeal_Type:" & itxtDeal_Type &"<br>"
			'Response.write "itxtSales_Grp:" & itxtSales_Grp &"<br>"
			'Response.write "itxtPlantCd:" & itxtPlantCd &"<br>"
			'Response.write "itxtvalid_dt:" & itxtvalid_dt &"<br>"
			'response.end

			'Response.write "--------------------------" &"<br>"

			On Error Resume Next                                                             '��: Protect system from crashing
			Err.Clear

			Call ObjPSMG111.S_UPDATE_MC_PO_STS_SOMK(gStrGlobalCollection,	iCFM_YN, _
											iPO_COMPANY, _
											iSO_COMPANY, _
											iPO_NO, _
											itxtSo_Type, _
											itxtSo_dt, _
											itxtDeal_Type, _
											itxtSales_Grp, _
											itxtPlantCd, _
											itxtvalid_dt, _
											iErrorPosition)

			'-----------------------
			'Com action result check area(DB,internal)
			'-----------------------
			If CheckSYSTEMError2(Err, True, iNumRow & "��:", "", "", "", "") = True Then
			    	Err.Clear
			    	If LngRow = LngMaxRow Then
			    		Exit For
			    	End If
				'ó���� �Ϸ�Ȱ��� Check Box �� Ǯ��.
				Response.Write "<Script language=vbscript> "		& vbCr
				Response.Write "	Dim iBln "				& vbCr
				Response.Write "            iBln = MsgBox (""��������Ͻðڽ��ϱ�?"", vbYesNo, """") "				& vbCr
				Response.Write "            If iBln = vbNo Then   "				& vbCr
				Response.Write "	       Parent.DbSaveOk    "				& vbCr
				Response.Write "	    End If"						& vbCr
				Response.Write "</Script> "
			Else
				'ó���� �Ϸ�Ȱ��� Check Box �� Ǯ��.
				Response.Write "<Script language=vbscript> "		& vbCr
				Response.Write "On error resume Next"				& vbCr
				Response.Write "	with Parent.frm1.vspdData"      & vbCr
				Response.Write "		Dim iIndex, iRowNo	"		& vbCr
				Response.Write "		for iIndex = 1 to .MaxRows	"      & vbCr
				Response.Write "			.Col = Parent.C_PO_NO	"      & vbCr
				Response.Write "			.Row = iIndex	"		& vbCr
				Response.Write "			If Trim(.text) = """	&  iPO_NO & """ then "     & vbCr
				Response.Write "				iRowNo = iIndex	"   & vbCr
				Response.Write "			End if	"				& vbCr
				Response.Write "		Next	"					& vbCr
				Response.Write "		.Col = parent.C_CfmFlg	"   & vbCr
				Response.Write "		.Row = iRowNo "				& vbCr
				Response.Write "		.Text = 0 "					& vbCr
				Response.Write "	end with "						& vbCr
				Response.Write "</Script> "

			End If
	Next

	If NOT(ObjPSMG111 is Nothing) Then
		Set ObjPSMG111 = Nothing
	End If

    Response.Write "<Script language=vbs> " & vbCr
    Response.Write " Parent.DbSaveOk "      & vbCr							'��: ȭ�� ó�� ASP �� ��Ī�� 
    Response.Write "</Script> "


End Sub

'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Const C_SHEETMAXROWS_D  = 100
	Dim iLoopCount
	Dim iRowStr
	Dim ColCnt

	Const C_PO_NO		= 0
	Const C_IMPORT_FLG	= 1
	Const C_PO_CUR		= 2
	Const C_PO_DOC_AMT	= 3
	Const C_PO_VAT_DOC_AMT	= 4
	Const C_PO_VAT_TOT_AMT	= 5
	Const C_PO_VAT_TYPE	= 6
	Const C_PO_VAT_TYPE_NM	= 7
	Const C_PO_VAT_RT	= 8
	Const C_PO_PAY_METH	= 9
	Const C_PO_PAY_METH_NM	= 10
	Const C_PO_INCOTERMS	= 11
	Const C_PO_INCOTERMS_NM	= 12

	Const C_PO_COMPANY	= 13
	Const C_SO_COMPANY	= 14

	lgDataExist    = "Yes"

	If CLng(lgPageNo) > 0 Then
		rs0.Move     	= CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
		intTRows	= CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)
	End If

	'----- ���ڵ�� Į�� ���� ----------
	'-----------------------------------
	iLoopCount = 0

    	ReDim PvArr(C_SHEETMAXROWS_D - 1)

	Do while Not (rs0.EOF Or rs0.BOF)
		iLoopCount =  iLoopCount + 1
		iRowStr = ""
		If ConvSPChars(rs0(C_IMPORT_FLG))="N" Then
			istr = "����"
		Else
			istr = "����"
		End If


		iRowStr = iRowStr & Chr(11) & ""
		iRowStr = iRowStr & Chr(11) & ""
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_NO))
		iRowStr = iRowStr & Chr(11) & istr
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_CUR))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_DOC_AMT))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_VAT_DOC_AMT))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_VAT_TOT_AMT))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_VAT_TYPE))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_VAT_TYPE_NM))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_VAT_RT))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_PAY_METH))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_PAY_METH_NM))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_INCOTERMS))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_INCOTERMS_NM))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_COMPANY))
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_SO_COMPANY))

		iRowStr = iRowStr & Chr(11) & iLngMaxRow + iLoopCount

		If iLoopCount - 1 < C_SHEETMAXROWS_D Then
		   istrData = istrData & iRowStr & Chr(11) & Chr(12)

    	Else
		   istrData = istrData & iRowStr & Chr(11) & Chr(12)
		   lgPageNo = lgPageNo + 1
		   Exit Do
		End If

		rs0.MoveNext
	Loop


	intARows = iLoopCount
	If iLoopCount < C_SHEETMAXROWS_D Then                                      '��: Check if next data exists
	  lgPageNo = ""
	End If

	rs0.Close                                                       '��: Close recordset object
	Set rs0 = Nothing	                                            '��: Release ADF
End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Function SetConditionData()
    On Error Resume Next

    SetConditionData = false

	If Not(rs1.EOF Or rs1.BOF) Then
		SupplierNM = rs1("BP_FULL_NM")
		Set rs1 = Nothing

	Else
		Set rs1 = Nothing
		If Len(Request("txtSupplierCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "���ֹ���", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		    Exit Function
		End If
	End If

    SetConditionData = True
End Function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

   	Dim strVal
	ReDim UNISqlId(1)                                                     '��: SQL ID ������ ���� ����Ȯ�� 
	'--------------- ������ coding part(�������,Start)----------------------------------------------------
	Redim UNIValue(1,4)                                                  '��: DB-Agent�� ���۵� parameter�� ���� ���� 
                                                                          '    parameter�� ���� ���� ������ 
	strVal = ""
	UNISqlId(0) = "SM111MA1"		'����Splead Query
	UNISqlId(1) = "SM111MA101"		'���ֹ��� PopUp

	UNIValue(1,0) = "'zzzz'"

	'//UNIValue(0,0) = "^"

	'���ֹ��� 
	If Trim(Request("txtSupplierCd")) <> "" Then
		UNIValue(0,0) = " '"& FilterVar(Trim(UCase(Request("txtSupplierCd"))), " " , "SNM") & "' "
		UNIValue(1,0) = " '"& FilterVar(Trim(UCase(Request("txtSupplierCd"))), " " , "SNM") & "' "
	Else
	    	UNIValue(0,0) = "|"

	End If

	'��������(From)
	If Trim(Request("txtFrDt")) <> "" Then
		UNIValue(0,1) = " '"& FilterVar(Trim(UCase(Request("txtFrDt"))), " " , "SNM") & "' "
	Else
	    	UNIValue(0,1) = "|"
	End If

	'��������(To)
	If Trim(Request("txtToDt")) <> "" Then
		UNIValue(0,2) = " '"& FilterVar(Trim(UCase(Request("txtToDt"))), " " , "SNM") & "' "
	Else
	    	UNIValue(0,2) = "|"
	End If

	'������ó�� ���� 
	If Trim(Request("rdoPostFlag2")) <> "" Then
		UNIValue(0,3) = " '"& FilterVar(Trim(UCase(Request("rdoPostFlag2"))), " " , "SNM") & "' "
	Else
	    	UNIValue(0,3) = "|"
	End If

	'�����ֹ�ȣ 
	If Trim(Request("txtPO_NO")) <> "" Then
		UNIValue(0,4) = " '"& FilterVar(Trim(UCase(Request("txtPO_NO"))), " " , "SNM") & "' "
	Else
	    	UNIValue(0,4) = "|"
	End If

	UNILock = DISCONNREAD :	UNIFlag = "1"                                 '��: set ADO read mode
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
	On Error Resume Next
	Dim lgstrRetMsg                                             '�� : Record Set Return Message �������� 
	Dim lgADF                                                   '�� : ActiveX Data Factory ���� �������� 
	Dim iStr

	Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")

	lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)

	Set lgADF   = Nothing

	iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
		Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	End If

	'�˾��ʵ� üũ 
	If Setconditiondata = False Then Exit Sub

	If  rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
		rs0.Close
		Set rs0 = Nothing
	Else

		Call  MakeSpreadSheetData()

	End If
End Sub



%>
