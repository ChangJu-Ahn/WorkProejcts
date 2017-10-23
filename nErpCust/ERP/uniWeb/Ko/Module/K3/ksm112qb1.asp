<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : MC
'*  2. Function Name        :
'*  3. Program ID           : sm112qb1
'*  4. Program Name         : ��Ƽ���۴ϼ�����������ȸ(���ֺ�)
'*  5. Program Desc         : ��Ƽ���۴ϼ�����������ȸ(���ֺ�)-�̱� 
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
Dim index,Count     	' ���� �� Return ���� ���� ������ ���� ���� 
Dim lgDataExist
Dim lgPageNo


Dim SupplierNM			'�� : ���ֹ��� 

Dim lgOpModeCRUD
Dim Inti
Dim intARows
Dim intTRows
intARows=0
intTRows=0
On Error Resume Next                                                             '��: Protect system from crashing
Err.Clear                                                                        '��: Clear Error status

Dim strSpread																'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 
Call HideStatusWnd                                                               '��: Hide Processing message

lgOpModeCRUD  = Request("txtMode")
'Dim aaa,bbb,ccc,ddd,eee
'
'	aaa = " '"& FilterVar(Trim(UCase(Request("txtSpplCd"))), " " , "SNM") & "' "
'	bbb = " '"& FilterVar(Trim(UCase(Request("txtSupplierCd"))), " " , "SNM") & "' "
'	ccc = " '"& FilterVar(Trim(UCase(Request("txtSo_Frdt"))), " " , "SNM") & "' "
'	ddd = " '"& FilterVar(Trim(UCase(Request("txtSo_Todt"))), " " , "SNM") & "' "
'	eee = " '"& FilterVar(Trim(UCase(Request("rdoPostFlag2"))), " " , "SNM") & "' "
'
'
'response.write "txtSpplCd" & aaa &"<br>"
'response.write "txtSupplierCd" & bbb &"<br>"
'response.write "txtSo_Frdt" & ccc &"<br>"
'response.write "txtSo_Todt" & ddd &"<br>"
'response.write "rdoPostFlag2" & eee &"<br>"

Select Case lgOpModeCRUD
	Case CStr(UID_M0001)
		Call  SubBizQueryMulti()
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

			.frm1.txtSupplierCd.focus
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

	Dim iPO_COMPANY
	Dim iSO_COMPANY
	Dim iPO_NO
	Dim ObjPSMG111


	LngMaxRow = CInt(Request("txtMaxRows"))								'��: �ִ� ������Ʈ�� ���� 
	arrTemp = Split(Request("txtSpread"), gRowSep)									'��: Spread Sheet ������ ��� �ִ� Element�� 

	lGrpCnt = 0

	Set ObjPSMG111 = Server.CreateObject ("PSMG111.CMaintMcCustPoSoSvr")

	If CheckSYSTEMError(Err,True) = true then
		Set ObjPSMG111 = Nothing
		Exit Sub
	End If

    For LngRow = 1 To LngMaxRow
		Err.Clear
		lGrpCnt = lGrpCnt 														'��: Group Count

		arrVal = Split(arrTemp(LngRow-1), gColSep)

		iPO_COMPANY	= arrVal(2)
		iSO_COMPANY	= arrVal(3)
		iPO_NO 		= arrVal(4)

		Call ObjPSMG111.S_UPDATE_MC_PO_STS_SOMK(gStrGlobalCollection,iPO_COMPANY,iSO_COMPANY,iPO_NO,iErrorPosition)

		'-----------------------
		'Com action result check area(DB,internal)
		'-----------------------
		If CheckSYSTEMError2(Err, True, LngRow & "��:", "", "", "", "") = True Then
		    	Err.Clear
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

	Const C_PO_COMPANY	        = 0	'���ֹ��� 
	Const C_PO_COMPANY_NM	        = 1	'���ֹ��� 
	Const C_SO_NO		        = 2	'���ֹ�ȣ 
	Const C_SO_SEQ_NO		= 3	'���ּ��� 
	Const C_ITEM_CD			= 4	'ǰ�� 
	Const C_ITEM_NM			= 5	'ǰ��� 
	Const C_SPEC		        = 6	'ǰ��԰� 
	Const C_PO_STS			= 7	'���ֹ��λ��� 
	Const C_SO_STS			= 8	'���ֹ��λ��� 
	Const C_UNIT		        = 9	'���� 
	Const C_PO_QTY			= 10	'���ּ��� 
	Const C_SO_QTY			= 11	'���ּ��� 
	Const C_PO_LC_QTY		= 12	'����L/C���� 
	Const C_SO_LC_QTY		= 13	'����L/C���� 
	Const C_SO_REQ_QTY	        = 14	'���Ͽ�û���� 
	Const C_SO_ISSUE_QTY	        = 15	'������ 
	Const C_SO_CC_QTY		= 16	'����������� 
	Const C_PO_CC_QTY		= 17	'����������� 
	Const C_PO_RCPT_QTY	        = 18	'�԰���� 
	Const C_SO_BILL_QTY	        = 19	'������� 
	Const C_PO_IV_QTY		= 20	'���Լ��� 
	Const C_PO_NO		        = 21	'���ֹ���ȣ 
	Const C_PO_SEQ_NO		= 22	'���� 
	Const C_BP_ITEM_CD	        = 23	'��ǰ�� 
	Const C_BP_ITEM_NM	        = 24	'��ǰ��� 

	lgDataExist    = "Yes"

	If CLng(lgPageNo) > 0 Then
		rs0.Move     	= CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
		intTRows	= CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)
	End If

	'//Response.end

	'----- ���ڵ�� Į�� ���� ----------
	'-----------------------------------
	iLoopCount = 0

    	ReDim PvArr(C_SHEETMAXROWS_D - 1)

	Do while Not (rs0.EOF Or rs0.BOF)
		iLoopCount =  iLoopCount + 1
		iRowStr = ""


		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_COMPANY))	        '���ֹ��� 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_COMPANY_NM))	        '���ֹ��θ� 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_SO_NO))		        '���ֹ�ȣ 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_SO_SEQ_NO))		'���ּ��� 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_ITEM_CD))		'ǰ�� 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_ITEM_NM))		'ǰ��� 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_SPEC))		        'ǰ��԰� 

		'=====================================================================================
		'	���ֹ��λ��� 
		'=====================================================================================
		If ConvSPChars(rs0(C_PO_STS))="PO-CFM" Then
			iRowStr = iRowStr & Chr(11) & "����Ȯ��"		'���ֹ��λ��� 
		ElseIf ConvSPChars(rs0(C_PO_STS))="PO-GR" Then
			iRowStr = iRowStr & Chr(11) & "�����԰�"		'���ֹ��λ��� 
		ElseIf ConvSPChars(rs0(C_PO_STS))="PO-IVMK" Then
			iRowStr = iRowStr & Chr(11) & "���Ե��"		'���ֹ��λ��� 
		ElseIf ConvSPChars(rs0(C_PO_STS))="PO-IVCF" Then
			iRowStr = iRowStr & Chr(11) & "����Ȯ��"		'���ֹ��λ��� 
		ElseIf ConvSPChars(rs0(C_PO_STS))="PO-LC" Then
			iRowStr = iRowStr & Chr(11) & "L/C"			'���ֹ��λ��� 
		ElseIf ConvSPChars(rs0(C_PO_STS))="PO-BL" Then
			iRowStr = iRowStr & Chr(11) & "B/L"			'���ֹ��λ��� 
		ElseIf ConvSPChars(rs0(C_PO_STS))="PO-CC" Then
			iRowStr = iRowStr & Chr(11) & "�������"		'���ֹ��λ��� 
		Else
			iRowStr = iRowStr & Chr(11) & ""			'���ֹ��λ��� 
		End If

		'=====================================================================================
		'	���ֹ��λ��� 
		'=====================================================================================
		If ConvSPChars(rs0(C_SO_STS))="SO-MK" Then
			iRowStr = iRowStr & Chr(11) & "���ֻ���"	'���ֹ��λ��� 
		ElseIf ConvSPChars(rs0(C_SO_STS))="SO-CFM" Then
			iRowStr = iRowStr & Chr(11) & "����Ȯ��"	'���ֹ��λ��� 
		ElseIf ConvSPChars(rs0(C_SO_STS))="SO-REQ" Then
			iRowStr = iRowStr & Chr(11) & "���Ͽ�û"	'���ֹ��λ��� 
		ElseIf ConvSPChars(rs0(C_SO_STS))="SO-GI" Then
			iRowStr = iRowStr & Chr(11) & "���"	'���ֹ��λ��� 
		ElseIf ConvSPChars(rs0(C_SO_STS))="SO-BILL" Then
			iRowStr = iRowStr & Chr(11) & "����"	'���ֹ��λ��� 
		ElseIf ConvSPChars(rs0(C_SO_STS))="SO-TAX" Then
			iRowStr = iRowStr & Chr(11) & "���ݰ�꼭����"	'���ֹ��λ��� 
		ElseIf ConvSPChars(rs0(C_SO_STS))="SO-LC" Then
			iRowStr = iRowStr & Chr(11) & "L/C"		'���ֹ��λ��� 
		ElseIf ConvSPChars(rs0(C_SO_STS))="SO-CC" Then
			iRowStr = iRowStr & Chr(11) & "�������"	'���ֹ��λ��� 
		Else
			iRowStr = iRowStr & Chr(11) & ""		'���ֹ��λ��� 
		End If

		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_UNIT))		        '���� 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_QTY))		'���ּ��� 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_SO_QTY))		'���ּ��� 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_LC_QTY))		'����L/C���� 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_SO_LC_QTY))		'����L/C���� 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_SO_REQ_QTY))	        '���Ͽ�û���� 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_SO_ISSUE_QTY))	        '������ 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_SO_CC_QTY))		'����������� 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_CC_QTY))		'����������� 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_RCPT_QTY))	        '�԰���� 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_SO_BILL_QTY))	        '������� 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_IV_QTY))		'���Լ��� 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_NO))		        '���ֹ���ȣ 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_PO_SEQ_NO))		'���� 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_BP_ITEM_CD))	        '��ǰ�� 
		iRowStr = iRowStr & Chr(11) & ConvSPChars(rs0(C_BP_ITEM_NM))	        '��ǰ��� 

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
	UNISqlId(0) = "SM112QA1"		'����Splead Query
	UNISqlId(1) = "SM111MA101"		'���ֹ��� PopUp

	UNIValue(1,0) = "'zzzz'"

	'//UNIValue(0,0) = "^"

'	aaa = " '"& FilterVar(Trim(UCase(Request("txtSpplCd"))), " " , "SNM") & "' "
'	bbb = " '"& FilterVar(Trim(UCase(Request("txtSupplierCd"))), " " , "SNM") & "' "
'	ccc = " '"& FilterVar(Trim(UCase(Request("txtSo_Frdt"))), " " , "SNM") & "' "
'	ddd = " '"& FilterVar(Trim(UCase(Request("txtSo_Todt"))), " " , "SNM") & "' "
'	eee = " '"& FilterVar(Trim(UCase(Request("rdoPostFlag2"))), " " , "SNM") & "' "

	'���ֹ��� 
	If Trim(Request("txtSpplCd")) <> "" Then
		UNIValue(0,0) = "|"
'		UNIValue(0,0) = " '"& FilterVar(Trim(UCase(Request("txtSpplCd"))), " " , "SNM") & "' "
	Else
	    	UNIValue(0,0) = "|"
	End If

	'���ֹ��� 
	If Trim(Request("txtSupplierCd")) <> "" Then
		UNIValue(0,1) = " '"& FilterVar(Trim(UCase(Request("txtSupplierCd"))), " " , "SNM") & "' "
		UNIValue(1,0) = " '"& FilterVar(Trim(UCase(Request("txtSupplierCd"))), " " , "SNM") & "' "
	Else
	    	UNIValue(0,1) = "|"

	End If


	'������(From)
	If Trim(Request("txtSo_Frdt")) <> "" Then
		UNIValue(0,2) = " '"& FilterVar(Trim(UCase(Request("txtSo_Frdt"))), " " , "SNM") & "' "
	Else
	    	UNIValue(0,2) = "|"
	End If

	'������(To)
	If Trim(Request("txtSo_Todt")) <> "" Then
		UNIValue(0,3) = " '"& FilterVar(Trim(UCase(Request("txtSo_Todt"))), " " , "SNM") & "' "
	Else
	    	UNIValue(0,3) = "|"
	End If

	'������ó������ 
	If Trim(Request("rdoPostFlag2")) <> "" Then
		UNIValue(0,4) = " '"& FilterVar(Trim(UCase(Request("rdoPostFlag2"))), " " , "SNM") & "' "
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
