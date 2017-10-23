<%@ LANGUAGE=VBSCript %>
<%'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b3b22mb2.asp 
'*  4. Program Name         : Called By B3B22MA1 (Class Management)
'*  5. Program Desc         : Manage Class Information
'*  6. Modified date(First) : 2003/02/06
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : Ryu Sung Won
'*  9. Modifier (Last)      : 
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<%
Call LoadBasisGlobalInf

Dim oPB3S222

Dim strMode													'�� : Lookup �� �ڵ� ���� ���� 
Dim lgIntFlgMode
Dim iCommandSent

Dim ADF														'ActiveX Data Factory ���� �������� 
Dim strRetMsg												'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag					'DBAgent Parameter ���� 
Dim rs0, rs1												'DBAgent Parameter ���� 

Dim strClassCd
Dim strClassNm
Dim strCharCd1
Dim strCharCd2
Dim strCharMgr

Dim I1_B_Class

ReDim I1_B_Class(4)
Const C_I1_Class_Cd = 0
Const C_I1_Class_Nm = 1
Const C_I1_Char_Cd1 = 2
Const C_I1_Char_Cd2 = 3
Const C_I1_Char_Mgr = 4

Call HideStatusWnd

On Error Resume Next
Err.Clear  

	strMode = Request("txtMode")							'�� : ���� ���¸� ���� 
	lgIntFlgMode = CInt(Request("txtFlgMode"))				'��: ����� Create/Update �Ǻ� 
	
	strClassCd = Request("txtClassCd1")
	strClassNm = Request("txtClassNm1")
	strCharCd1 = Request("txtCharCd1")
	strCharCd2 = Request("txtCharCd2")
	strCharMgr = Request("cboClassMgr")
	
	If lgIntFlgMode = OPMD_CMODE Then
		iCommandSent = "CREATE"
    ElseIf lgIntFlgMode = OPMD_UMODE Then
		iCommandSent = "UPDATE"
    End If

	'--------------------------------------------
	' CHAR_CD�� �����ϴ��� üũ 
	'--------------------------------------------
	Redim UNISqlId(1)
	Redim UNIValue(1, 0)

	UNISqlId(0) = "b3b28mb1a"
	UNISqlId(1) = "b3b28mb1a"
	
	UNIValue(0, 0) = " " & FilterVar(UCase(strCharCd1), "''", "S") & " "
	UNIValue(1, 0) = " " & FilterVar(UCase(strCharCd2), "''", "S") & " "

	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	Set ADF = Server.CreateObject("prjPublic.cCtlTake")
	strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1)

	'------------------------------------------------------------
	' ����׸�(Char_Cd1) üũ 
	'------------------------------------------------------------
	If (rs0.EOF And rs0.BOF) Then
		Call DisplayMsgBox("122630", vbOKOnly, "", "", I_MKSCRIPT)	'/// ����׸��� �������� �ʽ��ϴ�.
		rs0.Close
		Set rs0 = Nothing
		Set ADF = Nothing
		Response.Write "<Script Language=vbscript>			" & vbCr
		Response.Write "	parent.frm1.txtCharCd1.focus()	" & vbCr																
		Response.Write "</Script>							" & vbCr
		Response.End
	End If
	'------------------------------------------------------------
	' ����׸�(Char_Cd2) üũ 
	'------------------------------------------------------------
	If Trim(strCharCd2) <> "" Then
		If (rs1.EOF And rs1.BOF) Then
			Call DisplayMsgBox("122630", vbOKOnly, "", "", I_MKSCRIPT)	'/// ����׸��� �������� �ʽ��ϴ�.
			rs1.Close
			Set rs1 = Nothing
			Set ADF = Nothing
			Response.Write "<Script Language=vbscript>			" & vbCr
			Response.Write "	parent.frm1.txtCharCd2.focus()	" & vbCr																
			Response.Write "</Script>							" & vbCr
			Response.End
		End If
	End If
	'------------------------------------------------------------
	' ���� ����׸����� üũ 
	'------------------------------------------------------------
	If UCase(Trim(strCharCd1)) = UCase(Trim(strCharCd2)) Then
		Call DisplayMsgBox("122656", vbOKOnly, "", "", I_MKSCRIPT)	'/// ����׸�1,2�� �����մϴ�.
		Response.Write "<Script Language=vbscript>			" & vbCr
		Response.Write "	parent.frm1.txtCharCd2.focus()	" & vbCr																
		Response.Write "</Script>" & vbCr
		Response.End
	End If
	
	I1_B_Class(C_I1_Class_Cd) = UCase(Trim(strClassCd))
	I1_B_Class(C_I1_Class_Nm) = strClassNm
	I1_B_Class(C_I1_Char_Cd1) = UCase(Trim(strCharCd1))
	I1_B_Class(C_I1_Char_Cd2) = UCase(Trim(strCharCd2))
	I1_B_Class(C_I1_Char_Mgr) = UCase(Trim(strCharMgr))

	Set oPB3S222 = Server.CreateObject("PB3S222.cBMngClass")

    If CheckSYSTEMError(Err,True) = True Then
       Response.End 
    End If

    Call oPB3S222.B_MANAGE_CLASS(gStrGlobalCollection, _
								iCommandSent, _
								I1_B_Class)

    If CheckSYSTEMError(Err,True) = True Then
       Set oPB3S222 = Nothing		                                                 '��: Unload Comproxy DLL
       Response.End		
    End If
	
	Set oPB3S222 = Nothing
	
	'-----------------------
	'Result data display area
	'----------------------- 
	Response.Write "<Script Language=vbscript>	" & vbCr
	Response.Write "	With parent				" & vbCr																
	Response.Write "		.DbSaveOk			" & vbCr
	Response.Write "	End With				" & vbCr
	Response.Write "</Script>					" & vbCr
					

	Response.End																				'��: Process End

	'==============================================================================
	' ����� ���� ���� �Լ� 
	'==============================================================================

	Response.Write "<SCRIPT LANGUAGE=VBSCRIPT RUNAT=SERVER>	" & vbCr
	Response.Write "										" & vbCr
	Response.Write "</SCRIPT>								" & vbCr

	'==============================================================================
	' Function : IsVaildCodeLength (User Defined Function)
	' Description : ���ڿ��� ���̸� ���ϴ� ���̿� ���Ͽ� �۰ų� ������ True ���� 
	'==============================================================================
	Function IsVaildCodeLength(Byval iStr, Byval iDigit)
		Dim intLength
		Dim intIdx
		Dim intAsc
		Dim intSum
		
		IsVaildCodeLength = True
		
		intSum = 0
		intLength = Len(iStr)
		
		For intIdx=0 To intLength-1
			intAsc = ASC(Mid(iStr,intIdx+1,1))
			If CInt(intAsc) < 0 Or CInt(intAsc) > 255 Then
				intSum = intSum + 2
			Else
				intSum = intSum + 1
			End If
		Next
		
		If intSum > CInt(iDigit) Then
			IsVaildCodeLength = False
		End If
	End Function
%>
