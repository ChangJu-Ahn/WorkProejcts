<%@ LANGUAGE=VBSCript%>
<%
'**********************************************************************************************
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S1411MB1
'*  4. Program Name         : ���Ű����׷��� 
'*  5. Program Desc         : ���Ű����׷��� 
'*  6. Comproxy List        : PS1G111.dll, PS1G112.dll, PS1G113.dll
'*  7. Modified date(First) : 2000/08/05
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : Mr Cho
'* 10. Modifier (Last)      : sonbumyeol
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change"
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              : 2002/11/22 : Grid���� ����, Kang Jun Gu
'*                            2002/12/10 : INCLUDE �ٽ� ���� ����, Kang Jun Gu
'**********************************************************************************************
%>
<% Option Explicit %>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../ComASP/LoadInfTb19029.asp" -->

<%
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

	Call LoadBasisGlobalInf()
    Call LoadInfTB19029B("I","*","NOCOOKIE","MB")
	Call LoadBNumericFormatB("I", "*","NOCOOKIE","MB")
	
    Call HideStatusWnd                                                               '��: Hide Processing message    
    Dim lgOpModeCRUD, lgIntFlgMode
    Dim lgStrData    
        
    lgOpModeCRUD  = Request("txtMode")                                           '��: Read Operation Mode (CRUD)

    Select Case lgOpModeCRUD
        Case CStr(UID_M0001)                                                         '��: Query
             Call SubBizQuery()
        Case CStr(UID_M0002)                                                         '��: Save,Update
             Call SubBizSave()
        Case CStr(UID_M0003)                                                         '��: Delete
             Call SubBizDelete()
    End Select

'============================================================================================================
Sub SubBizQuery()

    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
	Dim objPS1G112	
	Dim iArrRsOut

	'output param ��� 
	Const C_CREDIT_GRP = 0
	Const C_CREDIT_GRP_NM = 1
	Const C_CUR = 2
	Const C_CREDIT_LIMIT_AMT = 3
	Const C_SO_CHK_FLAG = 4
	Const C_DN_CHK_FLAG = 5
	Const C_GI_CHK_FLAG = 6
	Const C_CHK_TYPE = 7
	Const C_CHK_TYPE_NM = 8
	Const C_SO_CHK_TYPE = 9
	Const C_SO_CHK_TYPE_NM = 10
	Const C_UNFAITH_FLAG = 11

    If Request("txtCreditGrp") = "" Then										'��: ��ȸ�� ���� ���� ���Դ��� üũ 
		Call ServerMesgBox("��ȸ ���ǰ��� ����ֽ��ϴ�!", vbInformation, I_MKSCRIPT)              
		Exit Sub 
	End If
	
    '-----------------------
    'Data manipulate  area(import view match)
    '-----------------------
    Set objPS1G112 = Server.CreateObject("PS1G112.CLookupCreditLimitSvr")    
    Call objPS1G112.LookupCreditLimitSvr(gStrGlobalCollection, Trim(Request("txtCreditGrp")), iArrRsOut)

	Set objPS1G112 = Nothing																	'��: ComProxy UnLoad
	If CheckSYSTEMError(Err,True) = True Then
		Response.Write "<Script language=vbs> " & vbCr   
		Response.Write " Parent.frm1.txtCreditGrpNm.value =  """"" & vbCr
		Response.Write " Call Parent.SetToolbar(""1111100000011111"") " & vbCr   
		Response.Write "</Script> "
		Exit Sub	
  	End If

  	Response.Write "<Script language=vbs> " & vbCr   
    Response.Write " With Parent.frm1 " & vbCr
    Response.Write " .txtCreditGrp.value		=  """ & ConvSPChars(iArrRsOut(C_CREDIT_GRP, 0)) & """" & vbCr		' ��ȸ���� 
    Response.Write " .txtCreditGrpNm.value		=  """ & ConvSPChars(iArrRsOut(C_CREDIT_GRP_NM, 0)) & """" & vbCr	' ��ȸ���� 
	Response.Write " .txtCreditGrpCd.value		= """ & ConvSPChars(iArrRsOut(C_CREDIT_GRP, 0)) & """" & vbCr
	Response.Write " .txtCreditGrpName.value	=  """ & ConvSPChars(iArrRsOut(C_CREDIT_GRP_NM, 0)) & """" & vbCr
	Response.Write " .txtCreditLmtAmt.text		= """ & UNINumClientFormat(iArrRsOut(C_CREDIT_LIMIT_AMT, 0), ggAmtOfMoney.DecPoint, 0) & """" & vbCr
	Response.Write " .txtCreditChkCd.value		= """ & ConvSPChars(iArrRsOut(C_CHK_TYPE, 0)) & """" & vbCr
	Response.Write " .txtCreditChkNm.value		= """ & ConvSPChars(iArrRsOut(C_CHK_TYPE_NM, 0)) & """" & vbCr
	Response.Write " .txtCreditSoChkCd.value	= """ & ConvSPChars(iArrRsOut(C_SO_CHK_TYPE, 0)) & """" & vbCr
	Response.Write " .txtCreditSoChkNm.value	= """ & ConvSPChars(iArrRsOut(C_SO_CHK_TYPE_NM, 0)) & """" & vbCr
	Response.Write " .txtHCreditGrp.value		= """ & ConvSPChars(iArrRsOut(C_CREDIT_GRP, 0)) & """" & vbCr
    
    If iArrRsOut(C_SO_CHK_FLAG, 0) = "N" Then
		Response.Write " .rdoSoChkFlag1.checked = True" & vbCr
	ElseIf iArrRsOut(C_SO_CHK_FLAG, 0)= "W" Then
		Response.Write " .rdoSoChkFlag2.checked = True" & vbCr
	ElseIf iArrRsOut(C_SO_CHK_FLAG, 0) = "E" Then
		Response.Write " .rdoSoChkFlag3.checked = True" & vbCr
	End If
				
	If iArrRsOut(C_GI_CHK_FLAG, 0) = "N" Then
		Response.Write " .rdoGiChkFlag1.checked = True" & vbCr
	ElseIf iArrRsOut(C_GI_CHK_FLAG, 0) = "W" Then
		Response.Write " .rdoGiChkFlag2.checked = True" & vbCr
	ElseIf iArrRsOut(C_GI_CHK_FLAG, 0) = "E" Then
		Response.Write " .rdoGiChkFlag3.checked = True" & vbCr
	End If
		
	If iArrRsOut(C_UNFAITH_FLAG, 0) = "Y" Then
		Response.Write " .chkBadCreditFlg.checked = True" & vbCr
	Else
		Response.Write " .chkBadCreditFlg.checked = False" & vbCr
	End If
    Response.Write " End With " & vbCr
	
	Response.Write " Parent.DbQueryOk "	& vbCr																		    	& vbCr	
	Response.Write "</Script>" & vbCr
					
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
    Call SubBizQueryMulti(iArrRsOut(C_SO_CHK_TYPE, 0), iArrRsOut(C_CHK_TYPE, 0), CDbl(iArrRsOut(C_CREDIT_LIMIT_AMT, 0)))
End Sub    

'============================================================================================================
Sub SubBizSave()
	Dim LngRow		                                                                    
	Dim iPS1G111
	Dim iErrorPosition
	Dim iArrSCreditLimit
	Dim iStrCUD
		    
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    If Request("txtCreditGrpCd") = "" Then												'��: ������ ���� ���� ���Դ��� üũ 
		Call ServerMesgBox("TXTFLGMODE ���ǰ��� ����ֽ��ϴ�!", vbInformation, I_MKSCRIPT)              
		Response.End 
	End If
	
	' ���� �׷� ������ ����� ��� 
	If Request("txtCreditLimitChgFalg") = "Y" Then	
		' ���� �׷� ���� 
		Redim iArrSCreditLimit(17)	
			
		' ���ǻ��� : �ε��� ���� �������� ���� 
		iArrSCreditLimit(0) = UCase(Trim(Request("txtCreditGrpCd")))			' ���� ���� �׷� 
	    iArrSCreditLimit(1) = Trim(Request("txtCreditGrpName"))					' ���� ���� �׷�� 
	    iArrSCreditLimit(2) = Trim(Request("txtLocCurrency1"))					' ȭ����� 
	    If Len(Request("txtCreditLmtAmt")) Then									' ���� �ѵ��ݾ� 
			iArrSCreditLimit(3) = UNIConvNum(Request("txtCreditLmtAmt"),0)
		End If
	    iArrSCreditLimit(4) = Trim(Request("rdoSOChkFlag"))						' ���ֽ� üũ��� 
	    iArrSCreditLimit(6) = Trim(Request("rdoGIChkFlag"))						' ���� üũ��� 
	    iArrSCreditLimit(7) = Trim(Request("txtCreditChkCd"))					' ���� ���� üũ 
	    iArrSCreditLimit(8) = Trim(Request("txtCreditSoChkCd"))					' ���ֽ� ���� üũ 
	    iArrSCreditLimit(9) = Trim(Request("txtBadCreditFlg"))					' �νǿ��� �׷� ���� 
	End If

	lgIntFlgMode = CInt(Request("txtFlgMode"))
	
    If lgIntFlgMode = OPMD_CMODE Then
		iStrCUD = "C"
    ElseIf lgIntFlgMode = OPMD_UMODE Then		
		iStrCUD = "U"
    End If

    ' ���� �׷������� ���� �� ��� 
	Set iPS1G111 = Server.CreateObject("PS1G111.CMaintCreditLimitSvr")             
	call iPS1G111.MaintCreditLimitSvr(gStrGlobalCollection, iStrCUD, iArrSCreditLimit , Request("txtSpread"), iErrorPosition)
			
	Set iPS1G111 = Nothing
	    
	If iErrorPosition > 0 Then
		If CheckSYSTEMError2(Err, True, iErrorPosition & "��","","","","") = True Then Exit Sub
	Else
		If CheckSYSTEMError(Err,True) = True Then Exit Sub
	End If
	
	Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DBSaveOk "      & vbCr   
    Response.Write "</Script> "  
			
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'========================================================================================================
Sub SubBizDelete()
	Dim iPS1G111
	Dim iErrorPosition
	Dim iStrCUD
	Dim iArrSCreditLimit
	
    On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status

    '---------- Developer Coding part (Start) ---------------------------------------------------------------
    Redim iArrSCreditLimit(0)
	
    If Request("txtCreditGrp") = "" Then										'��: ������ ���� ���� ���Դ��� üũ 
		Call ServerMesgBox("���� ���ǰ��� ����ֽ��ϴ�!", vbInformation, I_MKSCRIPT)              
		Exit Sub
	End If
    
    iStrCUD = "D"
    iArrSCreditLimit(0) = Trim(Request("txtCreditGrp"))
    
    Set iPS1G111 = Server.CreateObject("PS1G111.CMaintCreditLimitSvr")             
	call iPS1G111.MaintCreditLimitSvr(gStrGlobalCollection, iStrCUD, iArrSCreditLimit, "", iErrorPosition)
	
	Set iPS1G111 = Nothing
    If CheckSYSTEMError(Err,True) = True Then
		Exit Sub
	End If 
    '-----------------------
	'Result data display area
	'----------------------- 
	Response.Write "<Script language=vbs> " & vbCr         
    Response.Write " Parent.DbDeleteOk "      & vbCr   
    Response.Write "</Script> "  
    '---------- Developer Coding part (End  ) ---------------------------------------------------------------
End Sub

'============================================================================================================
' Name : SubBizQueryMulti
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti(byVal pvStrSoChkType,  byVal pvStrGiChkType, byVal pvDblCreditLimitAmt)
	Dim iLngRow	
    Dim iStrNextKey
    Dim iIntSheetMaxRows
    Dim iArrRsOut
    
	Dim iDblSoAmt, iDblDnAmt, iDblGiAmt, iDblExtAmt
	Dim iDblAvailableAmtForSo, iDblAvailableAmtForGi
    Dim iObjPS1G113
    
	'Constant for Detail
	Const C_BP_CD = 0
    Const C_BP_NM = 1    
    Const C_ASGN_AMT_LOC = 2
    Const C_SO_AMT = 3
    Const C_DN_AMT = 4
    Const C_GI_AMT = 5
    Const C_BILL_AMT = 6
    Const C_AR_AMT = 7
    Const C_NOTE_AMT = 8
    Const C_PRRCPT_AMT = 9
    Const C_OVER_DUE_AMT = 10
    
	On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
        
	Set iobjPS1G113 = Server.CreateObject("PS1G113.CListSUsedCreditLimit")
		
	If CheckSYSTEMError(Err,True) = True Then
       Exit Sub
    End If

	' ��� Row�� �� �� ���������� �Ѵ�(����� �������� �� ��)
	iIntSheetMaxRows = 0
	iStrNextKey = ""
    Call iObjPS1G113.ListRows (gStrGlobalCollection, iIntSheetMaxRows, Trim(Request("txtCreditGrp")), iStrNextKey, iArrRsOut)
    
	Set iobjPS1G113 = Nothing																	'��: ComProxy UnLoad	
    
    ' ������ �߻��� ��� �Ǵ� �������� ���� ��� 
	If CheckSYSTEMError(Err,True) Or UBound(iArrRsOut) < 0 Then
		Response.Write("<Script Language = vbscript>" & vbCr)
		Response.Write " Parent.frm1.txtAvailableAmtForSO.text = """ & UNINumClientFormat(pvDblCreditLimitAmt, ggAmtOfMoney.DecPoint, 0) & """" & vbCr
		Response.Write " Parent.frm1.txtAvailableAmtForGI.text = """ & UNINumClientFormat(pvDblCreditLimitAmt, ggAmtOfMoney.DecPoint, 0) & """" & vbCr
		Response.Write("Call Parent.SetToolbar(""1111100000011111"")" & vbCr)
		Response.Write("</Script>" & vbCr)
		Exit Sub
  	End If

	iDblSoAmt = 0
	iDblDnAmt = 0
	iDblGiAmt = 0
	iDblExtAmt = 0
	
	For iLngRow = 0 To Ubound(iArrRsOut, 2) 
		lgstrData = lgstrData & Chr(11) & ConvSPChars(iArrRsOut(C_BP_CD, iLngRow))					'�� 
	    lgstrData = lgstrData & Chr(11) & ConvSPChars(iArrRsOut(C_BP_NM, iLngRow))					' ���� 
	    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(iArrRsOut(C_ASGN_AMT_LOC, iLngRow), ggAmtOfMoney.DecPoint, 0)
	    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(iArrRsOut(C_SO_AMT, iLngRow), ggAmtOfMoney.DecPoint, 0)
	    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(iArrRsOut(C_DN_AMT, iLngRow), ggAmtOfMoney.DecPoint, 0)
	    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(iArrRsOut(C_GI_AMT, iLngRow), ggAmtOfMoney.DecPoint, 0)
	    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(iArrRsOut(C_BILL_AMT, iLngRow), ggAmtOfMoney.DecPoint, 0)
	    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(iArrRsOut(C_AR_AMT, iLngRow), ggAmtOfMoney.DecPoint, 0)
	    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(iArrRsOut(C_NOTE_AMT, iLngRow), ggAmtOfMoney.DecPoint, 0)
	    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(iArrRsOut(C_PRRCPT_AMT, iLngRow), ggAmtOfMoney.DecPoint, 0)
	    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(iArrRsOut(C_OVER_DUE_AMT, iLngRow), ggAmtOfMoney.DecPoint, 0)
	    lgstrData = lgstrData & Chr(11) & UNINumClientFormat(iArrRsOut(C_OVER_DUE_AMT, iLngRow), ggAmtOfMoney.DecPoint, 0)
	    lgstrData = lgstrData & Chr(11) & iLngRow + 1
	    lgstrData = lgstrData & Chr(11) & Chr(12)    
	    
	    iDblSoAmt = iDblSoAmt + CDbl(iArrRsOut(C_SO_AMT, iLngRow))
	    iDblDnAmt = iDblDnAmt + CDbl(iArrRsOut(C_DN_AMT, iLngRow))
	    iDblGiAmt = iDblGiAmt + CDbl(iArrRsOut(C_GI_AMT, iLngRow))
	    iDblExtAmt = iDblExtAmt + CDbl(iArrRsOut(C_BILL_AMT, iLngRow)) _
							+ CDbl(iArrRsOut(C_AR_AMT, iLngRow)) _
							+ CDbl(iArrRsOut(C_NOTE_AMT, iLngRow)) _
							+ CDbl(iArrRsOut(C_OVER_DUE_AMT, iLngRow)) _
							- CDbl(iArrRsOut(C_PRRCPT_AMT, iLngRow))
	Next

	' ���ֽ� Check�� ���� ���Ż�밡�ɱݾ� ��� 
	Select Case UCase(pvStrSoChkType)
		Case "SO"
			iDblAvailableAmtForSo = pvDblCreditLimitAmt - (iDblSoAmt + iDblDnAmt + iDblGiAmt + iDblExtAmt)
		Case "DN"
			iDblAvailableAmtForSo = pvDblCreditLimitAmt - (iDblDnAmt + iDblGiAmt + iDblExtAmt)
		Case "GI"
			iDblAvailableAmtForSo = pvDblCreditLimitAmt - (iDblGiAmt + iDblExtAmt)
	End Select
	
	' ���� Check�� ���� ���Ż�밡�ɱݾ� ��� 
	Select Case UCase(pvStrGiChkType)
		Case "SO"
			iDblAvailableAmtForGi = pvDblCreditLimitAmt - (iDblSoAmt + iDblDnAmt + iDblGiAmt + iDblExtAmt)
		Case "DN"
			iDblAvailableAmtForGi = pvDblCreditLimitAmt - (iDblDnAmt + iDblGiAmt + iDblExtAmt)
		Case "GI"
			iDblAvailableAmtForGi = pvDblCreditLimitAmt - (iDblGiAmt + iDblExtAmt)
	End Select

   	Response.Write "<Script language=vbs> " & vbCr 
   	Response.Write "With parent " & vbCr 
   	Response.Write " .frm1.txtAvailableAmtForSo.text = """ & UNINumClientFormat(iDblAvailableAmtForSo, ggAmtOfMoney.DecPoint, 0) & """" & vbCr
	Response.Write " .frm1.txtAvailableAmtForGI.text = """ & UNINumClientFormat(iDblAvailableAmtForGi, ggAmtOfMoney.DecPoint, 0) & """" & vbCr
    Response.Write " .ggoSpread.Source = .frm1.vspdData" & vbCr
    Response.Write " .ggoSpread.SSShowDataByClip """ & lgstrData & """" & vbCr
    Response.Write " .lgStrPrevKey = """ & iStrNextKey & """" & vbCr  
    Response.Write " .frm1.vspdData.ReDraw = False " & vbCr  
    Response.Write " .SetSpreadColor -1, -1 " & vbCr     
    Response.Write " .frm1.vspdData.ReDraw = True " & vbCr  
    Response.Write " .DbQueryOk " & vbCr    
    Response.Write "End With " & vbCr 
   	Response.Write "</Script> "
End Sub    

%>

