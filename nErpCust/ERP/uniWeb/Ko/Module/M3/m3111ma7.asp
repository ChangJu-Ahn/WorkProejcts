<%@ LANGUAGE="VBSCRIPT" %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : m3111ma7
'*  4. Program Name         : ��ǰ���ֵ�� 
'*  5. Program Desc         : ��ǰ���ֵ�� 
'*  6. Component List       : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2003/05/23
'*  9. Modifier (First)     : Min, HJ
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(��) means that "Do not change" 
'*                            this mark(��) Means that "may  change"
'*                            this mark(��) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--'#########################################################################################################
'												1. �� �� �� 
'##########################################################################################################!-->
<!-- '******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'********************************************************************************************************* !-->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'==========================================================================================================!-->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'==========================================================================================================!-->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs"></SCRIPT> 

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit          

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
Const TAB1 = 1									

'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
Const BIZ_PGM_ID 					= "m3111mb7.asp"											
Const BIZ_PGM_JUMP_ID_PO_DTL 		= "M3112MA7"
'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'=========================================================================================================
Dim lgBlnFlgChgValue				'��: Variable is for Dirty flag
Dim lgIntGrpCount				'��: Group View Size�� ������ ���� 
Dim lgIntFlgMode					'��: Variable is for Operation Status

Dim lgNextNo						'��: ȭ���� Single/SingleMulti �ΰ�츸 �ش� 
Dim lgPrevNo						' ""
'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 
'----------------  ���� Global ������ ����  ----------------------------------------------------------- 

'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++
Dim lgMpsFirmDate, lgLlcGivenDt								
Dim gSelframeFlg

Dim cboOldVal          
Dim IsOpenPop          
Dim lblnWinEvent
Dim lgCboKeyPress      
Dim lgOldIndex								
Dim lgOldIndex2 
Dim lgOpenFlag  
Dim lgTabClickFlag  
Dim arrCollectVatType
Dim StartDate
Dim EndDate
Dim iDBSYSDate

iDBSYSDate = "<%=GetSvrDate%>"
EndDate   = UniConvDateAToB(iDBSYSDate  ,Parent.gServerDateFormat,Parent.gDateFormat)

'==========================================================================================
'   Event Name : SupplierLookUp
'   Event Desc : 
'==========================================================================================
Function SupplierLookup()
    Err.Clear                                          
    
    If CheckRunningBizProcess = True Then
		Exit Function
	End If
    
    Dim strVal
    
    if Trim(frm1.txtSupplierCd.Value) = "" then
    	Exit Function
    End if
    
	lgBlnFlgChgValue = true
	
    strVal = BIZ_PGM_ID & "?txtMode=" & "SupplierLookupAfterOnline"			
    strVal = strVal & "&txtSupplierCd=" & Trim(frm1.txtSupplierCd.value)	
    
    If LayerShowHide(1) = False Then Exit Function
    
	Call RunMyBizASP(MyBizASP, strVal)										
End Function
'==========================================================================================
'   Event Name : ChangePotype
'   Event Desc : txtPotypeCd Chagne Event
'==========================================================================================
Sub ChangePotype()
	frm1.hdnpotype.value = ""
	If gLookUpEnable = False Then
		Exit Sub
	End If
	
	Call PotypeRef()
End Sub

'==========================================================================================
'   Event Name : ChangeSupplier
'   Event Desc : txtSupplierCd Chagne Event
'==========================================================================================
Sub ChangeSupplier()
	If gLookUpEnable = False Then
		Exit Sub
	End If

	Call SpplRef()
End Sub

'==========================================   PotypeRef()  ======================================
'	Name : PotypeRef()
'	Description : 
'=========================================================================================================
Sub PotypeRef()
    Err.Clear
    
    If CheckRunningBizProcess = True Then
		Exit Sub
	End If
    
    Dim strVal
    
    if Trim(frm1.txtPotypeCd.Value) = "" then
		Call DisplayMsgBox("205152", "X", "��ǰ����", "X")
    	Exit Sub
    End if
    
	if lgIntFlgMode <> Parent.OPMD_UMODE Then 
		lgBlnFlgChgValue = true
	end if		
	
    strVal = BIZ_PGM_ID & "?txtMode=" & "LookUpPoType"			
    strVal = strVal & "&txtPoTypeCd=" & Trim(frm1.txtPoTypeCd.value)
    strVal = strVal & "&txtTabClickFlag=" & lgTabClickFlag

    If LayerShowHide(1) = False Then Exit Sub
    
	Call RunMyBizASP(MyBizASP, strVal)								
End Sub

'==========================================   SpplRef()  ======================================
'	Name : SpplRef()
'	Description : It is Call at txtSupplier Change Event
'=========================================================================================================
Sub SpplRef()
    Err.Clear
    
    Dim strVal
    
    if Trim(frm1.txtSupplierCd.Value) = "" then
    	Exit Sub
    End if
    
	lgBlnFlgChgValue = true
	
    strVal = BIZ_PGM_ID & "?txtMode=" & "LookUpSupplier"			
    strVal = strVal & "&txtSupplierCd=" & Trim(frm1.txtSupplierCd.value)	
	
    If LayerShowHide(1) = False Then Exit Sub
    
	Call RunMyBizASP(MyBizASP, strVal)										
End Sub
'==========================================   Cfm()  ======================================
'	Name : Cfm()
'	Description : Ȯ����ư,Ȯ����ҹ�ư�� Event �ռ� 
'=========================================================================================================
Sub Cfm()
    Dim IntRetCD 
    
    Err.Clear                                                               
    
    if lgBlnFlgChgValue = True	then
		Call DisplayMsgBox("189217", "X", "X", "X")
		Exit sub
	End if
	
	if frm1.rdoRelease(0).checked = True then
		IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Sub
		End If
		Call DbSave("Cfm")				                                    
					                                                 
	elseif frm1.rdoRelease(1).checked = True then
		IntRetCD = DisplayMsgBox("900018", Parent.VB_YES_NO, "x", "x")
		If IntRetCD = vbNo Then
			Exit Sub
		End If
		Call DbSave("UnCfm")
	End if
End Sub

'-------------------------------------------------------------------
'		Ȯ�����ο� ���� Field�� �Ӽ��� Protect�� ��ȯ,���� ��Ű�� �Լ� 
'--------------------------------------------------------------------
Function ChangeTag(Byval Changeflg)
	with frm1
		if Changeflg = true then
			'ù��° Tab
			ggoOper.SetReqAttr	.txtPoTypeCd, "Q"
			ggoOper.SetReqAttr	.txtPoDt, "Q"
			ggoOper.SetReqAttr	.txtGroupCd, "Q"
			ggoOper.SetReqAttr	.txtSupplierCd, "Q"
			ggoOper.SetReqAttr	.txtCurr, "Q"
			ggoOper.SetReqAttr	.txtSuppSalePrsn, "Q"
			ggoOper.SetReqAttr	.txtTel, "Q"
			ggoOper.SetReqAttr	.txtRemark, "Q"
		else
			'ù��° Tab
			ggoOper.SetReqAttr	.txtPoNo2, "D"
			ggoOper.SetReqAttr	.txtPoTypeCd, "N"
			ggoOper.SetReqAttr	.txtPoDt, "N"
			ggoOper.SetReqAttr	.txtGroupCd, "N"
			ggoOper.SetReqAttr	.txtSupplierCd, "N"
			ggoOper.SetReqAttr	.txtCurr, "N"
			ggoOper.SetReqAttr	.txtSuppSalePrsn, "D"
			ggoOper.SetReqAttr	.txtTel, "D"
			ggoOper.SetReqAttr	.txtRemark, "D"
		End if 
	end with
End Function 

'--------------------------------------------------------------------
'		Cookie ����Լ� 
'--------------------------------------------------------------------
Function CookiePage(Byval Kubun)
	Dim strTemp, arrVal
	Dim IntRetCD

	If Kubun = 0 Then
		strTemp = ReadCookie("PoNo")
		
		If strTemp = "" then Exit Function
		
		frm1.txtPoNo.value = strTemp
		
		WriteCookie "PoNo" , ""
		
		Call MainQuery()
	elseIf Kubun = 1 Then
	    If lgIntFlgMode <> Parent.OPMD_UMODE Then                           
	        Call DisplayMsgBox("900002", "X", "X", "X")
	        Exit Function
	    End If
		
	    If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If
    	
		WriteCookie "PoNo" , frm1.txtPoNo.value
		
		Call PgmJump(BIZ_PGM_JUMP_ID_PO_DTL)
	elseIf Kubun = 2 Then
	    If lgIntFlgMode <> Parent.OPMD_UMODE Then                           
	        Call DisplayMsgBox("900002", "X", "X", "X")
	        Exit Function
	    End If
	    	
	    If lgBlnFlgChgValue = True Then
			IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
			If IntRetCD = vbNo Then
				Exit Function
			End If
	    End If
    	
	    WriteCookie "Process_Step" , "PO"
		WriteCookie "Po_No" , Trim(frm1.txtPoNo.value)
		WriteCookie "Pur_Grp", Trim(frm1.txtGroupCd.Value)
		WriteCookie "Po_Cur", Trim(frm1.txtCurr.Value)
		WriteCookie "Po_Xch", Trim(frm1.txtXch.Value)
		
		Call PgmJump(BIZ_PGM_JUMP_ID_PUR_CHARGE)
	End IF
End Function

'------------------------------------------------------------------------------------------
'Radio���� Click�� �� ��� flag�� Setting
'------------------------------------------------------------------------------------------
Sub Setchangeflg()
	lgBlnFlgChgValue = True	
End Sub

'------------------------------------------------------------------------------------------
'����ڰ� Radio Button�� Click�� �� ���� ������ hdnRelease�� Setting
'------------------------------------------------------------------------------------------
Sub Changeflg()
	if frm1.rdoRelease(0).checked = true then
		frm1.hdnRelease.value= "N"
	else
		frm1.hdnRelease.value= "Y"
	end if 
End Sub

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'=========================================================================================================
Sub InitVariables()
    lgIntFlgMode = Parent.OPMD_CMODE                                        
    lgBlnFlgChgValue = False                                         
    lgIntGrpCount = 0                                                
    IsOpenPop = False												
    lgCboKeyPress = False
    lgOldIndex = -1
    lgOldIndex2 = -1
    lgMpsFirmDate=""
    lgLlcGivenDt=""
End Sub

'=============================== 2.1.2 LoadInfTB19029() ========================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<% Call loadInfTB19029A("I", "*","NOCOOKIE","MA") %>
<% Call LoadBNumericFormatA("I", "*","NOCOOKIE","MA") %>
End Sub

'==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'=========================================================================================================
Sub SetDefaultVal()
	lgOpenFlag	= False
	lgTabClickFlag	= False
    Call SetToolbar("1110100000001111")
    frm1.rdoRelease(0).checked = true
    frm1.txtPoDt.text = EndDate
    frm1.hdnCurr.value = Parent.gCurrency   
    frm1.btnCfm.disabled = true
    frm1.btnSend.disabled = true
    frm1.txtGroupCd.Value = Parent.gPurGrp
	frm1.btnCfm.value = "Ȯ��"
	frm1.hdnMergPurFlg.Value = "N"
	frm1.txtPoNo.focus
	Set gActiveElement = document.activeElement
End Sub

'------------------------------------------  SetClickflag, ResetClickflag()  -----------------------------
'	Name : SetClickflag, ResetClickflag()
'	Description :  
'---------------------------------------------------------------------------------------------------------
Function SetClickflag()
	lgTabClickFlag = True	
End Function

Function ResetClickflag()
	lgTabClickFlag = False
End Function

'------------------------------------------  OpenPoRef()  -------------------------------------------------
'	Name : OpenPoRef()
'	Description : OpenPoRef PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPoRef()
	Dim strRet
	Dim arrParam(12)
	Dim iCalledAspName
	Dim IntRetCD
	Dim strVal

	If lgIntFlgMode = Parent.OPMD_UMODE then
		Call DisplayMsgBox("17A012", "X","�űԵ���� �ƴ� ���","��������" )
		Exit Function
	End if 

	If Trim(frm1.txtPotypeCd.Value) = "" then
		Call DisplayMsgBox("17A002","X" , "��ǰ����","X")
		frm1.txtPotypeCd.focus
		Set gActiveElement = document.activeElement
		Exit Function
	End if

	If frm1.hdnpotype.value = "" Then 'park.j.k
		Call DisplayMsgBox("17A003","X" , "��ǰ����","X")
		frm1.txtPotypeCd.focus
		Exit Function
	End If 
	
	if frm1.hdnRelease.Value = "Y" then
		Call DisplayMsgBox("17a008", "X", "X", "X")
		frm1.txtPotypeCd.focus
		Exit Function
	End if
	
'	if UCase(frm1.hdnRetflg.Value) = "Y" and UCase(frm1.hdnRcptflg.Value) = "Y" then
'		Call DisplayMsgBox("17A012", "X","��������" & frm1.txtPotypeCd.Value & "(" & frm1.txtPoTypeNm.value & ")","��������" )
'		frm1.txtPotypeCd.focus
'		Exit Function
'	End if
	
	If lblnWinEvent = True Then Exit Function
		
	lblnWinEvent = True
	
	arrParam(0) = Trim(frm1.hdnRcptflg.value)
	arrParam(1) = Trim(frm1.hdnIvflg.value)
	arrParam(2) = Trim(frm1.hdnSubcontraflg.value) 	'���ְ������� �߰� 
	
	iCalledAspName = AskPRAspName("M3111PA9")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3111PA9", "X")
		lblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
	
	lblnWinEvent = False
	If strRet(0) = "" Then
		Exit Function
	Else	
		frm1.txtPoNo1.value = UCase(Trim(strRet(0)))
		
		if LayerShowHide(1) =false then
		    exit Function
		end if

		strVal = BIZ_PGM_ID & "?txtMode=" & "afterPORef"	
		strVal = strVal & "&txtPoNo=" & UCase(Trim(strRet(0)))
	
		Call RunMyBizASP(MyBizASP, strVal)	

		frm1.txtPoNo2.focus
		Set gActiveElement = document.activeElement
	End If	
			
End Function

'------------------------------------------  OpenPoNo()  -------------------------------------------------
'	Name : OpenPoNo()
'	Description : OpenPoNo PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPoNo()
	Dim strRet
	Dim iCalledAspName
	Dim IntRetCD
		
	If lblnWinEvent = True Or UCase(frm1.txtPoNo.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function
		
	lblnWinEvent = True
		
	iCalledAspName = AskPRAspName("M3111PA7")
	
	If Trim(iCalledAspName) = "" Then
		IntRetCD = DisplayMsgBox("900040", parent.VB_INFORMATION, "M3111PA7", "X")
		lblnWinEvent = False
		Exit Function
	End If
	
	strRet = window.showModalDialog(iCalledAspName, Array(window.parent,"Y"), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")

	lblnWinEvent = False
	
	If strRet(0) = "" Then
		frm1.txtPoNo.focus
		Exit Function
	Else
		frm1.txtPoNo.value = strRet(0)
		frm1.txtPoNo.focus
	End If	
End Function

'------------------------------------------  OpenPoType()  -------------------------------------------------
'	Name : OpenPoType()
'	Description : OpenPoType PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenPotype()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtPotypeCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "��ǰ����"					
	arrParam(1) = "M_CONFIG_PROCESS"			
	
	arrParam(2) = Trim(frm1.txtPotypeCd.Value)	
	'arrParam(3) = Trim(frm1.txtPotypeNm.Value)	
	
	arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") & "  and Ret_FLG = " & FilterVar("Y", "''", "S") & " "							
	arrParam(5) = "��ǰ����"					
	
    arrField(0) = "PO_TYPE_CD"					
    arrField(1) = "PO_TYPE_NM"					
    
    arrHeader(0) = "��ǰ����"				
    arrHeader(1) = "��ǰ���¸�"				
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtPoTypeCd.focus
		Exit Function
	Else
		frm1.txtPoTypeCd.Value    = arrRet(0)		
		frm1.txtPoTypeNm.Value    = arrRet(1)
		lgBlnFlgChgValue = True

		Call PotypeRef()
		frm1.txtPoTypeCd.focus
		Set gActiveElement = document.activeElement
	End If	
End Function

'------------------------------------------  OpenCurr()  -------------------------------------------------
'	Name : OpenCurr()
'	Description : OpenCurr PopUp
'---------------------------------------------------------------------------------------------------------
Function OpenCurr()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtCurr.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "ȭ��"						
	arrParam(1) = "B_Currency"					
	
	arrParam(2) = Trim(frm1.txtCurr.Value)	
	'arrParam(3) = Trim(frm1.txtItemNm2.Value)	
		
	arrParam(4) = ""							
	arrParam(5) = "ȭ��"						
	
    arrField(0) = "Currency"					
    arrField(1) = "Currency_Desc"				
    
    arrHeader(0) = "ȭ��"					
    arrHeader(1) = "ȭ���"					
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtCurr.focus
		Exit Function
	Else
		frm1.txtCurr.Value    = arrRet(0)		
		frm1.txtCurrNm.Value  = arrRet(1)		
		Call ChangeCurr()
		lgBlnFlgChgValue = True
		frm1.txtCurr.focus
		Set gActiveElement = document.activeElement
	End If	
End Function

Sub ChangeCurr()
	if UCase(Trim(frm1.txtCurr.value)) = UCase(Parent.gCurrency) then
		frm1.txtXch.value = 1
		Call ggoOper.SetReqAttr(frm1.txtXch,"Q")
	else
		Call ggoOper.SetReqAttr(frm1.txtXch,"D")
		frm1.txtXch.value = 0
	end if 
	
End Sub

Function OpenSupplier()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True  Or UCase(frm1.txtSupplierCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "����ó"					
	arrParam(1) = "B_BIZ_PARTNER"				

	arrParam(2) = Trim(frm1.txtSupplierCd.Value)
	'arrParam(3) = Trim(frm1.txtSupplierNm.Value)
	
	arrParam(4) = "BP_TYPE In (" & FilterVar("S", "''", "S") & " ," & FilterVar("CS", "''", "S") & ") And usage_flag=" & FilterVar("Y", "''", "S") & "  AND IN_OUT_FLAG = " & FilterVar("O", "''", "S") & " "		
	arrParam(5) = "����ó"						
	
    arrField(0) = "BP_Cd"					
    arrField(1) = "BP_NM"					
    
    arrHeader(0) = "����ó"				
    arrHeader(1) = "����ó��"			
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
		frm1.txtSupplierCd.focus
		Exit Function
	Else
		frm1.txtSupplierCd.Value    = arrRet(0)		
		frm1.txtSupplierNm.Value    = arrRet(1)		
		lgBlnFlgChgValue = True
		Call SpplRef()
		frm1.txtSupplierCd.focus
	End If	
End Function

Function OpenGroup()
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Or UCase(frm1.txtGroupCd.className) = UCase(Parent.UCN_PROTECTED) Then Exit Function

	IsOpenPop = True

	arrParam(0) = "���ű׷�"	
	arrParam(1) = "B_Pur_Grp"				
	
	arrParam(2) = Trim(frm1.txtGroupCd.Value)
'	arrParam(3) = Trim(frm1.txtGroupNm.Value)	
	
	arrParam(4) = "USAGE_FLG=" & FilterVar("Y", "''", "S") & " "			
	arrParam(5) = "���ű׷�"			
	
    arrField(0) = "PUR_GRP"	
    arrField(1) = "PUR_GRP_NM"	
    
    arrHeader(0) = "���ű׷�"		
    arrHeader(1) = "���ű׷��"
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtGroupCd.focus
		Exit Function
	Else
		frm1.txtGroupCd.Value= arrRet(0)
		frm1.txtGroupNm.Value= arrRet(1)
		lgBlnFlgChgValue = True
		frm1.txtGroupCd.focus
	End If	
End Function 

'===================================== CurFormatNumericOCX()  =======================================
'	Name : CurFormatNumericOCX()
'	Description : ȭ�鿡�� �ϰ������� Rounding�Ǵ� Numeric OCX
'====================================================================================================
Sub CurFormatNumericOCX()

	With frm1
		ggoOper.FormatFieldByObjectOfCur .txtPoAmt, .txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
	    ggoOper.FormatFieldByObjectOfCur .txtNetPoAmt, .txtCurr.value, Parent.ggAmtOfMoneyNo, gBCurrency, gBDataType, gBDecimals, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec
	End With

End Sub

'========================================================================================
' Function Name : Sending
' Function Desc : 
'========================================================================================
Function Sending()
    Err.Clear                                           
    
    Sending = False                                     
    
	If LayerShowHide(1) = False Then Exit Function

    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & "SendingB2B"	
    strVal = strVal & "&txtPoNo=" & Trim(frm1.txtPoNo.value)		
    
	Call RunMyBizASP(MyBizASP, strVal)								
	
    Sending = True                                                  
End Function

Function SendingOK()
	'msgbox "������ �Ϸ� �Ǿ����ϴ�."	
End Function

'==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'=========================================================================================================
Sub Form_Load()
	Call LoadInfTB19029
    Call AppendNumberRange("0","0","999")
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
'    Call ggoOper.FormatField(Document, "3",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec,,,ggStrMinPart,ggStrMaxPart)
    Call ggoOper.LockField(Document, "N")							
    Call SetDefaultVal
    Call InitVariables
    '----------  Coding part  -------------------------------------------------------------
    Call Changeflg
    Call CookiePage(0)
    
	gIsTab     = "Y"
	gTabMaxCnt = 2
End Sub

'==========================================================================================
'   Event Name : OCX Event
'   Event Desc :
'==========================================================================================
Sub txtPoDt_DblClick(Button)
	if Button = 1 then
		frm1.txtPoDt.Action = 7
	    Call SetFocusToDocument("M")  
        frm1.txtPoDt.Focus
	End if
	'-- Modify for issue 8875 by Byun Jee Hyun 2004-11-18
	Call ChangeCurr()
	'-- End of issue 8875
End Sub

Sub txtPoDt_Change()
	lgBlnFlgChgValue = true	
	'-- Modify for issue 8875 by Byun Jee Hyun 2004-11-18
	Call ChangeCurr()
	'-- End of issue 8875
End Sub

'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================
Function FncQuery() 
    Dim IntRetCD 
    
    FncQuery = False                                                
    
    Err.Clear                                                       

    If lgBlnFlgChgValue = True Then
    	
		IntRetCD = DisplayMsgBox("900013", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "2")							
    Call InitVariables												

    If Not chkField(Document, "1") Then								
        If gPageNo > 0 Then
            gSelframeFlg = gPageNo
        End If
            
        Exit Function
    End If 

    If DbQuery = False Then Exit Function
    Call Changeflg       
      
	Set gActiveElement = document.activeElement
    FncQuery = True													
End Function

'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False                                                  
    
    If lgBlnFlgChgValue = True Then
       IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call ggoOper.ClearField(Document, "1")                          
    Call ggoOper.ClearField(Document, "2")                          
    Call ggoOper.ClearField(Document, "3")                          
    Call ggoOper.LockField(Document, "N")                              
    Call ChangeTag(False)
    Call SetDefaultVal
    Call InitVariables													

    frm1.txtPoNo.focus
	Set gActiveElement = document.activeElement
    
    FncNew = True														
End Function

'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================
Function FncDelete() 
	Dim IntRetCD

    FncDelete = False
    
    IntRetCD = DisplayMsgBox("900003", Parent.VB_YES_NO, "X", "X")
    
    If IntRetCD = vbNo Then Exit Function
    														
    
    If lgIntFlgMode <> Parent.OPMD_UMODE Then                      
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If
    
    If DbDelete = False Then Exit Function
    
	Set gActiveElement = document.activeElement
    FncDelete = True                                        
End Function

Function FncSave() 
    Dim IntRetCD 
    
    FncSave = False                                         
    
    Err.Clear
    
    If CheckRunningBizProcess = True Then
		Exit Function
	End If                                               
    
    If lgBlnFlgChgValue = False Then
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")
        Exit Function
    End If
    
    if Trim(frm1.hdnPoDt.value) <> "" then
        if (UniConvDateToYYYYMMDD(frm1.hdnPoDt.value,Parent.gDateFormat,"") > Parent.UniConvDateToYYYYMMDD(frm1.txtPoDt.text,Parent.gDateFormat,"")) then	
		    Call DisplayMsgBox("972001", "X","�����", "�������ֹ�ȣ���ֵ����")			
		    Exit Function
	    End if  
    end if

	IF frm1.hdnImportflg.value="Y" then
	
	    If Not chkField(Document, "2") Then                 
	        If gPageNo > 0 Then
	            gSelframeFlg = gPageNo
	        End If
	        
	        Exit Function
	    End If

	    If Not chkField(Document, "3") Then                            
	        If gPageNo > 0 Then
	            gSelframeFlg = gPageNo
	        End If
	        
	        Exit Function
	    End If

	else
	    If Not chkField(Document, "2") Then                            
	        If gPageNo > 0 Then
	            gSelframeFlg = gPageNo
	        End If
	        
	        Exit Function
	    End If

	End if
	
	Call Changeflg
    
	If DbSave("ToolBar") = False Then Exit Function
    
	Set gActiveElement = document.activeElement
    FncSave = True                                                     
End Function

'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================
Function FncCopy() 
	Dim IntRetCD
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900017", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    lgIntFlgMode = Parent.OPMD_CMODE											
  
    <% ' ���Ǻ� �ʵ带 �����Ѵ�. %>
    Call ggoOper.ClearField(Document, "1")                              
    Call ggoOper.LockField(Document, "N")								
    Call Changeflg
    Call ChangeTag(False)
    
    frm1.rdoRelease(0).checked = true
    Call SetToolbar("11101000000111")

	'frm1.txtPoAmt.Text		= UniNumClientFormat(0,ggAmtOfMoney.DecPoint,0)
	'frm1.txtPoLocAmt.Text	= UniNumClientFormat(0,ggAmtOfMoney.DecPoint,0)
	frm1.txtPoDt.text 		= EndDate
	frm1.txtNetPoAmt.Text	= 0
	frm1.txtNetPoLocAmt.Text = 0
	frm1.txtPoAmt.Text		= 0
	frm1.txtPoLocAmt.Text	= 0
	frm1.txtPoNo2.value 	= ""
	frm1.btnCfm.disabled 	= True
    frm1.btnSend.disabled 	= True
   
	Set gActiveElement = document.activeElement
    lgBlnFlgChgValue = True
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================
Function FncPrint() 
	Call parent.FncPrint()
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================
Function FncExcel() 
    Call parent.FncExport(Parent.C_SINGLE)										
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================
Function FncFind() 
    Call parent.FncFind(Parent.C_SINGLE , False)                               
	Set gActiveElement = document.activeElement
End Function

'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================
Function FncExit()
	Dim IntRetCD
	FncExit = False
    If lgBlnFlgChgValue = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO, "X", "X")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
	Set gActiveElement = document.activeElement
    FncExit = True
End Function

'========================================================================================
' Function Name : DbDelete
' Function Desc : This function delete data
'========================================================================================
 Function DbDelete() 
    Err.Clear                                                           
    
    DbDelete = False													
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0003						
    strVal = strVal & "&txtPoNo=" & Trim(frm1.txtPoNo.value)		
    
	If LayerShowHide(1) = False Then Exit Function    
	
	Call RunMyBizASP(MyBizASP, strVal)								
	
    DbDelete = True                                                 
End Function

'========================================================================================
' Function Name : DbDeleteOk
' Function Desc : DbDelete�� �������϶� ���� 
'========================================================================================
Function DbDeleteOk()												
	lgBlnFlgChgValue = False
	Call MainNew()
End Function

'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================
Function DbQuery() 
    Err.Clear                                                       
    
    DbQuery = False                                                 
    
    Dim strVal
    
    strVal = BIZ_PGM_ID & "?txtMode=" & Parent.UID_M0001					
    strVal = strVal & "&txtPoNo=" & Trim(frm1.txtPoNo.value)		
    
    If LayerShowHide(1) = False Then Exit Function
    
	Call RunMyBizASP(MyBizASP, strVal)								
	
    DbQuery = True                                                  
End Function
'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk()												
    '-----------------------
    'Reset variables area
    '-----------------------
    Call ggoOper.LockField(Document, "Q")							
	
    frm1.btnCfm.disabled = False
    
    if frm1.rdoRelease(1).checked = true then
		Call ChangeTag(true)
		Call SetToolbar("11100000001111")
    	frm1.btnSend.disabled = False
    	frm1.btnCfm.value = "Ȯ�����"
	Else
		Call ChangeTag(False)
		Call SetToolbar("11111000001111")
		frm1.btnCfm.value = "Ȯ��"
    	frm1.btnSend.disabled = True
	end if
	
	ggoOper.SetReqAttr	frm1.txtPoNo2, "Q"
	' ��ǰ ��ȣ ��ȸ �� �Ϻ� �ʵ� Lock 
	ggoOper.SetReqAttr	frm1.txtPotypeCd, "Q"
	If UNIConvNum(isPoDetail(),0) > 0 Then
		ggoOper.SetReqAttr	frm1.txtCurr, "Q"
	else
		ggoOper.SetReqAttr	frm1.txtCurr, "N"	
	end if		
	ggoOper.SetReqAttr	frm1.txtSupplierCd, "Q"
	
	lgIntFlgMode = Parent.OPMD_UMODE										
    lgBlnFlgChgValue = False
    
	frm1.txtPoDt.focus    
	Set gActiveElement = document.activeElement
End Function

Function isPoDetail()
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6, i
	Dim iPoDetailRsCnt

    Err.Clear
	
	isPoDetail=0
	
	Call CommonQueryRs(" Count(*) ", " M_PUR_ORD_DTL ", " PO_NO = " & FilterVar(frm1.txtPoNo.value, "''", "S") & "", lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    iPoDetailRsCnt = Split(lgF0, Chr(11))

	If Err.number <> 0 Then
		MsgBox Err.description, vbInformation, parent.gLogoName 
		Err.Clear 
		Exit Function
	End If

	If Ubound(iPoDetailRsCnt) > 0 Then
		isPoDetail = iPoDetailRsCnt(0)
	End If
End Function

'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================
Function DbQueryOk1()												
    '-----------------------
    'Reset variables area
    '-----------------------
    Call ggoOper.LockField(Document, "Q")							
	
    frm1.btnCfm.disabled = False
    
	Call ChangeTag(False)
	Call SetToolbar("11111000001111")
	frm1.btnCfm.value = "Ȯ��"
	
	ggoOper.SetReqAttr	frm1.txtPotypeCd, "Q"
	ggoOper.SetReqAttr	frm1.txtCurr, "N"
	ggoOper.SetReqAttr	frm1.txtSupplierCd, "Q"
		
	lgIntFlgMode = Parent.OPMD_CMODE										
    lgBlnFlgChgValue = True
End Function
'========================================================================================
' Function Name : DBSave
' Function Desc : ���� ���� ������ ���� , �������̸� DBSaveOk ȣ��� 
'========================================================================================
Function DbSave(byval btnflg) 
    Err.Clear														

	DbSave = False													
    
    Dim strVal

	With frm1
		.txtUpdtUserId.value = Parent.gUsrID
		.txtInsrtUserId.value = Parent.gUsrID
		.txtMode.value = Parent.UID_M0002									
		.txtFlgMode.value = lgIntFlgMode
		
		if btnflg = "Cfm" then
			.txtMode.value = "Release" 				
		elseif btnflg = "UnCfm" then
			.txtMode.value = "UnRelease" 			
		end if
		
		If LayerShowHide(1) = False Then Exit Function
    	
		Call ExecMyBizASP(frm1, BIZ_PGM_ID)										
	End With
	
    DbSave = True                                   
End Function

'========================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'========================================================================================
Function DbSaveOk()									
	lgBlnFlgChgValue = False
	Call MainQuery()	
End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 >
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>��ǰ���</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=* align=right><A href="vbscript:OpenPoRef">��������</A> 
					<TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>��ǰ��ȣ</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT NAME="txtPoNo" SIZE=32  MAXLENGTH=18 ALT="���ֹ�ȣ" tag="12NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnORGCd1" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenPoNo()"></TD>
									<TD CLASS=TD6></TD>
									<TD CLASS=TD6></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR height="*">
					<TD WIDTH=100% valign=top>
						<DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; OVERFLOW: auto; WIDTH: 100%;" SCROLL=no>
							<TABLE <%=LR_SPACE_TYPE_60%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>��ǰ��ȣ</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="��ǰ��ȣ" NAME="txtPoNo2"  MAXLENGTH=18 SIZE=34 tag="21XXXU"></TD>
									<TD CLASS="TD5" NOWRAP>Ȯ������</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=radio AlT="��ǰȮ��" NAME="rdoRelease" CLASS="RADIO" checked tag="24" ONCLICK="vbscript:SetChangeflg()"><label for="rdoRelease">&nbsp;��Ȯ��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</label>
														   <INPUT TYPE=radio AlT="����Ȯ��" NAME="rdoRelease" CLASS="RADIO" ONCLICK="vbscript:setChangeflg()" tag="24"><label for="rdoRelease">&nbsp;Ȯ��&nbsp;</label></TD>
								</TR>
									<TD CLASS="TD5" NOWRAP>��ǰ����</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="��ǰ����" NAME="txtPotypeCd"  MAXLENGTH=5 SIZE=10 tag="23NXXU" ONChange="vbscript:ChangePotype()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenPotype()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
														   <INPUT TYPE=TEXT AlT="��ǰ����" NAME="txtPotypeNm" SIZE=20 tag="24X" ></TD>
									<TD CLASS="TD5" NOWRAP>�����</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m3111ma7_fpDateTime1_txtPoDt.js'></script></TD>					   
									
								<TR>
									<TD CLASS="TD5" NOWRAP>����ó</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="����ó" NAME="txtSupplierCd" MAXLENGTH=10 SIZE=10 tag="22NXXU" ONChange="vbscript:ChangeSupplier()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroup1Cd" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenSupplier()" OnMouseOver="vbscript:PopUpMouseOver()" OnMouseOut="vbscript:PopUpMouseOut()">
														   <INPUT TYPE=TEXT AlT="����ó" ID="txtSupplierNm" NAME="arrCond" tag="24X"></TD>
									<TD CLASS="TD5" NOWRAP>���ű׷�</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="���ű׷�" NAME="txtGroupCd" MAXLENGTH=4 SIZE=10 tag="22NXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenGroup()">
														   <INPUT TYPE=TEXT AlT="���ű׷�" NAME="txtGroupNm" SIZE=20 tag="24X" ></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>ȭ��</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="ȭ��" NAME="txtCurr" MAXLENGTH=3 SIZE=10 tag="22NXXU" onChange="ChangeCurr()"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnGroupCd2" align=top TYPE="BUTTON"ONCLICK="vbscript:OpenCurr()">
														   <INPUT TYPE=HIDDEN AlT="ȭ��" NAME="txtCurrNm" tag="24X"></TD>
									<TD CLASS=TD5 NOWRAP>�������ֹ�ȣ</TD>
								<TD CLASS=TD6 NOWRAP>
								    <INPUT NAME="txtPoNo1" ALT="���ֹ�ȣ" TYPE=TEXT MAXLENGTH=35 SIZE=25 TAG="24XXXU">
								    <INPUT TYPE=CHECKBOX NAME="chkPoNo" tag="23" STYLE="BORDER-BOTTOM: 0px solid; BORDER-LEFT: 0px solid; BORDER-RIGHT: 0px solid; BORDER-TOP: 0px solid"><LABEL FOR="chkPoNo">���ֹ�ȣ����</LABEL></TD>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>��ǰ���ݾ�</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m3111ma7_fpDoubleSingle1_txtNetPoAmt.js'></script></TD>
									<TD CLASS="TD5" NOWRAP>��ǰ���ڱ��ݾ�</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m3111ma7_fpDoubleSingle1_txtNetPoLocAmt.js'></script></TD>									
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>��ǰ�ݾ�</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m3111ma7_fpDoubleSingle1_txtPoAmt.js'></script></TD>
									<TD CLASS="TD5" NOWRAP>��ǰ�ڱ��ݾ�</TD>
									<TD CLASS="TD6" NOWRAP><script language =javascript src='./js/m3111ma7_fpDoubleSingle1_txtPoLocAmt.js'></script></td>
								</TR>
								<TR>
									<TD CLASS="TD5" NOWRAP>����ó�������</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="����ó�������" NAME="txtSuppSalePrsn" MAXLENGTH=50 SIZE=34 tag="21"></TD>
									<TD CLASS="TD5" NOWRAP>��޿���ó</TD>
									<TD CLASS="TD6" NOWRAP><INPUT TYPE=TEXT AlT="��޿���ó" NAME="txtTel" MAXLENGTH=30 SIZE=34 tag="21"></TD>
								</TR>
								<TR>
									<TD CLASS="TD5">���</TD>
									<TD CLASS=TD6 Colspan=3 WIDTH=100% NOWRAP><INPUT TYPE=TEXT  NAME="txtRemark" ALT="���" tag = "21" SIZE=90 MAXLENGTH=70></TD>
								</TR>
								<%Call SubFillRemBodyTD5656(7)%>
							</TABLE>
						</div>
					</TD>	
				</TR>
			</table>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
    <tr HEIGHT="20">
		<td WIDTH="100%">
			<table <%=LR_SPACE_TYPE_30%>>
				<tr>
					<TD WIDTH=10>&nbsp;</TD>
					<td align="Left"><a><button name="btnCfmSel" id="btnCfm" class="clsmbtn" ONCLICK="vbscript:Cfm()">Ȯ��</button></a>
									 <Div  STYLE="DISPLAY: none"><a><button name="btnSend" id="btnSend" class="clsmbtn" ONCLICK="Sending()">�ֹ����߼�</button></a></Div>
					</td>   
					<td WIDTH="*" align="right"><a href="VBSCRIPT:CookiePage(1)">��ǰ���ֳ������</a></td>
					<TD WIDTH=10>&nbsp;</TD>
				</tr>
			</table>
		</td>
    </tr>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX=-1></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRelease" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnCurr" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnImportflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnBLflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnCCflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRcptflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnSubcontraflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRetflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIvflg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnRcptType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIssueType" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnIvType" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMaintNo" tag="24">
<INPUT TYPE=HIDDEN NAME="txtPotype" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnxchrateop" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnMergPurFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="txtXch" tag="24">		
<INPUT TYPE=HIDDEN NAME="txtPayTermCd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtPayDur" tag="24">	
<INPUT TYPE=HIDDEN NAME="txtPayTermstxt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtPayTypeCd" tag="24">
<INPUT TYPE=HIDDEN NAME="txtVattype" tag="24">
<INPUT TYPE=HIDDEN NAME="txtVatrt" tag="24">
<INPUT TYPE=HIDDEN NAME="txtApplicantCd" tag="24">
<INPUT TYPE=HIDDEN NAME="hdvatFlg" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnpotype" tag="24">
<INPUT TYPE=HIDDEN NAME="hdnPoDt" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT" STYLE="visible:true">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>
