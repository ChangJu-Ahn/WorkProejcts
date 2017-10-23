
<%@ LANGUAGE="VBSCRIPT" %>
<%Response.Expires = -1%>
<!--'**********************************************************************************************
'*  1. Module��          : ȸ����� 
'*  2. Function��        : �ڻ���� 
'*  3. Program ID        : a7121ma1.asp
'*  4. Program �̸�      : �����ڻ����Check LIst ��� 
'*  5. Program ����      : 
'*  6. Comproxy ����Ʈ   : 
'*  7. ���� �ۼ������   : 2000/12/16
'*  8. ���� ���������   : 2004/01/08
'*  9. ���� �ۼ���       : ������ 
'* 10. ���� �ۼ���       : Kim Chang Jin
'* 11. ��ü comment      :
'* 12. ���� Coding Guide : this mark(��) means that "Do not change"
'*                         this mark(��) Means that "may  change"
'*                         this mark(��) Means that "must change"
'* 13. History           :
'********************************************************************************************** -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--
'##########################################################################################################
'												1. �� �� �� 
'##########################################################################################################
'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/incSvrCcm.inc"  -->
<!-- #Include file="../../inc/incSvrHTML.inc"  -->
<!--
'===============================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="JavaScript" SRC="../../inc/TabScript.js"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/Cookie.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit																	'��: indicates that All variables must be declared in advance


'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************

'��: �����Ͻ� ���� ASP�� 
'Const BIZ_PGM_ID = "a7120mb1.asp"			'��: �����Ͻ� ���� ASP�� 


'==========================================  1.2.1 Global ��� ����  ======================================
'==========================================================================================================
'��: Grid Columns


'==========================================  1.2.2 Global ���� ����  =====================================
'	1. ���� ǥ�ؿ� ����. prefix�� g�� �����.
'	2.Array�� ���� ()�� �ݵ�� ����Ͽ� �Ϲ� ������ ������ �� 
'========================================================================================================= 

'Dim lgBlnFlgChgValue           ' Variable is for Dirty flag 
'Dim lgIntFlgMode               ' Variable is for Operation Status 


'==========================================  1.2.3 Global Variable�� ����  ===============================
'========================================================================================================= 

'----------------  ���� Global ������ ����  ----------------------------------------------------------- 
Dim IsOpenPop

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ�		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)				
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 


'++++++++++++++++  Insert Your Code for Global Variables Assign  ++++++++++++++++++++++++++++++++++++++ 

<!-- #Include file="../../inc/lgvariables.inc" --> 
'#########################################################################################################
'												2. Function�� 
'
'	���� : �����ڰ� ������ �Լ�, �� Event���� �Լ��� ������ ��� ����� ���� �Լ� �⽽ 
'	�������� ���� ���� : 1. Sub �Ǵ� Function�� ȣ���� �� �ݵ�� Call�� ����.
'		     	     	 2. Sub, Function �̸��� _�� ���� �ʵ��� �Ѵ�. (Event�� �����ϱ� ����) 
'######################################################################################################### 

'==========================================  2.1.1 InitVariables()  ======================================
'	Name : InitVariables()
'	Description : ���� �ʱ�ȭ(Global ����, �ʱ�ȭ�� �ʿ��� ���� �Ǵ� Flag���� Setting�Ѵ�.)
'========================================================================================================= 
Sub InitVariables()

    lgIntFlgMode = parent.OPMD_CMODE                   'Indicates that current mode is Create mode
    lgBlnFlgChgValue = False                    'Indicates that no value changed

    '---- Coding part--------------------------------------------------------------------    
End Sub


'******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************* 

'========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 

Sub SetDefaultVal()

	Dim strYear, strMonth, strDay, dtToday, EndDate, StartDate
	EndDate = "<%=GetSvrDate%>"
	Call ExtractDateFrom(EndDate, parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)

	StartDate = UniConvYYYYMMDDToDate(gDateFormat, strYear, strMonth, "01")
	EndDate = UniConvYYYYMMDDToDate(gDateFormat, strYear, strMonth, strDay)

	frm1.fpDateTime1.Text = StartDate
	frm1.fpDateTime2.Text = EndDate

End Sub


'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 

Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call LoadInfTB19029A("Q", "A", "NOCOOKIE", "OA") %>
End Sub


'******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'********************************************************************************************************* 

'========================================== 2.4.2 Open???()  =============================================
'	Name : Open???()
'	Description : �ߺ��Ǿ� �ִ� PopUp�� ������, �����ǰ� �ʿ��� ���� �ݵ�� CommonPopUp.vbs �� 
'				  ManufactPopUp.vbs ���� Copy�Ͽ� �������Ѵ�.
'========================================================================================================= 

'+++++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 

Function OpenPopup(Byval strCode, Byval Cond)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select case Cond

	Case "FrAcqNo", "ToAcqNo"
		arrParam(0) = "�ڻ�����ȣ�˾�"			' �˾� ��Ī 
		arrParam(1) = "A_ASSET_ACQ"					' TABLE ��Ī 
		arrParam(2) = strCode      					' Code Condition
		arrParam(3) = ""								' Name Cindition

		arrParam(4) = " 1 = 1 "								' Where Condition

		' ���Ѱ��� �߰� 
		If lgAuthBizAreaCd <> "" Then
			arrParam(4) = arrParam(4) & " AND BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
		End If
		If lgInternalCd <> "" Then
			arrParam(4) = arrParam(4) & " AND INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")			' Where Condition
		End If
		If lgSubInternalCd <> "" Then
			arrParam(4) = arrParam(4) & " AND INTERNAL_CD like " & FilterVar(lgSubInternalCd & "%", "''", "S")		' Where Condition
		End If
		If lgAuthUsrID <> "" Then
			arrParam(4) = arrParam(4) & " AND INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")			' Where Condition
		End If

		arrParam(5) = "�ڻ�����ȣ"				' �����ʵ��� �� ��Ī 
				
		arrField(0) = "Acq_No"						' Field��(0)
			   
		arrHeader(0) = "�ڻ�����ȣ"			' Header��(0)
			    

	Case "Biz", "Biz1"
	    
		arrParam(0) = "������ڵ��˾�"			' �˾� ��Ī 
		arrParam(1) = "B_BIZ_AREA"					' TABLE ��Ī 
		arrParam(2) = strCode      					' Code Condition
		arrParam(3) = ""								' Name Cindition

		' ���Ѱ��� �߰� 
		If lgAuthBizAreaCd <> "" Then
			arrParam(4) = " BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")			' Where Condition
		Else
			arrParam(4) = ""
		End If

		arrParam(5) = "������ڵ�"				' �����ʵ��� �� ��Ī	   
				
		arrField(0) = "Biz_Area_Cd"					' Field��(0)
		arrField(1) = "Biz_Area_Nm"					' Field��(1)
			    
		arrHeader(0) = "������ڵ�"				' Header��(0)
		arrHeader(1) = "������"

	End select
	    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
		
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetReturnVal(arrRet, Cond)		
	End If	
	frm1.txtFrYymm.focus
End Function

'==========================================  2.4.3 Set???()  =============================================
'	Name : Set???()
'	Description : PopUp���� ���۵� ���� Ư�� Tag Object�� ���� 
'========================================================================================================= 
'------------------------------------------  SetReturnVal()  ---------------------------------------------
'	Name : SetReturnVal()
'	Description : Account Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function SetReturnVal(ByVal arrRet, ByVal field_fg)	
	Select case field_fg
		case "FrAcqNo"
			frm1.txtFrAcqNo.Value		= arrRet(0)
			
		case "ToAcqNo"
			frm1.txtToAcqNo.Value		= arrRet(0)	
			
		case "Biz"  	'biz Area Popup
			frm1.txtBizAreacd.Value		= arrRet(0)
			frm1.txtBizAreanm.Value		= arrRet(1)
		
		case "Biz1"  	'biz Area Popup
			frm1.txtBizAreacd1.Value	= arrRet(0)
			frm1.txtBizAreanm1.Value	= arrRet(1)
				
	End select	
End Function

'++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

'#########################################################################################################
'												3. Event�� 
'	���: Event �Լ��� ���� ó�� 
'	����: Windowó��, Singleó��, Gridó�� �۾�.
'         ���⼭ Validation Check, Calcuration �۾��� ������ Event�� �߻�.
'         �� Object������ Grouping�Ѵ�.
'##########################################################################################################

'******************************************  3.1 Window ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'********************************************************************************************************* 

'==============================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029                           '��: Load table , B_numeric_format

    Call ggoOper.ClearField(Document, "1")        '��: Condition field clear
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,gDateFormat,parent.gComNum1000,parent.gComNumDec)
    
    Call ggoOper.LockField(Document, "N")         '��: ���ǿ� �´� Field locking    
    
    Call InitVariables                            '��: Initializes local global Variables
    Call SetDefaultVal
    
    Call SetToolbar("10000000000011")				'��: ��ư ���� ���� 
	
	frm1.txtBizAreaCd.focus	

	' ���Ѱ��� �߰� 
	Dim xmlDoc

	Call GetDataAuthXML(parent.gUsrID, gStrRequestMenuID, xmlDoc)

	' �����		
	lgAuthBizAreaCd	= xmlDoc.selectSingleNode("/root/data/data_biz_area_cd").Text
	lgAuthBizAreaNm	= xmlDoc.selectSingleNode("/root/data/data_biz_area_nm").Text

	' ���κμ�		
	lgInternalCd	= xmlDoc.selectSingleNode("/root/data/data_internal_cd").Text
	lgDeptCd		= xmlDoc.selectSingleNode("/root/data/data_dept_cd").Text
	lgDeptNm		= xmlDoc.selectSingleNode("/root/data/data_dept_nm").Text

	' ���κμ�(��������)		
	lgSubInternalCd	= xmlDoc.selectSingleNode("/root/data/data_sub_internal_cd").Text
	lgSubDeptCd		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_cd").Text
	lgSubDeptNm		= xmlDoc.selectSingleNode("/root/data/data_sub_dept_nm").Text

	' ����						
	lgAuthUsrID		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_id").Text
	lgAuthUsrNm		= xmlDoc.selectSingleNode("/root/data/data_auth_usr_nm").Text

	Set xmlDoc = Nothing

End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub


'**************************  3.2 HTML Form Element & Object Eventó��  **********************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'********************************************************************************************************* 


'******************************  3.2.1 Object Tag ó��  *********************************************
'	Window�� �߻� �ϴ� ��� Event ó��	
'********************************************************************************************************* 

'======================================================================================================
'   Event Name : txtFrYymm_DblClick
'   Event Desc : �޷� Popup�� ȣ�� 
'=======================================================================================================
Sub txtFrYymm_DblClick(Button)
    If Button = 1 Then
        frm1.fpDateTime1.Action = 7
    End If
End Sub
Sub txtToYymm_DblClick(Button)
    If Button = 1 Then
        frm1.fpDateTime2.Action = 7
    End If
End Sub
Sub txtBizAreacd_onBlur
	if frm1.txtBizAreacd.value = "" then
		frm1.txtBizAreanm.value = ""
	end if
End Sub
'======================================================================================================
' Function Name : SetPrintCond
' Function Desc : This function is related to Print/Preview Button
'=======================================================================================================
Function SetPrintCond(StrEbrFile, StrUrl)

	Dim	varBizArea, varBizArea1, varFromAcq, varToAcq, varFromdt, varTodt
	Dim	strAuthCond

	Dim IntRetCd

	SetPrintCond = False

    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Function
    End If	

	If CompareDateByFormat(frm1.txtFrYymm.text,frm1.txtToYymm.text,frm1.txtFrYymm.Alt,frm1.txtToYymm.Alt, _
        	               "970025",frm1.txtFrYYMM.UserDefinedFormat,parent.gComDateType, true) = False Then
	   frm1.txtFrYymm.focus
	   Exit Function
	End If

	If Trim(frm1.txtFrAcqNo.value) <> "" and Trim(frm1.txtToAcqNo.value) <> "" Then
		If Trim(frm1.txtFrAcqNo.value) > Trim(frm1.txtToAcqNo.value) Then
			IntRetCd = DisplayMsgBox("970025", "X", frm1.txtFrAcqNo.Alt, frm1.txtToAcqNo.Alt)
			frm1.txtFrAcqNo.focus
			Exit Function
		End If
	End If


	StrEbrFile = "a7121ma1"

	If Len(frm1.txtBizAreacd.value) < 1 Then
		varBizArea    = " "
	Else		
		varBizArea    = FilterVar(Trim(frm1.txtBizAreaCd.value),"","SNM")
	End If
	
	If Len(frm1.txtBizAreacd1.value) < 1 Then
		varBizArea1    = "ZZZZZZZZZZ"
	Else		
		varBizArea1    = FilterVar(Trim(frm1.txtBizAreaCd1.value),"","SNM")
	End If
	
	If Len(frm1.txtFrAcqNo.value) < 1 Then
		VarFromAcq = " "
	Else
		VarFromAcq   = FilterVar(Trim(frm1.txtFrAcqNo.value),"","SNM")
	End If
	
	If Len(frm1.txtToAcqNo.value) < 1 Then
		VarToAcq = "zzzzzzzzzzzzzzzzzz"
	Else
		VarToAcq	   = FilterVar(Trim(frm1.txtToAcqNo.value),"","SNM")
	End If	
	
	if Trim(frm1.fpDateTime1.Text) = "" then
		varFromDt = UniConvDateToYYYYMMDD(parent.gServerBaseDate, gDateFormat,"")
	else		
		VarFromDt = UniConvDateToYYYYMMDD(frm1.fpDateTime1.Text, gDateFormat,"")
	end if
	
	if Trim(frm1.fpDateTime1.Text) = "" then
		varToDt   = UniConvDateToYYYYMMDD(parent.gServerBaseDate, gDateFormat,"")
	else
		VarToDt   = UniConvDateToYYYYMMDD(frm1.fpDateTime2.Text, gDateFormat,"")
	end if

	' ���Ѱ��� �߰� 
	strAuthCond		= "	"
	
	If lgAuthBizAreaCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND A_ASSET_ACQ.BIZ_AREA_CD = " & FilterVar(lgAuthBizAreaCd, "''", "S")  		
	End If			

	If lgInternalCd <> "" Then			
		strAuthCond		= strAuthCond	& " AND A_ASSET_ACQ.INTERNAL_CD = " & FilterVar(lgInternalCd, "''", "S")  		
	End If			

	If lgSubInternalCd <> "" Then	
		strAuthCond		= strAuthCond	& " AND A_ASSET_ACQ.INTERNAL_CD LIKE " & FilterVar(lgSubInternalCd & "%", "''", "S")  
	End If	

	If lgAuthUsrID <> "" Then	
		strAuthCond		= strAuthCond	& " AND A_ASSET_ACQ.INSRT_USER_ID = " & FilterVar(lgAuthUsrID, "''", "S")  
	End If	


	StrUrl = StrUrl & "varFromAcq|"		& varFromAcq
	StrUrl = StrUrl & "|varToAcq|"		& varToAcq
	StrUrl = StrUrl & "|varFromDt|"		& varFromDt
	StrUrl = StrUrl & "|varToDt|"		& varToDt
	StrUrl = StrUrl & "|varBizArea|"	& varBizArea
	StrUrl = StrUrl & "|varBizArea1|"	& varBizArea1

	StrUrl = StrUrl & "|strAuthCond|"	& strAuthCond

	SetPrintCond = True


End Function

'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : This function is related to Print Button
'=======================================================================================================
Function FncBtnPrint() 

    Dim StrEbrFile, StrUrl, ObjName
	
	If Not SetPrintCond(StrEbrFile, strUrl) Then
		Exit Function
	End If
	    
	ObjName = AskEBDocumentName(StrEbrFile,"ebr")

	Call FncEBRPrint(EBAction,ObjName,StrUrl)		

End Function


'========================================================================================
' Function Name : FncBtnPreview
' Function Desc : This function is related to Preview Button
'========================================================================================

Function FncBtnPreview() 

    Dim StrEbrFile, StrUrl, ObjName
	
	If Not SetPrintCond(StrEbrFile, strUrl) Then
		Exit Function
	End If
	    
	ObjName = AskEBDocumentName(StrEbrFile,"ebr")

	Call FncEBRPreview(ObjName,StrUrl)	
	
End Function


'#########################################################################################################
'												4. Common Function�� 
'	���: Common Function
'	����: ȯ��ó���Լ�, VAT ó�� �Լ� 
'######################################################################################################### 


'#########################################################################################################
'												5. Interface�� 
'	���: Interface
'	����: ������ Toolbar�� ���� ó���� ���Ѵ�. 
'	      Toolbar�� ��ġ������� ����ϴ� ������ �Ѵ�. 
'	<< ���뺯�� ���� �κ� >>
' 	���뺯�� : Global Variables�� �ƴ����� ������ Sub�� Function���� ���� ����ϴ� ������ �������� 
'				�����ϵ��� �Ѵ�.
' 	1. ������Ʈ���� Call�ϴ� ���� 
'    	   ADF (ADS, ADC, ADF�� �״�� ���)
'    	   - ADF�� Set�ϰ� ����� �� �ٷ� Nothing �ϵ��� �Ѵ�.
' 	2. ������Ʈ�ѿ��� Return�� ���� �޴� ���� 
'    		strRetMsg
'######################################################################################################### 


'********************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* 



'========================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'========================================================================================

Function FncQuery() 
End Function


'========================================================================================
' Function Name : FncNew
' Function Desc : This function is related to New Button of Main ToolBar
'========================================================================================

Function FncNew() 
End Function


'========================================================================================
' Function Name : FncDelete
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncDelete() 
End Function


'========================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'========================================================================================

Function FncSave() 
End Function


'========================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================

Function FncCopy() 
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
End Function


'========================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================

Function FncInsertRow() 
End Function


'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call Parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================

Function FncPrev() 
    On Error Resume Next                        '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

Function FncNext() 
    On Error Resume Next                        '��: Protect system from crashing
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : 
'=======================================================================================================
Function FncFind() 
	Call parent.FncFind(parent.C_SINGLE, False)                                         '��:ȭ�� ����, Tab ���� 
End Function


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function

'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	

</HEAD>
<!--
'#########################################################################################################
'       					6. Tag�� 
'#########################################################################################################  -->
<BODY TABINDEX="-1" SCROLL="NO">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE CLASS="BatchTB2" CELLSPACING=0 CELLPADDING=0 >
	<TR>
		<TD <%=HEIGHT_TYPE_00%>><% ' ���� ���� %></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD>
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
			
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_60%>>
							<TR>
								<TD CLASS="TD5" NOWRAP>������ڵ�</TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreacd" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="������ڵ�" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCd.Value, 'Biz')"> 
													   <INPUT TYPE="Text" NAME="txtBizAreanm" SIZE=20 MAXLENGTH=30 tag="14" ALT="������">&nbsp;~</TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtBizAreacd1" SIZE=10 MAXLENGTH=10 tag="11XXXU" ALT="������ڵ�" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtBizAreaCd1.Value, 'Biz1')"> 
													   <INPUT TYPE="Text" NAME="txtBizAreanm1" SIZE=20 MAXLENGTH=30 tag="14" ALT="������"></TD>
							</TR>					
							<TR>
								<TD CLASS="TD5" NOWRAP>�������</TD>
								<TD CLASS="TD6" NOWRAP>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtFrYymm" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME" ALT=����������� id=fpDateTime1> </OBJECT>');</SCRIPT> ~
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtToYymm" CLASS=FPDTYYYYMMDD tag="12X1" Title="FPDATETIME" ALT=����������� id=fpDateTime2> </OBJECT>');</SCRIPT>
								</TD>								
							</TR>				
   							<TR>
           						<TD CLASS="TD5" NOWRAP>����ȣ</TD>           						
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtFrAcqNo" SIZE=15 MAXLENGTH=18 tag="11XXXU" ALT="��������ȣ" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtFrAcqNo.Value, 'FrAcqNo')">&nbsp;~&nbsp;</TD>
							</TR>	
							<TR>
								<TD CLASS="TD5" NOWRAP></TD>
								<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtToAcqNo" SIZE=15 MAXLENGTH=18 tag="11XXXU" ALT="��������ȣ" STYLE="TEXT-ALIGN:left"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnAcctCdFR" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenPopup(frm1.txtToAcqNo.Value, 'ToAcqNo')"></TD>
                    	        </TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%>></TD>
	</TR>
	<TR HEIGHT=20>
		<TD WIDTH=100%>
			<TABLE <%=LR_SPACE_TYPE_30%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD><BUTTON NAME="btnpre" CLASS="CLSSBTN" OnClick="VBScript:Call FncBtnPreview()" Flag=1>�̸�����</BUTTON>
					 &nbsp<BUTTON NAME="btn" CLASS="CLSSBTN" OnClick="VBScript:Call FncBtnPrint()" Flag=1>�μ�</BUTTON></TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 TABINDEX = "-1" ></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX = "-1" >
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24" TABINDEX = "-1" >
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname" TABINDEX = "-1" >
	<INPUT TYPE="HIDDEN" NAME="dbname" TABINDEX = "-1" >
	<INPUT TYPE="HIDDEN" NAME="filename" TABINDEX = "-1" >
	<INPUT TYPE="HIDDEN" NAME="condvar" TABINDEX = "-1" >
	<INPUT TYPE="HIDDEN" NAME="date" TABINDEX = "-1" >	
</FORM>
</BODY>
</HTML>

