

<%@ LANGUAGE="VBSCRIPT" %>
<!--===================================================================================================
'*  1. Module Name          : Accounting
'*  2. Function Name        : Treasury - Budget
'*  3. Program ID           : f2107ma1
'*  4. Program Name         : �������������ȸ 
'*  5. Program Desc         : Query of Budget Result by Monthly
'*  6. Comproxy List        : ADO
'*  7. Modified date(First) : 2001.02.13
'*  8. Modified date(Last)  : 2001.03.05
'*  9. Modifier (First)     : Song, Mun Gil
'* 10. Modifier (Last)      : Song, Mun Gil
'* 11. Comment              :
'======================================================================================================= -->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>
<!--########################################################################################################
'												1. �� �� �� 
'###########################################################################################################
'******************************************  1.1 Inc ����   **********************************************
'	���: Inc. Include
'*********************************************************************************************************  -->
<!-- #Include file="../../inc/incSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->
<!--'==========================================  1.1.1 Style Sheet  ======================================
'========================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">

<!--'==========================================  1.1.2 ���� Include   ======================================
'========================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAMain.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAEvent.vbs">  </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliVariables.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"  SRC="../../inc/incCliMAOperation.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentA.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliDBAgentVariables.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/Cookie.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript">

Option Explicit													'��: indicates that All variables must be declared in advance
<!-- #Include file="../../inc/lgvariables.inc" -->	

'******************************************  1.2 Global ����/��� ����  ***********************************
'	1. Constant�� �ݵ�� �빮�� ǥ��.
'**********************************************************************************************************
'Dim lgBlnFlgChgValue                                        '��: Dirtry tag                              
'Dim lgStrPrevKey                                            '��: Next Key tag                              
'Dim lgSortKey                                               '��: Sort���� ���庯��                       
Dim lgIsOpenPop                                             '��: Popupȭ���� ���� ���庯��               


Dim lgPopUpR                                                '��: Orderby default ��                      
Dim lgMark                                                  '��: ��ũ                                    

Dim lgSaveRow                                               '��: Cookie���� ����                          

' ���Ѱ��� �߰� 
Dim lgAuthBizAreaCd, lgAuthBizAreaNm			' ����� 
Dim lgInternalCd, lgDeptCd, lgDeptNm			' ���κμ�		
Dim lgSubInternalCd, lgSubDeptCd, lgSubDeptNm	' ���κμ�(��������)				
Dim lgAuthUsrID, lgAuthUsrNm					' ���� 


'--------------- ������ coding part(��������,Start)-----------------------------------------------------------
Const BIZ_PGM_ID 		= "f2107mb1.asp"                         '��: �����Ͻ� ���� ASP�� 

Const C_MaxKey          = 2		                                 '�١١١�: Max key value

'--------------- ������ coding part(��������,End)------------------------------------------------------------------------------

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
    lgStrPrevKey     = ""
	lgBlnFlgChgValue = False
    lgSortKey        = 1
    lgSaveRow        = 0
End Sub


 '******************************************  2.2 ȭ�� �ʱ�ȭ �Լ�  ***************************************
'	���: ȭ���ʱ�ȭ 
'	����: ȭ���ʱ�ȭ, Combo Display, ȭ�� Clear �� ȭ�� �ʱ�ȭ �۾��� �Ѵ�. 
'********************************************************************************************************* 
 '==========================================  2.2.1 SetDefaultVal()  ========================================
'	Name : SetDefaultVal()
'	Description : ȭ�� �ʱ�ȭ(���� Field�� �� �� ȭ���� �� �� Default���� ������� �ϴ� Field�� Setting)
'========================================================================================================= 
Sub SetDefaultVal()

'--------------- ������ coding part(�������,Start)--------------------------------------------------

	frm1.txtBdgYear.Text = UniConvDateAToB("<%=GetSvrDate%>" ,parent.gServerDateFormat,parent.gDateFormat) 
	frm1.hOrgChangeId.value = parent.gChangeOrgId
	
'--------------- ������ coding part(�������,End)----------------------------------------------------

End Sub

'======================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'========================================================================================
Sub LoadInfTB19029()
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
<% Call loadInfTB19029A("Q", "A","NOCOOKIE","QA") %>
	<% Call LoadBNumericFormatA("Q", "A", "NOCOOKIE", "QA") %>                           
End Sub


'======================= 2.2.3 InitSpreadSheet() ========================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================
Sub InitSpreadSheet()
    Call SetZAdoSpreadSheet("F2107MA1","G","A","V20021211",parent.C_GROUP_DBAGENT,frm1.vspdData, C_MaxKey, "X","X")
    Call SetSpreadLock("A")
End Sub



'================================== 2.2.4 SetSpreadLock() ==================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================
Sub SetSpreadLock(Byval iOpt )
    If iOpt = "A" Then
		With frm1
			.vspdData.ReDraw = False
			ggoSpread.Source = .vspdData 
			ggoSpread.SpreadLockWithOddEvenRowColor()
			.vspdData.ReDraw = True
		End With
    End If
End Sub

'================================== 2.2.5 SetSpreadColor() ==================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'========================================================================================

 '******************************************  2.4 POP-UP ó���Լ�  ****************************************
'	���: ���� POP-UP
'   ����: ���� POP-UP�� ���� Open�� Include�Ѵ�. 
'	      �ϳ��� ASP���� Popup�� �ߺ��Ǹ� �ϳ��� ��ũ���� ����ϰ� �������� �������Ͽ� ����Ѵ�.
'********************************************************************************************************* 
Function OpenPopUp(Byval strCode, Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If lgIsOpenPop = True Then Exit Function

	Select Case iWhere

'	   Case "DeptCd"
'			arrParam(0) = "�μ��ڵ� �˾�"								' �˾� ��Ī 
'			arrParam(1) = "B_ACCT_DEPT A "									' TABLE ��Ī 
'			arrParam(2) = strCode											' Code Condition
'			arrParam(3) = ""												' Name Cindition
'			arrParam(4) = "A.ORG_CHANGE_ID =  " & FilterVar(parent.gChangeOrgId , "''", "S") & ""
'			arrParam(5) = "�μ��ڵ�"									' �����ʵ��� �� ��Ī 
'
'			arrField(0) = "A.DEPT_CD"
'			arrField(1) = "A.DEPT_NM"
'			
'			arrHeader(0) = "�μ��ڵ�"									' Header��(0)
'			arrHeader(1) = "�μ���"										' Header��(1)
'		
		Case Else
			Exit Function
	End Select	
	
	lgIsOpenPop = True
	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

End Function

 '++++++++++++++++  Insert Your Code for PopUp(Open)  ++++++++++++++++++++++++++++++++++++++++++++++++++ 
'------------------------------------------  SetPopUp()  --------------------------------------------------
'	Name : SetPopUp()
'	Description : Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 
Function SetPopUp(Byval arrRet, Byval iWhere)
	With frm1
		Select Case iWhere
		    
			Case "DeptCd"
				.txtDeptCd.value = arrRet(0)
				.txtDeptNm.value = arrRet(1)
			
		End Select
    
	End With

End Function
 '------------------------------------------  OpenDept()  ---------------------------------------
'	Name : OpenDept()
'	Description : OpenAp Popup���� Return�Ǵ� �� setting
'--------------------------------------------------------------------------------------------------------- 

Function OpenDept()
	Dim arrRet
	Dim arrParam(8)
	Dim Temp
	Dim strYear, strMonth, strDay
	If lgIsOpenPop = True Then Exit Function

	lgIsOpenPop = True

	Call parent.ExtractDateFrom("<%=GetSvrDate%>", Parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)
	arrParam(0) = frm1.txtDeptCd.value		            '  Code Condition
	arrParam(0) = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtBdgYear.text, "01", "01")
	arrParam(1) = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtBdgYear.text, "12", "31")
	arrParam(2) = lgUsrIntCd                            ' �ڷ���� Condition  
	arrParam(3) = frm1.txtDeptCd.value
	arrParam(4) = "T"									' �������� ���� Condition  

	' ���Ѱ��� �߰� 
	arrParam(5)		= lgAuthBizAreaCd
	arrParam(6)		= lgInternalCd
	arrParam(7)		= lgSubInternalCd
	arrParam(8)		= lgAuthUsrID

	
	arrRet = window.showModalDialog("../../comasp/DeptPopupOrg.asp", Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	lgIsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetDept(arrRet)
	End If	
			
End Function
'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup���� Return�Ǵ� �� setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(ByVal arrRet)
		
	With frm1
		.txtDeptCd.focus
		.txtDeptCd.value = Trim(arrRet(0))
		.txtDeptNm.value = arrRet(1)
		.hOrgChangeId.value=arrRet(2)
	End With
End Function  
'===========================================================================
' Function Name : PopZAdoConfigGrid
' Function Desc : 
'===========================================================================
Function PopZAdoConfigGrid()

	Dim arrRet
	Dim gPos
	
	Select Case UCase(Trim(gActiveSpdSheet.Name))
	       Case "VSPDDATA"
	            gPos = "A"
	       End Select     
	       
	If lgIsOpenPop = True Then Exit Function
	lgIsOpenPop = True

	arrRet = window.showModalDialog("../../ComAsp/ZADOSortPopup.asp",Array(ggoSpread.GetXMLData(gPos),gMethodText),"dialogWidth=" & parent.SORTW_WIDTH & "px; dialogHeight=" & parent.SORTW_HEIGHT & "px;; center: Yes; help: No; resizable: No; status: No;")

	lgIsOpenPop = False
	
	If arrRet(0) = "X" Then
	   Exit Function
	Else
	   Call ggoSpread.SaveXMLData(gPos,arrRet(0),arrRet(1))
       Call InitVariables
       Call InitSpreadSheet()       
   End If
End Function

 '++++++++++++++++++++++++++++++++++++++++++  2.5 ������ ���� �Լ�  +++++++++++++++++++++++++++++++++++++++
'    ������ ���α׷� ���� �ʿ��� ������ ���� Procedure (Sub, Function, Validation & Calulation ���� �Լ�)
'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++ 

 '==========================================   CookiePage()  ======================================
'	Name : CookiePage()
'	Description : JUMP�� Loadȭ������ ���Ǻη� Value
'==================================================================================================== 
Function CookiePage(Byval Kubun)

	Const CookieSplit = 4877						'Cookie Split String : CookiePage Function Use
	Dim   strCookie
	Dim   ii,jj,kk
    Dim   iSeq
    Dim   IntRetCD
    
    
	If  lgSaveRow <  1 Then
		IntRetCD = DisplayMsgBox("900002",parent.VB_YES_NO,"X","X")
		Exit Function
	End If	
	
    Redim  lgMark(UBound(lgFieldNM)) 

	strCookie  = ""
    iSeq       = 0
    
    For ii = 0 to parent.C_MaxSelList - 1 
        For jj = 0 to UBound(lgFieldNM) -1
            If lgPopUpR(ii,0) = lgFieldCD(jj) Then
               iSeq = iSeq + 1
               lgMark(jj) = "X"
               strCookie = strCookie & Trim(lgFieldNM(jj)) & parent.gRowSep
               frm1.vspdData.Row = lgSaveRow
               frm1.vspdData.Col = iSeq
               strCookie = strCookie & frm1.vspdData.Text & parent.gRowSep
       
               kk = CInt(lgNextSeq(jj)) 
               If kk > 0 And kk <= UBound(lgFieldNM) Then 
                  lgMark(kk - 1) = "X"
                  iSeq = iSeq + 1
          
                  strCookie = strCookie & Trim(lgFieldNM(kk-1)) & parent.gRowSep
                  frm1.vspdData.Row = lgSaveRow
                  frm1.vspdData.Col = iSeq
                  strCookie = strCookie & frm1.vspdData.Text & parent.gRowSep
               End If    
               jj =  UBound(lgFieldNM)  + 100
             End If    
        Next
    Next      
    
	WriteCookie CookieSplit , strCookie
	Call PgmJump(BIZ_PGM_JUMP_ID)

End Function

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
 '==========================================  3.1.1 Form_Load()  ======================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� �Լ��� Call�ϴ� �κ� 
'========================================================================================================= 
Sub Form_Load()

    Call LoadInfTB19029														'��: Load table , B_numeric_format
	Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.LockField(Document, "N")                                   '��: Lock  Suitable  Field
    
	Call InitVariables														'��: Initializes local global variables
	Call SetDefaultVal	
	Call ggoOper.FormatDate(frm1.txtBdgYear, parent.gDateFormat, 3)
	Call InitSpreadSheet()
    
'--------------- ������ coding part(�������,Start)----------------------------------------------------
	Call FncSetToolBar("New")
	frm1.fpBdgYear.focus
	
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


'--------------- ������ coding part(�������,End)------------------------------------------------------

End Sub

'==========================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'==========================================================================================

Sub Form_QueryUnload(Cancel , UnloadMode )
End Sub

 '*****************  3.2 HTML Form Element & Object Eventó��  **************************
'	Document�� TAG���� �߻� �ϴ� Event ó��	
'	Event�� ��� �Ʒ��� ����� Event�̿��� ����� �����ϸ� �ʿ�� �߰� �����ϳ� 
'	Event�� �浹�� ����Ͽ� �ۼ��Ѵ�.
'****************************************************************************************** 

 '******************************  3.2.1 Object Tag ó��  ********************************
'	Window�� �߻� �ϴ� ��� Even ó��	
'****************************************************************************************** 
Sub txtBdgYear_DblClick(Button)
    If Button = 1 Then
       frm1.txtBdgYear.Action = 7
       Call SetFocusToDocument("M")	
       frm1.txtBdgYear.Focus       
    End If
End Sub

Sub txtBdgYear_KeyPress(KeyAscii)
	If KeyAscii = 13 Then 
	   Call MainQuery
	End If   
End Sub
'==========================================================================================
'   Event Name : txtDeptCd_Onchange
'   Event Desc : 
'==========================================================================================
Sub txtDeptCD_OnChange()
        
    Dim strSelect, strFrom, strWhere 	
    Dim IntRetCD 
	Dim arrVal1, arrVal2
	Dim ii, jj
	'Dim lgF2By2
	Dim strStartDt, strEndDt
	Dim strYear, strMonth, strDay

	if frm1.txtDeptCd.value = "" then
		frm1.txtDeptNm.value = ""
	end if
	
    lgBlnFlgChgValue = True
	'strStartDt = UniConvDateAToB(frm1.txtBdgYear,parent.gDateFormatYYYY,parent.gServerDateFormat)


	If TRim(frm1.txtDeptCd.value) <>"" Then
		'Call parent.ExtractDateFrom("<%=GetSvrDate%>", Parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)
			strStartDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtBdgYear.text, "01", "01")  
			strEndDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtBdgYear.text, "12", "31")  
			strStartDt =  UNIConvDateToYYYYMMDD(strStartDt, gDateFormat,Parent.gServerDateType)  
			strEndDt =  UNIConvDateToYYYYMMDD(strEndDt, gDateFormat,Parent.gServerDateType)  
		'----------------------------------------------------------------------------------------
			strSelect = "dept_cd, ORG_CHANGE_ID"
			strFrom =  " B_ACCT_DEPT "
			strWhere = " ORG_CHANGE_DT >= "
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(strStartDt , "''", "S") & ")"
			strWhere = strWhere & " AND ORG_CHANGE_DT <= " 
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(strEndDt , "''", "S") & ") "
			strWhere =	strWhere & " AND dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")

			'msgbox "Select " & strSelect& " from " &strFrom & " where "&strWhere
		If CommonQueryRs2by2(strSelect, strFrom ,  strWhere , lgF2By2) = False Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
			frm1.txtDeptCd.value = ""
			frm1.txtDeptNm.value = ""
			frm1.hOrgChangeId.value = ""
			frm1.txtDeptCd.focus

		Else 
		
			arrVal1 = Split(lgF2By2, Chr(11) & Chr(12))			
			jj = Ubound(arrVal1,1)
					
			For ii = 0 to jj - 1
				arrVal2 = Split(arrVal1(ii), chr(11))			
				frm1.hOrgChangeId.value = Trim(arrVal2(2))
				
			Next	
			
		End If
	End IF
		'----------------------------------------------------------------------------------------
End Sub



'==========================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'==========================================================================================

Sub vspdData_GotFocus()
    ggoSpread.Source = frm1.vspdData
End Sub


'==========================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc :
'==========================================================================================

	Function vspdData_DblClick(ByVal Col, ByVal Row)
		If frm1.vspdData.MaxRows > 0 Then
			If frm1.vspdData.ActiveRow = Row Or frm1.vspdData.ActiveRow > 0 Then
			'	Call CookiePage(1)
			End If
		End If
	End Function
	
'======================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'=======================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
    Call SetPopupMenuItemInf("00000000001")
    gMouseClickStatus = "SPC"	'Split �����ڵ� 
    Set gActiveSpdSheet = frm1.vspdData
    If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If    
        Exit Sub
    End If

'--------------- ������ coding part(�������,Start)----------------------------------------------------
   'lgSaveRow = Row                                                 '�� 'Cookie���� ���� 
'--------------- ������ coding part(�������,End)------------------------------------------------------
    'Call parent.SetSpreadColumnValue("A",frm1.vspdData,Col,Row)
End Sub
	
'==========================================================================================
'   Event Desc : Spread Split �����ڵ� 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button = 2 And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If
End Sub

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'=======================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
	if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then	'��: ������ üũ'
    	If lgStrPrevKey <> "" Then								'���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
 			Call DbQuery
    	End If
    End If
    
End Sub

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
 '*******************************  5.1 Toolbar(Main)���� ȣ��Ǵ� Function *******************************
'	���� : Fnc�Լ��� ���� �����ϴ� ��� Function
'********************************************************************************************************* 
Function FncQuery() 
	On Error Resume Next

    FncQuery = False                                                        '��: Processing is NG
    
    Err.Clear                                                               '��: Protect system from crashing

    '-----------------------
    'Erase contents area
    '-----------------------
    Call ggoOper.ClearField(Document, "2")									'��: Clear Contents  Field
    ggoSpread.Source = frm1.vspdData
    ggoSpread.ClearSpreadData
    Call InitVariables 														'��: Initializes local global variables
    
    '-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then								'��: This function check indispensable field
       Exit Function
    End If
	IF NOT CheckOrgChangeId Then
			IntRetCD = DisplayMsgBox("124600","X","X","X")  
		Exit Function
	End if

    '-----------------------
    'Query function call area
    '-----------------------
	Call FncSetToolBar("New")
    Call DbQuery															'��: Query db data

    FncQuery = True															'��: Processing is OK

End Function


'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint() 
    Call parent.FncPrint()
End Function


'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
	Call parent.FncExport(parent.C_MULTI)
End Function


'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(parent.C_MULTI , False)                                     '��:ȭ�� ����, Tab ���� 
End Function

'========================================================================================
' Function Name : FncSplitColumn
' Function Desc : 
'========================================================================================
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub


'========================================================================================
' Function Name : FncExit
' Function Desc : 
'========================================================================================

Function FncExit()
    FncExit = True
End Function

 '*******************************  5.2 Fnc�Լ����� ȣ��Ǵ� ���� Function  *******************************
'	���� : 
'********************************************************************************************************* 


'========================================================================================
' Function Name : DbQuery
' Function Desc : This function is data query and display
'========================================================================================

Function DbQuery() 
	Dim strVal

    DbQuery = False
    
    Err.Clear                                                               '��: Protect system from crashing
	Call LayerShowHide(1)
    
    With frm1
'--------------- ������ coding part(�������,Start)----------------------------------------------------------------------------
	    strVal = BIZ_PGM_ID & "?txtBdgYear=" & ggoOper.RetFormat(Trim(.fpBdgYear.Text), "yyyy")
    	strVal = strVal & "&txtDeptCd=" & Trim(.txtDeptCd.value)
'--------------- ������ coding part(�������,End)------------------------------------------------------------------------------

'		strVal = strVal & "&lgStrPrevKey="   & lgStrPrevKey                      '��: Next key tag		
'		strVal = strVal & "&lgSelectListDT=" & GetSQLSelectListDataType("A")
'		strVal = strVal & "&lgTailList="     & MakeSQLGroupOrderByList("A")
'		strVal = strVal & "&lgSelectList="   & EnCoding(GetSQLSelectList("A"))


		With frm1
			.txtPrevKey.value       = lgStrPrevKey
			.txtSelectListDT.value  = GetSQLSelectListDataType("A")
			.txtTailList.value      = MakeSQLGroupOrderByList("A")
			.txtSelectList.value    = EnCoding(GetSQLSelectList("A"))
			.txtDeptCd_Alt.value    = .txtDeptCd.Alt
			.hOrgChangeId.value    = Trim(frm1.hOrgChangeId.Value)

		' ���Ѱ��� �߰� 
			.txtAuthBizAreaCd.value		= lgAuthBizAreaCd
			.txtInternalCd.value		= lgInternalCd
			.txtSubInternalCd.value		= lgSubInternalCd
			.txtAuthUsrID.value			= lgAuthUsrID
		End With

        'Call RunMyBizASP(MyBizASP, strVal)		'��: �����Ͻ� ASP �� ���� 
        Call ExecMyBizASP(frm1, BIZ_PGM_ID)		'��: �����Ͻ� ASP �� ���� 
    End With
    
    DbQuery = True

End Function


'========================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'========================================================================================

Function DbQueryOk()														'��: ��ȸ ������ ������� 

    '-----------------------
    'Reset variables area
    '-----------------------
	lgBlnFlgChgValue = False
    lgSaveRow        = 1
    Call ggoOper.LockField(Document, "Q")									'��: This function lock the suitable field
	Call FncSetToolBar("Query")
	
	'SetGridFocus	
	
	Set gActiveElement = document.activeElement 
End Function


'########################################################################################
'########################################################################################
'# Area Name   : User-defined Method Part
'# Description : This part declares user-defined method
'########################################################################################
'########################################################################################

'==========================================================
'���ٹ�ư ���� 
'==========================================================
Function FncSetToolBar(Cond)
	Select Case UCase(Cond)
	Case "NEW"
		Call SetToolbar("1100000000001111")
	Case "QUERY"
		Call SetToolbar("1100000000011111")
	End Select
End Function


'=======================================================================================================
'   Event Name : SetGridFocus
'   Event Desc :
'=======================================================================================================    
Sub SetGridFocus()	
   
	Frm1.vspdData.Row = 1
	Frm1.vspdData.Col = 1
	Frm1.vspdData.Action = 1
		
End Sub
'==========================================================================================
'   Event Name : CheckOrgChangeId
'   Event Desc : 
'==========================================================================================
Function CheckOrgChangeId()
	Dim strSelect, strFrom, strWhere 	
	Dim IntRetCD 
	Dim ii, jj
	Dim arrVal1, arrVal2
	Dim strStartDt, strEndDt
	Dim strYear, strMonth, strDay
 
	CheckOrgChangeId = True
 
	With frm1
	
		If LTrim(RTrim(.txtDeptCd.value)) <> "" Then
			'Call parent.ExtractDateFrom("<%=GetSvrDate%>", Parent.gServerDateFormat, parent.gServerDateType, strYear, strMonth, strDay)
			strStartDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtBdgYear.text, "01", "01")  
			strEndDt = UniConvYYYYMMDDToDate(parent.gDateFormat, frm1.txtBdgYear.text, "12", "31")  
			strStartDt =  UNIConvDateToYYYYMMDD(strStartDt, gDateFormat,Parent.gServerDateType)  
			strEndDt =  UNIConvDateToYYYYMMDD(strEndDt, gDateFormat,Parent.gServerDateType)  
			'----------------------------------------------------------------------------------------
			strSelect = "Distinct ORG_CHANGE_ID"
			strFrom =  " B_ACCT_DEPT "
			strWhere = " ORG_CHANGE_DT >= "
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(strStartDt , "''", "S") & ")"
			strWhere = strWhere & " AND ORG_CHANGE_DT <= " 
			strWhere = strWhere & " (select max(org_change_dt) from B_ACCT_DEPT where org_change_dt<= " & FilterVar(strEndDt , "''", "S") & ") "
			strWhere = strWhere & " AND ORG_CHANGE_ID =  " & FilterVar(.hOrgChangeId.value , "''", "S") & ""
			strWhere =	strWhere & " AND dept_Cd = " & FilterVar(LTrim(RTrim(frm1.txtDeptCd.value)), "''", "S")
'		msgbox "Select " & strSelect& " from " &strFrom & " where "&strWhere

			IntRetCD= CommonQueryRs(strSelect,strFrom,strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
					
			If IntRetCD = False  OR Trim(Replace(lgF0,Chr(11),"")) <> Trim(.hOrgChangeId.value) Then
					.txtDeptCd.value = ""
					.txtDeptNm.value = ""
					.hOrgChangeId.value = ""
					.txtDeptCd.focus
					CheckOrgChangeId = False
			End if
		End If
	End With
		'----------------------------------------------------------------------------------------

End Function


'��: �Ʒ� OBJECT Tag�� InterDev ����ڸ� ���Ѱ����� ���α׷��� �ϼ��Ǹ� �Ʒ� Include �ڵ�� ��ü�Ǿ�� �Ѵ� 
</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>
<!-- '#########################################################################################################
'       					6. Tag�� 
'#########################################################################################################  -->
<BODY TABINDEX="-1" SCROLL="no">
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
				<TR>
					<TD <%=HEIGHT_TYPE_02%>></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS="TD5" NOWRAP>����⵵</TD>
									<TD CLASS="TD6" NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtBdgYear" CLASS=FPDTYYYY tag="12" Title="FPDATETIME" ALT=����⵵ id=fpBdgYear></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS="TD5" NOWRAP>�μ�</TD>
									<TD CLASS="TD6" NOWRAP><INPUT NAME="txtDeptCd" MAXLENGTH="10" SIZE=10  ALT ="�μ��ڵ�" tag="11XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDeptCd" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:Call OpenDept()">
														   <INPUT NAME="txtDeptNm" MAXLENGTH="40" SIZE=20 STYLE="TEXT-ALIGN:left" ALT ="�μ���" tag="24X"></TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%" COLSPAN=2>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% TITLE="SPREAD" id=vaSpread1 tag="2"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
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
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>>
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no noresize framespacing=0 tabIndex="-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA class=hidden name=txtSelectList tag=""></TEXTAREA>
<INPUT TYPE=hidden NAME="txtPrevKey" tag="">
<INPUT TYPE=hidden NAME="txtMaxRows" tag="">
<INPUT TYPE=hidden NAME="txtTailList" tag="">
<INPUT TYPE=HIDDEN NAME="hOrgChangeId" tag="">
<INPUT TYPE=hidden NAME="txtSelectListDT" tag="">
<INPUT TYPE=hidden NAME="txtDeptCd_Alt" tag="">

<INPUT TYPE=hidden NAME="txtAuthBizAreaCd"	tag="">
<INPUT TYPE=hidden NAME="txtInternalCd"		tag="">
<INPUT TYPE=hidden NAME="txtSubInternalCd"	tag="">
<INPUT TYPE=hidden NAME="txtAuthUsrID"		tag="">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

