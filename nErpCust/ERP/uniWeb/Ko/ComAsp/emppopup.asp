<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
'======================================================================================================
'*  1. Module Name          : Employee Popup
'*  2. Function Name        : Employee Popup
'*  3. Program ID           : EmpPopup.asp
'*  4. Program Name         : EmpPopup.asp
'*  5. Program Desc         :
'*  6. Modified date(First) : 
'*  7. Modified date(Last)  : 2001/05/08
'*  8. Modifier (First)     : Hwang Jeong-won
'*  9. Modifier (Last)      : Lee Seok Min
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'=======================================================================================================-->
<HTML>
<HEAD>
<!--
'=======================================================================================================
'												1. �� �� �� 
'=======================================================================================================
'=======================================================================================================
'                                               1.1 Inc ����   
'	���: Inc. Include
'=======================================================================================================-->
<!-- #Include file="../inc/IncSvrCcm.inc" -->
<!-- #Include file="../inc/IncSvrHTML.inc" -->
<!--
'============================================  1.1.1 Style Sheet  =======================================
'========================================================================================================-->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../inc/SheetStyle.css">
<!--
'============================================  1.1.2 ���� Include  ======================================
'========================================================================================================-->

<SCRIPT LANGUAGE="VBScript" SRC="../inc/IncCliPAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../inc/IncCliPAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../inc/IncCliPAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../inc/IncCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript" SRC="../inc/IncHRQuery.vbs"></SCRIPT>
<Script Language="JavaScript" SRC="../inc/incImage.js"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript">
Option Explicit                                                        '��: indicates that All variables must be declared in advance
'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================
	Const BIZ_PGM_ID = "EmpPopupBiz.asp"							'��: �����Ͻ� ���� ASP�� 
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
	Const C_SHEETMAXROWS = 100								'--- ��ȭ�鿡 ���ϼ� �ִ� �ִ� Row �� 
	
	Const CODE_CON       = 0								'--- Index of Code Condition value 
	Const NAME_CON       = 1								'--- Index of Name Condition value 
	Const INTERNAL       = 2
	Const TYPE_CON       = 3
	Const WHERE_CON      = 4
		
	Dim C_EmpNo
	Dim C_EmpNm
	Dim C_DeptNm
	Dim C_Role  
	Dim C_Grade 
	Dim C_EnterDt
	Dim C_RetireDt
	Dim C_Paycd 'cyc	
	
	
<!-- #Include file="../inc/lgvariables.inc" -->		
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================

	Dim lgQueryFlag				'--- 1:New Query 0:Continuous Query

	Dim lgCode					'--- Next code
	Dim lgName					'--- Next name
	Dim lgType                   ' query type
	Dim lgStdDt
    Dim arrParent
	Dim arrParam				'--- First Parameter Group		
	Dim arrGridHdr				'--- Third Parameter Group(Column Captions of the SpreadSheet) 
	Dim arrReturn				'--- Return Parameter Group
	Dim gintDataCnt				'--- Data Counts to Query
	Dim lgInternal
	Dim whereCon
				
	arrParent = window.dialogArguments
	arrParam = arrParent(1)
	Set PopupParent = arrParent(0)
	top.document.title = "��� Popup"

'########################################################################################################
'#                       5.Method Declaration Part
'########################################################################################################

'========================================================================================================
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================

Sub InitVariables()
		
	gintDataCnt      = 8 '7 --> 8 cyc
    vspdData.MaxRows = 0

	lgQueryFlag      = "1"
			
End Sub
'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Initialize value
'========================================================================================================
Sub SetDefaultVal()
	On error resume next  ' ������ �� �� 
	
	txtCd.value = arrParam(CODE_CON)
	txtNm.value = arrParam(NAME_CON)
	lgInternal  = arrParam(INTERNAL)
	whereCon    = arrParam(WHERE_CON)	
	
	'txtDate.text =UNIDateClientFormat("<%=GetSvrDate%>") 
	txtDate.text = UNIConvDateAToB("<%=GetSvrDate%>" ,popupParent.gServerDateFormat,gDateFormat)
	
	' ���� ��ư ������ ���� �κ�. ���� �ݿ����� 
'	lgType = "0"
'	lgType      = arrParam(TYPE_CON) ' �����߻��� 0 ���� ���� 
'	Select case lgType
'		case "0"
'		case "1" ' �����ڸ� 
'			retire_check1.checked = true
'			retire_check0.disabled = true
'			retire_check1.disabled = true
'			retire_check2.disabled = true
'		case "2" ' �����ڸ� 
'			retire_check2.checked = true
'			retire_check0.disabled = true
'			retire_check1.disabled = true
'			retire_check2.disabled = true
'
'	End Select


	Self.Returnvalue = Array("")
End Sub

'========================================================================================================
' Function Name : InitSpreadPosVariables
' Function Desc : 
'========================================================================================================
sub InitSpreadPosVariables()
	C_EmpNo        = 1
	C_EmpNm        = 2
	C_DeptNm       = 3
	C_Role         = 4
	C_Grade        = 5
	C_EnterDt      = 6
	C_RetireDt     = 7
	C_Paycd        = 8	'cyc

end sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            C_EmpNo        = iCurColumnPos(1)
			C_EmpNm        = iCurColumnPos(2)
			C_DeptNm       = iCurColumnPos(3)
			C_Role         = iCurColumnPos(4)
			C_Grade        = iCurColumnPos(5)
			C_EnterDt      = iCurColumnPos(6)
			C_RetireDt     = iCurColumnPos(7)
			C_Paycd        = iCurColumnPos(8)			'cyc

    End Select    
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
    
    call InitSpreadPosVariables()
    
    vspdData.ReDraw = False
	    
    ggoSpread.Source       = vspdData
	ggoSpread.Spreadinit "V20021212", , Popupparent.gAllowDragDropSpread
    vspdData.OperationMode = 3
		
    vspdData.MaxCols = C_Paycd + 1
    vspdData.MaxRows = 0
	vspdData.Col = vspdData.MaxCols
    vspdData.ColHidden = True
    vspdData.lock = false    
	Call GetSpreadColumnPos("A")
		
	ggoSpread.SSSetEdit C_EmpNo, "���", 12    	    
    ggoSpread.SSSetEdit C_EmpNm, "����", 12
	ggoSpread.SSSetEdit C_DeptNm, "�μ���", 30
	ggoSpread.SSSetEdit C_Role, "����", 15
	ggoSpread.SSSetEdit C_Grade, "��ȣ", 15
	ggoSpread.SSSetDate C_EnterDt, "�Ի���"  , 12, 2, PopupParent.gDateFormat
	ggoSpread.SSSetDate C_RetireDt, "�����"  , 12, 2, PopupParent.gDateFormat
	ggoSpread.SSSetEdit C_Paycd, "�޿�����", 12 ' cyc
	
	ggoSpread.SSSetProtected	-1,-1,-1  		
    ggoSpread.Source = vspdData                         '2003/06/26 leejinsoo
    ggoSpread.SpreadLockWithOddEvenRowColor()
	
	vspdData.ReDraw = True
End Sub

'========================================================================================
' Function Name : LoadInfTB19029
' Function Desc : This method loads format inf
'======================================================================================== 
Sub LoadInfTB19029()
	<!-- #Include file="LoadInfTB19029.asp" -->
	<%Call loadInfTB19029A("Q", "H", "NOCOOKIE", "PA")%>
End Sub
	
'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()
	
	' �̹��� ȿ�� �ڹٽ�ũ��Ʈ �Լ� ȣ�� 
	Call MM_preloadImages("../../Cshared/image/Query.gif", "../../Cshared/image/OK.gif", "../../Cshared/image/Cancel.gif")

    Call LoadInfTB19029                           '��: Load table , B_numeric_format
		
    Call ggoOper.FormatField(Document, "1",PopupParent.ggStrIntegeralPart, PopupParent.ggStrDeciPointPart,PopupParent.gDateFormat,PopupParent.gComNum1000,PopupParent.gComNumDec)
    Call ggoOper.LockField(Document, "N")

	Call InitVariables
    Call SetDefaultVal()
	Call InitSpreadSheet()
	lgCode = txtCd.value
	lgName = txtNm.value 
	Call DbQuery()

End Sub
	
'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
End Sub



'========================================================================================
' Function Name : PopSaveSpreadColumnInf()
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopSaveSpreadColumnInf()
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
End Sub


'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)

    Call SetPopupMenuItemInf("0000111111")	
    gMouseClickStatus = "SPC" 
	
    Set gActiveSpdSheet = vspdData
	if vspddata.MaxRows <= 0 then
		exit sub
	end if
	
	if Row <=0 then
		ggoSpread.Source = vspdData
		if lgSortkey = 1 then
			ggoSpread.SSSort Col
			lgSortKey = 2
		else
			ggoSpread.SSSort Col, lgSortkey
			lgSortKey = 1
		end if
		Exit sub
	
	end if
	'vspdData.Row = Row
	
End Sub


'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = vspdData
    call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub


'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName
    if Row <= 0 then
		exit sub
	end if
	if vspdData.MaxRows = 0 then
		exit sub
	end if
	If vspdData.ActiveRow = Row Or vspdData.ActiveRow > 0 Then
		  Call OKClick()
    End If
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = vspdData
End Sub

'========================================================================================================
'   Event Name : vspdData_MouseDown
'   Event Desc : 
'========================================================================================================
Sub vspdData_MouseDown(Button , Shift , x , y)
    If Button = 2 And gMouseClickStatus = "SPC" Then
       gMouseClickStatus = "SPCR"
    End If
End Sub 

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("A")
End Sub 
Sub FncSplitColumn()

    If UCase(Trim(TypeName(gActiveSpdSheet))) = "EMPTY" Then
       Exit Sub
    End If

    ggoSpread.Source = gActiveSpdSheet
    ggoSpread.SSSetSplit(gActiveSpdSheet.ActiveCol)   

End Sub

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : 
' Desc : 
'========================================================================================================

'========================================================================================================
' Name : 
' Desc : 
'========================================================================================================
Function vspdData_KeyPress(KeyAscii)
    On Error Resume Next
    If KeyAscii = 13 And vspdData.ActiveRow > 0 Then
        Call OKClick()
    ElseIf KeyAscii = 27 Then
        Call CancelClick()
    End If
End Function

Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    '��: ���� Ű ���� ������ �� �̻� ��������ASP�� ȣ������ ���� 
    If vspdData.MaxRows < NewTop + C_SHEETMAXROWS Then
       If lgCode <> "" Or lgName <> "" Then
          Call DbQuery
       End If
    End if
End Sub
	
'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    Dim strVal, strId
    Dim arrStrDT
			
    If Not chkField(Document, "1") Then								<%'This function check indispensable field%>
		Exit Function
	End If
	arrStrDT = ""
	arrStrDT = arrStrDT & "ED" & PopupParent.gColSep  '1
	arrStrDT = arrStrDT & "ED" & PopupParent.gColSep  '2
	arrStrDT = arrStrDT & "ED" & PopupParent.gColSep  '3
	arrStrDT = arrStrDT & "ED" & PopupParent.gColSep  '4
	arrStrDT = arrStrDT & "ED" & PopupParent.gColSep  '5
	arrStrDT = arrStrDT & "DT" & PopupParent.gColSep  '6
	arrStrDT = arrStrDT & "DT" & PopupParent.gColSep  '7
	arrStrDT = arrStrDT & "ED" & PopupParent.gColSep  '8
	arrStrDT = arrStrDT & "ED" & PopupParent.gColSep  '9 cyc	

	
    DbQuery = False                                                         '��: Processing is NG


    strVal = BIZ_PGM_ID & "?txtTable=" & "haa010t , hdf020t "    
    strVal = strVal & "&txtCode=" & lgCode
    strVal = strVal & "&txtName=" & lgName
    strVal = strVal & "&txtMaxRows= " & vspdData.MaxRows

	If trim(whereCon) <>"" Then
		whereCon = " and " & whereCon
	End If
	    
    if retire_check0.checked = true then   ' ��ü  
		 strVal = strVal & "&txtWhere=" & " (haa010t.emp_no = hdf020t.emp_no and haa010t.internal_cd like " & FilterVar(lgInternal & "%" , "''", "S") & whereCon & ")"
	end if
    
    if retire_check1.checked = true then  ' �����ڸ� 
		strVal = strVal & "&txtWhere= (haa010t.retire_dt is null or haa010t.retire_dt > " &FilterVar( UniConvDateAToB(Trim(txtDate.text),gDateFormat,PopupParent.gServerDateFormat) , "''", "S") & _
             ")and haa010t.emp_no = hdf020t.emp_no and haa010t.internal_cd like " & FilterVar(lgInternal & "%" , "''", "S") & whereCon  
	end if
    
    if retire_check2.checked = true then  ' �����ڸ� 
		strVal = strVal & "&txtWhere=" & " (haa010t.retire_dt is not null AND haa010t.retire_dt <= " &FilterVar( UniConvDateAToB(Trim(txtDate.text),gDateFormat,PopupParent.gServerDateFormat) , "''", "S")& _
             ") and haa010t.emp_no = hdf020t.emp_no and haa010t.internal_cd like " &  FilterVar(lgInternal & "%" , "''", "S")  & whereCon
	end if


	strVal = strVal & "&arrField1=" & "haa010t.emp_no"
	strVal = strVal & "&arrField2=" & "haa010t.name"		
'	strVal = strVal & "&arrField3=" & "dbo.ufn_getDeptName(dept_CD, '" & UniConvDateAToB(Trim(txtDate.text),gDateFormat,PopupParent.gServerDateFormat) & "')"
	strVal = strVal & "&arrField3=" & "dbo.ufn_getDeptName(dbo.ufn_H_get_dept_cd(haa010t.emp_no, " & FilterVar( UniConvDateAToB(Trim(txtDate.text),gDateFormat,PopupParent.gServerDateFormat), "''", "S") & "), " & FilterVar( UniConvDateAToB(Trim(txtDate.text),gDateFormat,PopupParent.gServerDateFormat), "''", "S") & ")"
	strVal = strVal & "&arrField4=" & "dbo.ufn_getCodeName(" & FilterVar( "H0002", "''", "S") & ", haa010t.roll_pstn)"
	strVal = strVal & "&arrField5=" & "dbo.ufn_getCodeName(" & FilterVar( "H0001", "''", "S") & ", haa010t.pay_grd1)"
	strVal = strVal & "&arrField6=" & "CONVERT(VARCHAR(40), haa010t.entr_dt)"
	strVal = strVal & "&arrField7=" & "CONVERT(VARCHAR(40), haa010t.retire_dt)"
	strVal = strVal & "&arrField8=" & "dbo.ufn_getCodeName(" & FilterVar( "H0005", "''", "S") & ", hdf020t.pay_cd)"  'cyc
	strVal = strVal & "&arrField9=" & "dbo.ufn_H_get_dept_cd(haa010t.emp_no, " & FilterVar(UniConvDateAToB(Trim(txtDate.text),gDateFormat,PopupParent.gServerDateFormat), "''", "S") & ")"
	strVal = strVal & "&arrStrDT="  & arrStrDT
	
	strVal = strVal & "&txtCd="  & Trim(txtCd.value)
	strVal = strVal & "&txtNm="  & Trim(txtNm.value)
	
	strVal = strVal & "&Flag=" & lgQueryFlag
	
	
	Call LayerShowHide(1)
	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 
		
    DbQuery = True                                                          '��: Processing is NG
End Function
'=======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function
'=======================================================================================================
Function DbQueryOk()
   Dim IntRetCD

   If vspdData.MaxRows = 0 Then

      IntRetCD = DisplayMsgBox("900014","X","X","X") 
      If Trim(txtCd.value) > "" Then
         txtCd.Select 
         txtCd.Focus
      Else   
         txtNm.Select 
         txtNm.Focus
     End If
   End If 
   
End Function	


'========================================================================================================
' Function Name :
' Function Desc :
'========================================================================================================
Function CancelClick()
	Self.Close()
End Function

'========================================================================================================
' Name : 
' Desc : 
'========================================================================================================
Sub ConditionKeypress()
	If window.event.keyCode = 13 Then
		Call FncQuery()
	End If
End sub

'========================================================================================================
' Name : 
' Desc : 
'========================================================================================================
Function Document_onkeypress()
	If window.event.keyCode = 27 Then
        Call CancelClick()
    End If
End Function
'========================================================================================================
' Function Name : MousePointer
' Function Desc : 
'========================================================================================================
Function MousePointer(pstr1)
      Select case UCase(pstr1)
           case "PON"
				window.document.search.style.cursor = "wait"
            case "POFF"
				window.document.search.style.cursor = ""
      End Select
End Function
'========================================================================================================
' Function Name :
' Function Desc :
'========================================================================================================
Function OKClick()
	Dim intColCnt
	Dim iCurColumnPos
	If vspdData.MaxRows < 1 Then
		self.close()
		Exit Function
	End If
		
	If vspdData.ActiveRow > 0 Then	
		Redim arrReturn(vspdData.MaxCols - 1)
		
		vspdData.Row = vspdData.ActiveRow
		Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
		For intColCnt = 0 To vspdData.MaxCols - 1
			vspdData.Col = iCurColumnPos(intColCnt + 1)
			arrReturn(intColCnt) = vspdData.Text
		Next
			
		Self.Returnvalue = arrReturn
	End If
	set PopupParent = nothing
	Self.Close()
End Function

'========================================================================================================
' Function Name : FncQuery
' Function Desc : 
'========================================================================================================
Function FncQuery()

    vspdData.MaxRows = 0

	lgQueryFlag = "1"
	lgCode = Trim(txtCd.value)
	lgName = Trim(txtNm.value)
	lgStdDt = Trim(txtDate.text)

	Call DbQuery()

End Function

'========================================================================================================
'   Event Name : txtCd_onChange             '<==�λ縶���Ϳ� �ִ� ������� Ȯ�� 
'   Event Desc :
'========================================================================================================
Function txtCd_onChange()
    Dim iDx
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd
    Dim strVal
    
    If txtCd.value = "" Then
	    txtNm.value = ""
    Else
	    IntRetCd = FuncGetEmpInf2(txtCd.value,lgInternal,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			'Call PopupParent.DisplayMsgBox("800048","X","X","X")	'�ش����� �������� �ʽ��ϴ�.
            else
                Call PopupParent.DisplayMsgBox("800454","X","X","X")	'�ڷῡ ���� ������ �����ϴ�.
            end if
			txtNm.value = ""
			'txtCd.value = ""
            txtCd.focus
            txtCd_Onchange = true
            Exit Function      
        Else
            txtNm.value = strName
        End if 
    End if  
    
End Function 

'========================================================================================================
' Name : txtDate_DblClick
' Desc : �޷� Popup�� ȣ�� 
'========================================================================================================
Sub txtDate_DblClick(Button)
	If Button = 1 Then
		txtDate.Action = 7
	End If
End Sub

'=======================================================================================================
'   Event Name : txtDate_Keypress(Key)
'   Event Desc : enter key down�ÿ� ��ȸ�Ѵ�.
'=======================================================================================================
Sub txtDate_Keypress(Key)
    On Error Resume Next
	If Key = 27 Then
		Call CancelClick()
	ElseIf Key = 13 Then
		Call FncQuery()
	End If
End Sub

</SCRIPT>
<!-- #Include file="../inc/uni2kcmcom.inc" -->	
</HEAD>
<!--
======================================================================================================
'#						6. Tag ��																		#
'=======================================================================================================-->
<BODY SCROLL=NO TABINDEX="-1">
<TABLE CELLSPACING=0 CLASS="basicTB" ID="Table1">
	<TR><TD HEIGHT=40>
		<FIELDSET>		
		<TABLE CLASS="basicTB" CELLSPACING=0 HEIGHT=100% ID="Table2">
			<TR>			
				<TD CLASS="TD5" NOWRAP>��������</TD>
				<TD CLASS="TD6" NOWRAP>
					<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME=txtDate style="HEIGHT: 20px; WIDTH: 100px" tag="12" Title="FPDATETIME" ALT="��������" ID="Object1"></OBJECT>');</SCRIPT>
				</TD>
				<TD CLASS="TD5" NOWRAP>��������</TD>				
				<TD CLASS="TD6" NOWRAP><INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtretire_check" TAG="1X" VALUE="��ü" ID="retire_check0"><LABEL FOR="retire_check0">��ü</LABEL>&nbsp;
				                       <INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtretire_check" TAG="1X" VALUE="������" CHECKED ID="retire_check1" ><LABEL FOR="retire_check1">������</LABEL>
				                       <INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtretire_check" TAG="1X" VALUE="������" ID="retire_check2"><LABEL FOR="retire_check2">������</LABEL></TD>
			</TR>
			<TR>
				<TD CLASS="TD5" NOWRAP>���</TD>
				<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" Name="txtCd" SIZE=20 MAXLENGTH=13 tag="11XXXU" onkeypress="ConditionKeypress" ID="Text1"></TD>
				<TD CLASS="TD5" NOWRAP>����</TD>				
				<TD CLASS="TD6" NOWRAP><INPUT TYPE="Text" NAME="txtNm" SIZE=20 MAXLENGTH=30 tag="11" onkeypress="ConditionKeypress" ID="Text2"></TD>				
			</TR>		
		</TABLE>
		</FIELDSET>
	</TD></TR>
	<TR><TD HEIGHT=100%>
			<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"> <PARAM NAME="MaxRows" VALsUE="0"> <PARAM NAME="ReDraw" VALUE="0"> <PARAM NAME="FontSize" VALUE="10"> </OBJECT>');</SCRIPT>
	</TD></TR>
	<TR><TD HEIGHT=20>
		<TABLE CLASS="basicTB" CELLSPACING=0 ID="Table3">
			<TR>
				<TD WIDTH=70% NOWRAP>&nbsp;&nbsp;
				<IMG SRC="../../Cshared/image/query_d.gif"  Style="CURSOR: hand" ALT="Search" NAME="Search" ONCLICK="FncQuery()" onMouseOut="javascript:MM_swapImgRestore()"  onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../Cshared/image/Query.gif',1)"></IMG></TD>
				<TD WIDTH=30% ALIGN=RIGHT>
				<IMG SRC="../../Cshared/image/ok_d.gif" Style="CURSOR: hand" ALT="OK" NAME="pop1" ONCLICK="OkClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../Cshared/image/OK.gif',1)"></IMG>&nbsp;&nbsp;
				<IMG SRC="../../Cshared/image/cancel_d.gif" Style="CURSOR: hand" ALT="CANCEL" NAME="pop2" ONCLICK="CancelClick()" onMouseOut="javascript:MM_swapImgRestore()" onMouseOver="javascript:MM_swapImage(this.name,'','../../Cshared/image/Cancel.gif',1)"></IMG>&nbsp;&nbsp;</TD>
			</TR>
		</TABLE>
	</TD></TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>