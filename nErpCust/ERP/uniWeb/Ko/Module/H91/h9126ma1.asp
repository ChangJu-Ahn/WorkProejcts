<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : �������� 
*  2. Function Name        : �Ƿ�����޸����Ű� 
*  3. Program ID           : H9126ma1
*  4. Program Name         : �Ƿ�����޸����Ű� 
*  5. Program Desc         : �Ƿ�����޸����Ű� 
*  6. Comproxy List        :
*  7. Modified date(First) : 2004/12/07
*  8. Modified date(Last)  : 2004/12/13
*  9. Modifier (First)     : �ֿ�ö 
* 10. Modifier (Last)      : �ֿ�ö 
* 11. Comment              :
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/incSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/IncEB.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit

Const BIZ_PGM_ID      = "h9126mb1.asp"						           '��: Biz Logic ASP Name
Const BIZ_PGM_ID2     = "h9126mb2.asp"                                 '��: File Creation Asp Name
'Const C_SHEETMAXROWS  = 10                                      '��: Visble row
'Const C_SHEETMAXROWS1 = 10                                      '��: Visble row

'========================================================================================================
'  �������� ����Ұ� :   emp_no, max_row
'
'  emp_no  = "%"      '������������: default�� ��ȸ���Ǿ��� ������� �������.
'  max_row = 7        '������������: ��¹��� �ѷ��� �� Row �� ����  
'  med_sub = 2000000  '������������: ��������� �Ƿ�� �����ݾ��� 200�����̻��� ����� ������ ������ 
'========================================================================================================
Dim emp_no, max_row
	emp_no  = "%"
	max_row = 6
	
'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop
Dim lgStrComDateType		                                            'Company Date Type�� ����(��� Mask�� �����.)
Dim lgStrPrevKey1,lgStrPrevKey2,lgStrPrevKey3
Dim topleftOK

Dim C_RECORD_TYPE
Dim C_DATA_TYPE
Dim C_TAX
Dim C_NO
Dim C_PROV_DT
Dim C_OWN_RGST_NO_01
Dim C_HOMETAX_ID
Dim C_MAG_NO
Dim C_OWN_RGST_NO_02
Dim C_CUST_NM_FULL

Dim C_RES_NO
Dim C_NAT1
Dim C_NAME
Dim C_MED_RGST_NO
Dim C_MED_NAME
Dim C_B_COUNT
Dim C_MED_AMT
dim C_B_COUNT2
dim C_MED_AMT2

Dim C_FAMILY_REL
Dim C_FAMILY_RES_NO
Dim C_NAT2
Dim C_FAMILY_TYPE

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column  value
'========================================================================================================
Sub initSpreadPosVariables(spd) 

		C_RECORD_TYPE		= 1
		C_DATA_TYPE			= 2
		C_TAX				= 3
		C_NO				= 4
		C_PROV_DT			= 5
		C_OWN_RGST_NO_01	= 6
		C_HOMETAX_ID		= 7
		C_MAG_NO			= 8
		C_OWN_RGST_NO_02    = 9
		C_CUST_NM_FULL		= 10

		C_RES_NO			= 11
		C_NAT1				= 12
		C_NAME				= 13
		C_MED_RGST_NO		= 14
		C_MED_NAME			= 15
		C_B_COUNT			= 16
		C_MED_AMT			= 17
		C_B_COUNT			= 18
		C_MED_AMT			= 19
		C_FAMILY_REL		= 20
		C_FAMILY_RES_NO		= 21
		C_NAT2				= 22
		C_FAMILY_TYPE		= 23

End Sub
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode       = parent.OPMD_CMODE						        '��: Indicates that current mode is Create mode
	lgBlnFlgChgValue   = False								    '��: Indicates that no value changed
	lgIntGrpCount      = 0										'��: Initializes Group View Size
    lgStrPrevKey       = ""                                     '��: initializes Previous Key
    lgStrPrevKey1       = ""                                     '��: initializes Previous Key
    lgStrPrevKey2       = ""                                     '��: initializes Previous Key
    lgStrPrevKey3       = ""                                     '��: initializes Previous Key            
    lgSortKey          = 1                                      '��: initializes sort direction		
End Sub
'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================	
Sub SetDefaultVal()
 
    Dim strYear,strMonth,strDay
    Call ExtractDateFrom("<%=GetsvrDate%>",parent.gServerDateFormat , parent.gServerDateType ,strYear,strMonth,strDay)
 
    frm1.txtDt.year = strYear
    frm1.txtDt.month = "12"
    frm1.txtDt.day = "31"

    frm1.txtBas_dt.year = strYear
    frm1.txtBas_dt.month = "12"
    frm1.txtBas_dt.day = "31" 
     
End Sub	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H","NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
	Dim lgF0    

    lgKeyStream       = Trim(Frm1.txtGubun.value) & parent.gColSep       'You Must append one character(parent.gColSep)

	If Frm1.txtGubun.value = 2 Then '���νŰ��ϰ�� �������� ����ڹ�ȣ�� �Ű������� ������ �Ѵ�.
		Call CommonQueryRs("OWN_RGST_NO","HFA100T","year_area_cd = '" & frm1.txtComp_cd.value & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
		frm1.txtOwn_rgst_no.value = Trim(Replace(lgF0,Chr(11),""))
		lgKeyStream       = lgKeyStream & Trim(frm1.txtOwn_rgst_no.value) & parent.gColSep
	Else
		lgKeyStream       = lgKeyStream & Trim(frm1.txtGubun_Comp.value) & parent.gColSep    
	End If

    lgKeyStream       = lgKeyStream & Trim(frm1.txtDt.year & right("0"&frm1.txtDt.month,2)& right("0"&frm1.txtDt.day,2))& parent.gColSep
    lgKeyStream       = lgKeyStream & Trim(frm1.txtHometax_id.value) & parent.gColSep
    lgKeyStream       = lgKeyStream & Trim(frm1.txtFile.value) & parent.gColSep
    lgKeyStream       = lgKeyStream & Trim(frm1.txtBas_dt.text) & parent.gColSep
        
	IF (frm1.txtComp_type1.checked = True) Then '�����Ű��̸� ���õ� ����� �ڵ�� 
		lgKeyStream       = lgKeyStream & Trim(frm1.txtComp_cd.value) & parent.gColSep
	Else
		lgKeyStream       = lgKeyStream & "%"  & parent.gColSep           '���սŰ��̸� ��ü "%" �� 
	End If		

	Call CommonQueryRs("med_sub","HFA020t","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	frm1.txtmed_sub.value = Trim(Replace(lgF0,Chr(11),""))

	lgKeyStream       = lgKeyStream & Trim(frm1.txtmed_sub.value) & parent.gColSep

End Sub
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iNameArr , iNameArr1 , iNameArr2
    Dim iCodeArr , iCodeArr1 , iCodeArr2         
    '������ ���� 
    Call CommonQueryRs("MINOR_NM,MINOR_CD","B_MINOR","MAJOR_CD = 'H0118'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iNameArr = lgF0
    iCodeArr = lgF1   
    Call SetCombo2(frm1.txtGubun,iCodeArr,iNameArr,Chr(11)) 
        
    frm1.txtGubun.value = 2    
    Call ggoOper.SetReqAttr(frm1.txtGubun_Comp, "Q")

    '�Ű����� 
    Call CommonQueryRs("YEAR_AREA_NM,YEAR_AREA_CD","HFA100T","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iNameArr2 = lgF0
    iCodeArr2 = lgF1
    Call SetCombo2(frm1.txtComp_cd,iCodeArr2,iNameArr2,Chr(11))  


    Call change_Attr2()
End Sub
'========================================================================================================
' Name : OnChange()
' Desc : 
'========================================================================================================
Sub txtComp_cd_OnChange()
	'lgBlnFlgChgValue = True
    Call change_Attr2
End Sub
Sub change_Attr2()
	Call CommonQueryRs("HOMETAX_ID","HFA100T","year_area_cd = '" & frm1.txtComp_cd.value & "'",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	frm1.txtHomeTax_id.value = Trim(Replace(lgF0,Chr(11),""))
End Sub   

Sub txtGubun_OnChange()
	'lgBlnFlgChgValue = True
    frm1.txtGubun_Comp.value = ""
    Call change_Attr
End Sub

Sub change_Attr()
    IF frm1.txtGubun.value = 1 OR frm1.txtGubun.value = 3 Then
       Call ggoOper.SetReqAttr(frm1.txtGubun_Comp, "N")
	Else
       Call ggoOper.SetReqAttr(frm1.txtGubun_Comp, "Q")	
    End If
End Sub    

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet(strSPD)
	Dim strMaskYM
	If parent.gComDateType = "/" Then 
		lgStrComDateType = "/" & parent.gComDateType
	Else
		lgStrComDateType = parent.gComDateType
	End If
	strMaskYM = "9999" & lgStrComDateType & "99"
	
	call InitSpreadPosVariables(strSPD )

	With Frm1.vspdData
	    ggoSpread.Source = Frm1.vspdData
		ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread    		    
	   .ReDraw = false			
	   .MaxCols = C_FAMILY_TYPE + 1                                                   '��: Add 1 to Maxcols
	   .Col = .MaxCols                                                          '��: Hide maxcols
	   .ColHidden = True                                                        '��:    

	   .MaxRows = 0

		Call GetSpreadColumnPos("A")  

		'[�ڷ������ȣ]
	    ggoSpread.SSSetEdit      C_RECORD_TYPE,     "���ڵ屸��",             10
	    ggoSpread.SSSetEdit      C_DATA_TYPE,       "�ڷᱸ��",                8
	    ggoSpread.SSSetEdit      C_TAX,             "������",                 10
	    ggoSpread.SSSetEdit      C_NO,              "�Ϸù�ȣ",                8
	    ggoSpread.SSSetEdit      C_PROV_DT,         "���⿬����",             10
		    
        '[������(�븮�α���)]
	    ggoSpread.SSSetEdit      C_OWN_RGST_NO_01,  "����ڵ�Ϲ�ȣ",         13  '�ڷ��������� ����ڵ�Ϲ�ȣ 
	    ggoSpread.SSSetEdit      C_HOMETAX_ID,      "Ȩ�ؽ�ID",               11  
	    ggoSpread.SSSetEdit      C_MAG_NO,          "�������α׷��ڵ�",       14
		    
        '[��õ¡���ǹ���]
	    ggoSpread.SSSetEdit      C_OWN_RGST_NO_02,  "����ڵ�Ϲ�ȣ",         13
	    ggoSpread.SSSetEdit      C_CUST_NM_FULL,    "��ü��",                 18

        '[�ҵ���(���������û��)]
	    ggoSpread.SSSetEdit      C_RES_NO,          "�ҵ����ֹε�Ϲ�ȣ",      15
	    ggoSpread.SSSetEdit      C_NAT1,            "���ܱ��α����ڵ�",			8
	    ggoSpread.SSSetEdit      C_NAME,            "����",                     8

        '[�Ƿ�� ���޳���]
	    ggoSpread.SSSetEdit      C_MED_RGST_NO,     "����ó����ڵ�Ϲ�ȣ",    16  '����ó�� ����ڵ�Ϲ�ȣ 
	    ggoSpread.SSSetEdit      C_MED_NAME,        "����ó��ȣ",              16
	    ggoSpread.SSSetEdit      C_B_COUNT,         "�ſ�ī�� �� ���ްǼ�",                 20  '�����Ǽ���� 
	    ggoSpread.SSSetEdit      C_MED_AMT,         "�ſ�ī�� �� ���ޱݾ�",                20
	    
	    ggoSpread.SSSetEdit      C_B_COUNT2,         "�������ްǼ�",                 20  '�����Ǽ���� 
	    ggoSpread.SSSetEdit      C_MED_AMT2,         "�������ޱݾ�",                20
	  
	  
	    ggoSpread.SSSetEdit      C_FAMILY_REL,      "����",                     5
	    ggoSpread.SSSetEdit      C_FAMILY_RES_NO,   "�ֹε�Ϲ�ȣ",            13
	    ggoSpread.SSSetEdit      C_NAT2,            "���ܱ��α����ڵ�",			8
	    ggoSpread.SSSetEdit      C_FAMILY_TYPE,     "���ε��ش翩��",  18

	   .ReDraw = true
 
	   Call SetSpreadLock 
    
	End With
 	
End Sub
'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
 
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

	C_RECORD_TYPE		= iCurColumnPos(1)
	C_DATA_TYPE			= iCurColumnPos(2)
	C_TAX				= iCurColumnPos(3)
	C_NO				= iCurColumnPos(4)
	C_PROV_DT			= iCurColumnPos(5)
	C_OWN_RGST_NO_01	= iCurColumnPos(6)
	C_HOMETAX_ID		= iCurColumnPos(7)
	C_MAG_NO			= iCurColumnPos(8)
	C_OWN_RGST_NO_02    = iCurColumnPos(9)
	C_CUST_NM_FULL		= iCurColumnPos(10)
            
	C_RES_NO			= iCurColumnPos(11)
	C_NAT1				= iCurColumnPos(12)
	C_NAME				= iCurColumnPos(13)
	C_MED_RGST_NO		= iCurColumnPos(14)
	C_MED_NAME			= iCurColumnPos(15)
	C_B_COUNT			= iCurColumnPos(16)
	C_MED_AMT			= iCurColumnPos(17)
	C_B_COUNT2			= iCurColumnPos(18)
	C_MED_AMT2			= iCurColumnPos(19)
	C_FAMILY_REL		= iCurColumnPos(20)
	C_FAMILY_RES_NO		= iCurColumnPos(21)
	C_NAT2				= iCurColumnPos(22)
	C_FAMILY_TYPE		= iCurColumnPos(23)
      
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
 
     With frm1 
		.vspdData.ReDraw = False
		ggoSpread.SpreadLock      -1,-1,-1
		ggoSpread.SSSetProtected  .vspdData.MaxCols   , -1, -1
		.vspdData.ReDraw = True
    End With
                  
End Sub
'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)

End Sub
 
'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '��: Clear err status
	Call LoadInfTB19029                                                             '��: Load table , B_numeric_format		

    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'��: Lock Field

    Call InitSpreadSheet("ALL")                                                             'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables

	frm1.txtDt.focus 											'��: Set ToolBar    
	Call SetDefaultVal
	Call SetToolbar("1100000000001111")	
	Call InitComboBox
	
End Sub	
'========================================================================================================
' Name : Form_QueryUnload
' Desc : developer describe this line Called by Window_OnUnLoad() evnt
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)

End Sub
'========================================================================================================
' Name : FncQuery
' Desc : developer describe this line Called by MainQuery in Common.vbs
'========================================================================================================
Function FncQuery()
    Dim IntRetCD 
    Dim RetStatus
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    
    FncQuery = False                                                            '��: Processing is NG    
    Err.Clear                                                                   '��: Protect system from crashing

    ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgbox("900013", parent.VB_YES_NO,"X","X")			        '��: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    
	
    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    Call InitVariables															'��: Initializes local global variables
    
    If Not chkField(Document, "1") Then									        '��: This function check indispensable field
       Exit Function
    End If

    If CompareDateByFormat(frm1.txtBas_dt.Text,frm1.txtDt.Text,frm1.txtBas_dt.Alt,frm1.txtDt.Alt,"970023",frm1.txtBas_dt.UserDefinedFormat,parent.gComDateType,True) = False Then
        frm1.txtDt.focus()
        Set gActiveElement = document.activeElement
        Exit Function
    End If

    lgCurrentSpd = "A"
	topleftOK = false        

    Call MakeKeyStream(lgCurrentSpd)

    If DbQuery = False Then  
		Exit Function
	End If
       
    FncQuery = True																'��: Processing is OK
   
End Function	
'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
 															 '��: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
                                                           '��: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
                                                                  '��: Processing is OK    
End Function
'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()
    FncCopy = False    
    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If

    Set gActiveElement = document.ActiveElement   
End Function
'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel() 
    ggoSpread.EditUndo
End Function
'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow()  

    Set gActiveElement = document.ActiveElement   
End Function
'========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow()     
   Set gActiveElement = document.ActiveElement   
End Function
'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                       '��: Protect system from crashing
End Function
'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport(parent.C_SINGLE)
End Function
'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind(parent.C_SINGLE, True)
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

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()

	Dim IntRetCD
	FncExit = False
	ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgbox("900016", parent.VB_YES_NO,"X","X")			 '��: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True

End Function
 
'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    Dim strVal
    Dim strEmpno
    Dim strNo
    Dim i
    Err.Clear                                                                        '��: Clear err status

    DbQuery = False                                                                  '��: Processing is NG
    
    Call LayerShowHide(1)

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                         '��: Query
    strVal = strVal     & "&lgCurrentSpd="       & lgCurrentSpd                      '��: Next key tag
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '��: Query Key
    strVal = strVal     & "&lgStrPrevKey="       &  lgStrPrevKey

    Call RunMyBizASP(MyBizASP, strVal)                                               '��:  Run biz logic
	
    DbQuery = True                                                                   '��: Processing is NG
End Function
 
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
	Dim i
    Err.Clear                                                                    '��: Clear err status

    If  frm1.vspdData.MaxRows <= 0  Then
		Call DisplayMsgbox("900014", "X","X","X")			                            '��: ��ȸ�� �����ϼ���		
    End If	
    
    Call SetToolbar("1100000000011111")
	Call ggoOper.LockField(Document, "Q")
    Call change_Attr
	frm1.vspdData.focus
 
End Function
 
'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc :
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111")

    gMouseClickStatus = "SPC"   
    Set gActiveSpdSheet = frm1.vspdData
    
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
     If Row = 0 Then
        ggoSpread.Source = frm1.vspdData
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If
    End If
	frm1.vspdData.Row = Row 
End Sub
 
'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
		If lgStrPrevKey <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			topleftOK = true	
			lgCurrentSpd = "A"		
			
'			If DBQuery = False Then
'				Call RestoreToolBar()
'				Exit Sub
'			End If
		End If
	End If  
End Sub

'========================================================================================================
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
		Exit Sub
	End if
	
	If frm1.vspdData.MaxRows = 0 then
		Exit Sub
	End if
	
End Sub

'========================================================================================================
'   Event Name : vspdData_GotFocus
'   Event Desc : This event is spread sheet data changed
'========================================================================================================
Sub vspdData_GotFocus()
    ggoSpread.Source = Frm1.vspdData
End Sub
'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("A")
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
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)

    Dim iDx
    Dim IntRetCD
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col

    If Frm1.vspdData.CellType = parent.SS_CELL_TYPE_FLOAT Then
       If UNICDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
          Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
       End If
    End If
	
    ggoSpread.Source = frm1.vspdData
    ggoSpread.UpdateRow Row

End Sub
 
'=======================================================================================================
'   Event Name : txtDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtDt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtDt.Action = 7
        frm1.txtDt.focus
    End If
End Sub
'=======================================================================================================
'   Event Name : txtBas_dt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtBas_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtBas_dt.Action = 7
        frm1.txtBas_dt.focus
    End If
End Sub

'======================================================================================================
' Function Name : btnCb_print_onClick
' Function Desc : ����ǥ ��� 
'=======================================================================================================
Sub btnCb_print_onClick()
	Dim RetFlag
    	
    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Sub
    End If
    
    RetFlag = DisplayMsgbox("900018", parent.VB_YES_NO,"x","x")                                '�� �۾��� ����Ͻðڽ��ϱ�?
	If RetFlag = VBNO Then
		Exit Sub
	End IF
    
    Call FncBtnPreview() 
End Sub
'======================================================================================================
' Function Name : FncBtnPreview
' Function Desc : ����ǥ �̸����� 
'=======================================================================================================
Function FncBtnPreview() 
	Dim strUrl
    Dim StrEbrFile
	Dim objName
	
	Dim prov_dt, year_yy, year_area_cd
	
	StrEbrFile = "h9126oa1"

    prov_dt = UniConvDateAToB(frm1.txtDt.text,parent.gDateFormat, parent.gServerDateFormat)
	year_yy = frm1.txtBas_dt.year

	IF (frm1.txtComp_type1.checked = True) Then '�����Ű��̸� ���õ� ����� �ڵ�� 
		year_area_cd  = Trim(frm1.txtComp_cd.value) 
	Else
		year_area_cd  = "%"    '���սŰ��̸� ��ü "%" �� 
	End If		

	strUrl = "emp_no|"  & emp_no '�������� 
	strUrl = strUrl & "|max_row|" & max_row  '�������� 
	strUrl = strUrl & "|med_sub|" & Trim(frm1.txtmed_sub.value)
	strUrl = strUrl & "|prov_dt|" & prov_dt
	strUrl = strUrl & "|year_yy|" & year_yy
	strUrl = strUrl & "|year_area_cd|" & year_area_cd

	objname = AskEBDocumentName(StrEbrFile,"EBR")
	Call FncEBRPreview(objname,strUrl)
End Function

'==========================================================================================
'   Event Name : btnCb_creation_OnClick
'   Event Desc : ���ϻ���(Server)
'==========================================================================================
Function btnCb_creation_OnClick()
	Dim RetFlag
	Dim strVal
	Dim intRetCD

    Err.Clear                                                                           '��: Clear err status
    
    If Not chkField(Document, "1") Then                                                 'Required�� ǥ�õ� Element���� �Է� [��/��]�� Check �Ѵ�.
       Exit Function                            
    End If
    
    If frm1.vspdData.MaxRows <= 0  Then
		Call DisplayMsgbox("800167", "X","X","X")			                            '��: ��ȸ�� �����ϼ��� 
		Exit Function		
    End If
 
	RetFlag = DisplayMsgbox("900018", parent.VB_YES_NO,"x","x")                                '�� �۾��� ����Ͻðڽ��ϱ�?
	If RetFlag = VBNO Then
		Exit Function
	End IF

    With frm1
        Call LayerShowHide(1)					 
        lgCurrentSpd = "A"		
	    Call MakeKeyStream(lgCurrentSpd)    
	    strVal = BIZ_PGM_ID2    & "?txtMode="           & parent.UID_M0001						'��: �����Ͻ� ó�� ASP�� ���� 	    	    		    
        strVal = strVal         & "&lgCurrentSpd="      & lgCurrentSpd                  '��: Mulit�� ���� 
        strVal = strVal         & "&txtKeyStream="      & lgKeyStream                   '��: Query Key	
	   
		Call RunMyBizASP(MyBizASP, strVal)

    End With    
End Function
'==========================================================================================
'   Event Name : subVatDiskOK
'   Event Desc : ���ϻ���(Client)
'==========================================================================================
Function subVatDiskOK(ByVal pFileName) 
Dim strVal
    Err.Clear                                                                           '��: server�� ������� file�̸� 
    If Trim(pFileName) <> "" Then
	    strVal = BIZ_PGM_ID2 & "?txtMode=" & parent.UID_M0002							        '��: �����Ͻ� ó�� ASP�� ���� 
	    strVal = strVal & "&txtFileName=" & pFileName							        '��: ��ȸ ���� ����Ÿ	
	    Call RunMyBizASP(MyBizASP, strVal)										        '��: �����Ͻ� ASP �� ���� 
    End If
End Function


'=======================================================================================================
'   Event Name : txtDt_Keypress(Key)
'   Event Desc : enter key down�ÿ� ��ȸ�Ѵ�.
'=======================================================================================================
Sub txtDt_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub


'=======================================================================================================
'   Event Name : txtBas_dt_Keypress(Key)
'   Event Desc : enter key down�ÿ� ��ȸ�Ѵ�.
'=======================================================================================================
Sub txtBas_dt_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR>
		<TD <%=HEIGHT_TYPE_00%>></TD>
	</TR>
	<TR HEIGHT=23>
		<TD WIDTH="100%">
			<TABLE <%=LR_SPACE_TYPE_10%>>
				<TR>
					<TD WIDTH=10>&nbsp;</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0>
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>�Ƿ�����޸����Ű�</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* ALIGN=LIGHT>&nbsp;</TD>
					<TD WIDTH=10>&nbsp;</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
		
    <TR HEIGHT=*>
		<TD WIDTH=100% CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR><TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD></TR>
				<TR>
					<TD HEIGHT=20>
						<FIELDSET CLASS="CLSFLD">
						<TABLE <%=LR_SPACE_TYPE_40%>>
			            	<TR>
								<TD CLASS="TD5" NOWRAP>����ڵ�Ϲ�ȣ</TD>
								<TD CLASS="TD6" NOWRAP><SELECT NAME="txtGubun" ALT="�����ڱ���" STYLE="WIDTH: 100px" TAG="12N"></SELECT>
								                       <INPUT TYPE=TEXT ID="txtGubun_Comp" MAXLENGTH=50 NAME="txtGubun_Comp" SIZE=20 tag="12X2Z" ALT="�����ڻ���ڵ�Ϲ�ȣ">&nbsp;������(�븮��)</TD>
								<TD CLASS=TD5  NOWRAP>Ȩ�ؽ�ID</TD>
								<TD CLASS=TD6  NOWRAP><INPUT TYPE=TEXT ID="txtHomeTax_id" MAXLENGTH=50 NAME="txtHomeTax_id" SIZE=20 tag="11XXX" ALT="Ȩ�ؽ�ID"></TD>
							</TR>
							<TR>
								<TD CLASS=TD5  NOWRAP></TD>
								<TD CLASS=TD6  NOWRAP></TD>
								<TD CLASS=TD5  NOWRAP>���⿬����</TD>
								<TD CLASS=TD6  NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="��������" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>	
				            <TR>
								<TD CLASS="TD5" NOWRAP>�Ű�����</TD>
								<TD CLASS="TD6" NOWRAP><SELECT NAME="txtComp_cd" ALT="�Ű�����" STYLE="WIDTH: 150px" TAG="12N"></SELECT>
													   <INPUT TYPE="RADIO" CLASS="RADIO" ID="txtComp_type1" NAME="txtComp_type" TAG="21X" VALUE="Y" CHECKED><LABEL FOR="txtComp_type1">����尳���Ű�</LABEL>
													   <INPUT TYPE="RADIO" CLASS="RADIO" ID="txtComp_type2" NAME="txtComp_type" TAG="21X" VALUE="N"><LABEL FOR="txtComp_type2">��������սŰ�</LABEL></TD>
								<TD CLASS=TD5  NOWRAP>���ؿ�����</TD>
								<TD CLASS=TD6  NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtBas_dt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="���ؿ�����" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>							
								<INPUT TYPE=HIDDEN ID="txtFile" NAME="txtFile" SIZE=15 tag="14XXXU" ALT="�������ϰ��">								
						</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR><TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD></TR>
				<TR >
					<TD WIDTH="100%" HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
                            <TR HEIGHT="25%">
				            	<TD WIDTH="100%" HEIGHT=* VALIGN=TOP>
				            		<TABLE WIDTH="100%" HEIGHT="100%">
				            			<TR>
				            				<TD HEIGHT="100%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT></TD>
					            		</TR>
					            	</TABLE>
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
	        <TABLE <%=LR_SPACE_TYUPE_30%>>
	            <TR>
	                <TD WIDTH=10>&nbsp;</TD>
	                <TD>
						<BUTTON NAME="btnCb_creation" CLASS="CLSMBTN">���ϻ���</BUTTON>&nbsp;
						<BUTTON TYPE=HIDDEN NAME="btnCb_print" CLASS="CLSMBTN">���޸������</BUTTON></TD>
	                <TD WIDTH=* ALIGN="right"></TD>
	                <TD WIDTH=10>&nbsp;</TD>
	            </TR>
	        </TABLE>
	    </TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=0><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=0 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioKind" tag="24">
<INPUT TYPE=HIDDEN NAME="txtRadioType" tag="24">
<INPUT TYPE=HIDDEN NAME="txtCheck" tag="24">
<INPUT TYPE=HIDDEN NAME="txtOwn_rgst_no" tag="24">
<INPUT TYPE=HIDDEN NAME="txtmed_sub" tag="24">

</FORM>

<DIV ID="MousePT" NAME="MousePT">
	<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>

<FORM NAME=EBAction TARGET="MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">	
</FORM>

</BODY>
</HTML>
