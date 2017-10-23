<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
'*  1. Module Name          : Human Resource
'*  2. Function Name        : ����������� 
'*  3. Program ID           : h9124ma1.asp
'*  4. Program Name         : h9124ma1.asp
'*  5. Program Desc         : �����ڷ���(�Ƿ��)
'*  6. Modified date(First) : 2001/05/17
'*  7. Modified date(Last)  : 2003/06/13
'*  8. Modifier (First)     : Bong-kyu Song
'*  9. Modifier (Last)      : Lee SiNa
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
=======================================================================================================-->
<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!-- #Include file="../../inc/IncSvrCcm.inc" -->
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/IncHRQuery.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/IncCliRdsQuery.vbs">   </SCRIPT>
<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const CookieSplit	= 1233
Const BIZ_PGM_ID	= "h9124mb1.asp"						           '��: Biz Logic ASP Name

Const TAB1 = 1										                   'Tab�� ��ġ 
Const TAB2 = 2
Const C_SHEETMAXROWS    = 15	                                      '��: Visble row

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim gSelframeFlg                                                       '���� TAB�� ��ġ�� ��Ÿ���� Flag %>
Dim gblnWinEvent                                                       'ShowModal Dialog(PopUp) Window�� ���� �� �ߴ� ���� �����ϱ� ���� 
Dim lgBlnFlawChgFlg	
Dim gtxtChargeType
Dim IsOpenPop
Dim lsInternal_cd

Dim C_MED_DT
Dim C_MED_NAME
Dim C_MED_RGST_NO
Dim C_FAMILY_NAME
Dim C_FAMILY_NAME_POP
Dim C_FAMILY_REL_CD
Dim C_FAMILY_REL_NM
Dim C_FAMILY_RES_NO
Dim C_FAMILY_TYPE_CD
Dim C_FAMILY_TYPE_NM
Dim C_MED_AMT
Dim C_PROV_CNT
Dim C_MED_TEXT
dim C_subMitCd
dim C_subMitCdNm

Dim C_YEAR_FLAG
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      =  Parent.OPMD_CMODE						        '��: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '��: Indicates that no value changed
	lgIntGrpCount     = 0										'��: Initializes Group View Size
    lgStrPrevKey      = ""                                      '��: initializes Previous Key
    lgSortKey         = 1                                       '��: initializes sort direction
		
	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	Dim strYear,strMonth,strDay

	frm1.txtYear.focus	
	Call  ggoOper.FormatDate(frm1.txtYear,  Parent.gDateFormat, 3)	
    Call  ExtractDateFrom("<%=GetSvrDate%>", Parent.gServerDateFormat, Parent.gServerDateType,strYear,strMonth,strDay)
    frm1.txtYear.Year = strYear
    frm1.txtYear.Month = strMonth
    frm1.txtYear.Day = strDay
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H",  "NOCOOKIE","MA") %>
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup���� Return�Ǵ� �� setting
'========================================================================================================
Function CookiePage(ByVal flgs)
End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)

    Dim strYear,strMonth,strDay
    lgKeyStream       = frm1.txtYear.Year & Parent.gColSep		                 'You Must append one character( Parent.gColSep)
    lgKeyStream       = lgKeyStream & Frm1.txtEmp_no.Value & Parent.gColSep         'You Must append one character( Parent.gColSep)
End Sub        

'======================================================================================================
'	Name : InitComboBox()
'	Description : Combo Display
'=======================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr
    Dim iWhere
    
     ggoSpread.Source = frm1.vspdData
     
	 ggoSpread.SetCombo "Y" & vbtab & "N"  , C_subMitCd
     ggoSpread.SetCombo "����û�ڷ�" & vbtab & "�׹����ڷ�",C_subMitCdNM
End Sub


'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
    Dim intRow
    Dim intIndex 
    Dim family_nm


    For intRow = 1 To Frm1.vspdData.MaxRows			
	    Frm1.vspdData.Row = intRow
		    
	     Frm1.vspdData.Col = C_FAMILY_NAME
	     family_nm = Frm1.vspdData.Text
		     
	     Call CommonQueryRs("  FAMILY_RES_NO, FAMILY_REL, dbo.ufn_GetCodeName('H0140',FAMILY_REL) " &_
	     "  ,CASE 	WHEN PARIA_YN ='Y' THEN 'A'  "&_
	     " 		WHEN   2005 -  CONVERT(INT,case when SUBSTRING (replace(FAMILY_RES_NO,'-',''),7,1) in(1,2) then 1900 else 2000 end  +LEFT(FAMILY_RES_NO,2))  >= 65 THEN 'B' "&_
	     " 		ELSE '' END "&_
	     "  ,CASE 	WHEN PARIA_YN ='Y' THEN '�����'  "&_
	     " 		WHEN   2005 -  CONVERT(INT,case when SUBSTRING (replace(FAMILY_RES_NO,'-',''),7,1) in(1,2) then 1900 else 2000 end  +LEFT(FAMILY_RES_NO,2))  >= 65 THEN '�����' "&_
	     " 		ELSE '' END ",_
	     "  HFA150T ",_
	     "EMP_NO =" & FilterVar(frm1.txtEmp_no.value, "''", "S")  & " AND YEAR_YY = " & FilterVar(frm1.txtYear.Year, "''", "S") & " AND FAMILY_NAME= " & FilterVar(family_nm, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	    Frm1.vspdData.Col = C_FAMILY_RES_NO
	    Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
		       	    
	    Frm1.vspdData.Col = C_FAMILY_REL_CD
	    Frm1.vspdData.Text = Trim(Replace(lgF1,Chr(11),""))

	    Frm1.vspdData.Col = C_FAMILY_REL_NM 
	    Frm1.vspdData.Text = Trim(Replace(lgF2,Chr(11),""))

	    Frm1.vspdData.Col = C_FAMILY_TYPE_CD 
	    Frm1.vspdData.Text = Trim(Replace(lgF3,Chr(11),""))

	    Frm1.vspdData.Col = C_FAMILY_TYPE_NM
	    Frm1.vspdData.Text = Trim(Replace(lgF4,Chr(11),""))	
 
   Next	
	
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()

	call InitSpreadPosVariables()

	With frm1.vspdData
		.MaxCols = C_YEAR_FLAG + 1												

		.Col = .MaxCols                                                              ' ��:��: Hide maxcols
		.ColHidden = True                                                            ' ��:��:

		.MaxRows = 0
        ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20041201",, parent.gAllowDragDropSpread

		Call GetSpreadColumnPos("A")

		.ReDraw = false

		ggoSpread.SSSetDate		C_MED_DT,			"������",   10,2, gDateFormat
        ggoSpread.SSSetEdit		C_MED_NAME,			"����ó��ȣ", 15,,, 20,1   
        ggoSpread.SSSetEdit		C_MED_RGST_NO,      "����ó����ڹ�ȣ", 14,,, 10,1   

        ggoSpread.SSSetEdit     C_FAMILY_NAME,		"��������",      10,,, 30,2
        ggoSpread.SSSetButton   C_FAMILY_NAME_POP

		ggoSpread.SSSetEdit		C_FAMILY_REL_CD,	"���������ڵ�", 10     
		ggoSpread.SSSetEdit		C_FAMILY_REL_NM,	"��������", 10       
	    ggoSpread.SSSetEdit		C_FAMILY_RES_NO,	"�ֹι�ȣ" ,11,,, 14,2        		
		ggoSpread.SSSetEdit		C_FAMILY_TYPE_CD,	"����ڱ����ڵ�", 10     
		ggoSpread.SSSetEdit		C_FAMILY_TYPE_NM,	"����ڱ���", 10         
		ggoSpread.SSSetFloat	C_MED_AMT,			"���ޱݾ�", 10,"2", ggStrIntegeralPart,  ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
		ggoSpread.SSSetFloat	C_PROV_CNT,			"���ްǼ�", 10,"6", ggStrIntegeralPart,  ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"P"
        ggoSpread.SSSetEdit		C_MED_TEXT,			"�Ƿ�񳻿�", 20,,, 50,2 
		ggoSpread.SSSetCombo C_subMitCd    , "",5
		ggoSpread.SSSetCombo C_subMitCdNM    , "���ⱸ��", 20

		ggoSpread.SSSetEdit	 C_YEAR_FLAG      ,	"�ݿ�����", 5   

		call ggoSpread.MakePairsColumn(C_FAMILY_REL_CD,C_FAMILY_REL_NM,"1")
		call ggoSpread.MakePairsColumn(C_FAMILY_TYPE_CD,C_FAMILY_TYPE_NM,"1")

        Call ggoSpread.MakePairsColumn(C_FAMILY_NAME,C_FAMILY_NAME_POP)  
        
		call ggoSpread.SSSetColHidden(C_FAMILY_REL_CD, C_FAMILY_REL_CD, true)
		call ggoSpread.SSSetColHidden(C_FAMILY_TYPE_CD, C_FAMILY_TYPE_CD, true)
		call ggoSpread.SSSetColHidden(C_subMitCd, C_subMitCd, true)
		.ReDraw = true
	
    End With

    Call SetSpreadLock 

End Sub

'======================================================================================================
' Function Name : InitSpreadPosVariables
' Function Desc : 
'=======================================================================================================
sub InitSpreadPosVariables()

	C_MED_DT			= 1
	C_MED_NAME			= 2	
	C_MED_RGST_NO		= 3
	C_FAMILY_NAME		= 4
	C_FAMILY_NAME_POP	= 5	
	C_FAMILY_REL_CD		= 6	
	C_FAMILY_REL_NM		= 7	
	C_FAMILY_RES_NO		= 8
	C_FAMILY_TYPE_CD	= 9	
	C_FAMILY_TYPE_NM	= 10				
	C_MED_AMT			= 11
	C_PROV_CNT			= 12
	C_MED_TEXT			= 13
	C_subMitCd          = 14
	C_subMitCdNM        = 15
	C_YEAR_FLAG			= 16
end sub


'======================================================================================================
' Function Name : GetSpreadColumnPos
' Function Desc : 
'=======================================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
 
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_MED_DT			= iCurColumnPos(1)
			C_MED_NAME			= iCurColumnPos(2)
			C_MED_RGST_NO		= iCurColumnPos(3)			
			C_FAMILY_NAME		= iCurColumnPos(4)
			C_FAMILY_NAME_POP	= iCurColumnPos(5)		
			C_FAMILY_REL_CD		= iCurColumnPos(6)
			C_FAMILY_REL_NM		= iCurColumnPos(7)
			C_FAMILY_RES_NO		= iCurColumnPos(8)
			C_FAMILY_TYPE_CD	= iCurColumnPos(9)
			C_FAMILY_TYPE_NM	= iCurColumnPos(10)
			C_MED_AMT			= iCurColumnPos(11)
			C_PROV_CNT			= iCurColumnPos(12)
			C_MED_TEXT			= iCurColumnPos(13)
			C_subMitCd			= iCurColumnPos(14)
			C_subMitCdnm			= iCurColumnPos(15)
			C_YEAR_FLAG			= iCurColumnPos(16)
	End Select
End sub	

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
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
'	Call InitData()
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock()

    With frm1.vspdData
    
		.ReDraw = False

		 ggoSpread.SpreadLock    C_MED_DT,   -1, C_MED_DT
		 ggoSpread.SpreadLock    C_MED_NAME,  -1, C_MED_NAME
		 
		 ggoSpread.SpreadLock    C_FAMILY_NAME, -1, C_FAMILY_NAME	
		 ggoSpread.SpreadLock    C_FAMILY_NAME_POP, -1, C_FAMILY_NAME_POP		 
		 ggoSpread.SpreadLock    C_FAMILY_REL_CD, -1, C_FAMILY_REL_CD
		 ggoSpread.SpreadLock    C_FAMILY_REL_NM, -1, C_FAMILY_REL_NM	
		 ggoSpread.SpreadLock    C_FAMILY_RES_NO, -1, C_FAMILY_RES_NO
		 ggoSpread.SpreadLock    C_FAMILY_TYPE_CD, -1, C_FAMILY_TYPE_CD
		 ggoSpread.SpreadLock    C_FAMILY_TYPE_NM, -1, C_FAMILY_TYPE_NM
		 ggoSpread.SpreadLock    C_subMitCdNM, -1, C_subMitCdNM
		 ggoSpread.SpreadLock    C_YEAR_FLAG, -1, C_YEAR_FLAG		 		 		 
		 ggoSpread.SSSetRequired C_MED_AMT,    -1, C_MED_AMT
		 ggoSpread.SSSetRequired C_PROV_CNT,    -1, C_PROV_CNT		 
		 ggoSpread.SSSetProtected  .MaxCols   , -1, -1

		.ReDraw = True

    End With

End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow, ByVal pvEndRow)

	With frm1
    
		.vspdData.ReDraw = False

		 ggoSpread.SSSetRequired		C_MED_DT, pvStartRow, pvEndRow
		 ggoSpread.SSSetRequired		C_MED_AMT, pvStartRow, pvEndRow
		 ggoSpread.SSSetRequired		C_PROV_CNT, pvStartRow, pvEndRow
		 ggoSpread.SSSetRequired		C_MED_NAME, pvStartRow, pvEndRow
		 ggoSpread.SSSetRequired		C_FAMILY_NAME, pvStartRow, pvEndRow
		  ggoSpread.SSSetRequired		C_subMitCdNM, pvStartRow, pvEndRow		 
		 ggoSpread.SSSetProtected		C_FAMILY_REL_NM, pvStartRow, pvEndRow
		 ggoSpread.SSSetProtected		C_FAMILY_RES_NO, pvStartRow, pvEndRow
		 ggoSpread.SSSetProtected		C_FAMILY_TYPE_NM , pvStartRow, pvEndRow
		 ggoSpread.SSSetProtected		C_FAMILY_REL_NM , pvStartRow, pvEndRow
		 ggoSpread.SSSetProtected		C_YEAR_FLAG, pvStartRow, pvEndRow
		 		 		 		 		 		 
		.vspdData.ReDraw = True
    
	End With

End Sub

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(���� Include ���Ͽ� ����)�� �����ʱ�ȭ �� ȭ���ʱ�ȭ�� �ϱ� ���� 
'                 �Լ��� Call�ϴ� �κ� 
'======================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '��: Clear err status
	Call LoadInfTB19029                                                             '��: Load table , B_numeric_format
	Call AppendNumberPlace("6", "3", "0")
	Call  ggoOper.FormatField(Document, "A", ggStrIntegeralPart,  ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
	
    Call  ggoOper.FormatDate(frm1.txtYear,  Parent.gDateFormat, 3)	

	Call  ggoOper.LockField(Document, "N")											'��: Lock Field
           
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables

    Call  FuncGetAuth(gStrRequestMenuID,  Parent.gUsrID, lgUsrIntCd)                                ' �ڷ����:lgUsrIntCd ("%", "1%")
    
    Call SetDefaultVal
	Call SetToolbar("1100000000001111")												'��: Set ToolBar
    
    Call InitComboBox

End Sub

'======================================================================================================
'   Event Name : Form_QueryUnload
'   Event Desc :
'======================================================================================================
Sub Form_QueryUnload(Cancel , UnloadMode )

End Sub

'======================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'======================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )

    Dim iDx
    Dim value ,strRel
    Dim family_nm ,rel_cd

   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
         Case  C_FAMILY_NAME       '������ 
            iDx = trim(Frm1.vspdData.Text)
   	        Frm1.vspdData.Col = C_FAMILY_NAME
			family_nm = Frm1.vspdData.Text

            If family_nm = "" Then
  	            Frm1.vspdData.Col = C_FAMILY_REL_CD
                Frm1.vspdData.Text = ""
  	            Frm1.vspdData.Col = C_FAMILY_REL_NM
                Frm1.vspdData.Text = ""
                
            Else

                strRel = CommonQueryRs(" FAMILY_RES_NO"," HFA150T "," medi_yn='Y' AND FAMILY_NAME = " & FilterVar(iDx, "''", "S")  & " AND EMP_NO= " & FilterVar(frm1.txtEmp_no.value, "''", "S") & " AND YEAR_YY = " & FilterVar(frm1.txtYear.Year, "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

                If strRel = false then
	        		Call DisplayMsgBox("970029","X","�ξ簡�������ڵ�Ͽ� �ش��ڰ� �Ƿ�� üũ�� �Ǿ��ִ���","X")	 
  	                Frm1.vspdData.Col = C_FAMILY_REL_CD
                    Frm1.vspdData.Text = ""
  	                Frm1.vspdData.Col = C_FAMILY_REL_NM
                    Frm1.vspdData.Text = ""
  	                Frm1.vspdData.Col = C_FAMILY_RES_NO
                    Frm1.vspdData.Text = ""
  	                Frm1.vspdData.Col = C_FAMILY_TYPE_CD
                    Frm1.vspdData.Text = ""
  	                Frm1.vspdData.Col = C_FAMILY_TYPE_NM
                    Frm1.vspdData.Text = ""                                                            
                Else
'		       	    Frm1.vspdData.Col = C_FAMILY_REL_CD
'		       	    rel_cd = Trim(Replace(lgF0,Chr(11),""))
'		       	    Frm1.vspdData.Text = rel_cd

					Call CommonQueryRs("  FAMILY_RES_NO, FAMILY_REL, dbo.ufn_GetCodeName('H0140',FAMILY_REL) " &_
					"  ,CASE 	WHEN PARIA_YN ='Y' THEN 'A'  "&_
					" 		WHEN   2005 -  CONVERT(INT,case when SUBSTRING (replace(FAMILY_RES_NO,'-',''),7,1) in(1,2) then 1900 else 2000 end  +LEFT(FAMILY_RES_NO,2))  >= 65 THEN 'B' "&_
					" 		ELSE '' END "&_
					"  ,CASE 	WHEN PARIA_YN ='Y' THEN '�����'  "&_
					" 		WHEN   2005 -  CONVERT(INT,case when SUBSTRING (replace(FAMILY_RES_NO,'-',''),7,1) in(1,2) then 1900 else 2000 end  +LEFT(FAMILY_RES_NO,2))  >= 65 THEN '�����' "&_
					" 		ELSE '' END ",_
					"  HFA150T ",_
					" EMP_NO =" & FilterVar(frm1.txtEmp_no.value, "''", "S")  & " AND YEAR_YY = " & FilterVar(frm1.txtYear.Year, "''", "S") & " AND FAMILY_NAME= " & FilterVar(family_nm, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

		       	    Frm1.vspdData.Col = C_FAMILY_RES_NO
		       	    Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
		       	    
		       	    Frm1.vspdData.Col = C_FAMILY_REL_CD
		       	    Frm1.vspdData.Text = Trim(Replace(lgF1,Chr(11),""))

		       	    Frm1.vspdData.Col = C_FAMILY_REL_NM 
		       	    Frm1.vspdData.Text = Trim(Replace(lgF2,Chr(11),""))

		       	    Frm1.vspdData.Col = C_FAMILY_TYPE_CD 
		       	    Frm1.vspdData.Text = Trim(Replace(lgF3,Chr(11),""))

		       	    Frm1.vspdData.Col = C_FAMILY_TYPE_NM
		       	    Frm1.vspdData.Text = Trim(Replace(lgF4,Chr(11),""))	
                End if 
            End if 
         
         Case Else
    End Select    

   	If Frm1.vspdData.CellType =  Parent.SS_CELL_TYPE_FLOAT Then
      If  UNICDbl(Frm1.vspdData.text) < CDbl(Frm1.vspdData.TypeFloatMin) Then
         Frm1.vspdData.text = Frm1.vspdData.TypeFloatMin
      End If
	End If
	
	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row

End Sub

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("1101111111")
    gMouseClickStatus = "SPC" 
    Set gActiveSpdSheet = frm1.vspdData

	if frm1.vspddata.MaxRows <= 0 then
		exit sub
	end if
	
	if Row <=0 then
		ggoSpread.Source = frm1.vspdData
		if lgSortkey = 1 then
			ggoSpread.SSSort Col
			lgSortKey = 2
		else
			ggoSpread.SSSort Col, lgSortkey
			lgSortKey = 1
		end if
		Exit sub
	end if
	frm1.vspdData.Row = Row
	Call SetPopupMenuItemInf("1101111111")
End Sub

'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
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
	if Frm1.vspdData.MaxRows = 0 then
		exit sub
	end if
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
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

'======================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : Combo ���� �̺�Ʈ 
'=======================================================================================================
Private Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim intIndex
	With frm1.vspdData
		.Row = Row
		Select Case Col
		    Case C_subMitCdNM
		        .Col = Col
		        intIndex = .Value 
				.Col = C_subMitCd
				.Value = intIndex
		
		
		 Case C_subMitCd
		        .Col = Col
		        intIndex = .Value 
				.Col = C_subMitCdNM
				.Value = intIndex
				
		End Select
	END WITH
   	 ggoSpread.Source = frm1.vspdData
     ggoSpread.UpdateRow Row

End Sub

'======================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'======================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
		Exit Sub
	End If
'	If frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData,NewTop) Then
'		If lgStrPrevKey <> "" Then
'			If CheckRunningBizProcess = True Then
'				Exit Sub
'			End If	
'			
'			Call DisableToolBar(Parent.TBC_QUERY)
'			If DBQuery = False Then
'				Call RestoreToolBar()
'				Exit Sub
'			End If
'		End If
'	End If  
End Sub
'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : �÷���ư�� Ŭ���� ��� �߻� 
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 Then
			Select Case Col
                Case C_FAMILY_NAME_POP
				    .Col = C_FAMILY_NAME
				    .Row = Row
                    Call OpenCode("", C_FAMILY_NAME_POP, Row)
			End Select
		End If
	End With
    
End Sub
'======================================================================================================
' Function Name : FncQuery
' Function Desc : This function is related to Query Button of Main ToolBar
'======================================================================================================
Function FncQuery() 
    Dim IntRetCD 
    Dim RetStatus
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    
    FncQuery = False                                                        
    
    Err.Clear                                                               


    ggoSpread.Source = Frm1.vspdData

	If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900013",  Parent.VB_YES_NO, "X", "X")			    '����Ÿ�� ����Ǿ����ϴ�. ��ȸ�Ͻðڽ��ϱ�?
   		If IntRetCD = vbNo Then
  			Exit Function
   		End If
	End If

    ggoSpread.ClearSpreadData
    
    If Not chkField(Document, "1") Then		
       Exit Function
    End If

    If txtEmp_no_Onchange() Then        'enter key �� ��ȸ�� ����� check�� �ش���� ������ query����...
        Exit Function
    End if

    Call InitVariables                      
    Call MakeKeyStream("X")

    If DbQuery = False Then  
		Exit Function
	End If
       
    FncQuery = True                                                              '��: Processing is OK
End Function

'======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'======================================================================================================
Function FncSave() 
    Dim IntRetCD 
    Dim intRow
    Dim strYear
	Dim strMonth
	Dim strDay
	
    FncSave = False                                                         
    
    Err.Clear                                                               
    
     ggoSpread.Source = frm1.vspdData
    If  ggoSpread.SSCheckChange = False Then
        IntRetCD =  DisplayMsgBox("900001","X","X","X")  
        Exit Function
    End If
    
     ggoSpread.Source = frm1.vspdData
    If Not chkField(Document, "2") OR Not  ggoSpread.SSDefaultCheck Then                                  '��: Check contents area
       Exit Function
    End If

    If Not chkField(Document, "1") Then									         '��: This function check required field
       Exit Function
    End If

    For intRow = 1 To Frm1.vspdData.MaxRows			
	    Frm1.vspdData.Row = intRow
	    Frm1.vspdData.Col = 0

		Select Case Frm1.vspdData.Text 
			Case  ggoSpread.InsertFlag
			    
				Frm1.vspdData.Col = C_FAMILY_RES_NO

				If Frm1.vspdData.Text = "" Then
					Call DisplayMsgBox("800489","X","���������� �ֹι�ȣ","X")
					Frm1.vspdData.action = 0   
					Exit Function
				End If
		
				Frm1.vspdData.Col = C_FAMILY_REL_NM 

				If Frm1.vspdData.Text = "" Then
					Call DisplayMsgBox("800489","X","ȯ������ REFERENCE","X")
					Frm1.vspdData.action = 0   
					Exit Function
				End If

				Frm1.vspdData.Col = C_MED_AMT 

				If Frm1.vspdData.Text = 0 Then
					Call DisplayMsgBox("800484","X","���ޱݾ�","X")
					Frm1.vspdData.action = 0   
					Exit Function
				End If

				Frm1.vspdData.Col = C_MED_DT 

				'�̰� ������.Lws
				'Call ExtractDateFrom(Frm1.vspdData.Text,parent.gServerDateFormat , parent.gServerDateType ,strYear,strMonth,strDay)

				If cint(left(UniConvDateAToB(frm1.vspdData.Text,parent.gDateFormat,parent.gServerDateFormat),4)) <> cint(Frm1.txtYear.Year) Then
					Call DisplayMsgBox("970029","X","������ �⵵","X")
					Frm1.vspdData.action = 0   
					Exit Function
				End If			
			Case  ggoSpread.UpdateFlag
				Frm1.vspdData.Col = C_MED_AMT 

				If Frm1.vspdData.Text = 0 Then
					Call DisplayMsgBox("800484","X","���ޱݾ�","X")
					Frm1.vspdData.action = 0   
					Exit Function
				End If														
		End Select
   Next		
   
    If DbSave = False Then
		Exit Function
	End If			                                                
    
    FncSave = True                                                          
    
End Function

'======================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'======================================================================================================
Function FncCopy()
    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
    
    ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
            ggoSpread.CopyRow
            SetSpreadColor .ActiveRow, .ActiveRow
	
           .Col  = C_FAMILY_NAME
           .Text = ""
           .Col  = C_FAMILY_REL_CD
           .Text = ""
           .Col  = C_FAMILY_REL_NM
           .Text = ""

            .ReDraw = True
		    .Focus
		 End If
	End With
	
    Set gActiveElement = document.ActiveElement   

End Function

'======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'======================================================================================================
Function FncCancel() 

     ggoSpread.Source = frm1.vspdData	

	 ggoSpread.EditUndo                                                  '��: Protect system from crashing

'	Call InitData

End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow(ByVal PvRowCnt) 
	Dim IntRetCD
	Dim imRow
	
	On Error Resume Next
	
	FncInsertRow = false
	if IsNumeric(Trim(pvRowCnt)) Then
		imRow = Cint(pvRowCnt)
	else
		imRow = AskSpdSheetAddRowCount()
		if imRow = "" then
			Exit function
		end if
	end if

	With frm1
        .vspdData.ReDraw = False
        .vspdData.focus
        ggoSpread.Source = .vspdData
        ggoSpread.InsertRow .vspdData.ActiveRow, imRow
        SetSpreadColor .vspdData.ActiveRow, .vspdData.ActiveRow + imRow - 1
		.vspdData.col= C_subMitCd
        .vspdData.text= "N"
        call vspdData_ComboSelChange(C_subMitCd,  .vspdData.ActiveRow)
       .vspdData.ReDraw = True
    End With
    Set gActiveElement = document.ActiveElement   
 
    if Err.number = 0 then
		FncInsertRow = true
	end if
End Function

'========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow() 
    Dim lDelRows
    If Frm1.vspdData.MaxRows < 1 then
       Exit function
	End if	
    With Frm1.vspdData 
    	.focus
    	ggoSpread.Source = frm1.vspdData 
    	lDelRows = ggoSpread.DeleteRow
    End With
    Set gActiveElement = document.ActiveElement   
End Function
'======================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'======================================================================================================
Function FncPrint()
    Call parent.FncPrint()                                                   '��: Protect system from crashing
End Function

'======================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'======================================================================================================
Function FncExcel() 
    Call parent.FncExport( Parent.C_MULTI)											 
End Function

'======================================================================================================
' Function Name : FncFind
' Function Desc : 
'======================================================================================================
Function FncFind() 
    Call parent.FncFind( Parent.C_MULTI, False)                                      
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


'======================================================================================================
' Function Name : FncExit
' Function Desc : 
'======================================================================================================
Function FncExit()
Dim IntRetCD
	
	FncExit = False

     ggoSpread.Source = frm1.vspdData	

	If  ggoSpread.SSCheckChange = True Then
		IntRetCD =  DisplayMsgBox("900016",  Parent.VB_YES_NO, "X", "X")                '����Ÿ�� ����Ǿ����ϴ�. ���� �Ͻðڽ��ϱ�?
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If

	FncExit = True
End Function
'========================================================================================================
'	Name : OpenCode()
'	Description : Item Popup���� Return�Ǵ� �� setting
'========================================================================================================
Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
        
	    Case C_FAMILY_NAME_POP

	        arrParam(0) = "�������� �˾�"			' �˾� ��Ī 
	        arrParam(1) = " HFA150T  "				 		' TABLE ��Ī 
	        arrParam(2) = frm1.txtEmp_no.value               		    ' Code Condition
	        arrParam(3) = ""							' Name Cindition
	        arrParam(4) = " medi_yn='Y' and year_yy = " & FilterVar(frm1.txtYear.Year, "''", "S") & " and emp_no = " & FilterVar(frm1.txtEmp_no.value, "''", "S") 			' Where Condition
	        arrParam(5) = "���"			    ' TextBox ��Ī 
	
            arrField(0) =  "HH"  & parent.gcolsep  & " emp_no "										' Field��(0)
            arrField(1) =  "ED21" & parent.gcolsep & " family_name "								' Field��(0)
            arrField(2) =  "ED22" & parent.gcolsep & " dbo.ufn_GetCodeName('H0140',family_rel) "	' Field��(1)
    
            arrHeader(0) = "���"					' Header��(0)
            arrHeader(1) = "��������"				' Header��(0)
            arrHeader(2) = "��������"			    ' Header��(1)
	End Select

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
    If arrRet(0) = "" Then
		frm1.vspdData.Col = C_FAMILY_NAME
		frm1.vspdData.action =0	
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
       	ggoSpread.Source = frm1.vspdData
        ggoSpread.UpdateRow Row
	End If	

End Function

'========================================================================================================
'	Name : SetCode()
'	Description : Item Popup���� Return�Ǵ� �� setting
'========================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)
	Dim strRel
	Dim family_nm , rel_nm
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	With frm1

		Select Case iWhere
		    Case C_FAMILY_NAME_POP
		        .vspdData.Col = C_FAMILY_NAME
		        family_nm = arrRet(1) 
		    	.vspdData.text = family_nm

	'	    	.vspdData.Col = C_FAMILY_REL_NM
	'	    	rel_nm = arrRet(2)
	'	    	.vspdData.text = rel_nm
 
				Call CommonQueryRs("  FAMILY_RES_NO, FAMILY_REL, dbo.ufn_GetCodeName('H0140',FAMILY_REL) " &_
				"  ,CASE 	WHEN PARIA_YN ='Y' THEN 'A'  "&_
				" 		WHEN   2005 -  CONVERT(INT,case when SUBSTRING (replace(FAMILY_RES_NO,'-',''),7,1) in(1,2) then 1900 else 2000 end  +LEFT(FAMILY_RES_NO,2))  >= 65 THEN 'B' "&_
				" 		ELSE '' END "&_
				"  ,CASE 	WHEN PARIA_YN ='Y' THEN '�����'  "&_
				" 		WHEN   2005 -  CONVERT(INT,case when SUBSTRING (replace(FAMILY_RES_NO,'-',''),7,1) in(1,2) then 1900 else 2000 end  +LEFT(FAMILY_RES_NO,2))  >= 65 THEN '�����' "&_
				" 		ELSE '' END ",_
				"  HFA150T ",_
				" EMP_NO =" & FilterVar(frm1.txtEmp_no.value, "''", "S")  & " AND YEAR_YY = " & FilterVar(frm1.txtYear.Year, "''", "S") & " AND FAMILY_NAME= " & FilterVar(family_nm, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

		       	Frm1.vspdData.Col = C_FAMILY_RES_NO
		       	Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))
	       	    
		       	Frm1.vspdData.Col = C_FAMILY_REL_CD
		       	Frm1.vspdData.Text = Trim(Replace(lgF1,Chr(11),""))

		       	Frm1.vspdData.Col = C_FAMILY_REL_NM 
		       	Frm1.vspdData.Text = Trim(Replace(lgF2,Chr(11),""))
		       	    		       	    		       	                       
		       	Frm1.vspdData.Col = C_FAMILY_TYPE_CD 
		       	Frm1.vspdData.Text = Trim(Replace(lgF3,Chr(11),""))

		       	Frm1.vspdData.Col = C_FAMILY_TYPE_NM
		       	Frm1.vspdData.Text = Trim(Replace(lgF4,Chr(11),""))			        
        End Select

	End With

End Function

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    Dim strVal
    Dim strArr
    Err.Clear                                                                        '��: Clear err status

    DbQuery = False                                                                  '��: Processing is NG
    
    if LayerShowHide(1) = false then
	    Exit Function
	end if

    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001                         '��: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '��: Query Key
    
    strVal = strVal     & "&lgCurrentSpd="       & lgCurrentSpd                   '��: Next key tag
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey              '��: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows          '��: Max fetched data

    Call RunMyBizASP(MyBizASP, strVal)                                               '��:  Run biz logic
    DbQuery = True        
    
    Call SetToolbar("1100111100111111")										        '��ư ���� ���� 

End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()	
   lgIntFlgMode =  Parent.OPMD_UMODE    
    Call  ggoOper.LockField(Document, "Q")										'��: Lock field
 '   Call InitData()
    Call SetToolbar("1100111100111111")										        '��ư ���� ���� 
	frm1.vspdData.focus
End Function
'======================================================================================================
' Function Name : DbSave
' Function Desc : This function is data query and display
'======================================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
		
    Err.Clear                                                                    '��: Clear err status
		
	DbSave = False														         '��: Processing is NG
		
	 if LayerShowHide(1) = false then
	    Exit Function
	end if

  	With Frm1
		.txtMode.value      =  Parent.UID_M0002                                            '��: Delete
		.txtKeyStream.value = lgKeyStream
	End With

    strVal  = ""
    strDel  = ""
    lGrpCnt = 1

	With Frm1
    
	 ggoSpread.Source = .vspdData

	For lRow = 1 To .vspdData.MaxRows
    
	    .vspdData.Row = lRow
	    
'üũ ���� 
	    
	    .vspdData.Col = 0

		Select Case .vspdData.Text
			Case  ggoSpread.InsertFlag									  
                                                 strVal = strVal & "C" & Parent.gColSep                   '0
                                                 strVal = strVal & lRow & Parent.gColSep                   '1
                .vspdData.Col = C_MED_DT			: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   
                .vspdData.Col = C_MED_RGST_NO		: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   
                .vspdData.Col = C_MED_NAME			: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_FAMILY_NAME		: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   
                .vspdData.Col = C_FAMILY_REL_CD		: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep 
                .vspdData.Col = C_FAMILY_RES_NO		: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   
                .vspdData.Col = C_FAMILY_TYPE_CD    : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep  
                .vspdData.Col = C_MED_AMT			: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_PROV_CNT			: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep                    
                .vspdData.Col = C_MED_TEXT			: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep  
				.vspdData.Col = C_subMitCd			: strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep  
                lGrpCnt = lGrpCnt + 1

            Case  ggoSpread.UpdateFlag
                                                 strVal = strVal & "U" & Parent.gColSep                   '0
                                                 strVal = strVal & lRow & Parent.gColSep                   '1
                .vspdData.Col = C_MED_DT			: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   
                .vspdData.Col = C_MED_RGST_NO		: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   
                .vspdData.Col = C_MED_NAME			: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   
                .vspdData.Col = C_FAMILY_NAME		: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   
                .vspdData.Col = C_FAMILY_REL_CD		: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep                  
                .vspdData.Col = C_FAMILY_RES_NO		: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   
                .vspdData.Col = C_FAMILY_TYPE_CD    : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep  
                .vspdData.Col = C_MED_AMT			: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_PROV_CNT			: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep                    
                .vspdData.Col = C_MED_TEXT			: strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep
				.vspdData.Col = C_subMitCd			: strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep 
                lGrpCnt = lGrpCnt + 1
                    
            Case  ggoSpread.DeleteFlag									
                                                 strDel = strDel & "D" & Parent.gColSep                   '0
                                                 strDel = strDel & lRow & Parent.gColSep                   '1
                .vspdData.Col = C_MED_DT			: strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep   
                .vspdData.Col = C_MED_NAME			: strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep
                .vspdData.Col = C_FAMILY_RES_NO		: strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep   
                
                lGrpCnt = lGrpCnt + 1
                
		End Select
	Next
	
	.txtMaxRows.value = lGrpCnt-1	
	.txtSpread.value = strDel & strVal
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)									
	End With

    DbSave  = True                                                               '��: Processing is NG
    
End Function

'======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncSave�� �ִ°��� �ű� 
'======================================================================================================
Function DbSaveOk()													      
	Call InitVariables

    ggoSpread.Source = Frm1.vspdData
    Frm1.vspdData.MaxRows = 0

    Call  DisableToolBar( Parent.TBC_QUERY)
    If DBQuery = false Then
		Call  RestoreToolBar()
      	Exit Function
    End If

End Function

'========================================================================================================
' Name : OpenEmpName()
' Desc : developer describe this line(grid�ܿ��� ���) 
'========================================================================================================
Function OpenEmpName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

    If  iWhere = 0 Then
	    arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtName.value			' Name Cindition
	    arrParam(2) = lgUsrIntCd                    ' �ڷ���� Condition  
    Else 'spread
        frm1.vspdData.Col = C_EMP_NO
		arrParam(0) = frm1.vspdData.Text			' Code Condition
        frm1.vspdData.Col = C_NAME
	    arrParam(1) = ""'frm1.vspdData.Text			' Name Cindition
	    arrParam(2) = lgUsrIntCd                    ' �ڷ���� Condition  
	End If

	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent, arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 0 Then
			frm1.txtEmp_no.focus
		Else
			frm1.vspdData.Col = C_EMP_NO
			frm1.vspdData.action =0
		End If	
		Exit Function
	Else
		Call SubSetCondEmp(arrRet, iWhere)
	End If	
			
End Function

'======================================================================================================
'	Name : SetCondArea()
'	Description : Item Popup���� Return�Ǵ� �� setting
'=======================================================================================================
Sub SubSetCondEmp(Byval arrRet, Byval iWhere)
	With frm1
		If iWhere = 0 Then
			.txtEmp_no.value   = arrRet(0)
			.txtName.value     = arrRet(1)
			.txtDept_nm.value  = arrRet(2)
			.txtRollPstn.value = arrRet(3)
			.txtPay_grd.value  = arrRet(4)
			.txtEntr_dt.text   = arrRet(5)
			.txtEmp_no.focus
		Else
			.vspdData.Col = C_NAME
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_EMP_NO
			.vspdData.Text = arrRet(0)
			.vspdData.action =0
		End If

		ggoSpread.Source = Frm1.vspdData
		ggoSpread.ClearSpreadData

		lgBlnFlgChgValue = False
	End With
End Sub

'========================================================================================================
'   Event Name : txtEmp_no_change             '<==�λ縶���Ϳ� �ִ� ������� Ȯ�� 
'   Event Desc :
'========================================================================================================
Function txtEmp_no_Onchange()
    Dim IntRetCd
    Dim strVal
    Dim RetStatus
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd
    
    If frm1.txtEmp_no.value = "" Then
		frm1.txtName.value = ""
		frm1.txtDept_nm.value = ""
		frm1.txtRollpstn.value = ""
		frm1.txtEntr_dt.text = ""
		frm1.txtPay_grd.value = ""

		ggoSpread.Source = Frm1.vspdData
		ggoSpread.ClearSpreadData

    Else
        
	    IntRetCd =  FuncGetEmpInf2(frm1.txtEmp_no.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
	        if  IntRetCd = -1 then
    			Call  DisplayMsgBox("800048","X","X","X")	'�ش����� �������� �ʽ��ϴ�.
            else
                Call  DisplayMsgBox("800454","X","X","X")	'�ڷῡ ���� ������ �����ϴ�.
            end if
		    frm1.txtName.value = ""
		    frm1.txtDept_nm.value = ""
		    frm1.txtRollpstn.value = ""
		    frm1.txtEntr_dt.text = ""
		    frm1.txtPay_grd.value = ""
            frm1.txtEmp_no.focus
            Set gActiveElement = document.ActiveElement
            txtEmp_no_Onchange = true

			ggoSpread.Source = Frm1.vspdData
			ggoSpread.ClearSpreadData
        Else
            frm1.txtName.value = strName
            frm1.txtDept_nm.value = strDept_nm
            frm1.txtRollpstn.value = strRoll_pstn
            frm1.txtPay_grd.value = strPay_grd1 & "-" & strPay_grd2            
            frm1.txtEntr_dt.text =  UNIDateClientFormat(strEntr_dt)
        End if 
    End if  
    
End Function 


'========================================================================================================
' Name : txtYear_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtYear_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtYear.Action = 7 
        frm1.txtYear.focus
    End If
    lgBlnFlgChgValue = True
End Sub
Sub txtYear_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub
'======================================================================================================
' Function Name : Reflect
' Function Desc : �������� �ݿ� 
'=======================================================================================================
Function Reflect() 
	Dim strVal
	Dim strYyyymm
	Dim IntRetCD

	Reflect = False                                                          '��: Processing is NG

    If txtEmp_no_Onchange() Then         'ENTER KEY �� ��ȸ�� ����� ����� CHECK �Ѵ� 
        Exit Function
    End if

	IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO,"X","X")

'    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
'        Call DisplayMsgbox("900002","X","X","X")                                '��:
'        Exit Function
'    End If
    	
	If IntRetCD = vbNo Then
		Exit Function
	End If

    '���ʵ����� �ִ��� üũ 
    IntRetCd = CommonQueryRs(" EMP_NO "," HFA030T "," YY =  " & FilterVar(Frm1.txtYear.Year , "''", "S") & " AND EMP_NO =  " & FilterVar(Frm1.txtEmp_no.Value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    If IntRetCd = False then
		Call DisplayMsgBox("800430","X",Frm1.txtYear.Year & "��","X")        '�����ڷḦ ���� �Է��ϼ���/	
		Exit Function
    End If
    
	Call LayerShowHide(1)
    
	On Error Resume Next                                                   '��: Protect system from crashing

    Call MakeKeyStream("X")
    
	strVal = BIZ_PGM_ID & "?txtMode=REFLECT" 
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '��: Query Key

	Call RunMyBizASP(MyBizASP, strVal)										'��: �����Ͻ� ASP �� ���� 

	Reflect = True                                                           '��: Processing is NG

End Function

'======================================================================================================
' Function Name : ReflectOk
' Function Desc : Reflect�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function
'======================================================================================================
Function ReflectOk()													      
	Dim IntRetCD 

	IntRetCD =DisplayMsgBox("990000","X","X","X")
	window.status = "�۾� �Ϸ�"
	call FncQuery()
End Function

Function ReflectNO()				          
	Dim IntRetCD 

    Call DisplayMsgBox("800414","X","X","X")
	window.status = "�۾� ����"

End Function

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no">
<FORM NAME=frm1 TARGET="MyBizASP" METHOD="POST">
<TABLE <%=LR_SPACE_TYPE_00%>>
	<TR >
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
								<td background="../../../Cshared/Image/table/seltab_up_bg.gif"><img src="../../../Cshared/Image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../Cshared/Image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>�����ڷ���(�Ƿ��)</font></td>
								<td background="../../../Cshared/Image/table/seltab_up_bg.gif" align="right"><img src="../../../Cshared/Image/table/seltab_up_right.gif" width="10" height="23"></td>
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
					<TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD HEIGHT=20 WIDTH=100%>
						<FIELDSET CLASS="CLSFLD">
							<TABLE <%=LR_SPACE_TYPE_40%>>
								<TR>
									<TD CLASS=TD5 NOWRAP>���꿬��</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtYear" CLASS=FPDTYYYY tag="12X1" Title="FPDATETIME" ALT="���꿬��" id=fpDateTime1> </OBJECT>');</SCRIPT>
									</TD>	
									<TD CLASS=TD5 NOWRAP>���</TD>
			     					<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" ALT="���" TYPE="Text" SiZE=15 MAXLENGTH=13 tag="12XXXU"><IMG SRC="../../../Cshared/Image/btnPopup.gif" NAME="btnEmpNo" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenEmpName('0')">
									                     <INPUT NAME="txtName" MAXLENGTH="30" SIZE="20" ALT ="����" tag="14XXXU"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�μ���</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDept_nm" MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN:left" ALT ="�μ���" tag="14">&nbsp;</TD>
									<TD CLASS=TD5 NOWRAP>��  ��</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRollPstn" MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN:left" ALT ="����" tag="14">&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>�Ի���</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtEntr_dt CLASSID=<%=gCLSIDFPDT%> ALT="�Ի���" tag="14X1" VIEWASTEXT></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>��  ȣ</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPay_grd" MAXLENGTH="20" SIZE=20 ALT ="��ȣ" tag="14">&nbsp;</TD>
								</TR>
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR HEIGHT=200>
					<TD WIDTH=100% HEIGHT=100% valign=top>
                        <TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="100%">
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%></TD>
	</TR> 
	<TR>
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%>&nbsp;&nbsp;&nbsp;���Ƿ��� �ξ簡�������ڵ�� ȭ�鿡�� �Ƿ��� üũ�� ����� �Է°����մϴ�.</TD>
	</TR> 	
	<TR HEIGHT=55>
		<TD WIDTH=100%>
			<TABLE >
				<TR><TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=100>�Ϲ��Ƿ��:</TD>
					<TD WIDTH=100><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtSum1 NAME="txtSum1" CLASS=FPDS140 tag="24X2Z" ALT="���б�������" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
					<TD WIDTH=200>&nbsp;&nbsp;����/�����/������Ƿ��:</TD>
					<TD WIDTH=100><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtSum2 NAME="txtSum2" CLASS=FPDS140 tag="24X2Z" ALT="�ʵ������" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
				<TR><TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=30><BUTTON NAME="btnSplit" CLASS="CLSMBTN" onclick="Reflect()" Flag=1>��������ݿ�</BUTTON>&nbsp;</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>			
			</TABLE>
		</TD>
	</TR>		 
	<TR>
		<TD WIDTH=100% HEIGHT=0><IFRAME NAME="MyBizASP" WIDTH=100% HEIGHT=0 FRAMEBORDER=0 SCROLLING=YES noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="lgCurrentSpd"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' ����ó��ASP�� �ѱ�� ���� ������ ��� �ִ� Tag�� %>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

