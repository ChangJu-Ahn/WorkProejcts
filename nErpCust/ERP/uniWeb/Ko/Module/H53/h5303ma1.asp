<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        :
*  3. Program ID           : h5303ma1
*  4. Program Name         : �ǰ����迬�����꺸���Ѿ� ��ȸ 
*  5. Program Desc         : �ǰ����迬�����꺸���Ѿ� ��ȸ 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/30
*  8. Modified date(Last)  : 2003/06/11
*  9. Modifier (First)     : TGS(CHUN HYUNG WON)
* 10. Modifier (Last)      : Lee SiNa
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript"	SRC="../../inc/TabScript.js"></SCRIPT>

<Script Language="VBScript">
Option Explicit
'========================================================================================================
'=                       4.2 Constant variables
'========================================================================================================
Const BIZ_PGM_ID      = "h5303mb1.asp"						           '��: Biz Logic ASP Name
Const BIZ_PGM_ID2     = "h5303mb2.asp"						           '��: Disk make
Const BIZ_PGM_ID3     = "h5303mb3.asp"						           '��: getFile
'Const BIZ_PGM_ID4     = "h5303mb4.asp"						           '��: getFile

Const C_SHEETMAXROWS    = 23	                                      '��: Visble row

Const TAB1 = 1																		'��: Tab�� ��ġ 
Const TAB2 = 2
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
Dim lgOldRow
Dim lsInternal_cd

Dim C_COUNT          '���� 
Dim C_COMP_NO        '������ȣ 
Dim C_NUMBER         '���� 
Dim C_ACCOUNT        'ȸ�� 
Dim C_AREA           '��������� 
Dim C_MED_INSUR_NO   '����ȣ 
Dim C_NAME           '���� 
Dim C_RES_NO         '�ֹε�Ϲ�ȣ 
Dim C_MED_ACQ_DT     '�ڰ������ 
Dim C_SUB_TOT_CNT    '���⵵����ᳳ�ο��� 
Dim C_SUB_TOT_AMT    '���⵵����ᳳ���Ѿ� 
Dim C_INCOME_TOT_AMT '���⵵�����Ѿ� 
Dim C_WORK_MONTH_AMT '�ٹ����� 

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  

	C_COUNT          = 1 '����							   'Column constant for Spread Sheet
	C_COMP_NO        = 2 '������ȣ 
	C_NUMBER         = 3 '���� 
	C_ACCOUNT        = 4 'ȸ�� 
	C_AREA           = 5 '��������� 
	C_MED_INSUR_NO   = 6 '����ȣ 
	C_NAME           = 7 '���� 
	C_RES_NO         = 8 '�ֹε�Ϲ�ȣ 
	C_MED_ACQ_DT     = 9 '�ڰ������ 
	C_SUB_TOT_CNT    = 10 '���⵵����ᳳ�ο��� 
	C_SUB_TOT_AMT    = 11 '���⵵����ᳳ���Ѿ� 
	C_INCOME_TOT_AMT = 12 '���⵵�����Ѿ� 
	C_WORK_MONTH_AMT = 13 '�ٹ����� 
End Sub
'========================================================================================================
' Name : InitVariables()
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = Parent.OPMD_CMODE						        '��: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '��: Indicates that no value changed
	lgIntGrpCount     = 0										'��: Initializes Group View Size
    lgStrPrevKey      = ""                                      '��: initializes Previous Key
    lgSortKey         = 1                                       '��: initializes sort direction
	lgOldRow = 0
	lsInternal_cd     = ""

	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False
End Sub

'========================================================================================================
' Name : SetDefaultVal()
' Desc : Set default value
'========================================================================================================

Sub SetDefaultVal()
	Dim strYear
	Dim strMonth
	Dim strDay

	Call ExtractDateFrom("<%=GetsvrDate%>",Parent.gServerDateFormat , Parent.gServerDateType ,strYear,strMonth,strDay)
'	frm1.txtbase_yy.Focus
	
	frm1.txtbase_yy.Year = strYear 		'��� default value setting
	frm1.txtbase_yy.Month = strMonth 
	frm1.txtbase_yy.Day = strDay

	frm1.txtbase_yy1.Year = strYear 		'��� default value setting
	frm1.txtbase_yy1.Month = strMonth 
	frm1.txtbase_yy1.Day = strDay
	
'	frm1.txtbase_yy1.Year = "2005"
'	frm1.txtbase_yy.Year = "2005"
'	frm1.txtSect_cd.value = "1001"
'	frm1.txtSect_cd1.value = "1001"	
'	frm1.txtComp_no.value = "1"
'	frm1.txtNumber.value = "2"
'	frm1.txtAccount.value = "3"
'	frm1.txtArea.value = "4"
	
    lgBlnFlgChgValue = false
End Sub

'========================================================================================================
' Name : LoadInfTB19029()
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("Q", "H", "NOCOOKIE", "MA") %>
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
Sub MakeKeyStream(pOpt)
    
    If gSelframeFlg = TAB1 Then    
		lgKeyStream       = Trim(Frm1.txtbase_yy.Year) & Parent.gColSep       'You Must append one character(Parent.gColSep)     
		lgKeyStream       = lgKeyStream & Trim(Frm1.txtSect_cd.value) & Parent.gColSep
		lgKeyStream       = lgKeyStream & lgUsrIntCd & Parent.gColSep
		lgKeyStream       = lgKeyStream & Trim(Frm1.txtComp_no.value) & Parent.gColSep
		lgKeyStream       = lgKeyStream & Trim(Frm1.txtNumber.value) & Parent.gColSep
		lgKeyStream       = lgKeyStream & Trim(Frm1.txtAccount.value) & Parent.gColSep
		lgKeyStream       = lgKeyStream & Trim(Frm1.txtArea.value) & Parent.gColSep 
	Else
		lgKeyStream       = Trim(Frm1.txtbase_yy1.Year) & Parent.gColSep   
		lgKeyStream       = lgKeyStream & Trim(Frm1.txtSect_cd1.value) & Parent.gColSep				    	
		lgKeyStream       = lgKeyStream & Trim(Frm1.hFileName.value) & Parent.gColSep 	  
	End If
			
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()

	Call initSpreadPosVariables()   'sbk 

	With frm1.vspdData

        ggoSpread.Source = frm1.vspdData
        ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk

	   .ReDraw = false

       .MaxCols   = C_WORK_MONTH_AMT + 1                                                  ' ��:��: Add 1 to Maxcols
	   .Col       = .MaxCols                                                              ' ��:��: Hide maxcols
       .ColHidden = True                                                                  ' ��:��:

       .MaxRows = 0

       Call GetSpreadColumnPos("A") 'sbk
			Call AppendNumberPlace("6","13","0")
			Call AppendNumberPlace("7","15","0")
			
            ggoSpread.SSSetEdit     C_COUNT,            "����", 6,,,6,2
            ggoSpread.SSSetEdit     C_COMP_NO,          "������ȣ", 8,,,8,2
            ggoSpread.SSSetEdit     C_NUMBER,           "����", 3,,,1,2
            ggoSpread.SSSetEdit     C_ACCOUNT,          "ȸ��", 5,,,2,2
            ggoSpread.SSSetEdit     C_AREA,             "���������", 10,,,3,2
            ggoSpread.SSSetEdit     C_MED_INSUR_NO,     "����ȣ", 11,,,11,2
            ggoSpread.SSSetEdit     C_NAME,             "����", 20,,,16,2
	        ggoSpread.SSSetEdit		C_RES_NO,	        "�ֹε�Ϲ�ȣ", 13,,,13,2
	        ggoSpread.SSSetEdit     C_MED_ACQ_DT,       "�ڰ������", 8,,,8,2
            ggoSpread.SSSetEdit     C_SUB_TOT_CNT,       "���⵵����ᳳ�ο���", 17,,,2
            
'	        ggoSpread.SSSetEdit     C_SUB_TOT_AMT,       "���⵵����ᳳ���Ѿ�", 17,,,13
'	        ggoSpread.SSSetEdit     C_INCOME_TOT_AMT,    "���⵵�����Ѿ�", 14,,,15
			ggoSpread.SSSetFloat    C_SUB_TOT_AMT,		 "���⵵����ᳳ���Ѿ�" ,  17,"6",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"
			ggoSpread.SSSetFloat    C_INCOME_TOT_AMT,	 "���⵵�����Ѿ�" ,  15,"7",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gComNum1000,Parent.gComNumDec,,,"Z"	        	        
	        ggoSpread.SSSetEdit     C_WORK_MONTH_AMT,    "�ٹ�����", 10,,,2

	   .ReDraw = true

       Call SetSpreadLock

    End With

End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
      ggoSpread.Source = frm1.vspdData
'      ggoSpread.SpreadLockWithOddEvenRowColor()
	If gSelframeFlg = TAB1 Then   	
		ggoSpread.SpreadLock C_COUNT,			-1 ,  -1
		ggoSpread.SpreadLock C_COMP_NO,			-1 ,  -1
		ggoSpread.SpreadLock C_NUMBER,			-1 ,  -1
		ggoSpread.SpreadLock C_ACCOUNT,			-1 ,  -1
		ggoSpread.SpreadLock C_AREA,			-1 ,  -1
		ggoSpread.SpreadLock C_MED_INSUR_NO,	-1 ,  -1
		ggoSpread.SpreadLock C_NAME,			-1 ,  -1
		ggoSpread.SpreadLock C_RES_NO,			-1 ,  -1
		ggoSpread.SpreadLock C_MED_ACQ_DT,		-1 ,  -1

		ggoSpread.SpreadUnLock C_SUB_TOT_CNT,		-1 , -1
		ggoSpread.SpreadUnLock C_SUB_TOT_AMT,		-1 , -1
		ggoSpread.SpreadUnLock C_INCOME_TOT_AMT,	-1 , -1
		ggoSpread.SpreadUnLock C_WORK_MONTH_AMT,	-1 , -1
		ggoSpread.SSSetRequired C_SUB_TOT_CNT,		-1 , -1
		ggoSpread.SSSetRequired C_SUB_TOT_AMT,		-1 , -1
		ggoSpread.SSSetRequired C_INCOME_TOT_AMT,	-1 , -1
		ggoSpread.SSSetRequired C_WORK_MONTH_AMT,	-1 , -1
	Else
		ggoSpread.SpreadLock C_COUNT,			-1 ,  -1
		ggoSpread.SpreadLock C_COMP_NO,			-1 ,  -1
		ggoSpread.SpreadLock C_NUMBER,			-1 ,  -1
		ggoSpread.SpreadLock C_ACCOUNT,			-1 ,  -1
		ggoSpread.SpreadLock C_AREA,			-1 ,  -1
		ggoSpread.SpreadLock C_MED_INSUR_NO,	-1 ,  -1
		ggoSpread.SpreadLock C_NAME,			-1 ,  -1
		ggoSpread.SpreadLock C_RES_NO,			-1 ,  -1
		ggoSpread.SpreadLock C_MED_ACQ_DT,		-1 ,  -1
		ggoSpread.SpreadLock C_SUB_TOT_CNT,		-1 , -1
		ggoSpread.SpreadLock C_SUB_TOT_AMT,		-1 , -1

		ggoSpread.SSSetRequired C_INCOME_TOT_AMT,	-1 , -1
		ggoSpread.SSSetRequired C_WORK_MONTH_AMT,	-1 , -1	
	End If		
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1

    .vspdData.ReDraw = False
    .vspdData.ReDraw = True

    End With
End Sub

'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,Parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> Parent.UC_PROTECTED Then
              Frm1.vspdData.Col = iDx
              Frm1.vspdData.Row = iRow
              Frm1.vspdData.Action = 0 ' go to
              Exit For
           End If

       Next

    End If
End Sub

'========================================================================================
' Function Name : GetSpreadColumnPos
' Description   : 
'========================================================================================
Sub GetSpreadColumnPos(ByVal pvSpdNo)
    Dim iCurColumnPos
    
    Select Case UCase(pvSpdNo)
       Case "A"
            ggoSpread.Source = frm1.vspdData
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
            
			C_COUNT          = iCurColumnPos(1)
			C_COMP_NO        = iCurColumnPos(2)
			C_NUMBER         = iCurColumnPos(3)
			C_ACCOUNT        = iCurColumnPos(4)
			C_AREA           = iCurColumnPos(5)
			C_MED_INSUR_NO   = iCurColumnPos(6)
			C_NAME           = iCurColumnPos(7)
			C_RES_NO         = iCurColumnPos(8)
			C_MED_ACQ_DT     = iCurColumnPos(9)
			C_SUB_TOT_CNT    = iCurColumnPos(10)
			C_SUB_TOT_AMT    = iCurColumnPos(11)
			C_INCOME_TOT_AMT = iCurColumnPos(12)
			C_WORK_MONTH_AMT = iCurColumnPos(13)
            
    End Select    
End Sub
'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '��: Clear err status
	Call LoadInfTB19029                                                             '��: Load table , B_numeric_format

    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)
    Call ggoOper.FormatDate(frm1.txtbase_yy, Parent.gDateFormat, 3)
    Call ggoOper.FormatDate(frm1.txtbase_yy1, Parent.gDateFormat, 3)
	Call ggoOper.LockField(Document, "N")											'��: Lock Field
    Call InitSpreadSheet                                                            'Setup the Spread sheet

    Call InitVariables                                                              'Initializes local global variables
    Call FuncGetAuth(gStrRequestMenuID, Parent.gUsrID, lgUsrIntCd)     ' �ڷ����:lgUsrIntCd ("%", "1%")
    Call SetDefaultVal

	Call SetToolbar("1100000000011111")												'��: Set ToolBar
	Call ClickTab1 
	Call CookiePage (0)                                                             '��: Check Cookie

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

    FncQuery = False                                                            '��: Processing is NG
    Err.Clear                                                                   '��: Protect system from crashing
	ggoSpread.ClearSpreadData
 '   Call InitVariables															'��: Initializes local global variables

'    If Not chkField(Document, "1") Then									        '��: This function check indispensable field
'       Exit Function
 '   End If
 

	If gSelframeFlg = TAB1 Then   	        
 		If trim(frm1.txtbase_yy.text) = "" Then
			call DisplayMsgBox("970029","X" , frm1.txtbase_yy.Alt, "X") 
			frm1.txtbase_yy.focus	
			Exit Function
		End If	
			
		If trim(frm1.txtSect_cd.value) = "" Then
			call DisplayMsgBox("970029","X" , frm1.txtSect_cd.Alt, "X") 	
			frm1.txtSect_cd.focus
			Exit Function
		End If
			
		If trim(frm1.txtComp_no.value) = "" Then
			call DisplayMsgBox("970029","X" , frm1.txtComp_no.Alt, "X")
			frm1.txtComp_no.focus 	
			Exit Function
		End If
		
		If trim(frm1.txtNumber.value) = "" Then
			call DisplayMsgBox("970029","X" , frm1.txtNumber.Alt, "X") 
			frm1.txtNumber.focus	
			Exit Function
		End If

			If trim(frm1.txtAccount.value) = "" Then
			call DisplayMsgBox("970029","X" , frm1.txtAccount.Alt, "X") 
			frm1.txtAccount.focus	
			Exit Function
		End If	
		
		If trim(frm1.txtArea.value) = "" Then
			call DisplayMsgBox("970029","X" , frm1.txtArea.Alt, "X") 
			frm1.txtArea.focus
			Exit Function
		End If										
	Else
		If trim(frm1.txtbase_yy1.text) = "" Then
			call DisplayMsgBox("970029","X" , frm1.txtbase_yy1.Alt, "X") 	
			frm1.txtbase_yy1.focus
			Exit Function
		End If
			
		If trim(frm1.txtSect_cd1.value) = "" Then
			call DisplayMsgBox("970029","X" , frm1.txtSect_cd1.Alt, "X") 	
			frm1.txtSect_cd1.focus
			Exit Function
		End If			    
	End If
	
    Call MakeKeyStream("X")
  
	Call RemovedivTextArea 	
    If DbQuery = False Then
		Exit Function
	End If   																'��: Query db data

    FncQuery = True																'��: Processing is OK
End Function
'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    
    FncSave = False                                                              '��: Processing is NG
    
    Err.Clear                                                                    '��: Clear err status
    
     ggoSpread.Source = frm1.vspdData
    If lgBlnFlgChgValue = False And  ggoSpread.SSCheckChange = False Then
        IntRetCD =  DisplayMsgBox("900001","X","X","X")                           '��: No data changed!!
        Exit Function
    End If

	If gSelframeFlg = TAB1 Then   		
		If trim(frm1.txtbase_yy.text) = "" Then
			call DisplayMsgBox("970029","X" , frm1.txtbase_yy.Alt, "X") 
			frm1.txtbase_yy.focus	
			Exit Function
		End If
			
		If trim(frm1.txtSect_cd.value) = "" Then
			call DisplayMsgBox("970029","X" , frm1.txtSect_cd.Alt, "X") 	
			frm1.txtSect_cd.focus
			Exit Function
		End If
		
		If trim(frm1.txtComp_no.value) = "" Then
			call DisplayMsgBox("970029","X" , frm1.txtComp_no.Alt, "X")
			frm1.txtComp_no.focus 	
			Exit Function
		End If
		
		If trim(frm1.txtNumber.value) = "" Then
			call DisplayMsgBox("970029","X" , frm1.txtNumber.Alt, "X") 
			frm1.txtNumber.focus	
			Exit Function
		End If

			If trim(frm1.txtAccount.value) = "" Then
			call DisplayMsgBox("970029","X" , frm1.txtAccount.Alt, "X") 
			frm1.txtAccount.focus	
			Exit Function
		End If	
		
		If trim(frm1.txtArea.value) = "" Then
			call DisplayMsgBox("970029","X" , frm1.txtArea.Alt, "X") 
			frm1.txtArea.focus
			Exit Function
		End If	
				
		Dim strWhere, strEmp_no
	
		strWhere = " (med_insur_no is null or ltrim(med_insur_no) ='') "
		strWhere = strWhere & " AND RES_NO IN (SELECT RES_NO FROM HDB030T "
		strWhere = strWhere & "					WHERE DIV = " & FilterVar(gSelframeFlg, "''", "S")
		strWhere = strWhere & "						AND YEAR_YY = " & FilterVar(Frm1.txtbase_yy.Year, "''", "S")
		strWhere = strWhere & "						AND BIZ_AREA_CD	= " & FilterVar(Frm1.txtSect_cd.value, "''", "S") & ")"

 		IntRetCD = CommonQueryRs(" HDF020T.emp_no "," HDF020T JOIN HAA010T ON HDF020T.EMP_NO = HAA010T.EMP_NO ", strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

		If IntRetCD = True  Then
	
			strEmp_no = Trim(Replace(lgF0,Chr(11),","))
    
			Call DisplayMsgbox("971012", "X","�޿������� ȭ�鿡 " & left(strEmp_no,len(strEmp_no)-1) & " ����� �ǰ����� ��ȣ�� �ԷµǾ� �ִ���","X")	
			Exit Function
		End If	
										
	Else
		If trim(frm1.txtbase_yy1.text) = "" Then
			call DisplayMsgBox("970029","X" , frm1.txtbase_yy1.Alt, "X") 
			frm1.txtbase_yy1.focus	
			Exit Function
		End If
			
		If trim(frm1.txtSect_cd1.value) = "" Then
			call DisplayMsgBox("970029","X" , frm1.txtSect_cd1.Alt, "X") 	
			frm1.txtSect_cd1.focus
			Exit Function
		End If
				
'		If trim(frm1.txtFileName2.value) = "" Then
'			call DisplayMsgBox("970029","X" , frm1.txtFileName2.Alt, "X")
'			frm1.txtFileName2.focus 	
'			Exit Function
'		Else
'		
'			if (ggoSaveFile.fileExists(frm1.hFileName.value) = 0)  = false  then
'				IntRetCD = DisplayMsgBox("115191","x","x","x")                           '��:There is no picture
'				Exit Function
'			end if
'			
'		End If
	End If
	 
    If Not chkField(Document, "2") Then
       Exit Function
    End If
  
	ggoSpread.Source = frm1.vspdData
    If Not  ggoSpread.SSDefaultCheck Then                                         '��: Check contents area
       Exit Function
    End If
    
    Call MakeKeyStream("X")
    Call  DisableToolBar( parent.TBC_SAVE)
	If DBSave=False Then
	   Call  RestoreToolBar()
	   Exit Function
	End If
    
    FncSave = True                                                              '��: Processing is OK
    
End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
	Dim strVal, strDel
    Err.Clear                                                                    '��: Clear err status
		
	DbSave = False														         '��: Processing is NG
		
    If LayerShowHide(1)=False Then
		Exit Function
    End If

	With frm1
		.txtMode.value        = parent.UID_M0002                                        '��: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '��: Save Key
	End With

    strVal = ""
    strDel = ""
    lGrpCnt = 1

	With Frm1
   
       For lRow = 1 To .vspdData.MaxRows
    
           .vspdData.Row = lRow
           .vspdData.Col = 0
        
           Select Case .vspdData.Text
               Case ggoSpread.InsertFlag                                      '��: insert�߰� 
																		   strVal = strVal & "C" & parent.gColSep 'array(0)
																		   strVal = strVal & lRow & parent.gColSep
	If gSelframeFlg = TAB1 Then
                                                                           strVal = strVal & Trim(.txtbase_yy.year) &  parent.gColSep
                                                                           strVal = strVal & Trim(.txtSect_cd.value) &  parent.gColSep
	Else 
                                                                           strVal = strVal & Trim(.txtbase_yy1.year) &  parent.gColSep
                                                                           strVal = strVal & Trim(.txtSect_cd1.value) &  parent.gColSep
	End If
                    .vspdData.Col = C_COUNT								 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_COMP_NO							 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_NUMBER							 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_ACCOUNT							 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_AREA								 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_MED_INSUR_NO                       : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_NAME								 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_RES_NO							 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_MED_ACQ_DT						 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_SUB_TOT_CNT						 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_SUB_TOT_AMT						 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_INCOME_TOT_AMT                     : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_WORK_MONTH_AMT					 : strVal = strVal & Trim(.vspdData.Text) &  parent.gRowSep
                    lGrpCnt = lGrpCnt + 1                                                               
               Case  ggoSpread.UpdateFlag                                      '��: Update
                                                                           strVal = strVal & "U" &  parent.gColSep
                                                                           strVal = strVal & lRow &  parent.gColSep
	If gSelframeFlg = TAB1 Then
                                                                           strVal = strVal & Trim(.txtbase_yy.year) &  parent.gColSep
                                                                           strVal = strVal & Trim(.txtSect_cd.value) &  parent.gColSep
	Else 
                                                                           strVal = strVal & Trim(.txtbase_yy1.year) &  parent.gColSep
                                                                           strVal = strVal & Trim(.txtSect_cd1.value) &  parent.gColSep
	End If
                    .vspdData.Col = C_RES_NO							 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_MED_ACQ_DT						 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_SUB_TOT_CNT						 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_SUB_TOT_AMT						 : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_INCOME_TOT_AMT                     : strVal = strVal & Trim(.vspdData.Text) &  parent.gColSep
                    .vspdData.Col = C_WORK_MONTH_AMT					 : strVal = strVal & Trim(.vspdData.Text) &  parent.gRowSep
			 
                    lGrpCnt = lGrpCnt + 1
           End Select
       Next
	   .txtMode.value        = parent.UID_M0002
	   .txtMaxRows.value     = lGrpCnt-1	
	   .txtSpread.value      = strDel & strVal
	   .txtTab.value		 = gSelframeFlg
  
	End With
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '��: Processing is NG
End Function
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    ggoSpread.Source = Frm1.vspdData    
	ggoSpread.ClearSpreadData      
	Call RemovedivTextArea    	
    Call InitVariables															'��: Initializes local global variables
    
	Call DisableToolBar(parent.TBC_QUERY)
    If DbQuery = False Then
        Call RestoreToolBar()
        Exit Function
    End If
End Function
'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel()
    ggoSpread.Source = frm1.vspdData
    ggoSpread.EditUndo
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                    '��: Protect system from crashing
End Function

'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel()
	Call Parent.FncExport(Parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind()
	Call Parent.FncFind(Parent.C_SINGLE, True)
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
' Function Name : PopSaveSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopSaveSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.SaveSpreadColumnInf()
End Sub

'========================================================================================
' Function Name : PopRestoreSpreadColumnInf
' Description   : 
'========================================================================================
Sub PopRestoreSpreadColumnInf()
    ggoSpread.Source = gActiveSpdSheet
    Call ggoSpread.RestoreSpreadInf()
    Call InitSpreadSheet()      
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()

	Dim IntRetCD
	FncExit = False
	ggoSpread.Source = frm1.vspdData
    If  ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"X","X")			 '��: Data is changed.  Do you want to exit?
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
    Err.Clear                                                                    '��: Clear err status

    DbQuery = False                                                              '��: Processing is NG

    If LayerShowHide(1) = false Then
        Exit Function
    End If
    
    strVal =""
    
	strVal = BIZ_PGM_ID & "?txtMode="			& Parent.UID_M0001                     '��: Query
	strVal = strVal     & "&txtKeyStream="		& lgKeyStream                   '��: Query Key
	strVal = strVal     & "&lgStrPrevKey="		& lgStrPrevKey             '��: Next key tag
	strVal = strVal     & "&txtMaxRows="		& Frm1.vspdData.MaxRows         '��: Max fetched data
	strVal = strVal     & "&txtTab="			& gSelframeFlg 
	
    Call RunMyBizASP(MyBizASP, strVal)                                           '��:  Run biz logic

    DbQuery = True                                                               '��: Processing is NG
End Function
'======================================================================================================
' Function Name : DbQueryOk
' Function Desc : DbQuery�� �������� ��� MyBizASP ���� ȣ��Ǵ� Function, ���� FncQuery�� �ִ°��� �ű� 
'=======================================================================================================%>
Function DbQueryOk()													<%'��ȸ ������ ������� %>
	
    lgIntFlgMode = Parent.OPMD_UMODE
    Call ggoOper.LockField(Document, "Q")									<%'This function lock the suitable field%>
    Call SetToolbar("1100100000011111")	
	frm1.vspdData.focus	
End Function

'======================================================================================================
' Function Name : FileOK
' Function Desc : 
'=======================================================================================================%>
Function FileOK()													<%'��ȸ ������ ������� %>
	
    Dim lRow
    With Frm1

        .vspdData.ReDraw = false

        ggoSpread.Source = .vspdData
       For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0
            .vspdData.Text = ggoSpread.InsertFlag
        Next
 
'      ggoSpread.SpreadLock C_CHANG_DT, -1,C_CHANG_DT
    .vspdData.ReDraw = TRUE
    ggoSpread.ClearSpreadData "T"            
    End With    
    Call SetToolbar("1100100000011111")	
    lgStrPrevKey = ""
    Set gActiveElement = document.ActiveElement   	
End Function

'========================================================================================================
' Name : FncOpenPopup
' Desc : developer describe this line
'========================================================================================================
Function FncOpenPopup(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)
    Dim strFrom_dt
    Dim strTo_dt

	If IsOpenPop = True  Then
	   Exit Function
	End If

	IsOpenPop = True
	Select Case iWhere
	    Case "1"
	    	arrParam(2) = frm1.txtSect_cd.value        			' Code Condition
	    Case "2"
	    	arrParam(2) = frm1.txtSect_cd1.value        			' Code Condition
	End Select
	
	arrParam(0) = "������ڵ� �˾�"			        ' �˾� ��Ī 
	arrParam(1) = " HFA100T "						    ' TABLE ��Ī 
	arrParam(3) = ""'frm1.txtSect_nm.value				' Name Cindition	
	arrParam(4) = ""                      		    	' Where Condition
	arrParam(5) = "������ڵ�" 			            ' TextBox ��Ī 

	arrField(0) = "YEAR_AREA_CD"						    	' Field��(0)
	arrField(1) = "YEAR_AREA_NM"    					    	' Field��(1)

	arrHeader(0) = "������ڵ�"	   		    	    ' Header��(0)
	arrHeader(1) = "������"	    		            ' Header��(1)
	    	
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False


	If arrRet(0) = "" Then
		Select Case iWhere
		    Case "1"
		    	frm1.txtSect_cd.focus
		    Case "2"
		    	frm1.txtSect_cd1.focus
		End Select	
		
		Exit Function
	Else
		Select Case iWhere
		    Case "1"
		        Frm1.txtSect_cd.value = arrRet(0)
		        Frm1.txtSect_nm.value = arrRet(1)
		        Frm1.txtSect_cd.focus
		    Case "2"
		        Frm1.txtSect_cd1.value = arrRet(0)
		        Frm1.txtSect_nm1.value = arrRet(1)
		        Frm1.txtSect_cd1.focus		        
        End Select
	End If

End Function

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

End Sub

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row)

    Dim iDx

   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

   	If Frm1.vspdData.CellType = Parent.SS_CELL_TYPE_FLOAT Then
      If UNICDbl(Frm1.vspdData.text) < UNICDbl(Frm1.vspdData.TypeFloatMin) Then
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

    Call SetPopupMenuItemInf("0000111111")
    
    gMouseClickStatus = "SPC"   

    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
   	    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData
       
       If lgSortKey = 1 Then
           ggoSpread.SSSort Col               'Sort in ascending
           lgSortKey = 2
       Else
           ggoSpread.SSSort Col, lgSortKey    'Sort in descending 
           lgSortKey = 1
       End If
       
       Exit Sub
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
    End If
    
    If frm1.vspdData.MaxRows = 0 Then
        Exit Sub
    End If
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
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

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
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
'   	If OldLeft <> NewLeft Then
'		Exit Sub
'	End If
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

'======================================================================================================
'   Event Name : txtSect_cd_OnChange
'   Event Desc : ������ڵ尡 ����� ��� 
'=======================================================================================================
Function txtSect_cd_OnChange()
    Dim IntRetCd
    Dim strWhere

    If Trim(frm1.txtSect_cd.Value) = "" Then
        frm1.txtSect_nm.Value=""
    Else    
        strWhere = " biz_area_cd=" & FilterVar(frm1.txtSect_cd.Value, "''", "S")

        IntRetCD = CommonQueryRs(" biz_area_cd,biz_area_nm "," b_biz_area ",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCD=False And Trim(frm1.txtSect_cd.Value)<>""  Then
            Call DisplayMsgBox("800054","X","X","X")                         '�� : ��ϵ��� ���� �ڵ��Դϴ�.
            frm1.txtSect_nm.Value=""
            ggoSpread.ClearSpreadData
            frm1.txtSect_nm.focus 
            txtSect_cd_OnChange = True
        Else
            frm1.txtSect_nm.Value=Trim(Replace(lgF1,Chr(11),""))
        End If
    End If
   lgBlnFlgChgValue = true    
End Function

'======================================================================================================
'   Event Name : txtSect_cd1_OnChange
'   Event Desc : ������ڵ尡 ����� ��� 
'=======================================================================================================
Function txtSect_cd1_OnChange()
    Dim IntRetCd
    Dim strWhere

    If Trim(frm1.txtSect_cd1.Value) = "" Then
        frm1.txtSect_nm1.Value=""
    Else    
        strWhere = " biz_area_cd=" & FilterVar(frm1.txtSect_cd1.Value, "''", "S")

        IntRetCD = CommonQueryRs(" biz_area_cd,biz_area_nm "," b_biz_area ",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
        If IntRetCD=False And Trim(frm1.txtSect_cd1.Value)<>""  Then
            Call DisplayMsgBox("800054","X","X","X")                         '�� : ��ϵ��� ���� �ڵ��Դϴ�.
            frm1.txtSect_nm1.Value=""
            ggoSpread.ClearSpreadData
            frm1.txtSect_nm1.focus 
            txtSect_cd1_OnChange = True
        Else
            frm1.txtSect_nm1.Value=Trim(Replace(lgF1,Chr(11),""))
        End If
    End If
   lgBlnFlgChgValue = true    
End Function
'======================================================================================================
'   Event Name : txtbase_yy_Change
'   Event Desc : ����⵵�� ����� ��� 
'=======================================================================================================
Function txtbase_yy_Change()
    lgBlnFlgChgValue = true
End Function
'======================================================================================================
'   Event Name : txtComp_no_OnChange
'   Event Desc : ������ȣ�� ����� ��� 
'=======================================================================================================
Function txtComp_no_OnChange()
	lgBlnFlgChgValue = true    
End Function
'======================================================================================================
'   Event Name : txtNumber_OnChange 
'   Event Desc : ������ ����� ��� 
'=======================================================================================================
Function txtNumber_OnChange()
    lgBlnFlgChgValue = true
End Function
'======================================================================================================
'   Event Name : txtAccount_OnChange
'   Event Desc : ȸ�谡 ����� ��� 
'=======================================================================================================
Function txtAccount_OnChange()
    lgBlnFlgChgValue = true
End Function
'======================================================================================================
'   Event Name : txtArea_OnChange
'   Event Desc : ����������� ����� ��� 
'=======================================================================================================
Function txtArea_OnChange()
    lgBlnFlgChgValue = true
End Function
'=======================================================================================================
'   Event Name : txtYear_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================================================================================
Sub txtbase_yy_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtbase_yy.Action = 7
        frm1.txtbase_yy.focus
    End If
End Sub
'=======================================================================================================
'   Event Name : txtbase_yy_Keypress(Key)
'   Event Desc : 3rd party control���� Enter Ű�� ������ ��ȸ ���� 
'=======================================================================================================
Sub txtbase_yy_Keypress(Key)
    If Key = 13 Then
        Call MainQuery()
    End If
End Sub


'======================================================================================================
' Function Name : btnCb_print_onClick
' Function Desc : ����ǥ ��� 
'=======================================================================================================
Sub btnCb_print_onClick()
	Dim RetFlag ,RetFlag2

    If frm1.vspdData.MaxRows <= 0 Then
		Call DisplayMsgbox("800167", "X","X","X")			                            '��: ��ȸ�� �����ϼ��� 
		Exit Sub
    End If
    	
    'If Not chkField(Document, "1") Then									'��: This function check indispensable field
     '  Exit Sub
    'End If
    
    RetFlag = DisplayMsgbox("900018", parent.VB_YES_NO,"x","x")                                '�� �۾��� ����Ͻðڽ��ϱ�?
	If RetFlag = VBNO Then
		Exit Sub
	End IF
    
    Call FncBtnPrint() 
End Sub
'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : ����ǥ ��� 
'=======================================================================================================
Function FncBtnPrint() 
	Dim strUrl
	Dim StrEbrFile    
	Dim objName
    
	Dim base_yy, Sect_cd, Comp_no, Number, Account, Area
	
	StrEbrFile = "h5303oa1_1"

	base_yy  = frm1.txtbase_yy.Year
	Sect_cd = frm1.txtSect_cd.value
	Comp_no = frm1.txtComp_no.value
	Number  = frm1.txtNumber.value
	Account = frm1.txtAccount.value
	Area    = frm1.txtArea.value

	strUrl = "base_yy|" & base_yy
	strUrl = strUrl & "|Sect_cd|" & Sect_cd 
	strUrl = strUrl & "|Comp_no|" & Comp_no
'	strUrl = strUrl & "|Number|" & Number
	strUrl = strUrl & "|Account|" & Account
	strUrl = strUrl & "|Area|" & Area
		
    objname = AskEBDocumentName(StrEbrFile,"EBR")

'    Call FncEBRPrint(EBAction,objname,strUrl)        'prient
	Call FncEBRpreview(objname,strUrl)               'prewiew

End Function

'======================================================================================================
' Function Name : btnCb_print2_onClick
' Function Desc : ����ǥ����� ��� 
'=======================================================================================================
Sub btnCb_print2_onClick()
	Dim RetFlag , RetFlag2

    If frm1.vspdData.MaxRows <= 0 Then
		Call DisplayMsgbox("800167", "X","X","X")			                            '��: ��ȸ�� �����ϼ��� 
		Exit Sub
    End If
  
    RetFlag = DisplayMsgbox("900018", parent.VB_YES_NO,"x","x")                                '�� �۾��� ����Ͻðڽ��ϱ�?
	If RetFlag = VBNO Then
		Exit Sub
    Else
        Call FloppyDiskLabelForm()      '����ǥ����� 
	End IF
End Sub
'======================================================================================================
' Function Name : FloppyDiskLabelForm
' Function Desc : ����ǥ����� 
'=======================================================================================================
Function FloppyDiskLabelForm()
	Dim strUrl	
    Dim StrEbrFile
	Dim objName
	Dim base_yy, count
	
	StrEbrFile = "h5303oa1_2"

	base_yy  = frm1.txtbase_yy.Year
	count    = frm1.vspdData.MaxRows

	strUrl = "base_yy|" & base_yy
	strUrl = strUrl & "|Count|" & Count 

    objname = AskEBDocumentName(StrEbrFile,"EBR")

'   Call FncEBRPrint(EBAction,objname,strUrl)        'prient
	Call FncEBRpreview(objname,strUrl)               'prewiew

	
End Function
'==========================================================================================
'   Event Name : btnCb_creation_OnClick
'   Event Desc : ���ϻ���(Server)
'==========================================================================================
Function btnCb_creation_OnClick()
	Dim RetFlag ,RetFlag2
	Dim strVal
	Dim intRetCD

    Err.Clear                                                                           '��: Clear err status
'	If gSelframeFlg = TAB1 Then      
'		If Not chkField(Document, "1") Then                                                 'Required�� ǥ�õ� Element���� �Է� [��/��]�� Check �Ѵ�.
'		   Exit Function                            
'		End If
'	End If
	
    If frm1.vspdData.MaxRows <= 0 Then
		Call DisplayMsgbox("800167", "X","X","X")			                            '��: ��ȸ�� �����ϼ��� 
		Exit Function		
    End If
 
	RetFlag = DisplayMsgbox("900018", parent.VB_YES_NO,"x","x")                         '�� �۾��� ����Ͻðڽ��ϱ�?
	If RetFlag = VBNO Then
		Exit Function
	End IF

    With frm1
        Call LayerShowHide(1)					 
        lgCurrentSpd = "A"		
	    Call MakeKeyStream(lgCurrentSpd)  
	      
		If gSelframeFlg = TAB1 Then   	        
			strVal = BIZ_PGM_ID2    & "?txtMode="           & parent.UID_M0001						'��: �����Ͻ� ó�� ASP�� ���� 	    	    		    
			strVal = strVal         & "&lgCurrentSpd="      & lgCurrentSpd                  '��: Mulit�� ���� 
			strVal = strVal         & "&txtKeyStream="      & lgKeyStream                   '��: Query Key	
			strVal = strVal         & "&txtTab="      & gSelframeFlg
			Call RunMyBizASP(MyBizASP, strVal)			
	    Else
    
			frm1.txtMode.value  =  parent.UID_M0001
			frm1.txtTab.value	= gSelframeFlg			
			Call ExecMyBizASP(frm1, BIZ_PGM_ID2)
 		   				    
		End If
    End With    
End Function

'==========================================================================================
'   Event Name : btnCb_select_OnClick
'   Event Desc : ������ �������� 
'==========================================================================================
Function btnCb_select_OnClick()
	Dim RetFlag ,RetFlag2
	Dim strVal
	Dim intRetCD,strWhere, strEmp_no

    Err.Clear                                                                           '��: Clear err status
'	If gSelframeFlg = TAB1 Then      
'		If Not chkField(Document, "1") Then                                                 'Required�� ǥ�õ� Element���� �Է� [��/��]�� Check �Ѵ�.
'		   Exit Function                            
'		End If
'	End If
		
	If gSelframeFlg = TAB1 Then
		If trim(frm1.txtbase_yy.text) = "" Then
			call DisplayMsgBox("970029","X" , frm1.txtbase_yy.Alt, "X") 	
			frm1.txtbase_yy.focus
			Exit Function
		End If
	
		If trim(frm1.txtSect_cd.value) = "" Then
			call DisplayMsgBox("970029","X" , frm1.txtSect_cd.Alt, "X") 	
			frm1.txtSect_cd.focus
			Exit Function
		End If		
		
		strWhere = " YEAR_YY = " & FilterVar(Frm1.txtbase_yy.Year, "''", "S")
		strWhere = strWhere & " AND DIV = " & FilterVar(gSelframeFlg, "''", "S")
		strWhere = strWhere & " AND BIZ_AREA_CD	= " & FilterVar(Frm1.txtSect_cd.value, "''", "S")
			
	Else
		If trim(frm1.txtbase_yy1.text) = "" Then
			call DisplayMsgBox("970029","X" , frm1.txtbase_yy1.Alt, "X") 	
			frm1.txtbase_yy1.focus
			Exit Function
		End If
	
		If trim(frm1.txtSect_cd1.value) = "" Then
			call DisplayMsgBox("970029","X" , frm1.txtSect_cd1.Alt, "X") 	
			frm1.txtSect_cd1.focus
			Exit Function
		End If
			
		If trim(frm1.txtFileName2.value) = "" Then
			call DisplayMsgBox("970029","X" , frm1.txtFileName2.Alt, "X")
			frm1.txtFileName2.focus 	
			Exit Function
		Else
		
			if (ggoSaveFile.fileExists(frm1.hFileName.value) = 0)  = false  then
				IntRetCD = DisplayMsgBox("115191","x","x","x")                           '��:There is no picture
				Exit Function
			end if
			
		End If
		strWhere = " YEAR_YY = " & FilterVar(Frm1.txtbase_yy1.Year, "''", "S")
		strWhere = strWhere & " AND DIV = " & FilterVar(gSelframeFlg, "''", "S")
		strWhere = strWhere & " AND BIZ_AREA_CD	= " & FilterVar(Frm1.txtSect_cd1.value, "''", "S")		
	End If
	 
 	IntRetCD = CommonQueryRs(" * "," HDB030T ", strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 	
	If IntRetCD = True  Then
		        
		IntRetCD = DisplayMsgBox("800502", 35,"X","X")	    '�̹� ������ �ڷᰡ �ֽ��ϴ�.�����Ͻðڽ��ϱ�?
		If IntRetCD = vbCancel Then
		   	Exit Function
		End If
	End If
					
'	RetFlag = DisplayMsgbox("900018", parent.VB_YES_NO,"x","x")                         '�� �۾��� ����Ͻðڽ��ϱ�?
'	If RetFlag = VBNO Then
'		Exit Function
'	End IF
	ggoSpread.ClearSpreadData
	
    With frm1
        Call LayerShowHide(1)					 
        lgCurrentSpd = "A"		
	    Call MakeKeyStream(lgCurrentSpd)  
	      
		If gSelframeFlg = TAB1 Then   	        
			strVal = BIZ_PGM_ID    & "?txtMode="           & parent.UID_M0003						'��: �����Ͻ� ó�� ASP�� ���� 	    	    		    
			strVal = strVal         & "&lgCurrentSpd="      & lgCurrentSpd                  '��: Mulit�� ���� 
			strVal = strVal         & "&txtKeyStream="      & lgKeyStream                   '��: Query Key	
			Call RunMyBizASP(MyBizASP, strVal)
	    Else
			strVal = BIZ_PGM_ID3    & "?txtMode="           & parent.UID_M0003						'��: �����Ͻ� ó�� ASP�� ���� 	    	    		    
			strVal = strVal         & "&lgCurrentSpd="      & lgCurrentSpd                  '��: Mulit�� ���� 
			strVal = strVal         & "&txtKeyStream="      & lgKeyStream                   '��: Query Key	
			Call RunMyBizASP(MyBizASP, strVal)
		End If
    End With    
End Function

Sub DBAutoQueryOk()
    Dim lRow
    With Frm1

        .vspdData.ReDraw = false
        ggoSpread.Source = .vspdData
       For lRow = 1 To .vspdData.MaxRows
            .vspdData.Row = lRow
            .vspdData.Col = 0

            .vspdData.Text = ggoSpread.InsertFlag
        Next

'      ggoSpread.SpreadLock C_CHANG_DT, -1,C_CHANG_DT
    .vspdData.ReDraw = TRUE
    ggoSpread.ClearSpreadData "T"            
    End With    
    Call SetToolbar("1100100000011111")	
    lgStrPrevKey = ""
    Set gActiveElement = document.ActiveElement   
End Sub


'==========================================================================================
'   Event Name : subVatDiskOK
'   Event Desc : ���ϻ���(Client)
'==========================================================================================
Function subVatDiskOK(ByVal pFileName) 

	Dim strVal
    Err.Clear                                                                           '��: server�� ������� file�̸� 
    If Trim(pFileName) <> "" Then
    
		If gSelframeFlg = TAB1 Then   	        
			strVal = BIZ_PGM_ID2 & "?txtMode=" & parent.UID_M0002							        '��: �����Ͻ� ó�� ASP�� ���� 
			strVal = strVal & "&txtFileName=" & pFileName							        '��: ��ȸ ���� ����Ÿ	
			Call RunMyBizASP(MyBizASP, strVal)				
	    Else
			strVal = BIZ_PGM_ID2 & "?txtMode=" & parent.UID_M0002							        '��: �����Ͻ� ó�� ASP�� ���� 
			strVal = strVal & "&txtFileName=" & pFileName							        '��: ��ȸ ���� ����Ÿ	
			Call RunMyBizASP(MyBizASP, strVal)		    
		End If
			    
    End If
End Function

'========================================================================================
' Function Name : ClickTab1
' Function Desc : This function tab1 click
'========================================================================================
Function ClickTab1()
	If gSelframeFlg = TAB1 Then Exit Function
	Call changeTabs(TAB1)	 '~~~ ù��° Tab 
	gSelframeFlg = TAB1
	ggoSpread.ClearSpreadData
'	Call SetDefaultVal()
'	frm1.cboIOFlag.focus
	document.all("divButton").style.VISIBILITY = "visible"	
    Call SetSpreadLock
	frm1.txtSect_cd.value = frm1.txtSect_cd1.value
	frm1.txtSect_nm.value = frm1.txtSect_nm1.value
	frm1.txtbase_yy.text = frm1.txtbase_yy1.text	    
End Function

'========================================================================================
' Function Name : ClickTab2
' Function Desc : This function tab2 click
'========================================================================================
Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
	Call changeTabs(TAB2)	 '~~~ �ι�° Tab 
	gSelframeFlg = TAB2
	ggoSpread.ClearSpreadData
'	Call SetDefaultVal()
'	frm1.cboIOFlag2.focus 

	document.all("divButton").style.VISIBILITY = "hidden"
    Call SetSpreadLock	
	frm1.txtSect_cd1.value = frm1.txtSect_cd.value
	frm1.txtSect_nm1.value = frm1.txtSect_nm.value
	frm1.txtbase_yy1.text = frm1.txtbase_yy.text

End Function

'===============================================================================================
'   by Shin hyoung jae 
'	Name : GetOpenFilePath()
'	Description : GetTextFilePath	
'================================================================================================= 
Function GetOpenFilePath()
	Dim dlg
    Dim sPath

	On Error Resume Next
	Set dlg = CreateObject("uni2kCM.SaveFile")
	
	If Err.Number <> 0 Then
		Msgbox Err.Number & " : " & Err.Description
	End If
	
    sPath = dlg.GetOpenFilePath()

	If Err.Number <> 0 Then
		Msgbox Err.Number & " : " & Err.Description
	End If

	If gSelframeFlg = TAB1 Then 
		lgFilePath = sPath
		frm1.txtFileName.Value = ExtractFileName(sPath)
    ElseIf gSelframeFlg = TAB2 Then 
		lgFilePath2 = sPath
		frm1.txtFileName2.Value = ExtractFileName(sPath)
    End If
    Set dlg = Nothing
	frm1.hFileName.value = sPath		
End Function

Function ExtractFileName(byVal strPath)
	strPath = StrReverse(strPath)
	strPath = Left(strPath, InStr(strPath, "\") - 1)
	ExtractFileName = StrReverse(strPath)
End Function
'========================================================================================
 ' Function Name : RemovedivTextArea
 ' Function Desc : ������, �������� ������ HTML ��ü(TEXTAREA)�� Clear���� �ش�.
'========================================================================================
 Function RemovedivTextArea()
 
 	Dim ii
 		
 	For ii = 1 To divTextArea.children.length
 	    divTextArea.removeChild(divTextArea.children(0))
 	Next
 
 End Function
</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->
</HEAD>

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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>�ǰ����迬�������Ѿ׽Ű�</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 onclick="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><IMG src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>�������Ͽ�������Ű�</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><IMG src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23" ></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=*>&nbsp;</TD>					
				</TR>
			</TABLE>
		</TD>
		
	</TR>

    <TR HEIGHT=*>
		<TD CLASS="Tab11">
		
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR><TD <%=HEIGHT_TYPE_02%> WIDTH=100%></TD></TR>
				<TR>
					<TD HEIGHT=20>
					  <FIELDSET CLASS="CLSFLD">

		<!--ù��° TAB  -->
		<DIV ID="TabDiv"  SCROLL="no">
					   <TABLE <%=LR_SPACE_TYPE_40%>>
						    <TR>
							    <TD CLASS=TD5 NOWRAP>����⵵</TD>
			                    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtbase_yy" style="HEIGHT: 20px; WIDTH: 50px" TAG="12X1" ALT="����⵵" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>�Ű�����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSect_cd"  SIZE=10  MAXLENGTH=10  ALT ="�Ű�����"   TAG="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: FncOpenPopup(1)">
								                     <INPUT NAME="txtSect_nm"  SIZE=20  MAXLENGTH=100 ALT ="�Ű������" TAG="14XXXU"></TD>
							</TR>
							<TR>
							    <TD CLASS=TD5 NOWRAP>������ȣ</TD>
			                    <TD CLASS=TD6 NOWRAP><INPUT NAME="txtComp_no" TYPE = TEXT SIZE=10  MAXLENGTH=8 ALT ="������ȣ"   TAG="12XXXU"></TD>
								<TD CLASS=TD5 NOWRAP>����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtNumber" TYPE = TEXT SIZE=10  MAXLENGTH=1  ALT ="����"   TAG="12XXXU"></TD>
							</TR>
							<TR>
							    <TD CLASS=TD5 NOWRAP>ȸ��</TD>
			                    <TD CLASS=TD6 NOWRAP><INPUT NAME="txtAccount" TYPE = TEXT SIZE=10  MAXLENGTH=2  ALT ="ȸ��"   TAG="12XXXU"></TD>
								<TD CLASS=TD5 NOWRAP>���������</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtArea" TYPE = TEXT SIZE=10  MAXLENGTH=3  ALT ="���������"   TAG="12XXXU"></TD>
							</TR>
					  </TABLE>
		</div>
		<!--��° TAB  -->
		<DIV ID="TabDiv"  SCROLL="no">		
					   <TABLE <%=LR_SPACE_TYPE_40%>>
							    <TD CLASS=TD5 NOWRAP>����⵵</TD>
			                    <TD CLASS=TD6 NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 NAME="txtbase_yy1" style="HEIGHT: 20px; WIDTH: 50px" TAG="12X1" ALT="����⵵" Title="FPDATETIME"></OBJECT>');</SCRIPT></TD>
								<TD CLASS=TD5 NOWRAP>�Ű�����</TD>
								<TD CLASS=TD6 NOWRAP><INPUT NAME="txtSect_cd1"  SIZE=10  MAXLENGTH=10  ALT ="�Ű�����"   TAG="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnPlantCd1" ALIGN=top TYPE="BUTTON" ONCLICK="VBScript: FncOpenPopup(2)">
								                     <INPUT NAME="txtSect_nm1"  SIZE=20  MAXLENGTH=100 ALT ="�Ű������" TAG="14XXXU"></TD>

						    <TR>
									<TD CLASS="TD5">ȭ�ϸ�</TD>
									<TD CLASS="TD6"><INPUT TYPE=TEXT ID="txtFileName2" NAME="txtFileName2" SIZE=30 MAXLENGTH=100 STYLE="TEXT-ALIGN: left" ALT="ȭ�ϸ�" tag="12X" ><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnOpenFilePath" ALIGN=top TYPE="BUTTON" ONCLICK="vbscript:GetOpenFilePath()"></TD>
									<TD CLASS="TD5">&nbsp</TD>
									<TD CLASS="TD6">&nbsp</TD>
							</TR>	
					  </TABLE>
		</div>	

				     </FIELDSET>
				   </TD>
				</TR>
				<TR><TD <%=HEIGHT_TYPE_03%>></TD></TR>
				<TR>
					<TD WIDTH=100% HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_20%> >
							<TR>
								<TD HEIGHT=100%>
									<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
								</TD>
							</TR>
						</TABLE>
					</TD>	
				</TR>
			</TABLE>
		</TD>
	</TR>
	<TR HEIGHT=20>
	    <TD WIDTH=100%>
	        <TABLE <%=LR_SPACE_TYUPE_30%>>
	            <TR>
	                <TD><TABLE>
						<TR>
							<TD></TD>						
							<TD><BUTTON NAME="btnCb_select" CLASS="CLSMBTN">�����ͻ���</BUTTON>&nbsp;
								<BUTTON NAME="btnCb_creation" CLASS="CLSMBTN">���ϻ���</BUTTON>&nbsp;</TD>
							<TD><DIV ID="divButton" >
								<BUTTON NAME="btnCb_print2" CLASS="CLSMBTN">������ǥ�����</BUTTON>&nbsp;
								<BUTTON NAME="btnCb_print" CLASS="CLSMBTN">����ǥ���</BUTTON>&nbsp;
								
							</DIV></TD>							
						</TD></TR>
						</TABLE>
	                </TD>
	            </TR>
	        </TABLE>
	    </TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=0><IFRAME NAME="MyBizASP"  WIDTH=100% HEIGHT=0 FRAMEBORDER=0 SCROLLING=no noresize framespacing=0></IFRAME>
		</TD>
	</TR>
</TABLE>
<P ID="divTextArea"></P>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<INPUT TYPE=HIDDEN NAME="txtTab"    TAG="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' ����ó��ASP�� �ѱ�� ���� ������ ��� �ִ� Tag�� %>
<INPUT TYPE=HIDDEN NAME="hFileName" tag="14" TABINDEX="-1">
	
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
