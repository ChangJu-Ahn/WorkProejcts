<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
'*  1. Module Name          : Human Resource
'*  2. Function Name        : 연말정산관리 
'*  3. Program ID           : h9130ma1.asp
'*  4. Program Name         : h9130ma1.asp
'*  5. Program Desc         : 기초자료등록(보험료)
'*  6. Modified date(First) : 2006.12.06
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : lee wol san
'*  9. Modifier (Last)      : 
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
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
Const CookieSplit = 1233
Const BIZ_PGM_ID      = "h9130mb1.asp"						           '☆: Biz Logic ASP Name

Const TAB1 = 1										                   'Tab의 위치 
Const TAB2 = 2
Const C_SHEETMAXROWS    = 15	                                      '☜: Visble row

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim gSelframeFlg                                                       '현재 TAB의 위치를 나타내는 Flag %>
Dim gblnWinEvent                                                       'ShowModal Dialog(PopUp) Window가 여러 개 뜨는 것을 방지하기 위해 
Dim lgBlnFlawChgFlg	
Dim gtxtChargeType
Dim IsOpenPop
Dim lsInternal_cd

Dim C_FamName
Dim C_FamName_POP
Dim C_FamRelCd
Dim C_FamRelNm
Dim C_FamTypeCd
Dim C_FamTypeNm
Dim C_INSURAmt
dim C_subMitCd
dim C_subMitCdNm
Dim C_YEAR_FLAG
 
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      =  Parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction
		
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
'	Description : Item Popup에서 Return되는 값 setting
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
    
    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0024", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 
    iCodeArr = lgF0
    iNameArr = lgF1

	 ggoSpread.SetCombo Replace(iCodeArr,Chr(11),vbTab), C_FamTypeCd
	 ggoSpread.SetCombo Replace(iNameArr,Chr(11),vbTab), C_FamTypeNm
	 

	 ggoSpread.SetCombo "Y" & vbtab & "N"  , C_subMitCd
     ggoSpread.SetCombo "국세청자료" & vbtab & "그밖의자료",C_subMitCdNM


	  
     
End Sub

'========================================================================================================
' Name : InitData()
' Desc : Reset Combox
'========================================================================================================
Sub InitData()
    Dim intRow
    Dim intIndex 

	With frm1.vspdData
        For intRow = 1 To .MaxRows			
		    .Row = intRow

            .Col = C_FamTypeCd
            intIndex = .value
            .col = C_FamTypeNm
            .value = intindex					
       Next	
	End With
	
End Sub

'======================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'=======================================================================================================
Sub InitSpreadSheet()
	
	call InitSpreadPosVariables()
	With frm1.vspdData
		.MaxCols = C_YEAR_FLAG + 1												

		.Col = .MaxCols                                                              ' ☜:☜: Hide maxcols
		.ColHidden = True                                                            ' ☜:☜:

		.MaxRows = 0
        ggoSpread.Source = frm1.vspdData
		ggoSpread.Spreadinit "V20021126",, parent.gAllowDragDropSpread
		Call GetSpreadColumnPos("A")
		
		.ReDraw = false
		
        ggoSpread.SSSetEdit     C_FamName,       "가족성명",      20,,, 30,2
        ggoSpread.SSSetButton   C_FamName_POP

		ggoSpread.SSSetEdit		C_FamRelCd		, "가족관계코드", 12     
		ggoSpread.SSSetEdit		C_FamRelNm      , "가족관계", 12 
                        
		ggoSpread.SSSetCombo C_FamTypeCd    , "",5
		ggoSpread.SSSetCombo C_FamTypeNm    , "구분", 20
		ggoSpread.SSSetFloat C_INSURAmt       , "금액", 15,"2", ggStrIntegeralPart,  ggStrDeciPointPart, Parent.gComNum1000, Parent.gComNumDec,,,"Z"
		
		ggoSpread.SSSetCombo C_subMitCd    , "",5
		ggoSpread.SSSetCombo C_subMitCdNM    , "제출구분", 20


		ggoSpread.SSSetEdit  C_YEAR_FLAG	, "반영여부", 10
		
		call ggoSpread.MakePairsColumn(C_FamName,C_FamName_POP,"1")
		call ggoSpread.MakePairsColumn(C_FamRelCd,C_FamRelNm,"1")
		call ggoSpread.MakePairsColumn(C_FamTypeCd,C_FamTypeNm,"1")
		call ggoSpread.SSSetColHidden(C_FamRelCd, C_FamRelCd, true)
		call ggoSpread.SSSetColHidden(C_FamTypeCd, C_FamTypeCd, true)
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
	C_FamName		= 1
	C_FamName_POP   = 2
	C_FamRelCd		= 3
	C_FamRelNm		= 4
	C_FamTypeCd		= 5
	C_FamTypeNm		= 6
	C_INSURAmt		= 7
	C_subMitCd      = 8
	C_subMitCdNM    = 9
	C_YEAR_FLAG		= 10
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
			C_FamName = iCurColumnPos(1)
			C_FamName_POP   = iCurColumnPos(2)
			C_FamRelCd  = iCurColumnPos(3)
			C_FamRelNm  = iCurColumnPos(4)
			C_FamTypeCd = iCurColumnPos(5)
			C_FamTypeNm = iCurColumnPos(6)
			C_INSURAmt    = iCurColumnPos(7)
			C_subMitCd      = iCurColumnPos(8)
			C_subMitCdNM    = iCurColumnPos(9)
			C_YEAR_FLAG	= iCurColumnPos(10)			
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
	Call InitData()
End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'=======================================================================================================
Sub SetSpreadLock()

    With frm1.vspdData
    
		.ReDraw = False

		 ggoSpread.SpreadLock    C_FamName,   -1, C_FamName
		 ggoSpread.SpreadLock    C_FamName_POP,   -1, C_FamName_POP		 
		 ggoSpread.SpreadLock    C_FamRelNm,  -1, C_FamRelNm
		 ggoSpread.SpreadLock    C_FamTypeNm, -1, C_FamTypeNm
		 ggoSpread.SpreadLock    C_subMitCdNM, -1, C_subMitCdNM
		 ggoSpread.SSSetRequired C_INSURAmt,    -1, C_INSURAmt
		 ggoSpread.SpreadLock    C_YEAR_FLAG, -1, C_YEAR_FLAG			 
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

		 ggoSpread.SSSetRequired		C_FamName, pvStartRow, pvEndRow
		 ggoSpread.SSSetProtected		C_FamRelNm, pvStartRow, pvEndRow
		 'ggoSpread.SSSetRequired		C_FamTypeNm, pvStartRow, pvEndRow
		  ggoSpread.SSSetProtected		C_FamTypeNm, pvStartRow, pvEndRow
		 ggoSpread.SSSetRequired		C_subMitCdNM, pvStartRow, pvEndRow
		 ggoSpread.SSSetRequired		C_INSURAmt, pvStartRow, pvEndRow
		 ggoSpread.SSSetProtected		C_YEAR_FLAG, pvStartRow, pvEndRow
		 
		.vspdData.ReDraw = True
    
	End With

End Sub

'======================================================================================================
'	Name : Form_Load()
'	Description : Window On Load(공통 Include 파일에 선언)시 변수초기화 및 화면초기화를 하기 위해 
'                 함수를 Call하는 부분 
'======================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
	Call  ggoOper.FormatField(Document, "A", ggStrIntegeralPart,  ggStrDeciPointPart, Parent.gDateFormat, Parent.gComNum1000, Parent.gComNumDec)
    Call  ggoOper.FormatDate(frm1.txtYear,  Parent.gDateFormat, 3)	

	Call  ggoOper.LockField(Document, "N")											'⊙: Lock Field
            
    Call InitSpreadSheet                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables

    Call  FuncGetAuth(gStrRequestMenuID,  Parent.gUsrID, lgUsrIntCd)                                ' 자료권한:lgUsrIntCd ("%", "1%")
    
    Call SetDefaultVal
	Call SetToolbar("1100000000001111")												'⊙: Set ToolBar
    
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
    Dim value
    Dim strRel,family_nm
    
   	Frm1.vspdData.Row = Row
   	Frm1.vspdData.Col = Col

    Select Case Col
         Case  C_FamName
             iDx = trim(Frm1.vspdData.Text)
   	        Frm1.vspdData.Col = C_FamName
			family_nm = Frm1.vspdData.Text

            If family_nm = "" Then
  	            Frm1.vspdData.Col = C_FamRelCd
                Frm1.vspdData.Text = ""
  	            Frm1.vspdData.Col = C_FamRelNm
                Frm1.vspdData.Text = ""
            Else

                strRel = CommonQueryRs(" FAMILY_RES_NO"," HFA150T "," INSUR_YN = 'Y' AND FAMILY_NAME = " & FilterVar(iDx, "''", "S")  & " AND EMP_NO= " & FilterVar(frm1.txtEmp_no.value, "''", "S") & " AND YEAR_YY = " & FilterVar(frm1.txtYear.Year, "''", "S") ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

                If strRel = false then
	        		Call DisplayMsgBox("970029","X","부양가족공제자등록에 해당자가 보험료 체크가 되어있는지","X")	 
  	                Frm1.vspdData.Col = C_FamRelCd
                    Frm1.vspdData.Text = ""
  	                Frm1.vspdData.Col = C_FamRelNm
                    Frm1.vspdData.Text = ""
 
                Else
 	       	    
					Call CommonQueryRs(" FAMILY_REL, dbo.ufn_GetCodeName('H0140',FAMILY_REL) " ,_
 					" HFA150T ",_
					" INSUR_YN='Y' and EMP_NO =" & FilterVar(frm1.txtEmp_no.value, "''", "S")  & " AND YEAR_YY = " & FilterVar(frm1.txtYear.Year, "''", "S") & " AND FAMILY_NAME= " & FilterVar(family_nm, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
 
		       	    Frm1.vspdData.Col = C_FamRelCd
		       	    Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))

		       	    Frm1.vspdData.Col = C_FamRelNm 
		       	    Frm1.vspdData.Text = Trim(Replace(lgF1,Chr(11),""))
 
                End if 
            End if 
         Case  C_FamTypeNm
                 iDx = Frm1.vspdData.value
                 Frm1.vspdData.Col   = C_FamTypeCd
                 Frm1.vspdData.value = iDx
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
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 컬럼버튼을 클릭할 경우 발생 
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)

	With frm1.vspdData 
		ggoSpread.Source = frm1.vspdData
		If Row > 0 Then
			Select Case Col
                Case C_FamName_POP
				    .Col = C_FamName
				    .Row = Row
                    Call OpenCode("", C_FamName_POP, Row)
			End Select
		End If
	End With
    
End Sub
'========================================================================================================
'	Name : OpenCode()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
        
	    Case C_FamName_POP

	        arrParam(0) = "가족성명 팝업"			' 팝업 명칭 
	        arrParam(1) = " HFA150T  "				 		' TABLE 명칭 
	        arrParam(2) = frm1.txtEmp_no.value               		    ' Code Condition
	        arrParam(3) = ""							' Name Cindition
	        arrParam(4) = " INSUR_YN='Y' and year_yy = " & FilterVar(frm1.txtYear.Year, "''", "S") & " and emp_no = " & FilterVar(frm1.txtEmp_no.value, "''", "S") 			' Where Condition
	        arrParam(5) = "사번"			    ' TextBox 명칭 
	
            arrField(0) =  "HH"  & parent.gcolsep  & " emp_no "										' Field명(0)
            arrField(1) =  "ED21" & parent.gcolsep & " family_name "								' Field명(0)
            arrField(2) =  "ED22" & parent.gcolsep & " dbo.ufn_GetCodeName('H0140',family_rel) "	' Field명(1)
            arrField(3) =  "HH" & parent.gcolsep & " family_rel "	' Field명(1)

            arrHeader(0) = "사번"					' Header명(0)
            arrHeader(1) = "가족성명"				' Header명(0)
            arrHeader(2) = "가족관계"			    ' Header명(1)
            arrHeader(3) = "가족관계"			    ' Header명(1)
	End Select

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
    If arrRet(0) = "" Then
		frm1.vspdData.Col = C_FamName
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
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)
	Dim strRel,Row
	Dim family_nm , rel_nm, intIndex
	Dim lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6
	With frm1

		Select Case iWhere
		    Case C_FamName_POP
		        .vspdData.Col = C_FamName
		        family_nm = arrRet(1) 
		    	.vspdData.text = family_nm
 
 				Frm1.vspdData.Col = C_FamRelcd
		       	Frm1.vspdData.Text = arrRet(3) 
		       	Frm1.vspdData.Col = C_FamRelNm
		       	Frm1.vspdData.Text = arrRet(2) 
		      Row = frm1.vspdData.ActiveRow
				
			if (	CommonQueryRs(" supp_cd, dbo.ufn_GetCodeName('H0024',supp_cd) ,rel_cd", " HAA020T ",_
				" EMP_NO =" & FilterVar(frm1.txtEmp_no.value, "''", "S")  & " AND FAMILY_NM= " & FilterVar(family_nm, "''", "S"),lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) )  then
				
                if Trim(Replace(lgF0,Chr(11),""))<>"" then
		       		Frm1.vspdData.Col = C_FamTypeCd
		       		Frm1.vspdData.Text = Trim(Replace(lgF0,Chr(11),""))

		       		Frm1.vspdData.Col = C_FamTypeNm 
		       		Frm1.vspdData.Text = Trim(Replace(lgF1,Chr(11),""))
		       	
		       		ggoSpread.SSSetProtected	C_FamTypeNm, Row, C_FamTypeNm		       
		       	end if	
		     else
				ggoSpread.SSSetProtected	C_FamTypeNm, Row, C_FamTypeNm	
				'ggoSpread.SpreadUnLock		C_FamTypeNm, Row, C_FamTypeNm,Row
				'ggoSpread.SSSetRequired		C_FamTypeNm, C_FamTypeNm, C_FamTypeNm			
		     end if  	
 				
 				
        End Select

	End With

End Function

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
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
'   Event Desc : 컬럼을 클릭할 경우 발생 
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
'   Event Desc : Combo 변경 이벤트 
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
    End With

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
		IntRetCD =  DisplayMsgBox("900013",  Parent.VB_YES_NO, "X", "X")			    '데이타가 변경되었습니다. 조회하시겠습니까?
   		If IntRetCD = vbNo Then
  			Exit Function
   		End If
	End If

    ggoSpread.ClearSpreadData
    
    If Not chkField(Document, "1") Then		
       Exit Function
    End If

    If txtEmp_no_Onchange() Then        'enter key 로 조회시 사원을 check후 해당사항 없으면 query종료...
        Exit Function
    End if

    Call InitVariables                      
    Call MakeKeyStream("X")

    If DbQuery = False Then  
		Exit Function
	End If
       
    FncQuery = True                                                              '☜: Processing is OK
End Function

'======================================================================================================
' Function Name : FncSave
' Function Desc : This function is related to Delete Button of Main ToolBar
'======================================================================================================
Function FncSave() 
    Dim IntRetCD ,intRow

    FncSave = False                                                         
    
    Err.Clear                                                               
    
     ggoSpread.Source = frm1.vspdData
    If  ggoSpread.SSCheckChange = False Then
        IntRetCD =  DisplayMsgBox("900001","X","X","X")  
        Exit Function
    End If
    
     ggoSpread.Source = frm1.vspdData
    If Not chkField(Document, "2") OR Not  ggoSpread.SSDefaultCheck Then                                  '⊙: Check contents area
       Exit Function
    End If

    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
    
    For intRow = 1 To Frm1.vspdData.MaxRows			
	    Frm1.vspdData.Row = intRow
	    Frm1.vspdData.Col = 0

		Select Case Frm1.vspdData.Text 
			Case  ggoSpread.InsertFlag ,ggoSpread.UpdateFlag
				Frm1.vspdData.Col = C_FamRelNm 
			
				If Frm1.vspdData.Text = "" Then
					Call DisplayMsgBox("971012","X","가족성명","X")
					Frm1.vspdData.Col = C_FamRelNm - 3
					Frm1.vspdData.action = 0   
					Exit Function
				End If
				
				Frm1.vspdData.Col = C_INSURAmt 

				If Frm1.vspdData.Text = 0 Then
					Call DisplayMsgBox("800484","X","지급금액","X")
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
    With frm1.vspdData
		If .ActiveRow > 0 Then
			.focus
			.ReDraw = False
		
			 ggoSpread.Source = frm1.vspdData	
			 ggoSpread.CopyRow
			 SetSpreadColor .ActiveRow, .ActiveRow
   
			.ReDraw = True
		End If
	End With
	
End Function

'======================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'======================================================================================================
Function FncCancel() 

     ggoSpread.Source = frm1.vspdData	

	 ggoSpread.EditUndo                                                  '☜: Protect system from crashing

	Call InitData

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
    Call parent.FncPrint()                                                   '☜: Protect system from crashing
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
		IntRetCD =  DisplayMsgBox("900016",  Parent.VB_YES_NO, "X", "X")                '데이타가 변경되었습니다. 종료 하시겠습니까?
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
    Dim strArr
    Err.Clear                                                                        '☜: Clear err status

    DbQuery = False                                                                  '☜: Processing is NG
    
    if LayerShowHide(1) = false then
	    Exit Function
	end if

    strVal = BIZ_PGM_ID & "?txtMode="            & Parent.UID_M0001                         '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key
    
    strVal = strVal     & "&lgCurrentSpd="       & lgCurrentSpd                   '☜: Next key tag
    strVal = strVal     & "&lgStrPrevKey=" & lgStrPrevKey              '☜: Next key tag
    strVal = strVal     & "&txtMaxRows="         & Frm1.vspdData.MaxRows          '☜: Max fetched data
	strArr = Split(lgKeyStream, Parent.gColSep)
    
    Call RunMyBizASP(MyBizASP, strVal)                                               '☜:  Run biz logic
	
    DbQuery = True        
    
    Call SetToolbar("1100111100111111")										        '버튼 툴바 제어 

End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()													     
   lgIntFlgMode =  Parent.OPMD_UMODE    
    Call  ggoOper.LockField(Document, "Q")										'⊙: Lock field
    Call InitData()
    Call SetToolbar("1100111100111111")										        '버튼 툴바 제어 
    
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
		
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG
		
	 if LayerShowHide(1) = false then
	    Exit Function
	end if

  	With Frm1
		.txtMode.value      =  Parent.UID_M0002                                            '☜: Delete
		.txtKeyStream.value = lgKeyStream
	End With

    strVal  = ""
    strDel  = ""
    lGrpCnt = 1

	With Frm1
    
	 ggoSpread.Source = .vspdData

	For lRow = 1 To .vspdData.MaxRows
    
	    .vspdData.Row = lRow
	    .vspdData.Col = 0

		Select Case .vspdData.Text
			Case  ggoSpread.InsertFlag									  
                                                 strVal = strVal & "C" & Parent.gColSep                   '0
                                                 strVal = strVal & lRow & Parent.gColSep                   '1
                .vspdData.Col = C_FamName      : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '2
                .vspdData.Col = C_FamRelCd     : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '3
                .vspdData.Col = C_FamTypeCd    : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '4
                .vspdData.Col = C_INSURAmt       : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '5
                .vspdData.Col = C_subMitCd       : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep   '6

                lGrpCnt = lGrpCnt + 1

            Case  ggoSpread.UpdateFlag
                                                 strVal = strVal & "U" & Parent.gColSep                   '0
                                                 strVal = strVal & lRow & Parent.gColSep                   '1
                .vspdData.Col = C_FamName      : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '2
                .vspdData.Col = C_FamRelCd     : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '3
                .vspdData.Col = C_FamTypeCd    : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '4
                .vspdData.Col = C_INSURAmt       : strVal = strVal & Trim(.vspdData.Text) & Parent.gColSep   '5
                .vspdData.Col = C_subMitCd       : strVal = strVal & Trim(.vspdData.Text) & Parent.gRowSep   '6


                lGrpCnt = lGrpCnt + 1
                    
            Case  ggoSpread.DeleteFlag									
                                                 strDel = strDel & "D" & Parent.gColSep                   '0
                                                 strDel = strDel & lRow & Parent.gColSep                   '1
                .vspdData.Col = C_FamName      : strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep   '2
                .vspdData.Col = C_FamRelCd     : strDel = strDel & Trim(.vspdData.Text) & Parent.gColSep   '3
                .vspdData.Col = C_FamTypeCd    : strDel = strDel & Trim(.vspdData.Text) & Parent.gRowSep   '4
                
                lGrpCnt = lGrpCnt + 1
                
		End Select
	Next
	
	.txtMaxRows.value = lGrpCnt-1	
	.txtSpread.value = strDel & strVal
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)									
	End With

    DbSave  = True                                                               '☜: Processing is NG
    
End Function

'======================================================================================================
' Function Name : DbSaveOk
' Function Desc : DBSave가 성공적일 경우 MyBizASP 에서 호출되는 Function, 현재 FncSave에 있는것을 옮김 
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
' Desc : developer describe this line(grid외에서 사용) 
'========================================================================================================
Function OpenEmpName(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

    If  iWhere = 0 Then
	    arrParam(0) = frm1.txtEmp_no.value			' Code Condition
	    arrParam(1) = ""'frm1.txtName.value			' Name Cindition
	    arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  
    Else 'spread
        frm1.vspdData.Col = C_EMP_NO
		arrParam(0) = frm1.vspdData.Text			' Code Condition
        frm1.vspdData.Col = C_NAME
	    arrParam(1) = ""'frm1.vspdData.Text			' Name Cindition
	    arrParam(2) = lgUsrIntCd                    ' 자료권한 Condition  
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
'	Description : Item Popup에서 Return되는 값 setting
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
'   Event Name : txtEmp_no_change             '<==인사마스터에 있는 사원인지 확인 
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
    			Call  DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call  DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
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
' Function Desc : 연말정산 반영 
'=======================================================================================================
Function Reflect() 
	Dim strVal
	Dim strYyyymm
	Dim IntRetCD

	Reflect = False                                                          '⊙: Processing is NG

    If txtEmp_no_Onchange() Then         'ENTER KEY 로 조회시 사원과 사번을 CHECK 한다 
        Exit Function
    End if

	IntRetCD = DisplayMsgBox("900018",parent.VB_YES_NO,"X","X")

'    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
'        Call DisplayMsgbox("900002","X","X","X")                                '☆:
'        Exit Function
'    End If
    	
	If IntRetCD = vbNo Then
		Exit Function
	End If

    '기초데이터 있는지 체크 
    IntRetCd = CommonQueryRs(" EMP_NO "," HFA030T "," YY =  " & FilterVar(Frm1.txtYear.Year , "''", "S") & " AND EMP_NO =  " & FilterVar(Frm1.txtEmp_no.Value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    If IntRetCd = False then
		Call DisplayMsgBox("800430","X",Frm1.txtYear.Year & "년","X")        '기초자료를 먼저 입력하세요/	
		Exit Function
    End If
    
	Call LayerShowHide(1)
    
	On Error Resume Next                                                   '☜: Protect system from crashing

    Call MakeKeyStream("X")
    
	strVal = BIZ_PGM_ID & "?txtMode=REFLECT" 
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key

	Call RunMyBizASP(MyBizASP, strVal)										'☜: 비지니스 ASP 를 가동 

	Reflect = True                                                           '⊙: Processing is NG

End Function

'======================================================================================================
' Function Name : ReflectOk
' Function Desc : Reflect가 성공적일 경우 MyBizASP 에서 호출되는 Function
'======================================================================================================
Function ReflectOk()													      
	Dim IntRetCD 

	IntRetCD =DisplayMsgBox("990000","X","X","X")
	window.status = "작업 완료"
    call FncQuery()
End Function

Function ReflectNO()				          
	Dim IntRetCD 

    Call DisplayMsgBox("800414","X","X","X")
	window.status = "작업 실패"

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
								<td background="../../../Cshared/Image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
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
									<TD CLASS=TD5 NOWRAP>정산년도</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> NAME="txtYear" CLASS=FPDTYYYY tag="12X1" Title="FPDATETIME" ALT="정산년도" id=fpDateTime1> </OBJECT>');</SCRIPT>
									</TD>	
									<TD CLASS=TD5 NOWRAP>사번</TD>
			     					<TD CLASS=TD6 NOWRAP><INPUT NAME="txtEmp_no" ALT="사번" TYPE="Text" SiZE=15 MAXLENGTH=13 tag="12XXXU"><IMG SRC="../../../Cshared/Image/btnPopup.gif" NAME="btnEmpNo" align=top TYPE="BUTTON" ONCLICK="VBScript: OpenEmpName('0')">
									                     <INPUT NAME="txtName" MAXLENGTH="30" SIZE="20" ALT ="성명" tag="14XXXU"></TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>부서명</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtDept_nm" MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN:left" ALT ="부서명" tag="14">&nbsp;</TD>
									<TD CLASS=TD5 NOWRAP>직  위</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtRollPstn" MAXLENGTH="20" SIZE=20 STYLE="TEXT-ALIGN:left" ALT ="직위" tag="14">&nbsp;</TD>
								</TR>
								<TR>
									<TD CLASS=TD5 NOWRAP>입사일</TD>
									<TD CLASS=TD6 NOWRAP>
										<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT id=fpDateTime2 title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtEntr_dt CLASSID=<%=gCLSIDFPDT%> ALT="입사일" tag="14X1" VIEWASTEXT></OBJECT>');</SCRIPT>
									</TD>
									<TD CLASS=TD5 NOWRAP>급  호</TD>
									<TD CLASS=TD6 NOWRAP><INPUT NAME="txtPay_grd" MAXLENGTH="20" SIZE=20 ALT ="급호" tag="14">&nbsp;</TD>
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
		<TD <%=HEIGHT_TYPE_01%> WIDTH=100%>&nbsp;&nbsp;&nbsp;※보험료는 부양가족공제자등록 화면에서 보험료로 체크된 사람만 입력가능합니다.</TD>
	</TR> 		
	<TR HEIGHT=55>
		<TD WIDTH=100%>
			<TABLE >
				<TR><TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=110>장애인보험료:</TD>
					<TD WIDTH=100><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtSum1 NAME="txtSum1" CLASS=FPDS140 tag="24X2Z" ALT="장애인보험료" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
					<TD WIDTH=120>&nbsp;&nbsp;기타보험료:</TD>
					<TD WIDTH=100><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtSum2 NAME="txtSum2" CLASS=FPDS140 tag="24X2Z" ALT="기타보험료" Title="FPDOUBLESINGLE"></OBJECT>');</SCRIPT></TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>
				
				
												
				<TR><TD WIDTH=10>&nbsp;</TD>
					<TD WIDTH=30><BUTTON NAME="btnSplit" CLASS="CLSMBTN" onclick="Reflect()" Flag=1>연말정산반영</BUTTON>&nbsp;</TD>
					<TD WIDTH=*>&nbsp;</TD>
				</TR>			
			</TABLE>
		</TD>
	</TR>	  
	<TR>
		<TD WIDTH=100% HEIGHT=0><IFRAME NAME="MyBizASP" SRC="h9130mb1.asp" WIDTH=100% HEIGHT=1 FRAMEBORDER=0 SCROLLING=YES noresize framespacing=0></IFRAME>
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
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>


