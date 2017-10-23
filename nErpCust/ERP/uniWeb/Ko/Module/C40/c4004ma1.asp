<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!-- '======================================================================================================
'*  1. Module Name          : Cost
'*  2. Function Name        : 오더 배부규칙 등록 
'*  3. Program ID           : c4004ma1.asp
'*  4. Program Name         : 오더 배부규칙 등록 
'*  5. Program Desc         : 오더 배부규칙 등록 
'*  6. Modified date(First) : 2005-09-13
'*  7. Modified date(Last)  : 2005-09-27
'*  8. Modifier (First)     : HJO 
'*  9. Modifier (Last)      : HJO
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'======================================================================================================= -->

<HTML>
<HEAD>
<TITLE><%=Request("strASPMnuMnuNm")%></TITLE>

<!--
########################################################################################################
#						   3.    External File Include Part
########################################################################################################-->

<!--
========================================================================================================
=                          3.1 Server Side Script
========================================================================================================-->
<!-- #Include file="../../inc/IncSvrCcm.inc"  -->

<!--
========================================================================================================
=                          3.2 Style Sheet
======================================================================================================== -->
<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<!--
========================================================================================================
=                          3.3 Client Side Script
======================================================================================================== -->
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs">		</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incCliRdsQuery.vbs">	</SCRIPT>
<SCRIPT LANGUAGE="VBScript"		SRC="../../inc/incEB.vbs">	</SCRIPT>

<SCRIPT LANGUAGE=VBSCRIPT>
Option Explicit

'########################################################################################################
'#                       4.  Data Declaration Part
'########################################################################################################

'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID = "c4004mb1.asp"                               'Biz Logic ASP
Const BIZ_PGM_ID2 = "c4004mb2.asp"                               'Biz Logic ASP

' -- 
Dim C_VerCd			' -
Dim C_WcFlag
Dim C_WcFlagNM
Dim C_WcCd
Dim C_WcCdPop
Dim C_WcNm
Dim C_GpLevel
Dim C_GpLevelPop
Dim C_GpCd
Dim C_GpCdPop
Dim C_GpNm

Dim C_AcctCd
Dim C_AcctCdPop
Dim C_AcctNm
Dim C_AFctrCd
Dim C_AFctrCdPop
Dim C_AFctrNm
Dim C_SFctrCd
Dim C_SFctrCdPop
Dim C_SFctrNm


'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	


'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lgQueryFlag
Dim IsOpenPop          
Dim lgCurrGrid
Dim lgCopyVersion
Dim lgErrRow, lgErrCol

Dim lgCostConfig
'========================================================================================================
'                        5.1 Common Method-1
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column Const value
'========================================================================================================
Sub initSpreadPosVariables()  
	
	' -- 그리드1의 컬럼 정의 
	C_VerCd						= 1		
	C_WcFlag			= 2		
	C_WcFlagNM		=3
	C_WcCd            =4
	C_WcCdPop       =5
	C_WcNm           =6
	C_GpLevel         =7
	C_GpLevelPop    =8
	C_GpCd             =9
	C_GpCdPop        =10
	C_GpNm             =11

	C_AcctCd           =12
	C_AcctCdPop     =13
	C_AcctNm          =14
	C_AFctrCd            =15
	C_AFctrCdPop       =16
	C_AFctrNm            =17
	C_SFctrCd            =18
	C_SFctrCdPop       =19
	C_SFctrNm            =20

End Sub

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'======================================================================================================== 
Sub InitVariables()

    lgIntFlgMode = Parent.OPMD_CMODE
    lgBlnFlgChgValue = False 
    lgIntGrpCount = 0  
    
    lgStrPrevKey = ""	
    lgLngCurRows = 0 
	lgSortKey = 1

	lgCopyVersion = "" :frm1.versionFlag.value=""
	lgErrRow = 0 : lgErrCol = 0
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
Sub SetDefaultVal()
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	frm1.txtVER_CD.focus
   	Set gActiveElement = document.activeElement			    
	'------ Developer Coding part (End )   -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
' Para : 1. Currency
'        2. I(Input),Q(Query),P(Print),B(Bacth)
'        3. "*" is for Common module
'           "A" is for Accounting
'           "I" is for Inventory
'           ...
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp"  -->
	<%Call LoadInfTB19029A("I","*", "NOCOOKIE", "MA") %>
End Sub

'========================================================================================================
' Name : CookiePage()
' Desc : Write or Read cookie value 
'========================================================================================================
Sub CookiePage(Kubun)
   '------ Developer Coding part (Start ) -------------------------------------------------------------- 
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub

'========================================================================================================
' Name : MakeKeyStream
' Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   '------ Developer Coding part (Start ) --------------------------------------------------------------        
   '------ Developer Coding part (End   ) -------------------------------------------------------------- 
End Sub        

'========================================================================================================
' Name : InitSpreadSheet
' Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet()
	' -- initialize grid column 
	Call initSpreadPosVariables()    

	With frm1.vspdData
	
	.MaxCols = C_SFctrNm+1

	.Col = .MaxCols	
    .ColHidden = True
    
    ggoSpread.Source = frm1.vspdData
    ggoSpread.Spreadinit "V20021122",,parent.gAllowDragDropSpread 

	Call ggoSpread.ClearSpreadData()    '☜: Clear spreadsheet data 

	.ReDraw = false

	Call GetSpreadColumnPos("A")
	
    ggoSpread.SSSetEdit		C_VerCd				,"ver.",7,1    
  	ggoSpread.SSSetCombo		C_WcFlag,				"사내/외주가공구분",10
	ggoSpread.SSSetCombo		C_WcFlagNM,			"사내/외주가공구분명",10
	
    ggoSpread.SSSetEdit		C_WcCd	,		"공정(WC)/구매그룹" ,10,,, 20,2      
    ggoSpread.SSSetButton	C_WcCdPop    
    ggoSpread.SSSetEdit		C_WcNm			,"공정(WC)/구매그룹명" ,25
    ggoSpread.SSSetEdit		C_GpLevel			,"계정그룹(Level)" ,10,,, 20,2
	ggoSpread.SSSetButton	C_GpLevelPop    
    ggoSpread.SSSetEdit		C_GpCd				,"계정그룹" ,10,,, 20,2
    ggoSpread.SSSetButton	C_GpCdPop    
    ggoSpread.SSSetEdit		C_GpNm				,"계정그룹명",15
    ggoSpread.SSSetEdit		C_AcctCd			,"계정" ,10,,, 20,2
    ggoSpread.SSSetButton	C_AcctCdPop    
    ggoSpread.SSSetEdit		C_AcctNm			,"계정명",15
    ggoSpread.SSSetEdit		C_AFctrCd		,"실제원가배부요소" ,10,,,, 2
    ggoSpread.SSSetButton	C_AFctrCdPop    
    ggoSpread.SSSetEdit		C_AFctrNm		,"실제원가배부요소명",25
    ggoSpread.SSSetEdit		C_SFctrCd		,"표준원가배부요소" ,10,,,, 2
    ggoSpread.SSSetButton	C_SFctrCdPop    
    ggoSpread.SSSetEdit		C_SFctrNm		,"표준원가배부요소명",25

   call ggoSpread.MakePairsColumn(C_WcCd,C_WcCdPop)
   call ggoSpread.MakePairsColumn(C_GpLevel,C_GpLevelPop)
   call ggoSpread.MakePairsColumn(C_GpCd,C_GpCdPop)
   
   call ggoSpread.MakePairsColumn(C_AcctCd,C_AcctCdPop)   
   call ggoSpread.MakePairsColumn(C_AFctrCd,C_AFctrCdPop)
   call ggoSpread.MakePairsColumn(C_SFctrCd,C_SFctrCdPop)    
   
   Call ggoSpread.SSSetColHidden(C_SFctrCd,C_SFctrCd,True)
   Call ggoSpread.SSSetColHidden(C_SFctrCdPop,C_SFctrCdPop,True)
   Call ggoSpread.SSSetColHidden(C_SFctrNm,C_SFctrNm,True)      
   Call ggoSpread.SSSetColHidden(C_VerCd,C_WcFlag,True)   
	
	.rowheight(-1000) = 20	' 높이 재지정 
	ggoSpread.SSSetSplit2(C_WcCd)										'frozen 기능추가 
	
	.ReDraw = true
	
    Call SetSpreadLock     
    End With     
End Sub

'======================================================================================================
' Name : SetSpreadLock
' Desc : This method set color and protect cell in spread sheet
'======================================================================================================
Sub SetSpreadLock()
    With frm1.vspdData
    
    ggoSpread.Source = frm1.vspdData
    
    .ReDraw = False
    ggoSpread.SpreadLock		C_VerCd				,-1	,C_VerCd	
	ggoSpread.SSSetRequired		C_WcFlag		,-1	,-1
	ggoSpread.SSSetRequired		C_WcFlagNm		,-1	,-1
	ggoSpread.SSSetRequired		C_WcCd	,-1	,-1
	ggoSpread.SpreadLock		C_WcNm	,-1	,C_WcNm
	
	ggoSpread.SpreadLock		C_GpNm				,-1	,C_GpNm
	
	ggoSpread.SpreadLock		C_AcctNm			,-1	,C_AcctNm
	ggoSpread.SpreadLock		C_AFctrNm		,-1	,C_AFctrNm
	ggoSpread.SpreadLock		C_SFctrNm		,-1	,C_SFctrNm
    .ReDraw = True

    End With
End Sub


'======================================================================================================
' Name : SetSpreadColor
' Desc : This method set color and protect cell by cell in spread sheet
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1
    
    ggoSpread.Source = frm1.vspdData
    .vspdData.ReDraw = False
								      'Col          Row				Row2    
	ggoSpread.SSSetProtected	C_VerCd				,pvStartRow		,pvEndRow    
	ggoSpread.SSSetRequired		C_WcFlag		,pvStartRow		,pvEndRow
	ggoSpread.SSSetRequired		C_WcFlagNm		,pvStartRow		,pvEndRow
	ggoSpread.SSSetRequired		C_WcCd	,pvStartRow		,pvEndRow
	ggoSpread.SSSetProtected	C_WcNm	,pvStartRow		,pvEndRow    	
	ggoSpread.SSSetProtected	C_GpNm				,pvStartRow		,pvEndRow   	
	ggoSpread.SSSetProtected	C_AcctNm			,pvStartRow		,pvEndRow 
	ggoSpread.SSSetProtected	C_AFctrNm		,pvStartRow		,pvEndRow
	ggoSpread.SSSetProtected	C_SFctrNm		,pvStartRow		,pvEndRow        
	
    .vspdData.ReDraw = True
    
    End With
End Sub


'======================================================================================================
' Name : SetSpreadColorQuery
' Desc : This method set color and protect cell by cell in spread sheet
'======================================================================================================
Sub SetSpreadColorQuery()
    With frm1
    
    ggoSpread.Source = frm1.vspdData
    .vspdData.ReDraw = False
								      'Col          Row				Row2    
	ggoSpread.SSSetProtected		C_VerCd				,-1		,-1    
	ggoSpread.SSSetProtected		C_WcFlag			,-1		,-1    
	ggoSpread.SSSetProtected		C_WcFlagNm			,-1		,-1    
	ggoSpread.SSSetProtected		C_WcCd		,-1		,-1    
	ggoSpread.SSSetProtected		C_WcNm		,-1		,-1
	ggoSpread.SSSetProtected		C_GpCd					,-1		,-1        
	ggoSpread.SSSetProtected		C_GpNm					,-1		,-1    
	ggoSpread.SSSetProtected		C_AcctCd					,-1		,-1    
	ggoSpread.SSSetProtected		C_AcctNm				,-1		,-1    
	'ggoSpread.SSSetRequired		C_AFctrCd			,-1		,-1    
	ggoSpread.SSSetProtected		C_AFctrNm			,-1		,-1
	ggoSpread.SSSetProtected		C_SFctrNm			,-1		,-1        
	ggoSpread.SSSetProtected		C_GpLevel	,-1		,-1
	ggoSpread.SSSetProtected		C_GpLevelPop	,-1		,-1    
	ggoSpread.SSSetProtected		C_WcCdPop	,-1		,-1    
	ggoSpread.SSSetProtected		C_GpCdPop	,-1		,-1    
	ggoSpread.SSSetProtected		C_AcctCdPop	,-1		,-1    
	    
    .vspdData.ReDraw = True
    
    End With
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
            
			' -- 그리드1의 컬럼 정의 
			 C_VerCd						= iCurColumnPos(1)	
			 C_WcFlag			= iCurColumnPos(2)	' 
			 C_WcFlagNm			= iCurColumnPos(3)	' 
			 C_WcCd		= iCurColumnPos(4)		
			 C_WcCdPop					= iCurColumnPos(5)	' -
			 C_WcNm				= iCurColumnPos(6)		
			 C_GpLevel					= iCurColumnPos(7)		
			 C_GpLevelPop					= iCurColumnPos(8)		
			 C_GpCd				= iCurColumnPos(9)		
			 C_GpCdPop					= iCurColumnPos(10)	' -- 계정그룹 
			 C_GpNm				= iCurColumnPos(11)		
			 C_AcctCd					= iCurColumnPos(12)		
			 C_AcctCdPop					= iCurColumnPos(13)	' -- 계정 
			 C_AcctNm				= iCurColumnPos(14)		
			 C_AFctrCd					= iCurColumnPos(15)	
			 C_AFctrCdPop				= iCurColumnPos(16)	' -실제 배부요소 
			 C_AFctrNm			= iCurColumnPos(17)	
			 C_SFctrCd					= iCurColumnPos(18)	
			 C_SFctrCdPop				= iCurColumnPos(19)	' - 표준 배부요소 
			 C_SFctrNm			= iCurColumnPos(20)		
		
    End Select    
End Sub

'******************************************  2.4 POP-UP 처리함수  ****************************************
'	기능: 기준 POP-UP
'   설명: 기준 POP-UP에 관한 Open은 Include한다. 
'	      하나의 ASP에서 Popup이 중복되면 하나는 링크시켜 사용하고 나머지는 재정의하여 사용한다.
'********************************************************************************************************* 
 
'------------------------------------------  InitComboBox()  ----------------------------------------------
'	Name :InitComboBox
'	Description : 
'--------------------------------------------------------------------------------------------------------- 
Sub InitComboBox
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
   
    Call CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR ", "MAJOR_CD=" & FilterVar("P1003", "''", "S") & " AND MINOR_CD IN (" & FilterVar("M", "''", "S") & "," & FilterVar("O", "''", "S") & ") "  ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

	lgF0 = Replace(lgF0,Chr(11),vbTab) 
	lgF1 = Replace(lgF1,Chr(11),vbTab) 
	
	ggoSpread.Source = frm1.vspdData
    ggoSpread.SetCombo lgF0, C_WcFlag		:    ggoSpread.SetCombo lgF1, C_WcFlagNm 
    frm1.vspdData.Col=C_WcFlag : frm1.vspdData.Value=0
    frm1.vspdData.Col=C_WcFlagNm : frm1.vspdData.Value=0
	
End Sub
'------------------------------------------  OpenVersion()  ----------------------------------------------
'	Name :OpenVersion
'	Description : version popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenVersion(Byval iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
	Select Case iWhere
	Case 0
		If frm1.txtVER_CD.readOnly = True Then
			IsOpenPop = False
			Exit Function
		End If
		arrParam(0) = "VERSION팝업"
		arrParam(1) = " ufn_getListOfPopup_C4004MA1('6') "	
		arrParam(2) = Trim(frm1.txtVER_CD.Value)
		arrParam(3) = ""
		arrParam(4) = ""		
		arrParam(5) = "Version"
	
		arrField(0) = "CODE"

		arrHeader(0) = "Version"

    
		arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
			"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

		IsOpenPop = False

		If arrRet(0) = "" Then
			frm1.txtVER_CD.focus
			Exit Function
		Else
			Call SetVersion(arrRet, iWhere)
		End If
	Case 1

		'-----------------------
		'Check condition area
		'-----------------------
		If Not chkField(Document, "1") Then									'⊙: This function check indispensable field		
			IsOpenPop = False		  
		   Exit Function
		Else
    
			arrParam(0) = "Version 팝업"
			arrParam(1) = "ufn_getListOfPopup_C4004MA1('6')"	
			arrParam(2) = ""
			arrParam(3) = ""		
			If 	Trim(frm1.txtVER_CD.Value)	="" Then
			arrParam(4) = ""
			Else
			arrParam(4) = " CODE <> " & Filtervar(Trim(frm1.txtVER_CD.Value),"''","S")
			End If	
			arrParam(5) = "Version"
	
			arrField(0) = "CODE"
			
    
			arrHeader(0) = "Version"

			arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
				"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

			IsOpenPop = False

			If arrRet(0) = "" Then
				frm1.txtVER_CD.focus
				Exit Function
			Else
				Call SetVersion(arrRet, iWhere)
			End If
		
		End If
	End Select 	
End Function

'------------------------------------------  OpenPopUp()  ----------------------------------------------
'	Name :OpenPopUp
'	Description : open grid popup
'--------------------------------------------------------------------------------------------------------- 
Function OpenPopUp(Byval iWhere, Byval strCode, Byval strCode1)
	Dim arrRet, sTmp
	Dim arrParam(5), arrField(6), arrHeader(6)
	
	Dim tmpFlag, tmpLevel, tmpGpCd

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	With frm1
	
	ggoSpread.Source = frm1.vspdData
	.vspdData.Col = C_WcFlag	: tmpFlag = .vspdData.Text
	.vspdData.Col = C_GpLevel	: tmpLevel = .vspdData.Text
	.vspdData.Col = C_GpCd	: tmpGpCd = .vspdData.Text
	.vspdData.Col = iWhere-1		: strCode = .vspdData.Text		'code value
	
	Select Case iWhere		
		Case C_WcCdPop
			If tmpFlag="M" Then					'WC CD
				arrParam(0) = "공정/구매그룹"
				arrParam(1) = "dbo.ufn_getListOfPopup_C4004MA1('1')"	
				arrParam(2) = strCode
				arrParam(3) = ""	
				arrParam(4) =""			
							
				arrParam(5) = "공정/구매그룹" 

				arrField(0) = "ED12" & Parent.gColSep & "CODE"
				arrField(1) = "ED25"  & Parent.gColSep & "CD_NM"		
				
    
				arrHeader(0) = "공정/구매그룹"
				arrHeader(1) = "공정구매그룹명"
				
			ELSE											'PUR GROUP
				arrParam(0) = "공정/구매그룹"
				arrParam(1) = "dbo.ufn_getListOfPopup_C4004MA1('2')"	
				arrParam(2) = strCode
				arrParam(3) = ""			
				arrParam(4) =""		
				
				arrParam(5) = "공정/구매그룹" 

				arrField(0) = "ED12" & Parent.gColSep & "CODE"
				arrField(1) = "ED25"  & Parent.gColSep & "CD_NM"		
				
    
				arrHeader(0) = "공정/구매그룹"
				arrHeader(1) = "공정구매그룹명"
		
			End If

		Case C_GpLevelPop
			arrParam(0) = "계정그룹Level"						' 팝업 명칭 
			arrParam(1) = "( SELECT DISTINCT LEVEL_CD FROM  dbo.ufn_getListOfPopup_C4004MA1('3') ) AA "
			arrParam(2) = 	strCode ' Code Condition
			arrParam(3) = "" 	' Name Cindition
			arrParam(4) = ""							' Where Condition
			arrParam(5) = "계정그룹Level"							' TextBox 명칭	
	
			arrField(0) = "ED18" & Parent.gColSep & "LEVEL_CD"					' Field명(1)
			     
			arrHeader(0) = "계정그룹Level"						' Header명(0)

		
		Case C_GpCdPop
			arrParam(0) = "계정그룹"						' 팝업 명칭 
			arrParam(1) = " dbo.ufn_getListOfPopup_C4004MA1('4') "
			arrParam(2) = 	strCode ' Code Condition
			arrParam(3) = "" 	' Name Cindition
		
			If 	tmpLevel <> "" Then
				arrParam(4) = "LEVEL_CD=" & FilterVar(tmpLevel, "''", "S")
			Else 
				arrParam(4) = ""
			End If							' Where Condition
			arrParam(5) = "계정그룹"							' TextBox 명칭	
	
			arrField(0) = "ED18" & Parent.gColSep & "CODE"					' Field명(0)
			arrField(1) = "ED20" & Parent.gColSep & "CD_NM"					' Field명(1)
			arrField(2) = "ED15" & Parent.gColSep & "LEVEL_CD"					' Field명(2)			
			     
			arrHeader(0) = "계정그룹"						' Header명(0)
			arrHeader(1) = "계정그룹명"						' Header명(0)
			arrHeader(2) = "계정그룹Level"						' Header명(0)		


		Case C_AcctCdPop
			arrParam(0) = "계정 팝업"
			arrParam(1) = " (select distinct code, cd_nm  from dbo.ufn_getListOfPopup_C4004MA1('7') "
			If 	tmpLevel <> "" AND tmpGpCd<>""  Then
				arrParam(1) =  arrParam(1) & " where LEVEL_CD=" & FilterVar(tmpLevel, "''", "S")
				arrParam(1) = arrParam(1) & " AND TEMP_CD1=" & FilterVar(tmpGpCd, "''", "S")
			Else
				If 	tmpLevel <> "" Then
					arrParam(1) = arrParam(1)  & " where LEVEL_CD=" & FilterVar(tmpLevel, "''", "S")
				ElseIf 	tmpGpCd <> "" Then
					arrParam(1) = arrParam(1)  &  " where TEMP_CD1=" & FilterVar(tmpGpCd, "''", "S")
				End If
			End If
			arrParam(1) = arrParam(1) & ") a"
			arrParam(2) = 	strCode ' Code Condition
			arrParam(3) = "" 	' Name Cindition
			arrParam(4) =""
		
			arrParam(5) = "계정" 
			arrField(0) = "ED15" & Parent.gColSep & "CODE"
			arrField(1) = "ED25" & Parent.gColSep & "CD_NM"		
    
			arrHeader(0) = "계정"	
			arrHeader(1) = "계정명"
							
		Case C_AFctrCdPop
			arrParam(0) = "실제원가 배부요소 팝업"
			arrParam(1) = " dbo.ufn_getListOfPopup_C4004MA1('5') "
			arrParam(2) = strCode
			arrParam(3) = ""	
			arrParam(4) = ""
			arrParam(5) = "실제원가배부요소" 

			arrField(0) = "ED10"  & Parent.gColSep &"CODE"
			arrField(1) = "ED30"  & Parent.gColSep &"CD_NM"		
    
			arrHeader(0) = "실제원가배부요소"	
			arrHeader(1) = "실제원가배부요소명"
		Case C_SFctrCdPop
			arrParam(0) = "표준원가 배부요소 팝업"
			arrParam(1) = " dbo.ufn_getListOfPopup_C4004MA1('5') "
			arrParam(2) = strCode
			arrParam(3) = ""	
			arrParam(4) = ""
			arrParam(5) = "표준원가배부요소" 

			arrField(0) = "ED10"  & Parent.gColSep &"CODE"
			arrField(1) = "ED30"  & Parent.gColSep &"CD_NM"		
    
			arrHeader(0) = "표준원가배부요소"	
			arrHeader(1) = "표준원가배부요소명"

	End Select
    
	arrRet = window.showModalDialog("../../comasp/ADOCommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetPopUp(arrRet, iWhere)
	End If	

	End With
End Function
'------------------------------------------  SetPopUp()  ----------------------------------------------
'	Name :SetPopUp
'	Description : set grid popup
'--------------------------------------------------------------------------------------------------------- 
Function SetPopUp(Byval arrRet, Byval iWhere)
	Dim sTmp
	
	With frm1.vspdData
		ggoSpread.Source = frm1.vspdData
		Select Case iWhere
			
			Case C_WcCdPop
				.Col = C_WcCd	: .Text = arrRet(0)
				.Col = C_WcNm	: .Text = arrRet(1)
				Call SetActiveCell(frm1.vspdData,C_WcCd,.ActiveRow,"M","X","X")			
			Case C_GpLevelPop
				.Col = C_GpLevel	: .Text = arrRet(0)
				Call SetActiveCell(frm1.vspdData,C_GpLevel,.ActiveRow,"M","X","X")		
			Case C_GpCdPop
				.Col = C_GpCd			: .Text = arrRet(0)
				.Col = C_GpNm			: .Text = arrRet(1)
				.Col = C_GpLevel		: .Text = arrRet(2)
				Call SetActiveCell(frm1.vspdData,C_GpCd,.ActiveRow,"M","X","X")	
			Case C_AcctCdPop
				.Col = C_AcctCd		: .Text = arrRet(0)
				.Col =C_AcctNm			: .Text = arrRet(1)
				
				Call SetActiveCell(frm1.vspdData,C_AcctCd,.ActiveRow,"M","X","X")	
			Case  C_AFctrCdPop
				.Col = iWhere - 1		: .Text = arrRet(0)
				.Col = iWhere + 1		: .Text = arrRet(1)
				Call SetActiveCell(frm1.vspdData,C_AFctrCd,.ActiveRow,"M","X","X")	
			Case  C_SFctrCdPop
				.Col = iWhere - 1		: .Text = arrRet(0)
				.Col = iWhere + 1		: .Text = arrRet(1)
			Call SetActiveCell(frm1.vspdData,C_SFctrCd,.ActiveRow,"M","X","X")
		End Select			
		Call vspddata_Change(.ActiveCol, .ActiveRow)
		Set gActiveElement = document.activeElement
		lgBlnFlgChgValue = True
	End With
	
End Function
'------------------------------------------  SetVersion()  ----------------------------------------------
'	Name :SetVersion
'	Description : SetVersion
'--------------------------------------------------------------------------------------------------------- 
Function SetVersion(byval arrRet, Byval iWhere)

	If iWhere ="0" Then 
		frm1.txtVER_CD.focus
		frm1.txtVER_CD.Value    = arrRet(0)		
	Else
		lgCopyVersion = "Y"
		frm1.versionFlag.value=lgCopyVersion
		frm1.hVER_CD.value=ucase(arrRet(0))
		Call fncVersionCopy
	End If

End Function

'========================================================================================================
'========================================================================================================
'                        5.2 Common Method-2
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : Form_Load
' Desc : This sub is called from window_OnLoad in Common.vbs automatically
'========================================================================================================
Sub Form_Load()
	
    Call LoadInfTB19029
    Call ggoOper.LockField(Document, "N") 
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,Parent.gDateFormat,Parent.gComNum1000,Parent.gComNumDec)

    Call InitSpreadSheet
    Call InitVariables
    
	Call InitComboBox
    Call SetDefaultVal
    Call SetToolbar("111011010011111")	
    
     
End Sub

'========================================================================================================
' Name : Form_QueryUnload
' Desc : This sub is called from window_Unonload in Common.vbs automatically
'========================================================================================================
Sub Form_QueryUnload(Cancel, UnloadMode)
    '--------- Developer Coding Part (Start) ----------------------------------------------------------
    '--------- Developer Coding Part (End) ------------------------------------------------------------
End Sub

'========================================================================================================
'========================================================================================================
'                        5.4 User-defined Method
'========================================================================================================

'========================================================================================================
'========================================================================================================
'                        5.5 Tag Event
'========================================================================================================

'========================================================================================================
'   Event Name : vspdData_Click
'   Event Desc : 컬럼을 클릭할 경우 발생 
'========================================================================================================
'==========================================================================================
'   Event Name : vspdData_Click
'   Event Desc : This event is spread sheet data changed
'==========================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Dim IntRetCD
	
	If lgIntFlgMode = parent.OPMD_CMODE Then
		Call SetPopupMenuItemInf("0000110111")
	Else 	
		If frm1.vspdData.MaxRows = 0 Then 
			Call SetPopupMenuItemInf("1001111111")
		Else
			Call SetPopupMenuItemInf("1101111111") 
		End if			
	End If	
	
	gMouseClickStatus = "SPC"   
	Set gActiveSpdSheet = frm1.vspdData
	   
	If frm1.vspdData.MaxRows <= 0 Then
		Exit Sub
	End If
		
	If Row <= 0 Then
		ggoSpread.Source = frm1.vspdData 
		If lgSortKey = 1 Then
			ggoSpread.SSSort Col					'Sort in Ascending
			lgSortKey = 2
		Else
			ggoSpread.SSSort Col, lgSortKey		'Sort in Descending
			lgSortKey = 1
		End If
	End If
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
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
'   Event Name : vspdData_DblClick
'   Event Desc : 
'========================================================================================================
Sub vspdData_DblClick(ByVal Col, ByVal Row)				
    
	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 
    	
End Sub


'==========================================================================================
'   Event Desc : Spread Split 상태코드 
'==========================================================================================
Sub vspdData_MouseDown(Button, Shift, X, Y)
	If Button <> "1" And gMouseClickStatus = "SPC" Then
		gMouseClickStatus = "SPCR"
	End If

End Sub


'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_Change(ByVal Col, ByVal Row)
    Frm1.vspdData.Row = Row
    Frm1.vspdData.Col = Col
	
	Call CheckMinNumSpread(frm1.vspdData, Col, Row)

	ggoSpread.Source = frm1.vspdData
	ggoSpread.UpdateRow Row
	lgBlnFlgChgValue = True

	'------ Developer Coding part (Start ) -------------------------------------------------------------- 
	Dim sFromSQL, sWhereSQL, sVal, sCd, sCdNm, sTmp, sLvl
	Dim tmpLevel,tmpGpCd, tmpFlag
	
	sFromSQL = " dbo.ufn_getListOfPopup_C4004MA1"

	
	With frm1.vspdData
		.Col = C_GpLevel : tmpLevel =.Text
		.Col = C_GpCd     : tmpGpCd = .Text
		.Col = C_WcFlag     : tmpFlag = .Text
		.Row = Row	: .Col = Col : sVal = UCase(Trim(.Text))
		
		
		Select Case Col
			Case C_WcCd
				If tmpFlag="M" Then 
				sFromSQL = sFromSQL & "('1')"
				Else
				sFromSQL = sFromSQL & "('2')" 
				End If 
				
				sWhereSQL = "CODE = " & FilterVar(sVal,"''","S")

			Case C_GpLevel
				sFromSQL = sFromSQL & "('3')" 

				sWhereSQL = " LEVEL_CD = " & FilterVar(sVal, "''", "S")				
			Case C_GpCd
				sFromSQL = sFromSQL & "('4')" 
				sWhereSQL = " CODE = " & FilterVar(sVal, "''", "S")				
     			If tmpLevel <> "" Then
					sWhereSQL = sWhereSQL & " AND LEVEL_CD = " & FilterVar(tmpLevel, "''", "S")
				End If
			Case C_AcctCd
				sFromSQL=" (select distinct code, cd_nm ,'' level_cd  from dbo.ufn_getListOfPopup_C4004MA1('7') "
				If 	tmpLevel <> "" AND tmpGpCd<>""  Then
					sFromSQL =  sFromSQL& " where LEVEL_CD=" & FilterVar(tmpLevel, "''", "S")
					sFromSQL= sFromSQL & " AND TEMP_CD1=" & FilterVar(tmpGpCd, "''", "S")
				Else
					If 	tmpLevel <> "" Then
						sFromSQL = sFromSQL  & " where LEVEL_CD=" & FilterVar(tmpLevel, "''", "S")
					ElseIf 	tmpGpCd <> "" Then
						sFromSQL = sFromSQL &  " where TEMP_CD1=" & FilterVar(tmpGpCd, "''", "S")
					End If
				End If
				sFromSQL= sFromSQL & ") a"
				sWhereSQL = " CODE = " & FilterVar(sVal, "''", "S")		
			Case C_AFctrCd
				sFromSQL = sFromSQL & "('5')" 
				sWhereSQL = " CODE = " & FilterVar(sVal, "''", "S")
			Case C_SFctrCd
				sFromSQL = sFromSQL & "('5')" 
				sWhereSQL = " CODE = " & FilterVar(sVal, "''", "S")
		End Select
	   
	   .Row = Row
	   
		If sWhereSQL <> "" Then
			' -- DB 콜 
			If CommonQueryRs(" TOP 1 CODE, CD_NM, LEVEL_CD ", sFromSQL , sWhereSQL, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
				sCd		= Replace(lgF0, Chr(11), "")
				sCdNm	= Replace(lgF1, Chr(11), "")
				sLvl	= Replace(lgF2, Chr(11), "")
				
				' -- 존재시 코드명을 출력한다.
				Select Case Col
					Case C_WcCd
						.Col =C_WcNM
						.Text = sCdNm

					Case C_GpCd
						.Col = C_GpNm	: .Text = sCdNm
						.Col = C_GpLevel: .Text = sLvl
						
						.Col = C_AcctCd	: .Text = ""
						.Col = C_AcctNm	: .Text = ""

					Case  C_AFctrCd
						.Col = C_AFctrNM
						.Text = sCdNm
					Case  C_SFctrCd
						.Col = C_SFctrNm	
						.Text = sCdNm

					Case C_GpLevel
							.Col = C_GpCd	:			.Text = ""
							.Col = C_GpNm 	:			.Text = ""			
							.Col = C_AcctCd :			.Text = ""		
							.Col =C_AcctNm:       	.Text = ""		
				End Select
			Else
				' -- 비존재시 메시지 처리 
				If sVal <> "" Then
					Call DisplayMsgBox("970000", "x",sVal,"x")
					Call SetFocusToDocument("M")
					.Focus
				End If
				
				' -- 명 들을 지운다 
				Select Case Col
					Case C_WcCd, C_AFctrCd,C_SFctrCd
						.Col = Col		: .Text = ""
						.Col = C_WcCd+2	: .Text = ""

					Case C_GpLevel
						.Col = Col		: .Text = ""
						.Col = C_GpCd	: .Text = ""
						.Col = C_GpNm	: .Text = ""
						.Col = C_AcctCd	: .Text = ""
						.Col = C_AcctNm 	: .Text = ""
					Case C_GpCd
						.Col = Col		: .Text = ""
						.Col = C_GpNm	: .Text = ""
						.Col = C_AcctCd	: .Text = ""
						.Col = C_AcctNm 	: .Text = ""
				End Select		
			End If		
		End If
	End With
	'------ Developer Coding part (End   ) -------------------------------------------------------------- 	
End Sub


'========================================================================================================
'   Event Name : vspdData_ComboSelChange
'   Event Desc : This function is ComboSelChange with spread sheet 
'========================================================================================================
Sub vspdData_ComboSelChange(ByVal Col, ByVal Row)
	Dim index
	
	With frm1.vspdData

	    ggoSpread.Source = frm1.vspdData
		
		If Row = 0 Then Exit Sub
	
		.Row = Row
			
		Select Case Col
			Case C_WcFlagNm
				.Col = Col
				index = .Value
				
				.Col = C_WcFlag
				.Value = index

			.Col = C_WcCd : .Text =""
			.Col = C_WcNm : .Text =""
			Call SetActiveCell(frm1.vspdData,C_WcCd,.ActiveRow ,"M","X","X")   
		End Select

	End With	
End Sub

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : 
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	Dim sCode, sCode2
	
	With frm1 
	
	    ggoSpread.Source = frm1.vspdData
   
		If Row = 0 Then Exit Sub
		
		Select Case Col
			Case C_WcCdPop, C_GpLevelPop, C_GpCdPop, C_AcctCdPop, C_AFctrCdPop,C_SFctrCdPop
				.vspdData.Col = Col - 1
				.vspdData.Row = Row				
				sCode = UCase(Trim(.vspdData.Text))				
				Call OpenPopup(Col, sCode, sCode2)				
		End Select        
    End With
    
End Sub

'========================================================================================================
'   Event Name : vspdData_ScriptDragDropBlock
'   Event Desc : 
'========================================================================================================
Sub vspdData_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
  
   ggoSpread.Source = frm1.vspdData
   Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)
   Call GetSpreadColumnPos("A")
End Sub


Sub vspdData_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
    
    If OldLeft <> NewLeft Then
        Exit Sub
    End If
    
    
    '----------  Coding part  -------------------------------------------------------------   
    if frm1.vspdData.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData, NewTop)	Then
		If lgStrPrevKey <> "" Then							'⊙: 다음 키 값이 없으면 더 이상 업무로직ASP를 호출하지 않음 
			Call DisableToolBar(parent.TBC_QUERY)  
			If DbQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
    End if
    
End Sub

'========================================================================================================
' Name : FncQuery
' Desc : This function is called from MainQuery in Common.vbs
'========================================================================================================
Function FncQuery() 
  Dim IntRetCD 
	
    FncQuery = False															'⊙: Processing is NG

    Err.Clear																    '☜: Protect system from crashing

	'-----------------------
    'Check condition area
    '-----------------------
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field		
       Exit Function
    End If
    
	'IF ChkKeyField()=False Then Exit Function 
    '-----------------------
    'Check previous data area
    '-----------------------
    ggoSpread.Source = frm1.vspdData                          '⊙: Preset spreadsheet pointer 
    If ggoSpread.SSCheckChange = True Then                   '⊙: Check If data is chaged
		IntRetCD = DisplayMsgBox("900013", parent.VB_YES_NO,"X","X")					'⊙: "Will you destory previous data"
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    '-----------------------
    'Erase contents area
    '-----------------------    
	Call ggoSpread.ClearSpreadData
	Call InitVariables
    '-----------------------
    'Query function call area
    '-----------------------
    If DbQuery = False Then   
		Exit Function           
    End If     												'☜: Query db data

    FncQuery = True	
    
End Function

'========================================================================================================
' Name : FncNew
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncNew() 
    Dim IntRetCD 
    
    FncNew = False 
    
    Err.Clear     

    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
    	IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")   
			If IntRetCD = vbNo Then
				Exit Function
			End If
    End If
    
    Call SetToolbar("111011010011111")

    Call ggoOper.ClearField(Document, "1") 
    Call ggoOper.ClearField(Document, "2") 
    ggoSpread.Source = frm1.vspdData
	ggoSpread.ClearSpreadData
     
    Call ggoOper.LockField(Document, "N") 
    Call InitVariables 
    Call SetDefaultVal    
    FncNew = True 

End Function

'========================================================================================================
' Name : FncDelete
' Desc : This function is called from MainNew in Common.vbs
'========================================================================================================
Function FncDelete() 
    Dim IntRetCD 

    FncDelete = False
    If lgIntFlgMode <> parent.OPMD_UMODE Then
        Call DisplayMsgBox("900002", "X", "X", "X")
        Exit Function
    End If

    IntRetCD = DisplayMsgBox("900003",parent.VB_YES_NO, "X", "X")
    IF IntRetCD = vbNO Then
		Exit Function
	End IF
'check the version using or not
	If CommonQueryRs(" TOP 1 ver_cd  ", " C_WORK_VERSION_S " , " work_step ='07' and ver_cd=" & FilterVar(trim(frm1.txtVER_CD.Value),"''","S") , lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
		IntRetCD= DisplayMsgBox("236327",parent.VB_YES_NO, "X", "X")			
		If intRetCd=vbNo Then
			Exit Function
		End IF
	End If	
    Call DbDelete

    FncDelete = True
End Function

'========================================================================================================
' Name : FncSave
' Desc : This function is called from MainSave in Common.vbs
'========================================================================================================
Function FncSave() 
    Dim IntRetCD , blnChange1,  iRow, iSeqNo
    
    FncSave = False
    
    Err.Clear     
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field		
       Exit Function
    End If
    
    ggoSpread.Source = frm1.vspddata
    
        
	If ggoSpread.SSCheckChange = False Then								  '☜:match pointer
        IntRetCD = DisplayMsgBox("900001", "X", "X", "X")				  '☜:There is no changed data.  
        Exit Function
    End If    
        
    If Not ggoSpread.SSDefaultCheck Then      
       Exit Function
    End If
       
    IF DbSave = False Then
		Exit function
	END IF

    FncSave = True      
    
End Function

'========================================================================================================
' Name : FncCopy
' Desc : This function is called from MainCopy in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy() 
	   If frm1.vspdData.maxrows < 1 Then Exit Function
    
    frm1.vspdData.focus 
    Set gActiveElement = document.activeElement    
		    
	frm1.vspdData.ReDraw = False    	    
    ggoSpread.Source = frm1.vspdData	
        
    ggoSpread.CopyRow   
    
    With frm1			
   
		frm1.vspdData.ReDraw = True    
       
	    SetSpreadColor frm1.vspdData.ActiveRow, frm1.vspdData.ActiveRow   
	    .vspdData.Focus
    	Call SetActiveCell(frm1.vspdData,C_WcCd,frm1.vspdData.ActiveRow,"M","X","X")
    End With
	
End Function


'========================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================

Function FncCancel() 
	If frm1.vspdData.maxrows < 1 Then Exit Function
	
    ggoSpread.EditUndo                                                  '☜: Protect system from crashing
    
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : This function is called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow(ByVal pvRowCnt)
    Dim iIntReqRows
    Dim iIntCnt

    On Error Resume Next
    Err.Clear                                                                     '☜: Clear error status
    
    FncInsertRow = False                                                         '☜: Processing is NG

	If IsNumeric(Trim(pvRowCnt)) Then
		iIntReqRows = CInt(pvRowCnt)
	Else
		iIntReqRows = AskSpdSheetAddRowCount()
		If iIntReqRows = "" Then
		    Exit Function
		End If
	End If
	'-----------------------
    'Check condition area
    '-----------------------
   ' If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
  '     Exit Function
  '  End If
    
    With frm1			
		.vspdData.ReDraw = False
		.vspdData.focus

	    ggoSpread.Source = .vspdData
        ggoSpread.InsertRow , iIntReqRows

		Call SetSpreadColor(.vspdData.ActiveRow, .vspdData.ActiveRow + iIntReqRows - 1)	
		.vspdData.ReDraw = True
     
    End With    

    Set gActiveElement = document.activeElement 

	If Err.number = 0 Then
		FncInserRow = True
	End IF
End Function

'========================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================

Function FncDeleteRow() 
  
    Dim lDelRows
    Dim iDelRowCnt

    '----------------------
    ' 데이터가 없는 경우 
    '----------------------
    If frm1.vspdData.maxrows < 1 Then Exit Function
	
	ggoSpread.Source = frm1.vspdData 
	lDelRows = ggoSpread.DeleteRow
End Function

'========================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================

Function FncPrint()
    Call parent.FncPrint() 
End Function

'========================================================================================
' Function Name : FncPrev
' Function Desc : This function is related to Previous Button
'========================================================================================

Function FncPrev() 
End Function
'========================================================================================
' Function Name : FncNext
' Function Desc : This function is related to Next Button
'========================================================================================

Function FncNext() 
End Function
'========================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel 
'========================================================================================

Function FncExcel() 
    Call parent.FncExport(Parent.C_MULTI)
End Function
'========================================================================================
' Function Name : FncFind
' Function Desc : 
'========================================================================================

Function FncFind() 
    Call parent.FncFind(Parent.C_MULTI, False) 
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
    Call InitComboBox
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
	Call SetSpreadColorQuery
End Sub

'========================================================================================================
' Name : FncExit
' Desc : This function is called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
Dim IntRetCD
	
	FncExit = False
	
    ggoSpread.Source = frm1.vspdData	
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900016", Parent.VB_YES_NO,"x","x")
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True
End Function

'========================================================================================================
'========================================================================================================
'                        5.3 Common Method-3
'========================================================================================================
'========================================================================================================

'========================================================================================================
' Name : DbQuery
' Desc : This sub is called by FncQuery
'========================================================================================================
Function DbQuery() 
	Dim strVal

    DbQuery = False
    IF LayerShowHide(1) = False Then
		Exit Function
	END IF
    Err.Clear	

    With frm1
			strVal = BIZ_PGM_ID2 & "?txtMode=" & Parent.UID_M0001
			strVal = strVal & "&lgStrPrevKey=" & lgStrPrevKey				
			strVal = strVal & "&txtVER_CD=" & ucase(Trim(frm1.txtVER_CD.value))
			strVal = strVal & "&hVER_CD=" & ucase(Trim(frm1.hVER_CD.value)		)
			strVal = strVal & "&txtMaxRows=" & .vspdData.MaxRows
			strVal = strVal & "&lgCopyVersion=" & lgCopyVersion
	 
		Call RunMyBizASP(MyBizASP, strVal)   
    End With
    
    DbQuery = True

End Function

'========================================================================================================
' Name : DbQueryOk
' Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()	

	lgIntFlgMode = Parent.OPMD_UMODE    
	Call ggoOper.LockField(Document, "Q")
	call SetSpreadColorQuery
	Call SetToolbar("111111110011111")

	lgBlnFlgChgValue =False
	Frm1.vspdData.Focus   	
    Set gActiveElement = document.ActiveElement   
   	
End Function


'========================================================================================================
' Name : InitData()
' Desc : Reset ComboBox
'========================================================================================================
Sub InitData()
End Sub

'========================================================================================================
' Name : DbSave
' Desc : This sub is called by FncSave
'========================================================================================================
Function DbSave() 
    Dim lRow        
    Dim lGrpCnt     
	Dim strVal, strDel
    Dim iColSep 
    Dim iRowSep  
    Dim sSQLI1, sSQLI2, sSQLU1, sSQLU2, sSQLD1, sSQLD2, sVerCd
	
	Dim changeFlag
	Dim tmpA, tmpB
	
    DbSave = False 
    IF LayerShowHide(1) = False Then
		Exit Function
	END IF
	
	iColSep = Parent.gColSep
	iRowSep = Parent.gRowSep	

	lGrpCnt = 1
	strVal = ""
	strDel = ""
	sVerCd = UCase(Trim(frm1.txtVER_CD.value))
	tmpA="" :tmpB=""

	With frm1.vspdData
	
		ggoSpread.Source = frm1.vspdData	
		
		For lRow = 1 To .MaxRows    
			.Row = lRow	: .Col = 0        
			Select Case .Text

	            Case ggoSpread.InsertFlag	
					strVal = "C" & iColSep 
					strVal = strVal & lRow &  iColSep
					changeFlag="Y"
				Case ggoSpread.UpdateFlag		
	            
					strVal = "U" & iColSep 
					strVal = strVal & lRow &  iColSep
					changeFlag="Y"
				Case ggoSpread.DeleteFlag		

					strVal = "D" & iColSep 
					strVal = strVal & lRow &  iColSep
					changeFlag="Y"
				Case Else
				changeFlag="N"

				End Select 
				If changeFlag="Y" Then 
					.Col = C_VerCd				: strVal = strVal & Trim(.Text) & iColSep
					.Col = C_WcFlag		: strVal = strVal & Trim(.Text) & iColSep
					.Col = C_WcCd		: strVal = strVal & Trim(.Text) & iColSep
					.Col = C_GpLevel						
					If Trim(.Text)<>"" Then
					strVal = strVal & Trim(.Text) & iColSep
					Else
					strVal = strVal & "0" & iColSep
					End If
					.Col = C_GpCd				
					If Trim(.Text)<>"" Then
					strVal = strVal & Trim(.Text) & iColSep
					Else
					strVal = strVal & "*" & iColSep
					End If
					.Col = C_AcctCd			
					If Trim(.Text)<>"" Then
					strVal = strVal & Trim(.Text) & iColSep
					Else
					strVal = strVal & "*" & iColSep
					End If
					.Col = C_AFctrCd		: strVal = strVal & Trim(.Text) & iColSep 		 : tmpA=Trim(.Text)	
					.Col = C_SFctrCd		: strVal = strVal & Trim(.Text) & iColSep &  Parent.gRowSep :tmpB=Trim(.Text)
					
					If tmpA="" and tmpB="" then 						
						Call LayerShowHide(0)
						Call DisplayMsgBox("236324", "X","X","X")
						Call SetActiveCell(frm1.vspdData,C_AFctrCd,.ActiveRow,"M","X","X")	
						Exit Function 
					End If

					sSQLI1 = sSQLI1 + strVal
					lGrpCnt = lGrpCnt + 1
				End If                
		Next

	End With

			
	frm1.txtMode.value = Parent.UID_M0002
	frm1.txtMaxRows.value = lGrpCnt-1


	frm1.txtSpread.value =sSQLI1 
	
	
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
	
    DbSave = True    
    
End Function

'========================================================================================================
' Name : DbSaveOk
' Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()	

	Call InitVariables
	frm1.vspdData.MaxRows = 0

	Call MainQuery()

End Function


'========================================================================================
Function DbDelete() 
    Err.Clear
    DbDelete = False

    Dim strVal

    strVal = BIZ_PGM_ID & "?txtMode="	& parent.UID_M0003							'☜: 비지니스 처리 ASP의 상태 
    strVal = strVal     & "&txtVER_CD=" & frm1.txtVER_CD.value					    '☜: Query Key        
	
	Call RunMyBizASP(MyBizASP, strVal)

    DbDelete = True
End Function


'========================================================================================
Function DbDeleteOk()
	Call InitVariables
	Call FncNew()
End Function
'=================================================================================
'	Name : CheckPop()
'	Description : check the valid data
'=========================================================================================================
Function CheckPop(ByVal strLevel, Byval strGp, ByVal strAcct,ByVal iCol, Byval iRow)
	Dim strFrom
	Dim strWhere
	
	strFrom ="  (SELECT	B.GP_LVL, A.GP_CD, B.GP_NM,  A.ACCT_CD , A.ACCT_NM "
	strFrom= strFrom  & "	FROM	A_ACCT	A"
	strFrom = strFrom  & "	INNER JOIN A_ACCT_GP	B  ON A.GP_CD = B.GP_CD"
	strFrom= strFrom  & "		AND A.TEMP_FG_3 IN  ('M2', 'M3', 'M4') AND A.DEL_FG='N') AA	"
	
	strWhere= " ACCT_CD = " & FilterVar(strAcct,"''","S")
	If 	strLevel <> "" AND strGp<>""  Then
		strWhere =strWhere &  " AND GP_LVL=" & FilterVar(strLevel, "''", "S")
		strWhere = strWhere & " AND GP_CD=" & FilterVar(strGp, "''", "S")
	Else
		If 	strLevel <> "" Then
			strWhere=strWhere &  " AND GP_LVL=" & FilterVar(strLevel, "''", "S")
		ElseIf 	strGp <> "" Then
			strWhere = strWhere & " AND GP_CD=" & FilterVar(strGp, "''", "S")
		End If
	End If
	
	If CommonQueryRs(" TOP 1 GP_LVL, GP_CD, GP_NM, ACCT_NM ", strFrom , strWhere, lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6) Then
			with frm1.vspdData
				 .Row = iRow
				.Col = C_GpLevel : .Text =  Replace(lgF0, Chr(11), "")
				.Col = C_GpCd : .Text =  Replace(lgF1, Chr(11), "")
				.Col = C_GpNm : .Text =  Replace(lgF2, Chr(11), "")
				.Col = C_AcctNm : .Text =  Replace(lgF3, Chr(11), "")
			End With
	Else	
		Call DisplayMsgBox("970000", "x",strAcct,"x")
		Call SetFocusToDocument("M")
	End If 
	
End Function 
'=================================================================================
'	Name : ChkKeyField()
'	Description : check the valid data
'=========================================================================================================
Function ChkKeyField()
	Dim strDataCd, strDataNm
    Dim strWhere , strFrom 
    Dim lgF0, lgF1, lgF2, lgF3, lgF4, lgF5, lgF6
    
    Err.Clear                                       

	ChkKeyField = true		
'check version 
	If Trim(frm1.txtVER_CD.value) <> "" Then
		strWhere = " ver_cd= " & FilterVar(frm1.txtVER_CD.value, "''", "S") & "  "

		Call CommonQueryRs(" ver_cd ","	 C_MFC_DSTB_RULE_BY_ORDER_S ", strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
	
		IF Len(lgF0) < 1 Then 
			Call DisplayMsgBox("970000","X",frm1.txtVER_CD.alt,"X")
			frm1.txtVER_CD.focus 
			ChkKeyField = False
			Exit Function
		End If
	
		strDataNm = split(lgF0,chr(11))
		frm1.txtVER_CD.value = strDataNm(0)
	End If
End Function
'=================================================================================
'	Name : SheetFocus()
'	Description : set err. position 
'=========================================================================================================
Function SheetFocus(lRow, lCol)
	frm1.vspdData.focus
	frm1.vspdData.Row = lRow
	frm1.vspdData.Col = lCol
	frm1.vspdData.Action = 0
	frm1.vspdData.SelStart = 0
	frm1.vspdData.SelLength = len(frm1.vspdData.Text)
End Function
'=================================================================================
'	Name : fncVersionCopy()
'	Description : 
'=========================================================================================================
Function fncVersionCopy()
	IF lgBlnFlgChgValue THEN lgBlnFlgChgValue =False
	
	  If Not chkField(Document, "1") Then									'⊙: This function check indispensable field		
       Exit Function
    End If
	If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgBox("900015", Parent.VB_YES_NO,"x","x")   
			If IntRetCD = vbNo Then
				Exit Function
			End If
	End If
    ggoSpread.Source = frm1.vspddata
    Call ggoSpread.ClearSpreadData
        
    
	Call ExecMyBizASP(frm1, BIZ_PGM_ID2)
End Function 

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc"  -->	
</HEAD>

<BODY TABINDEX="-1" SCROLL="no" oncontextmenu="javascript:return false">
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white><%=Request("strASPMnuMnuNm")%></font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD WIDTH=* align=right><a onclick="vbscript:Call OpenVersion(1)">타 Version Copy</a>&nbsp;&nbsp;</TD>
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
									<TD CLASS="TD5">Version</TD>
									<TD CLASS="TD656"><INPUT ClASS="clstxt" TYPE=TEXT NAME="txtVER_CD" SIZE=10 MAXLENGTH=3 tag="12XXXU" ALT="Version"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDstbFctr" align=top TYPE="BUTTON"  ONCLICK="vbscript:Call OpenVersion(0)"  OnMouseOver="vbscript:PopUpMouseOver()"  OnMouseOut="vbscript:PopUpMouseOut()">
									</TD>
								</TR>    
							</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR>
					<TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD>
				</TR>
				<TR>
					<TD WIDTH=100% HEIGHT=100% valign=top>
						<TABLE <%=LR_SPACE_TYPE_20%>>
							<TR>
								<TD HEIGHT="50%" NOWRAP>
								<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> ID=vspdData NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD"> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"> </OBJECT>');</SCRIPT>
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
			<IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=no  noresize framespacing=0 TABINDEX= "-1"></IFRAME>
		</TD>
	</TR>
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24" TABINDEX= "-1"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="hVER_CD" tag="24" TABINDEX= "-1">
<INPUT TYPE=HIDDEN NAME="versionFlag" tag="24" TABINDEX= "-1">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
</BODY>
</HTML>

