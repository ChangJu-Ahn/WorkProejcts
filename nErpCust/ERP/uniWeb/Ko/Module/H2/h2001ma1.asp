<%@ LANGUAGE="VBScript" CODEPAGE=949 %>
<% session.CodePage=949 %>

<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : 인사마스타등록 
*  3. Program ID           : H2001ma1
*  4. Program Name         : H2001ma1
*  5. Program Desc         : 인사기본자료관리/인사마스타등록 
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/05/09
*  8. Modified date(Last)  : 2003/06/10
*  9. Modifier (First)     : YBI
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="javascript" SRC="../../inc/TabScript.js"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/Cookie.vbs">   </SCRIPT>

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>

<Script Language="VBScript">
Option Explicit                                                        '☜: indicates that All variables must be declared in advance
'========================================================================================================
'=                       4.2 Constant variables 
'========================================================================================================
Const BIZ_PGM_ID      = "h2001mb1.asp"						           '☆: Biz Logic ASP Name

Const BIZ_PGM_JUMP_ID  = "h2003ma1"                                         '가족사항등록 
Const BIZ_PGM_JUMP_ID1 = "h2004ma1"                                         '학력사항등록 
Const BIZ_PGM_JUMP_ID2 = "h2007ma1"                                         '경력사항등록 
Const BIZ_PGM_JUMP_ID3 = "h2005ma1"                                         '보증인등록 
Const BIZ_PGM_JUMP_ID4 = "h2008ma1"                                         '자격/면허등록 
Const BIZ_PGM_JUMP_ID5 = "h3001ma1"                                         '인사변동등록 
Const BIZ_PGM_JUMP_ID6 = "h3010ma1"                                         '교육사항등록 
Const BIZ_PGM_JUMP_ID7 = "h3009ma1"                                         '상벌등록 
Const BIZ_PGM_JUMP_ID8 = "h3012ma1"                                         '고과사항등록 
Const BIZ_PGM_JUMP_ID9 = "h3013ma1"                                         '어학자격등록 
Const BIZ_PGM_JUMP_ID10= "h2002ma1"                                        '사진등록 

Const BIZ_PGM_JUMP_ID11= "b2903ma1"                                         '조직도 

Const TAB1 = 1
Const TAB2 = 2
Const TAB3 = 3
Const TAB4 = 4

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

Dim IsOpenPop						                                    ' Popup
Dim lsGetsvrDate

Dim temp_txtDept_cd, temp_txtRoll_pstn, temp_txtFunc_cd
Dim temp_txtRole_cd, temp_txtPay_grd1,  temp_txtPay_grd2,  temp_txtSect_cd
Dim temp_flg_chk

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      =  parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '⊙: Indicates that no value changed
	lgIntGrpCount     = 0										'⊙: Initializes Group View Size
    lgStrPrevKey      = ""                                      '⊙: initializes Previous Key
    lgSortKey         = 1                                       '⊙: initializes sort direction

	gblnWinEvent      = False
	lgBlnFlawChgFlg   = False

	gIsTab = "Y"
	gTabMaxCnt = 4
End Sub

'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================
	
Sub SetDefaultVal()
	lsGetsvrDate = "<%=GetsvrDate%>"
End Sub
	
'========================================================================================================
' Name : LoadInfTB19029()	
' Desc : Set System Number format
'========================================================================================================
Sub LoadInfTB19029()
	<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
	<% Call loadInfTB19029A("I", "H","NOCOOKIE","MA") %> 
End Sub

'========================================================================================================
'	Name : CookiePage()
'	Description : Item Popup에서 Return되는 값 setting
'========================================================================================================
Function CookiePage(ByVal flgs)
	On Error Resume Next

	Const CookieSplit     = 4877						
	Const DeptCookieSplit = 5877						
	Dim strTemp

	If flgs = 1 Then
		 WriteCookie CookieSplit , frm1.txtEmp_no.Value
		 WriteCookie DeptCookieSplit , frm1.txtdept_cd.Value
	ElseIf flgs = 0 Then

		strTemp =  ReadCookie(CookieSplit)
			If strTemp = "" then Exit Function
			
		frm1.txtEmp_no1.value =  strTemp

		If Err.number <> 0 Then
			Err.Clear
			 WriteCookie CookieSplit , ""
			 WriteCookie DeptCookieSplit , ""
			Exit Function 
		End If

		 WriteCookie CookieSplit , ""
		 WriteCookie DeptCookieSplit , ""
		Call MainQuery()
			
	End If

End Function

'========================================================================================================
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pOpt)
   lgKeyStream = UCase(Frm1.txtEmp_no1.Value) &  parent.gColSep    ' 사번 
   lgKeyStream = lgKeyStream & lgUsrIntCd &  parent.gColSep ' 자료권한 
End Sub        
	
'========================================================================================================
' Name : InitComboBox()
' Desc : developer describe this line Initialize ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iCodeArr 
    Dim iNameArr

	' 최종학력 
    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0007", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call  SetCombo2(frm1.txtSch_ship, lgF0, lgF1, Chr(11))

	' 퇴직사유 
    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0025", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call  SetCombo2(frm1.txtRetire_resn, lgF0, lgF1, Chr(11))

    ' 결혼여부 
    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0105", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call  SetCombo2(frm1.txtmarry_cd, lgF0, lgF1, Chr(11))    

    ' 신고사업장 
    Call  CommonQueryRs(" YEAR_AREA_CD, YEAR_AREA_NM "," HFA100T ","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call  SetCombo2(frm1.txtYear_area_cd, lgF0, lgF1, Chr(11))    

    ' 혈액형1
    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0106", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call  SetCombo2(frm1.txtBlood_type1, lgF0, lgF1, Chr(11))    

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0107", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call  SetCombo2(frm1.txtBlood_type2, lgF0, lgF1, Chr(11))    

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0013", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call  SetCombo2(frm1.txtparia_cd, lgF0, lgF1, Chr(11))    

    Call  CommonQueryRs(" MINOR_CD,MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0014", "''", "S") & " ",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    Call  SetCombo2(frm1.txtrelief_cd, lgF0, lgF1, Chr(11))    


End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format
		
	Call  AppendNumberPlace("6", "6", "0")
	Call  AppendNumberPlace("7", "3", "2")
	Call  AppendNumberPlace("8", "1", "1")
    Call  AppendNumberPlace("9", "3", "0")
	Call  AppendNumberRange("0", "-12x34", "13x440")
	
	Call  ggoOper.FormatField(Document, "2", ggStrIntegeralPart,  ggStrDeciPointPart, gDateFormat, parent.gComNum1000, parent.gComNumDec)

	Call  ggoOper.LockField(Document, "N")											'⊙: Lock Field

    Call  FuncGetAuth(gStrRequestMenuID,  parent.gUsrID, lgUsrIntCd)     ' 자료권한:lgUsrIntCd ("%", "1%")
	
    Call SetDefaultVal()
    gSelframeFlg = TAB1
	Call SetToolBar("1110100000000111")												'⊙: Set ToolBar

	Call InitVariables

    Call changeTabs(TAB1)
	frm1.txtEmp_no1.focus
    gIsTab     = "Y" ' <- "Yes"의 약자 Y(와이) 입니다.[V(브이)아닙니다]
    gTabMaxCnt = 4   ' Tab의 갯수를 적어 주세요    

    Call InitComboBox
	Call CookiePage (0)                                                             '☜: Check Cookie

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
    
    FncQuery = False															 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    If  lgBlnFlgChgValue = True Then
		IntRetCD =  DisplayMsgBox("900013",  parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to display it? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    Call  ggoOper.ClearField(Document, "2")										 '☜: Clear Contents  Field
    
    If Not chkField(Document, "1") Then									         '☜: This function check required field
       Exit Function
    End If
    
    if  frm1.txtEmp_no1.value = "" AND frm1.txtName1.value <> "" then
        OpenEmpName(0)
        exit function
    end if

    Call InitVariables                                                           '⊙: Initializes local global variables
    Call MakeKeyStream("Q")
    
	Call  DisableToolBar( parent.TBC_QUERY)
    If DbQuery = False Then
        Call  RestoreToolBar()
        Exit Function
    End If
       
    FncQuery = True                                                              '☜: Processing is OK

End Function
	
'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    If lgBlnFlgChgValue = True Then
       IntRetCD =  DisplayMsgBox("900015",  parent.VB_YES_NO,"x","x")					 '☜: Data is changed.  Do you want to make it new? 
       If IntRetCD = vbNo Then
          Exit Function
       End If
    End If
    
    Call  ggoOper.ClearField(Document, "A")                                       '☜: Clear Condition Field
    Call  ggoOper.LockField(Document , "N")                                       '☜: Lock  Field
    
    Call SetToolBar("11101000000001")
    Call SetDefaultVal
    Call InitVariables                                                        '⊙: Initializes local global variables
    Call changeTabs(TAB1)
	Frm1.imgPhoto.src = ""    
    frm1.txtEmp_no.focus()    
    Set gActiveElement = document.ActiveElement   
    
    FncNew = True																 '☜: Processing is OK
End Function
	
'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
    Dim intRetCD
    
    FncDelete = False                                                            '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call  DisplayMsgBox("900002","x","x","x")                                
        Exit Function
    End If
    
    IntRetCD =  DisplayMsgBox("900003",  parent.VB_YES_NO,"x","x")                        '☜: Do you want to delete? 
	If IntRetCD = vbNo Then
		Exit Function
	End If
    
    Call MakeKeyStream("D")
    
	Call  DisableToolBar( parent.TBC_DELETE)
    If DbDelete = False Then
        Call  RestoreToolBar()
        Exit Function
    End If

    Set gActiveElement = document.ActiveElement   
    
    FncDelete = True                                                            '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    Dim res_no1, res_no2            ' 주민번호 
    Dim intChk, intMod, intDef      ' 주민번호 
    Dim strWhere
    Dim strTab

    FncSave = False                                                              '☜: Processing is NG
    
    Err.Clear                                                                    '☜: Clear err status
    If lgBlnFlgChgValue = False Then 
        IntRetCD =  DisplayMsgBox("900001","x","x","x")                           '☜:There is no changed data. 
        Exit Function
    End If
    
    If Not chkField(Document, "2") Then                                          '☜: Check contents area
       Exit Function
    End If

	strTab = gSelframeFlg

	  With Frm1
           if .txtNat_cd.value<>"" and .txtNat_cd_nm.value = "" then
                  Call  DisplayMsgBox("970000","X","국적코드","X")
                  Call ClickTab1()
                  .txtNat_cd_nm.value =""
                  .txtNat_cd.focus

                  exit function
            end if 
           if .txtNatv_state.value <> "" And .txtNatv_state_nm.value = "" then
                  Call  DisplayMsgBox("970000","X","출신도코드","X")
                  Call ClickTab1()
                  .txtNatv_state_nm.value =""
                  .txtNatv_state.focus
                  exit function
           end if 
           if .txtEntr_cd.value <> "" And .txtEntr_cd_nm.value = "" then
                  Call  DisplayMsgBox("970000","X","입사구분코드","X")
                  Call ClickTab1()
                  .txtEntr_cd_nm.value = ""
                  .txtEntr_cd.focus
                  exit function
           end if 
           if .txtapp_cd.value <> "" And .txtapp_cd_nm.value = "" then
                  Call  DisplayMsgBox("970000","X","채용구분코드","X")
                  Call ClickTab1()
                  .txtapp_cd_nm.value = ""
                  .txtapp_cd.focus
                  exit function
           end if 
           if .txtOcpt_type_nm.value = "" then
                  Call  DisplayMsgBox("970000","X","직종코드","X")
                  Call ClickTab1()
                  .txtOcpt_type_nm.value = ""
                  .txtOcpt_type.focus
                  Set gActiveElement = document.ActiveElement   
                  exit function
           end if 
           if .txthouse_cd.value <> "" And .txthouse_cd_nm.value = "" then
                  Call  DisplayMsgBox("970000","X","주거구분코드","X")
                  Call ClickTab1()
                  .txthouse_cd_nm.value = ""
                  .txthouse_cd.focus
                  exit function
           end if         
           if .txtMemo_cd.value <> "" And .txtMemo_cd_nm.value = "" then
                  Call  DisplayMsgBox("970000","X","기념일구분코드","X")
                  Call ClickTab1()
                  .txtMemo_cd_nm.value = ""
                  .txtMemo_cd.focus
                  exit function
           end if                                                                                                                                                                      
          if .txtDir_indir.value <> "" And .txtDir_indir_nm.value = "" then
                  Call  DisplayMsgBox("970000","X","직간접구분코드","X")
                  Call ClickTab1()
                  .txtDir_indir_nm.value = ""
                  .txtDir_indir.focus
                  exit function
           end if                 

           if .txtComp_cd.value <> "" And .txtComp_cd_nm.value = "" then
                  Call  DisplayMsgBox("970000","X","법인코드","X")
                  Call ClickTab2()
                  .txtComp_cd_nm.value = "" 
                  .txtComp_cd.focus
                  Set gActiveElement = document.ActiveElement   
                  exit function
           end if 
           if .txtSect_cd_nm.value = "" then
                  Call  DisplayMsgBox("970000","X","근무구역코드","X")
                  Call ClickTab2()
                  .txtSect_cd_nm.value = ""
                  .txtSect_cd.focus
                  exit function
           end if 
           if .txtDept_cd_nm.value = "" then
                  Call  DisplayMsgBox("970000","X","부서코드","X")
                  Call ClickTab2()
                  .txtDept_cd.focus
                  exit function
           end if 
           if .txtRoll_pstn_nm.value = "" then
                  Call  DisplayMsgBox("970000","X","직위코드","X")
                  Call ClickTab2()
					.txtRoll_pstn_nm.value = ""                  
                  .txtRoll_pstn.focus
                  exit function
           end if 
           if .txtRole_cd_nm.value = "" then
                  Call  DisplayMsgBox("970000","X","직책코드","X")
                  Call ClickTab2()
                  .txtRole_cd_nm.value = ""
                  .txtRole_cd.focus
                  exit function
           end if 
           if .txtFunc_cd_nm.value = "" then
                  Call  DisplayMsgBox("970000","X","직무코드","X")
                  Call ClickTab2()
                  .txtFunc_cd_nm.value = "" 
                  .txtFunc_cd.focus
                  exit function
           end if 
           if .txtPay_grd1_nm.value = "" then
                  Call  DisplayMsgBox("970000","X","급호코드","X")
                  Call ClickTab2()
                  .txtPay_grd1_nm.value = ""
                  .txtPay_grd1.focus
                  exit function
           end if 

           gSelframeFlg = TAB3
           
          if .txtMil_type.value <> "" And .txtMil_type_nm.value = "" then
                  Call  DisplayMsgBox("970000","X","병역구분코드","X")
                  Call ClickTab3()
                  .txtMil_type_nm.value = ""
                  .txtMil_type.focus
                  exit function
           end if                 
          if .txtMil_kind.value <> "" And .txtMil_kind_nm.value = "" then
                  Call  DisplayMsgBox("970000","X","병역군별코드","X")
                  Call ClickTab3()
                  .txtMil_kind_nm.value = "" 
                  .txtMil_kind.focus
                  exit function
           end if                 
          if .txtMil_grade.value <> "" And .txtMil_grade_nm.value = "" then
                  Call  DisplayMsgBox("970000","X","병역등급코드","X")
                  Call ClickTab3()
                  .txtMil_grade_nm.value = ""
                  .txtMil_grade.focus
                  exit function
           end if                 
          if .txtMil_branch.value <> "" And .txtMil_branch_nm.value = "" then
                  Call  DisplayMsgBox("970000","X","병역병과코드","X")
                  Call ClickTab3()
                  .txtMil_branch_nm.value = ""
                  .txtMil_branch.focus
                  exit function
           end if                 
          if .txtRelig_cd.value <> "" And .txtRelig_cd_nm.value = "" then
                  Call  DisplayMsgBox("970000","X","종교코드","X")
                  Call ClickTab3()
                  .txtRelig_cd_nm.value = "" 
                  .txtRelig_cd.focus
                  exit function
           end if       
'우편번호체크 
          If .txtZip_cd.value <> "" and trim(frm1.txtNat_cd.value)="KR"  then
              strWhere =                " ZIP_CD =  " & FilterVar(frm1.txtZip_cd.value , "''", "S") & ""
	          strWhere = strWhere & " AND COUNTRY_CD=  " & FilterVar(frm1.txtNat_cd.value , "''", "S") & ""

              if   CommonQueryRs(" COUNT(*) "," B_ZIP_CODE ",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true then
                  if  Replace(lgF0, Chr(11), "") <= 0 then
			          Call  DisplayMsgBox("800016","X","X","X")
                      Call ClickTab4()
	                  frm1.txtZip_cd.focus()
	                  exit function
	              end if
              end if
          End If
          
          If .txtCurr_zip_cd.value <> "" then
              strWhere =                " ZIP_CD =  " & FilterVar(frm1.txtCurr_zip_cd.value , "''", "S") & ""
'	          strWhere = strWhere & " AND COUNTRY_CD=  " & FilterVar(frm1.txtNat_cd.value , "''", "S") & ""

              if   CommonQueryRs(" COUNT(*) "," B_ZIP_CODE ",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true then
                  if  Replace(lgF0, Chr(11), "") <= 0 then
			          Call  DisplayMsgBox("800016","X","X","X")
                      Call ClickTab4()
	                  frm1.txtCurr_zip_cd.focus()
	                  exit function
	              end if
              end if
          End If
    end with

	gSelframeFlg = strTab

    strWhere = " PAY_GRD1 =  " & FilterVar(frm1.txtPay_grd1.value , "''", "S") & ""
	strWhere = strWhere & " AND PAY_GRD2 =  " & FilterVar(frm1.txtPay_grd2.value , "''", "S") & ""
    strWhere = strWhere & " AND APPLY_STRT_DT = (SELECT MAX(APPLY_STRT_DT) "
	strWhere = strWhere & "   FROM HDF010T "
    strWhere = strWhere & "  WHERE PAY_GRD1 =  " & FilterVar(frm1.txtPay_grd1.value , "''", "S") & ""
    strWhere = strWhere & "    AND PAY_GRD2 =  " & FilterVar(frm1.txtPay_grd2.value , "''", "S") & ""
    strWhere = strWhere & "    AND APPLY_STRT_DT <= GETDATE())"

    if   CommonQueryRs(" COUNT(*) "," HDF010T ",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true then
        if  Replace(lgF0, Chr(11), "") <= 0 then
			Call  DisplayMsgBox("800057","X","X","X")
	        call clicktab2()
	        frm1.txtPay_grd2.focus()
	        exit function
	    end if
	End If


'******************************************************************************************/
' 입사일 check !! 
'   입사일 <= 현재일 
'******************************************************************************************/
    
    IF  lsGetsvrDate <  UniConvDateAToB(frm1.txtEntr_dt.Text, gDateFormat, parent.gServerDateFormat) THEN   
        '입사일은 현재일보다 작거나 같아야합니다 
        ' // lsGetsvrDate는 GetsvrDate변수(Fixed Value : YYYY-MM-DD --> parent.gServerDateFormat)
        ' 그러므로 비교시에 YYYY-MM-DD로 만들어 비교..
		Call  DisplayMsgBox("800259", "x", "x", "x")
        Call ClickTab2()
        Frm1.txtEntr_dt.focus
        Set gActiveElement = document.ActiveElement   
        exit function
    END IF

'******************************************************************************************/
' 그룹입사일 check !!
' - 그룹입사일 <= 입사일 
'   그룹입사일 <= 수습만료일 
'******************************************************************************************/	
    IF  frm1.txtGroup_entr_dt.Text <> "" THEN
		If  ValidDateCheck(frm1.txtGroup_entr_dt, frm1.txtEntr_dt)=False Then
			Exit Function
		End if
    END IF


    IF  frm1.txtIntern_dt.Text <> "" THEN
		If  ValidDateCheck(frm1.txtGroup_entr_dt, frm1.txtIntern_dt)=False Then
			Exit Function
		End if
    END IF


'******************************************************************************************/	
' 수습만료일 check !!
' - 수습만료일 >= 입사일 
'******************************************************************************************/	
    
    IF  frm1.txtIntern_dt.Text <> "" THEN
		If  ValidDateCheck(frm1.txtEntr_dt, frm1.txtIntern_dt)=False Then
			Exit Function
		End if
    END IF
    

'******************************************************************************************/	
' 인사변동일 check !!
' - 인사변동일 >= 입사일 
'******************************************************************************************/	
   
    IF  frm1.txtOrder_change_dt.Text <> "" THEN
		If  ValidDateCheck(frm1.txtEntr_dt, frm1.txtOrder_change_dt)=False Then
			Exit Function
		End if
    END IF
    
    
'******************************************************************************************/	
' HelpDesk 결함번호 : 19980511006
' 최근승급일 check !!
' - 최근승급일 >= 그룹입사일(그룹입사일 있으면)
' - 최근승급일 >= 입사일    (그룹입사일 없으면)
'   입사일 <= 그룹입사일 <= 최근승급일 
'******************************************************************************************/	
    IF frm1.txtGroup_entr_dt.Text <> ""  THEN
        IF frm1.txtResent_promote_dt.Text <> ""  THEN

			If  ValidDateCheck(frm1.txtGroup_entr_dt, frm1.txtResent_promote_dt)=False Then
				Exit Function
			End if
        END IF
	ELSE
	    IF frm1.txtResent_promote_dt.Text <> ""  THEN

			If  ValidDateCheck(frm1.txtEntr_dt, frm1.txtResent_promote_dt)=False Then
				Exit Function
			End if
        END IF
	END IF

'***************************************************/
' 퇴직일 check !!
' - 퇴직일 >= 입사일 
'   퇴직일 not null -> 퇴직사유는 필수 
'   퇴직일 null     -> 퇴직사유는 null
'******************************************************************************************/

    IF  frm1.txtRetire_dt.Text <> "" THEN

		If  ValidDateCheck(frm1.txtEntr_dt, frm1.txtRetire_dt)=False Then
			Exit Function
		End if

		IF  frm1.txtRetire_Resn.value = "" THEN
            'MessageBox(This.Title, "퇴직자의 경우 퇴직사유는 입력항목입니다.", Exclamation!)
            Call  DisplayMsgBox("800255", "x", "x", "x")
            Call ClickTab2()
            Frm1.txtRetire_Resn.focus
            Set gActiveElement = document.ActiveElement   
            exit function
	   END IF
	   		
   END IF
	
   IF frm1.txtRetire_dt.Text = "" THEN
        IF frm1.txtRetire_Resn.value <> "" THEN
'	      MessageBox(This.Title, "재직자의 경우 퇴직사유는 입력할 수 없습니다.", Exclamation!)
            Call  DisplayMsgBox("800017", "x", "x", "x")
            Call ClickTab2()
            Frm1.txtRetire_Resn.focus
            Set gActiveElement = document.ActiveElement   
            exit function
        END IF
   END IF


'********* 복무기간 <= 현재일 -> OK !!
    IF lsGetsvrDate <  UniConvDateAToB(frm1.txtMil_start.Text, gDateFormat, parent.gServerDateFormat) or lsGetsvrDate <  UniConvDateAToB(frm1.txtMil_end.Text, gDateFormat, parent.gServerDateFormat) THEN
		'"복무기간은 현재일보다 작거나 같아야합니다."
        ' // lsGetsvrDate는 GetsvrDate변수(Fixed Value : YYYY-MM-DD --> parent.gServerDateFormat)
        ' 그러므로 비교시에 YYYY-MM-DD로 만들어 비교..
        Call  DisplayMsgBox("800010", "x", "x", "x")
        Call ClickTab3()
        Frm1.txtMil_start.focus
        Set gActiveElement = document.ActiveElement   
        exit function
	END IF
	

	If  ValidDateCheck(frm1.txtMil_start, frm1.txtMil_end)=False Then
		Exit Function
	End if

' 주민번호 Check **** Start
    If  UCase(frm1.txtNat_cd.value) = "KR" Then

        if  txtRes_no_Check() = true then

			res_no1 = Mid(Trim(Replace(frm1.txtRes_no.value,"-","")), 1, 6)

			res_no2 = Mid(Trim(Replace(frm1.txtRes_no.value,"-","")), 7, 7)
            ' 주민번호 Check **** End
            ' 남여구분 
            if  left(res_no2,1) = "1" OR left(res_no2,1) = "3" then
                if  frm1.txtSex_cd2.checked = true then ' 여 
	        	    Call  DisplayMsgBox("970027", "x", "남여구분", "x")
                    Call ClickTab1()
                    Frm1.txtSex_cd1.focus
                    Set gActiveElement = document.ActiveElement   
                    exit function
                end if
            else
                if  frm1.txtSex_cd1.checked = true then ' 여 
	        	    Call  DisplayMsgBox("970027", "x", "남여구분", "x")
                    Call ClickTab1()
                    Frm1.txtSex_cd1.focus
                    Set gActiveElement = document.ActiveElement   
                    exit function
                end if
            end if

            frm1.txtRes_no.value = res_no1 & res_no2
        else

            'msgbox "주민번호 이상입니다."
            IntRetCD =  DisplayMsgBox("800345",  parent.VB_YES_NO,"x","x")
            if  IntRetCD = VBNO then
                Call ClickTab1()
                Frm1.txtRes_no.focus
                Set gActiveElement = document.ActiveElement
                exit function
            else 
				frm1.txtRes_no.value= Trim(Replace(frm1.txtRes_no.value,"-",""))
            end if
        end if
    End If

    Call MakeKeyStream("S")
	Call  DisableToolBar( parent.TBC_SAVE)
    If DbSave = False Then
        Call  RestoreToolBar()
        Exit Function
    End If    

    FncSave = True                                                              '☜: Processing is OK
End Function


'========================================================================================================
' Name : FncCopy
' Desc : developer describe this line Called by MainSave in Common.vbs
' Keep : Make sure to clear primary key area
'========================================================================================================
Function FncCopy()
	Dim IntRetCD

    FncCopy = False                                                              '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
	
    If lgBlnFlgChgValue = True Then
		IntRetCD =  DisplayMsgBox("900017",  parent.VB_YES_NO,"x","x")				     '☜: Data is changed.  Do you want to continue? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    
    lgIntFlgMode =  parent.OPMD_CMODE												     '⊙: Indicates that current mode is Crate mode
    
    Call  ggoOper.ClearField(Document, "1")                                       '⊙: Clear Condition Field
    Call  ggoOper.LockField(Document, "N")									     '⊙: This function lock the suitable field
    Call SetToolbar("11101000000001")
 
    frm1.txtEmp_no.value = ""
    frm1.txtName.value = ""
    frm1.txtEng_name.value = ""

    Set gActiveElement = document.ActiveElement   
    FncCopy = True                                                            '☜: Processing is OK
    
End Function

'========================================================================================================
' Name : FncCancel
' Desc : developer describe this line Called by MainCancel in Common.vbs
'========================================================================================================
Function FncCancel() 
	On Error Resume Next                                                      '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncInsertRow
' Desc : developer describe this line Called by MainInsertRow in Common.vbs
'========================================================================================================
Function FncInsertRow()
	On Error Resume Next                                                      '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncDeleteRow
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncDeleteRow()
	On Error Resume Next                                                      '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncPrint
' Desc : developer describe this line Called by MainDeleteRow in Common.vbs
'========================================================================================================
Function FncPrint()
	Call Parent.FncPrint()                                                    '☜: Protect system from crashing
End Function

'========================================================================================================
' Name : FncPrev
' Desc : developer describe this line Called by MainPrev in Common.vbs
'========================================================================================================
Function FncPrev() 

    Dim strVal
    Dim IntRetCD

    FncPrev = False                                                              '☜: Processing is OK
    Err.Clear                                                                    '☜: Clear err status
    
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call  DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD =  DisplayMsgBox("900017",  parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    Call MakeKeyStream("P")
    Call  ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    
    Call InitVariables														 '⊙: Initializes local global variables

	If   LayerShowHide(1) = False Then
	     Exit Function
	End If

    strVal = BIZ_PGM_ID & "?txtMode="          &  parent.UID_M0001                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & "P"	                         '☜: Direction

	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz 

    FncPrev = True                                                               '☜: Processing is OK

End Function
'========================================================================================================
' Name : FncNext
' Desc : developer describe this line Called by MainNext in Common.vbs
'========================================================================================================
Function FncNext() 
    Dim strVal
    Dim IntRetCD

    FncNext = False                                                              '☜: Processing is OK
    Err.Clear                                                                    '☜: Clear err status
    
    If lgIntFlgMode <>  parent.OPMD_UMODE Then                                           '☜: Please do Display first. 
        Call  DisplayMsgBox("900002","x","x","x")
        Exit Function
    End If
	
	If lgBlnFlgChgValue = True Then
		IntRetCD =  DisplayMsgBox("900017",  parent.VB_YES_NO,"x","x")					 '☜: Will you destory previous data
		If IntRetCD = vbNo Then
			Exit Function
		End If
	End If
	
    Call MakeKeyStream("N")

    Call  ggoOper.ClearField(Document, "2")										 '⊙: Clear Contents Area
    
    Call InitVariables														     '⊙: Initializes local global variables

	If   LayerShowHide(1) = False Then
	     Exit Function
	End If


    strVal = BIZ_PGM_ID & "?txtMode="          &  parent.UID_M0001                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & "N"	                         '☜: Direction

	Call RunMyBizASP(MyBizASP, strVal)										     '☜: Run Biz 

    FncNext = True                                                               '☜: Processing is OK
	
End Function
'========================================================================================================
' Name : FncExcel
' Desc : developer describe this line Called by MainExcel in Common.vbs
'========================================================================================================
Function FncExcel() 
	Call Parent.FncExport( parent.C_SINGLE)
End Function

'========================================================================================================
' Name : FncFind
' Desc : developer describe this line Called by MainFind in Common.vbs
'========================================================================================================
Function FncFind() 
	Call Parent.FncFind( parent.C_SINGLE, True)
End Function

'========================================================================================================
' Name : FncExit
' Desc : developer describe this line Called by MainExit in Common.vbs
'========================================================================================================
Function FncExit()
	Dim IntRetCD

	FncExit = False
	If lgBlnFlgChgValue = True Then
		IntRetCD =  DisplayMsgBox("900016",  parent.VB_YES_NO,"x","x")			'⊙: Data is changed.  Do you want to exit? 
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
    Err.Clear                                                                    '☜: Clear err status

    DbQuery = False                                                              '☜: Processing is NG

	If   LayerShowHide(1) = False Then
	     Exit Function
	End If

    strVal = BIZ_PGM_ID & "?txtMode="          &  parent.UID_M0001                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & ""	                             '☜: Direction
    Call RunMyBizASP(MyBizASP, strVal)                                           '☜:  Run biz logic
    
    DbQuery = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()

    if  frm1.txtDept_cd.value  <> temp_txtDept_cd or frm1.txtRoll_pstn.value <> temp_txtRoll_pstn or frm1.txtFunc_cd.value <> temp_txtFunc_cd or frm1.txtPay_grd2.value <> temp_txtPay_grd2 or _
        frm1.txtRole_cd.value  <> temp_txtRole_cd or frm1.txtPay_grd1.value  <> temp_txtPay_grd1  or frm1.txtSect_cd.value <> temp_txtSect_cd then
		frm1.temp_flg_chk.value = "true"
	else 	
		frm1.temp_flg_chk.value = "false"
	end if	
	
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbSave = False														         '☜: Processing is NG
		
	If   LayerShowHide(1) = False Then
	     Exit Function
	End If

	With Frm1
		.txtMode.value        =  parent.UID_M0002                                        '☜: Delete
		.txtFlgMode.value     = lgIntFlgMode
        .txtKeyStream.Value   = lgKeyStream                                      '☜: Save Key
	End With

	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
		
    DbSave  = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
	Dim strVal
    Err.Clear                                                                    '☜: Clear err status
		
	DbDelete = False			                                                 '☜: Processing is NG
		
	If   LayerShowHide(1) = False Then
	     Exit Function
	End If
		
    strVal = BIZ_PGM_ID & "?txtMode="          &  parent.UID_M0003                       '☜: Query
    strVal = strVal     & "&txtKeyStream="     & lgKeyStream                     '☜: Query Key
    strVal = strVal     & "&txtPrevNext="      & ""	                             '☜: Direction

	Call RunMyBizASP(MyBizASP, strVal)                                           '☜: Run Biz logic
	
	DbDelete = True                                                              '⊙: Processing is NG
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
    Dim strVal

	lgIntFlgMode      =  parent.OPMD_UMODE                                               '⊙: Indicates that current mode is Create mode
    
    lgBlnFlgChgValue = false
    Frm1.txtName1.focus 

	Call SetToolbar("11111000111001")
	
	Call CommonQueryRs(" COUNT(*) "," HAA070T "," emp_no= " & FilterVar( Frm1.txtEmp_no.value, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    if   Replace(lgF0, Chr(11), "") > 0  then
		strVal = "../../ComASP/CPictRead.asp" & "?txtKeyValue=" & Frm1.txtEmp_no.value '☜: query key
		strVal = strVal     & "&txtDKeyValue=" & "default"                            '☜: default value
		strVal = strVal     & "&txtTable="     & "HAA070T"                            '☜: Table Name
		strVal = strVal     & "&txtField="     & "Photo"	                          '☜: Field
		strVal = strVal     & "&txtKey="       & "Emp_no"	                          '☜: Key
	else
		strVal = "../../../CShared/image/default_picture.jpg"
	end if

    Frm1.imgPhoto.src = strVal

	temp_txtDept_cd = frm1.txtDept_cd.value 
	temp_txtRoll_pstn = frm1.txtRoll_pstn.value 
	temp_txtFunc_cd =  frm1.txtFunc_cd.value
	temp_txtPay_grd2 = frm1.txtPay_grd2.value
	temp_txtRole_cd = frm1.txtRole_cd.value 
	temp_txtPay_grd1 = frm1.txtPay_grd1.value
	temp_txtSect_cd	 = frm1.txtSect_cd.value    

    Call  ggoOper.LockField(Document, "Q")
    Set gActiveElement = document.ActiveElement   
	If gSelframeFlg = TAB1 Then    
		frm1.txtName.focus
	elseIf gSelframeFlg = TAB2 Then    
		frm1.txtComp_cd.focus
	elseIf gSelframeFlg = TAB3 Then    
		frm1.txtMil_type.focus
	elseIf gSelframeFlg = TAB4 Then    
		frm1.txtDomi.focus
	end if
End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
    Call InitVariables
    Frm1.txtEmp_no1.value =  Frm1.txtEmp_no.value
    Frm1.txtName1.value =  Frm1.txtName.value    
    Call MainQuery()
End Function
	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()
	Call InitVariables()
	Call MainNew()	
	Frm1.imgPhoto.src = ""
End Function

'========================================================================================================
' Name : PgmJump1(PGM_JUMP_ID)
' Desc : developer describe this line 
'========================================================================================================

Function PgmJump1(PGM_JUMP_ID)
    Call BtnDisabled(1)
    Call CookiePage(1)  ' Write Cookie
    PgmJump(PGM_JUMP_ID)
    Call BtnDisabled(0)
End Function

'========================================================================================================
' Name : OpenEmp()
' Desc : developer describe this line 
'========================================================================================================
Function OpenEmp()
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = frm1.txtEmp_no1.value			' Code Condition
	arrParam(1) = ""'frm1.txtName1.value			' Name Cindition
    arrParam(2) = lgUsrIntCd

	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		frm1.txtEmp_no1.focus
		Exit Function
	Else
		Call SetEmp(arrRet)
	End If	
			
End Function

'======================================================================================================
'	Name : SetEmp()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetEmp(arrRet)
	With frm1
		.txtEmp_no1.value = arrRet(0)
		.txtName1.value = arrRet(1)
		Call  ggoOper.ClearField(Document, "2")					 '☜: Clear Contents  Field
		Set gActiveElement = document.ActiveElement
		.txtEmp_no1.focus
		lgBlnFlgChgValue = False
	End With
End Sub

'==========================================  2.3.1 Tab Click 처리  =================================================
'	기능: Tab Click
'	설명: Tab Click시 필요한 기능을 수행한다.
'===================================================================================================================

Function ClickTab1()
	If gSelframeFlg = TAB1 Then Exit Function
	
	Call changeTabs(TAB1)
	
	gSelframeFlg = TAB1
End Function

Function ClickTab2()
	If gSelframeFlg = TAB2 Then Exit Function
	
	Call changeTabs(TAB2)
	
	gSelframeFlg = TAB2
End Function

Function ClickTab3()
	If gSelframeFlg = TAB3 Then Exit Function
	
	Call changeTabs(TAB3)
	
	gSelframeFlg = TAB3
End Function

Function ClickTab4()
	If gSelframeFlg = TAB4 Then Exit Function
	
	Call changeTabs(TAB4)
	
	gSelframeFlg = TAB4
End Function
'========================================================================================================
' Name : SubOpenCollateralNoPop()
' Desc : developer describe this line Call Master L/C No PopUp
'========================================================================================================
Sub SubOpenCollateralNoPop()
	Dim strRet
	If gblnWinEvent = True Then Exit Sub
	gblnWinEvent = True
		
	strRet = window.showModalDialog("s1413pa1.asp", "", _
		"dialogWidth=780px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	gblnWinEvent = False
		
	If strRet = "" Then
       Exit Sub
	Else
       Call SetCollateralNo(strRet)
	End If	
End Sub

'===========================================================================
' Function Name : OpenSItemDC
' Function Desc : OpenSItemDC Reference Popup
'===========================================================================
Function OpenSItemDC(Byval strCode, Byval iWhere)

	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case 1  ' 급호 
	    	arrParam(1) = "B_minor"				            	' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtPay_grd1.Value)	        ' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0001", "''", "S") & ""		    		' Where Condition
	    	arrParam(5) = "급호"		    				    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)%>
    
	    	arrHeader(0) = "급호코드"			        		' Header명(0)%>
	    	arrHeader(1) = "급호명"	        					' Header명(1)%>

	    Case 2  ' 직위 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtRoll_pstn.Value)			' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0002", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "직위"    						    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "직위코드"			        		' Header명(0)
	    	arrHeader(1) = "직위명"	        					' Header명(1)

	    Case 3  ' 직종 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtOcpt_type.Value)			' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0003", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "직종"    						    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "직종코드"			        		' Header명(0)
	    	arrHeader(1) = "직종명"	        					' Header명(1)

	    Case 4  ' 직무 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtFunc_cd.Value)			' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0004", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "직무"    						    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "직무코드"			        		' Header명(0)
	    	arrHeader(1) = "직무명"	        					' Header명(1)

	    Case 5  ' 직책 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtRole_cd.Value)			' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0026", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "직책"    						    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "직책코드"			        		' Header명(0)
	    	arrHeader(1) = "직책명"	        					' Header명(1)

	    Case 13     ' 장애구분 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtParia_cd.Value)			' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0013", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "장애구분"    					    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "장애구분코드"		        		' Header명(0)
	    	arrHeader(1) = "장애구분"	       					' Header명(1)
	    	
	    Case 14     ' 보훈구분 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtRelief_cd.Value)			' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0014", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "보훈구분"    					    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "보훈구분코드"		        		' Header명(0)
	    	arrHeader(1) = "보훈구분"	       					' Header명(1)

	    Case 15     ' 주거구분 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtHouse_cd.Value)			' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0015", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "주거구분"    					    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "주거구분코드"		        		' Header명(0)
	    	arrHeader(1) = "주거구분"	       					' Header명(1)

	    Case 16     ' 입사구분 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtEntr_cd.Value)			' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0016", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "입사구분"    					    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "입사구분코드"		        		' Header명(0)
	    	arrHeader(1) = "입사구분"	       					' Header명(1)

	    Case 17     ' 채용구분 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtApp_cd.Value)			' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0017", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "채용구분"    					    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "채용구분코드"		        		' Header명(0)
	    	arrHeader(1) = "채용구분"	       					' Header명(1)

	    Case 18  ' 종교구분 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtRelig_cd.Value)			' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0018", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "종교구분"    					    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "종교구분코드"		        		' Header명(0)
	    	arrHeader(1) = "종교구분"             				' Header명(1)

	    Case 19  ' 병역구분 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtMil_type.Value)			' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0019", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "병역구분"    					    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "병역구분코드"		        		' Header명(0)
	    	arrHeader(1) = "병역명"             				' Header명(1)

	    Case 20  ' 병역군별 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtMil_kind.Value)			' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0020", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "병역군별"    					    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "병역군별코드"		        		' Header명(0)
	    	arrHeader(1) = "병역군별"             				' Header명(1)

	    Case 21  ' 병역계급 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtMil_grade.Value)			' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0021", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "병역계급"    					    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "병역계급코드"		        		' Header명(0)
	    	arrHeader(1) = "병역계급명"          				' Header명(1)
	    Case 22  ' 병역병과 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtMil_branch.Value)		' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0022", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "병역병과"    					    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "병역병과코드"	        			' Header명(0)
	    	arrHeader(1) = "병역병과명"          				' Header명(1)

	    Case 27  ' 출신도 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtNatv_state.Value)		' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0027", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "출신도"    	    				' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "지역코드"		        		' Header명(0)
	    	arrHeader(1) = "지역명"          				' Header명(1)

	    Case 28  ' 기념구분 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtMemo_cd.Value)		    ' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0028", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "기념구분"    	    			' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "기념구분코드"		        	' Header명(0)
	    	arrHeader(1) = "기념구분"          				' Header명(1)

	    Case 35  ' 근무구역 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtSect_cd.Value)			' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0035", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "근무구역"    	    			    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "근무구역코드"	        			' Header명(0)
	    	arrHeader(1) = "근무구역명"          				' Header명(1)
	    Case 36  ' 근무지 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtWk_area_cd.Value)		' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0036", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "근무지"    	    				    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "근무지코드" 	        			' Header명(0)
	    	arrHeader(1) = "근무지명"          					' Header명(1)

	    Case 71  ' 직간접구분 
	    	arrParam(1) = "B_minor"							    ' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtDir_indir.Value)		' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "major_cd=" & FilterVar("H0071", "''", "S") & ""			    	' Where Condition
	    	arrParam(5) = "직간접구분"    	    				    ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"							' Field명(0)
	    	arrField(1) = "minor_nm"    						' Field명(1)
    
	    	arrHeader(0) = "직간접구분코드" 	        			' Header명(0)
	    	arrHeader(1) = "직간접구분명"          					' Header명(1)

        Case 102  ' 국적코드 
            arrParam(1) = "B_COUNTRY"		    			    ' TABLE 명칭 
            arrParam(2) = Trim(frm1.txtNat_cd.Value)            ' Code Condition
            arrParam(3) = ""                                    ' Name Cindition
            arrParam(4) = ""                				    ' Where Condition
            arrParam(5) = "국가코드"	    				    ' TextBox 명칭 
	
            arrField(0) = "country_cd"	    				    ' Field명(0)
            arrField(1) = "country_nm"                          ' Field명(1)
    
            arrHeader(0) = "국가코드"                           ' Header명(0)
            arrHeader(1) = "국가명"                             ' Header명(1)
        Case 103  ' 법인코드 
	        arrParam(1) = "B_COMPANY"						' TABLE 명칭 
	        arrParam(2) = Trim(frm1.txtComp_cd.Value)							' Code Condition
	        arrParam(3) = ""								' Name Cindition
	        arrParam(4) = ""								' Where Condition
	        arrParam(5) = "법인"
	
            arrField(0) = "CO_CD"							' Field명(0)
            arrField(1) = "CO_FULL_NM"						' Field명(1)
    
            arrHeader(0) = "법인코드"					' Header명(0)
            arrHeader(1) = "법인명"					' Header명(1)
	    Case 105  ' 호봉 
			If frm1.txtPay_grd1.Value ="" Then
				Call  DisplayMsgBox("800489", "x","급호","x")
				IsOpenPop = False
				Exit Function
			End If
			
			Call CommonQueryRs(" top 1 dbo.ufn_H_GetCodeName( 'HDA010T', allow1_cd,''),dbo.ufn_H_GetCodeName( 'HDA010T', allow2_cd,''),dbo.ufn_H_GetCodeName( 'HDA010T', allow3_cd,'') "," hdf010t "," 1=1" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
			
	    	arrParam(1) = "hdf010t"				            	' TABLE 명칭 
	    	arrParam(2) = Trim(frm1.txtPay_grd2.Value)	        ' Code Condition
	    	arrParam(3) = ""									' Name Cindition
	    	arrParam(4) = "apply_strt_dt = (select max(apply_strt_dt) from hdf010t ) and  pay_grd1 = " & frm1.txtPay_grd1.value    		' Where Condition
	    	arrParam(5) = "호봉"		    				    ' TextBox 명칭 
	
	    	arrField(0) = "ED7" & Parent.gColSep &"pay_grd2"							' Field명(0)
	    	arrField(1) = "F212" & Parent.gColSep &"allow1"    							' Field명(1)%>
	    	arrField(2) = "F212" & Parent.gColSep &"allow2"    							' Field명(1)%>
	    	arrField(3) = "F212" & Parent.gColSep &"allow3"    							' Field명(1)%>
    
    
	    	arrHeader(0) = "호봉"			        		' Header명(0)%>
            arrHeader(1) = Replace(lgF0, Chr(11), "")               ' Header명(1)
            arrHeader(2) = Replace(lgF1, Chr(11), "")               ' Header명(2)
            arrHeader(3) = Replace(lgF2, Chr(11), "")               ' Header명(3)            
	End Select

    arrParam(3) = ""	
	arrParam(0) = arrParam(5)								    ' 팝업 명칭 

	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")

	IsOpenPop = False

	If arrRet(0) = "" Then
			With frm1

				Select Case iWhere
				Case 0
					.txtEmp_no.focus
				Case 1
					.txtPay_grd1.focus
				Case 2
					.txtRoll_pstn.focus
				Case 3  ' 직종 
					.txtOcpt_type.focus
				Case 4
					.txtFunc_cd.focus
				Case 5
					.txtRole_cd.focus
				Case 13
					.txtparia_cd.focus
				Case 14
					.txtrelief_cd.focus
				Case 15
					.txthouse_cd.focus
				Case 16
					.txtEntr_cd.focus
				Case 17
					.txtapp_cd.focus
				Case 18
					.txtrelig_cd.focus
				Case 19 '병역구분 
					.txtMil_type.focus
				Case 20
					.txtmil_kind.focus
				Case 21 '병역계급 
					.txtMil_grade.focus
				Case 22 '병역병과 
					.txtMil_Branch.focus
				Case 27 '출신도 
					.txtNatv_state.focus
				Case 28
					.txtmemo_cd.focus
				Case 35 '근무구역코드 
					.txtSect_cd.focus
				Case 36 '근무지코드 
					.txtWk_area_cd.focus
				Case 71 '직간접구분 
					.txtDir_indir.focus
				Case 102 '국적코드 
					.txtNat_cd.focus
				Case 103 '법인코드 
					.txtComp_cd.focus
				Case 105 '호봉 
					.txtPay_grd2.focus					
				End Select
			End With
		Exit Function
	Else
		Call SetSItemDC(arrRet, iWhere)
	End If	
	
End Function

'------------------------------------------  SetSItemDC()  --------------------------------------------------
'	Name : SetSItemDC()
'	Description : OpenSItemDC Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetSItemDC(Byval arrRet, Byval iWhere)

	With frm1

		Select Case iWhere
		Case 0
			.txtName.value = arrRet(0) 
			.txtEmp_no.value = arrRet(1)   
			.txtEmp_no.focus
		Case 1
			.txtPay_grd1.value = arrRet(0)
			.txtPay_grd1_nm.value = arrRet(1)  
			.txtPay_grd1.focus
		Case 2
			.txtRoll_pstn.value = arrRet(0)
			.txtRoll_pstn_nm.value = arrRet(1)  
			.txtRoll_pstn.focus
		Case 3  ' 직종 
			.txtOcpt_type.value = arrRet(0)
			.txtOcpt_type_nm.value = arrRet(1)  
			.txtOcpt_type.focus
		Case 4
			.txtFunc_cd.value = arrRet(0) 
			.txtFunc_cd_nm.value = arrRet(1)   
			.txtFunc_cd.focus
		Case 5
			.txtRole_cd.value = arrRet(0) 
			.txtRole_cd_nm.value = arrRet(1)   
			.txtRole_cd.focus
		Case 13
			.txtparia_cd.value = arrRet(0) 
			.txtparia_cd_nm.value = arrRet(1)   
			.txtparia_cd.focus
		Case 14
			.txtrelief_cd.value = arrRet(0) 
			.txtrelief_cd_nm.value = arrRet(1)   
			.txtrelief_cd.focus
		Case 15
			.txthouse_cd.value = arrRet(0) 
			.txthouse_cd_nm.value = arrRet(1)   
			.txthouse_cd.focus
		Case 16
			.txtEntr_cd.value = arrRet(0) 
			.txtEntr_cd_nm.value = arrRet(1)   
			.txtEntr_cd.focus
		Case 17
			.txtapp_cd.value = arrRet(0) 
			.txtapp_cd_nm.value = arrRet(1)   
			.txtapp_cd.focus
		Case 18
			.txtrelig_cd.value = arrRet(0) 
			.txtrelig_cd_nm.value = arrRet(1)   
			.txtrelig_cd.focus
		Case 19 '병역구분 
			.txtMil_type.value = arrRet(0) 
			.txtMil_type_nm.value = arrRet(1)   
			.txtMil_type.focus
		Case 20
			.txtmil_kind.value = arrRet(0) 
			.txtmil_kind_nm.value = arrRet(1)   
			.txtmil_kind.focus
		Case 21 '병역계급 
			.txtMil_grade.value = arrRet(0) 
			.txtMil_grade_nm.value = arrRet(1)   
			.txtMil_grade.focus
		Case 22 '병역병과 
			.txtMil_Branch.value = arrRet(0) 
			.txtMil_Branch_nm.value = arrRet(1)   
			.txtMil_Branch.focus
		Case 27 '출신도 
			.txtNatv_state.value = arrRet(0) 
			.txtNatv_state_nm.value = arrRet(1)   
			.txtNatv_state.focus
		Case 28
			.txtmemo_cd.value = arrRet(0)
			.txtmemo_cd_nm.value = arrRet(1)   
			.txtmemo_cd.focus
		Case 35 '근무구역코드 
			.txtSect_cd.value = arrRet(0) 
			.txtSect_cd_nm.value = arrRet(1)   
			.txtSect_cd.focus
		Case 36 '근무지코드 
			.txtWk_area_cd.value = arrRet(0) 
			.txtWk_area_cd_nm.value = arrRet(1)   
			.txtWk_area_cd.focus
		Case 71 '직간접구분 
			.txtDir_indir.value = arrRet(0) 
			.txtDir_indir_nm.value = arrRet(1)   
			.txtDir_indir.focus
		Case 102 '국적코드 
			.txtNat_cd.value = arrRet(0) 
			.txtNat_cd_nm.value = arrRet(1)   
			.txtNat_cd.focus
		Case 103 '법인코드 
			.txtComp_cd.value = arrRet(0) 
			.txtComp_cd_nm.value = arrRet(1)   
			.txtComp_cd.focus
		Case 105 '호봉 
			.txtPay_grd2.value = arrRet(0) 
			.txtPay_grd2.focus			
		End Select

		lgBlnFlgChgValue = True

	End With
	
End Function

'========================================================================================================
' Name : OpenEmpName()
' Desc : developer describe this line 
'========================================================================================================
Function OpenEmpName(iWhere)
	Dim arrRet
	Dim arrParam(1)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True
    If  iWhere = 0 Then
	    arrParam(0) = ""			' Code Condition
	    arrParam(1) = frm1.txtName1.value			' Name Cindition
    Else
	    arrParam(0) = frm1.txtEmp_no1.value			' Code Condition
	    arrParam(1) = ""			' Name Cindition
	End If

	arrRet = window.showModalDialog(HRAskPRAspName("EmpPopup"), Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetEmpName(arrRet)
	End If	
			
End Function

'======================================================================================================
'	Name : SetEmp()
'	Description : Item Popup에서 Return되는 값 setting
'=======================================================================================================
Sub SetEmpName(arrRet)
	With frm1
		.txtEmp_no1.value = arrRet(0)
		.txtName1.value = arrRet(1)
		Call  ggoOper.ClearField(Document, "2")					 '☜: Clear Contents  Field
		Set gActiveElement = document.ActiveElement

		lgBlnFlgChgValue = False
	End With
End Sub

'========================================================================================================
' Name : SetCurrency
' Desc : developer describe this line 
'========================================================================================================
Function SetCurrency(arrRet)
	frm1.txtCurrency.Value = arrRet(0)
	lgBlnFlgChgValue = True
End Function


'========================================================================================================
' Name : OpenDept
' Desc : 부서 POPUP
'========================================================================================================
Function OpenDept(iWhere)
	Dim arrRet
	Dim arrParam(2)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	If iWhere = 0 Then 'TextBox(Condition)
		arrParam(0) = frm1.txtDept_cd.value			' 조건부에서 누른 경우 Code Condition
	Else 'spread
		arrParam(0) = frm1.vspdData.Text			' Grid에서 누른 경우 Code Condition
	End If
	arrParam(1) = ""
	arrParam(2) = lgUsrIntCd

	arrRet = window.showModalDialog(HRAskPRAspName("DeptPopupDt"), Array(window.parent,arrParam), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		If iWhere = 0 Then 'TextBox(Condition)
			frm1.txtDept_cd.focus
		Else 'spread
			frm1.vspdData.Col = C_Dept
			frm1.vspdData.action =0				
		End If
		Exit Function
	Else
		Call SetDept(arrRet, iWhere)
	End If	
			
End Function

'------------------------------------------  SetDept()  ------------------------------------------------
'	Name : SetDept()
'	Description : Dept Popup에서 Return되는 값 setting
'---------------------------------------------------------------------------------------------------------
Function SetDept(Byval arrRet, Byval iWhere)
		
	With frm1
		If iWhere = 0 Then 'TextBox(Condition)
			.txtDept_cd.value = arrRet(0)
			.txtDept_cd_Nm.value = arrRet(1)
			lgBlnFlgChgValue = True
			.txtDept_cd.focus
		Else 'spread
			.vspdData.Col = C_DeptNm
			.vspdData.Text = arrRet(1)
			.vspdData.Col = C_Dept
			.vspdData.Text = arrRet(0)
			.vspdData.action =0				
		End If
	End With
End Function

'------------------------------------------  OpenZipCode()  ------------------------------------------------
'	Name : OpenZipCode()
'	Description : 
'---------------------------------------------------------------------------------------------------------
Function OpenZipCode(ByVal strCode, ByVal iWhere)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	arrParam(0) = strCode
	arrParam(1) = ""

	If Trim(frm1.txtNat_cd.value) = "" Then
		arrParam(2) = "KR"
	Else 
		If iWhere = 0 Then	
			arrParam(2) = Trim(frm1.txtNat_cd.value)
		ElseIf iWhere = 1 Then
			arrParam(2) = "KR"
		End If
	End If	

	arrRet = window.showModalDialog("../../comasp/ZipPopup.asp", Array(window.parent,arrParam), _
		"dialogWidth=760px; dialogHeight=420px; center: Yes; help: No; resizable: No; status: No;")
		
	IsOpenPop = False

	If arrRet(0) = "" Then
		If iWhere = 0 Then	
			frm1.txtZip_cd.focus
		ElseIf iWhere = 1 Then
			frm1.txtCurr_zip_cd.focus
		End If
	
	    Exit Function
	Else
		Call SetCurrencyInfo(arrRet,iWhere)
	End If	

End Function

'------------------------------------------  SetCurrencyInfo()  -----------------------------------------------
'	Name : SetCurrency()
'	Description : Popup에서 Return되는 값 setting
'--------------------------------------------------------------------------------------------------------- 
Function SetCurrencyInfo(arrRet, iWhere)'

	With frm1
		If iWhere = 0 Then
			.txtZip_cd.value = arrRet(0)
			.txtAddr.value   = arrRet(1)
			.txtZip_cd.focus
		ElseIf iWhere = 1 Then
			.txtCurr_zip_cd.value = arrRet(0)
			.txtCurr_addr.value   = arrRet(1)
			.txtCurr_zip_cd.focus
		End If
		lgBlnFlgChgValue = True
	End With

End Function

'========================================================================================================
'   Event Name : txtEmp_no1_Onchange             
'   Event Desc :
'========================================================================================================
Function txtEmp_no1_Onchange()
    Dim IntRetCd
    Dim strName
    Dim strDept_nm
    Dim strRoll_pstn
    Dim strPay_grd1
    Dim strPay_grd2
    Dim strEntr_dt
    Dim strInternal_cd
    Dim strVal

    frm1.txtName1.value = ""
    txtEmp_no1_Onchange = true
    Frm1.imgPhoto.src = ""

    If  frm1.txtEmp_no1.value = "" Then
    Else
	    IntRetCd =  FuncGetEmpInf2(frm1.txtEmp_no1.value,lgUsrIntCd,strName,strDept_nm,_
	                strRoll_pstn, strPay_grd1, strPay_grd2, strEntr_dt, strInternal_cd)
	    if  IntRetCd < 0 then
	    
	    	strVal = "../../../CShared/image/default_picture.jpg"
		    Frm1.imgPhoto.src = strVal  
		    
	        if  IntRetCd = -1 then
    			Call  DisplayMsgBox("800048","X","X","X")	'해당사원은 존재하지 않습니다.
            else
                Call  DisplayMsgBox("800454","X","X","X")	'자료에 대한 권한이 없습니다.
            end if
            
			frm1.txtName1.value = ""
            frm1.txtEmp_no1.focus
	        exit function            
        Else
			frm1.txtName1.value = strName

			Call CommonQueryRs(" COUNT(*) "," HAA070T "," emp_no= " & FilterVar( Frm1.txtEmp_no.value, "''", "S") & "" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

			if   Replace(lgF0, Chr(11), "") > 0  then
				strVal = "../../ComASP/CPictRead.asp" & "?txtKeyValue=" & Frm1.txtEmp_no.value '☜: query key
				strVal = strVal     & "&txtDKeyValue=" & "default"                            '☜: default value
				strVal = strVal     & "&txtTable="     & "HAA070T"                            '☜: Table Name
				strVal = strVal     & "&txtField="     & "Photo"	                          '☜: Field
				strVal = strVal     & "&txtKey="       & "Emp_no"	                          '☜: Key
			else
				strVal = "../../../CShared/image/default_picture.jpg"
			end if
			
            Frm1.imgPhoto.src = strVal
        End if 

    End if  
    
End Function

Sub txtSex_cd1_OnClick()
    lgBlnFlgChgValue = True
End Sub

Sub txtSex_cd2_OnClick()
    lgBlnFlgChgValue = True
End Sub

Sub txtso_lu_cd1_OnClick()
    lgBlnFlgChgValue = True
End Sub

Sub txtso_lu_cd2_OnClick()
    lgBlnFlgChgValue = True
End Sub

Function txtNat_cd_OnChange()
    txtNat_cd_OnChange = true
    If  frm1.txtNat_cd.value <> "" Then
        if   CommonQueryRs(" country_nm "," B_COUNTRY "," country_cd =  " & FilterVar(frm1.txtNat_cd.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtNat_cd_nm.value = ""
            Call  DisplayMsgBox("970000", "x","국적코드","x")
	        frm1.txtNat_cd.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtNat_cd_nm.value = Replace(lgF0, Chr(11), "")
	    End If
	else 
		 frm1.txtNat_cd_nm.value=""
    End If

End Function

Function txtNatv_state_OnChange()
    txtNatv_state_OnChange = true
    
    If  frm1.txtNatv_state.value = "" Then
        frm1.txtNatv_state_nm.value = ""
        frm1.txtNatv_state.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0027", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtNatv_state.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            Call  DisplayMsgBox("970000", "x","출신도코드","x")

            frm1.txtNatv_state_nm.value = ""
	        frm1.txtNatv_state.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtNatv_state_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If

End Function

Function txtComp_cd_OnChange()
    txtComp_cd_OnChange = true

    If  frm1.txtComp_cd.value = "" Then
        frm1.txtComp_cd_nm.value = ""
        frm1.txtComp_cd.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" CO_FULL_NM "," B_COMPANY "," CO_CD =  " & FilterVar(frm1.txtComp_cd.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtComp_cd_nm.value = ""
            Call  DisplayMsgBox("970000", "x","법인코드","x")
	        frm1.txtComp_cd.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtComp_cd_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If
	    
End Function

Function txtFunc_cd_OnChange()

    txtFunc_cd_OnChange = true
    
    If  frm1.txtFunc_cd.value = "" Then
        frm1.txtFunc_cd_nm.value = ""
        frm1.txtFunc_cd.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0004", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtFunc_cd.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtFunc_cd_nm.value = ""
            Call  DisplayMsgBox("970000", "x","직무코드","x")
	        frm1.txtFunc_cd.focus

	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtFunc_cd_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If
	    
End Function

Function txtPay_grd1_OnChange()

    txtPay_grd1_OnChange = true

    If  frm1.txtPay_grd1.value = "" Then
        frm1.txtPay_grd1_nm.value = ""
        frm1.txtPay_grd1.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0001", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtPay_grd1.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtPay_grd1_nm.value = ""
            Call  DisplayMsgBox("970000", "x","급호코드","x")
	        frm1.txtPay_grd1.focus
	        Set gActiveElement = document.ActiveElement
	    Else
	        frm1.txtPay_grd1_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If
	    
End Function

Function txtPay_grd2_OnChange()

    Dim strWhere

    txtPay_grd2_OnChange = true
    strWhere = " PAY_GRD1 =  " & FilterVar(frm1.txtPay_grd1.value , "''", "S") & ""
	strWhere = strWhere & " AND PAY_GRD2 =  " & FilterVar(frm1.txtPay_grd2.value , "''", "S") & ""
    strWhere = strWhere & " AND APPLY_STRT_DT = (SELECT MAX(APPLY_STRT_DT) "
	strWhere = strWhere & "   FROM HDF010T "
    strWhere = strWhere & "  WHERE PAY_GRD1 =  " & FilterVar(frm1.txtPay_grd1.value , "''", "S") & ""
    strWhere = strWhere & "    AND PAY_GRD2 =  " & FilterVar(frm1.txtPay_grd2.value , "''", "S") & ""
    strWhere = strWhere & "    AND APPLY_STRT_DT <= GETDATE())"

    if   CommonQueryRs(" COUNT(*) "," HDF010T ",strWhere,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = true then
        if  Replace(lgF0, Chr(11), "") <= 0 then
			Call  DisplayMsgBox("800057","X","X","X")

	        frm1.txtPay_grd2.focus()
	        Set gActiveElement = document.ActiveElement
	        exit function
	    end if
	End If

	    
End Function

Function txtSect_cd_OnChange()
    txtSect_cd_OnChange = true

    If  frm1.txtSect_cd.value = "" Then
        frm1.txtSect_cd_nm.value = ""
        frm1.txtSect_cd.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0035", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtSect_cd.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtSect_cd_nm.value = ""
            Call  DisplayMsgBox("970000", "x","근무구역코드","x")
	        frm1.txtSect_cd.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtSect_cd_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If
	    
End Function

Function txtWk_area_cd_OnChange()
    txtWk_area_cd_OnChange = true
    
    If  frm1.txtWk_area_cd.value = "" Then
        frm1.txtWk_area_cd_nm.value = ""
        frm1.txtWk_area_cd.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0036", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtWk_area_cd.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtWk_area_cd_nm.value = ""
            Call  DisplayMsgBox("970000", "x","근무지코드","x")
	        frm1.txtWk_area_cd.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtWk_area_cd_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If
	    
End Function

Function txtRoll_pstn_OnChange()
    txtRoll_pstn_OnChange = true
    
    If  frm1.txtRoll_pstn.value = "" Then
        frm1.txtRoll_pstn_nm.value = ""
        frm1.txtRoll_pstn.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0002", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtRoll_pstn.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtRoll_pstn_nm.value = ""
            Call  DisplayMsgBox("970000", "x","직위코드","x")
	        frm1.txtRoll_pstn.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtRoll_pstn_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If
	    
End Function

Function txtRole_cd_OnChange()
    txtRole_cd_OnChange = true
    
    If  frm1.txtRole_cd.value = "" Then
        frm1.txtRole_cd_nm.value = ""
        frm1.txtRole_cd.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0026", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtRole_cd.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtRole_cd_nm.value = ""
            Call  DisplayMsgBox("970000", "x","직책코드","x")
	        frm1.txtRole_cd.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtRole_cd_nm.value = Replace(lgF0, Chr(11), "")
        	lgBlnFlgChgValue = True
	    End If
    End If
End Function

Function txtMil_type_OnChange()
    txtMil_type_OnChange = true
    
    If  frm1.txtMil_type.value = "" Then
        frm1.txtMil_type_nm.value = ""
        frm1.txtMil_type.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0019", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtMil_type.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtMil_type_nm.value = ""
            Call  DisplayMsgBox("970000", "x","병역구분코드","x")
	        frm1.txtMil_type.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtMil_type_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If
	    
End Function

Function txtMil_kind_OnChange()
    txtMil_kind_OnChange = true
    
    If  frm1.txtMil_kind.value = "" Then
        frm1.txtMil_kind_nm.value = ""
        frm1.txtMil_kind.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0020", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtMil_kind.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtMil_kind_nm.value = ""
            Call  DisplayMsgBox("970000", "x","병역군별코드","x")
	        frm1.txtMil_kind.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtMil_kind_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If
	    
End Function

Function txtMil_grade_OnChange()
    txtMil_grade_OnChange = true

    If  frm1.txtMil_grade.value = "" Then
        frm1.txtMil_grade_nm.value = ""
        frm1.txtMil_grade.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0021", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtMil_grade.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtMil_grade_nm.value = ""
            Call  DisplayMsgBox("970000", "x","병역등급코드","x")
	        frm1.txtMil_grade.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtMil_grade_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If
	    
End Function


Sub txtrelief_cd_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtparia_cd_OnChange()
	lgBlnFlgChgValue = True
End Sub

Function txthouse_cd_OnChange()
    txthouse_cd_OnChange = true
    
    If  frm1.txthouse_cd.value = "" Then
        frm1.txthouse_cd_nm.value = ""
        frm1.txthouse_cd.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0015", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txthouse_cd.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txthouse_cd_nm.value = ""
            Call  DisplayMsgBox("970000", "x","주거구분코드","x")
	        frm1.txthouse_cd.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txthouse_cd_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If
	    
End Function

Function txtapp_cd_OnChange()
    txtapp_cd_OnChange = true

    If  frm1.txtapp_cd.value = "" Then
        frm1.txtapp_cd_nm.value = ""
        frm1.txtapp_cd.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0017", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtapp_cd.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtapp_cd_nm.value = ""
            Call  DisplayMsgBox("970000", "x","채용구분코드","x")
	        frm1.txtapp_cd.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtapp_cd_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If
	    
End Function

Function txtmil_kind_OnChange()
    txtmil_kind_OnChange = true

    If  frm1.txtmil_kind.value = "" Then
        frm1.txtmil_kind_nm.value = ""
        frm1.txtmil_kind.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0020", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtmil_kind.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtmil_kind_nm.value = ""
            Call  DisplayMsgBox("970000", "x","병역군별코드","x")
	        frm1.txtmil_kind.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtmil_kind_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If
	    
End Function

Function txtrelig_cd_OnChange()
    txtrelig_cd_OnChange = true

    If  frm1.txtrelig_cd.value = "" Then
        frm1.txtrelig_cd_nm.value = ""
        frm1.txtrelig_cd.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0018", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtrelig_cd.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtrelig_cd_nm.value = ""
            Call  DisplayMsgBox("970000", "x","종교구분코드","x")
	        frm1.txtrelig_cd.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtrelig_cd_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If
	    
End Function

Function txtmemo_cd_OnChange()
    txtmemo_cd_OnChange = true

    If  frm1.txtmemo_cd.value = "" Then
        frm1.txtmemo_cd_nm.value = ""
        frm1.txtmemo_cd.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0028", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtmemo_cd.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtmemo_cd_nm.value = ""
            Call  DisplayMsgBox("970000", "x","기념구분코드","x")
	        frm1.txtmemo_cd.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtmemo_cd_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If
	    
End Function


Function txtMil_branch_OnChange()
    txtMil_branch_OnChange = true

    If  frm1.txtMil_branch.value = "" Then
        frm1.txtMil_branch_nm.value = ""
        frm1.txtMil_branch.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0022", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtMil_branch.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtMil_branch_nm.value = ""
            Call  DisplayMsgBox("970000", "x","병역병과코드","x")
	        frm1.txtMil_branch.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtMil_branch_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If
	    
End Function

Function txtDept_cd_OnChange()
    Dim IntRetCd
    Dim strDept_nm
    Dim strInternal_cd
    
    txtDept_cd_OnChange = true

    If RTrim(frm1.txtDept_cd.value) = "" Then
        frm1.txtDept_cd_nm.value = ""
        frm1.txtDept_cd.focus()
    Else
        IntRetCd =  FuncDeptName(frm1.txtDept_cd.value,"",lgUsrIntCd,strDept_nm,strInternal_cd)
        if  IntRetCd < 0 then
            if  IntRetCd = -1 then
                Call  DisplayMsgBox("970000", "x","부서코드","x")
            else
                Call  DisplayMsgBox("800455", "x","x","x")   ' 자료권한이 없습니다.
            end if
            frm1.txtDept_cd_nm.value = ""
            frm1.txtDept_cd.focus()
            Set gActiveElement = document.ActiveElement
            exit function
        else
            frm1.txtDept_cd_nm.value = strDept_nm
        end if
    End if

End Function


Sub txtBlood_type1_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtBlood_type2_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtGroup_entr_dt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtEntr_dt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtIntern_dt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtCareer_mm_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtResent_promote_dt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtMarry_cd_OnChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtOrder_change_dt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtRetire_dt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtRest_month_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtTech_man_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub txtMil_start_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtMil_end_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtHgt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtWgt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtDalt_type_OnClick()
	lgBlnFlgChgValue = True
End Sub

Sub txtEyesgt_left_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtEyesgt_right_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtYear_area_cd_OnChange()
	lgBlnFlgChgValue = True
End Sub

'========================================================================================================
' Name : SELECT_OnChange
' Desc : developer describe this line 
'========================================================================================================
Function txtRes_no_Check()
    Dim IntRetCD
    Dim res_no1, res_no2 , tmp_no            ' 주민번호 
    Dim intChk, intMod, intDef      ' 주민번호 
    txtRes_no_Check = true
        ' 주민번호 Check **** Start
    If  UCase(frm1.txtNat_cd.value) = "KR" Then
        tmp_no  = Trim(Replace(frm1.txtRes_no.value,"-",""))
        res_no1 = Mid(tmp_no, 1, 6)
        res_no2 = Mid(tmp_no, 7, 7)
        if  Len(tmp_no) = 13 then
			On Error Resume Next          
            intChk = Cint(Mid(res_no1, 1, 1)) * 2 + Cint(Mid(res_no1, 2, 1)) * 3 + _
                     Cint(Mid(res_no1, 3, 1)) * 4 + Cint(Mid(res_no1, 4, 1)) * 5 + _
                     Cint(Mid(res_no1, 5, 1)) * 6 + Cint(Mid(res_no1, 6, 1)) * 7 + _
                     Cint(Mid(res_no2, 1, 1)) * 8 + Cint(Mid(res_no2, 2, 1)) * 9 + _
                     Cint(Mid(res_no2, 3, 1)) * 2 + Cint(Mid(res_no2, 4, 1)) * 3 + _
                     Cint(Mid(res_no2, 5, 1)) * 4 + Cint(Mid(res_no2, 6, 1)) * 5
			if err.number <>0  then
				Set gActiveElement = document.ActiveElement  									
                .vspdData.Action = 0
                txtRes_no_Check = false
                exit function									
			end if                     
            intMod = intChk Mod 11
            intDef = 11 - intMod
            If intDef = 10 Then
                intDef = 0
            ElseIf intDef = 11 Then
                intDef = 1
            End If
            If Cstr(intDef) <> Mid(res_no2, 7, 1) Then
					txtRes_no_Check = false            
                    exit function
            End If

        else
				txtRes_no_Check = false        
                exit function
        end if
		res_no1 = Mid(Trim(Replace(frm1.txtRes_no.value,"-","")),7,1)
		If  res_no1 = "1" or res_no1 = "3" Then
		    frm1.txtSex_cd(0).checked = true
		ElseIf res_no1 = "2" or res_no1 = "4" Then
		    frm1.txtSex_cd(1).checked = true
		End if
	End If
	lgBlnFlgChgValue = True
End Function

Sub txtSch_ship_onChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtRetire_resn_onChange()
	lgBlnFlgChgValue = True
End Sub

Sub txtBirt_Change()
	lgBlnFlgChgValue = True
End Sub

Sub txtMemo_dt_Change()
	lgBlnFlgChgValue = True
End Sub

Function txtEntr_cd_onChange()
    txtEntr_cd_onChange = true

    If  frm1.txtEntr_cd.value = "" Then
        frm1.txtEntr_cd_nm.value = ""
        frm1.txtEntr_cd.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0016", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtEntr_cd.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtEntr_cd_nm.value = ""
            Call  DisplayMsgBox("970000", "x","입사구분코드","x")
	        frm1.txtEntr_cd.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtEntr_cd_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If

End Function

Function txtOcpt_type_onChange()
    txtOcpt_type_onChange = true

    If  frm1.txtOcpt_type.value = "" Then
        frm1.txtOcpt_type_nm.value = ""
        frm1.txtOcpt_type.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0003", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtOcpt_type.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtOcpt_type_nm.value = ""
            Call  DisplayMsgBox("970000", "x","직종코드","x")
	        frm1.txtOcpt_type.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtOcpt_type_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If

End Function


Function txtDir_indir_onChange()
    txtDir_indir_onChange = true

    If  frm1.txtDir_indir.value = "" Then
        frm1.txtDir_indir_nm.value = ""
        frm1.txtDir_indir.focus
        Set gActiveElement = document.ActiveElement
    Else
        if   CommonQueryRs(" MINOR_NM "," B_MINOR "," MAJOR_CD = " & FilterVar("H0071", "''", "S") & " AND MINOR_CD =  " & FilterVar(frm1.txtDir_indir.value , "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6) = false then
            frm1.txtDir_indir_nm.value = ""
            Call  DisplayMsgBox("970000", "x","직간접구분코드","x")
	        frm1.txtDir_indir.focus
	        Set gActiveElement = document.ActiveElement
	        exit function
	    Else
	        frm1.txtDir_indir_nm.value = Replace(lgF0, Chr(11), "")
	    End If
    End If

End Function

'========================================================================================================
' Name : txtBirt_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtBirt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")  
        frm1.txtBirt.Action = 7 
        frm1.txtBirt.focus
    End If
End Sub

'========================================================================================================
' Name : txtGroup_entr_dt_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtGroup_entr_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")  
        frm1.txtGroup_entr_dt.Action = 7 
        frm1.txtGroup_entr_dt.focus
    End If
End Sub

'========================================================================================================
' Name : txtEntr_dt_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtEntr_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")      
        frm1.txtEntr_dt.Action = 7 
        frm1.txtEntr_dt.focus
    End If
End Sub

'========================================================================================================
' Name : txtIntern_dt_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtIntern_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")      
        frm1.txtIntern_dt.Action = 7 
        frm1.txtIntern_dt.focus
    End If
End Sub

'========================================================================================================
' Name : txtResent_promote_dt_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtResent_promote_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")      
        frm1.txtResent_promote_dt.Action = 7 
        frm1.txtResent_promote_dt.focus
    End If
End Sub

'========================================================================================================
' Name : txtOrder_change_dt_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtOrder_change_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")      
        frm1.txtOrder_change_dt.Action = 7 
        frm1.txtOrder_change_dt.focus
    End If
End Sub

'========================================================================================================
' Name : txtRetire_dt_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtRetire_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")      
        frm1.txtRetire_dt.Action = 7 
        frm1.txtRetire_dt.focus
    End If
End Sub

'========================================================================================================
' Name : txtMil_start_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtMil_start_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")      
        frm1.txtMil_start.Action = 7 
        frm1.txtMil_start.focus
    End If
End Sub

'========================================================================================================
' Name : txtMil_end_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtMil_end_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")      
        frm1.txtMil_end.Action = 7 
        frm1.txtMil_end.focus
    End If
End Sub

'========================================================================================================
' Name : txtMemo_dt_DblClick
' Desc : developer describe this line 
'========================================================================================================
Sub txtMemo_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")      
        frm1.txtMemo_dt.Action = 7 
        frm1.txtMemo_dt.focus
    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/UNI2KCM.inc" -->	

</HEAD>

<BODY SCROLL="NO" TABINDEX="-1">
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
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab1()">
							<TR>
								<td background="../../../CShared/image/table/seltab_up_bg.gif"><img src="../../../CShared/image/table/seltab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>기본인적사항</font></td>
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="right"><img src="../../../CShared/image/table/seltab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab2()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>회사근무사항</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab3()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>기타인적사항</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
						    </TR>
						</TABLE>
					</TD>
					<TD CLASS="CLSMTABP">
						<TABLE ID="MyTab" CELLSPACING=0 CELLPADDING=0 ONCLICK="ClickTab4()">
							<TR>
								<td background="../../../CShared/image/table/tab_up_bg.gif"><img src="../../../CShared/image/table/tab_up_left.gif" width="9" height="23"></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="center" CLASS="CLSMTAB"><font color=white>주소사항</font></td>
								<td background="../../../CShared/image/table/tab_up_bg.gif" align="right"><img src="../../../CShared/image/table/tab_up_right.gif" width="10" height="23"></td>
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
			        <TD <%=HEIGHT_TYPE_02%> WIDTH=100% COLSPAN=2></TD>
			    </TR>
			    <TR>
					<TD HEIGHT=20 WIDTH=100% COLSPAN=2>
                        <FIELDSET CLASS="CLSFLD">
					    <TABLE <%=LR_SPACE_TYPE_40%>>
					  		<TR>
					  			<TD CLASS=TD5 NOWRAP>사원</TD>
					  			<TD CLASS=TD6 NOWRAP>
					  				<INPUT NAME="txtEmp_no1" MAXLENGTH=13 SIZE=13 ALT ="사원" tag="12XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenEmp"></TD>
					  		    <TD CLASS=TD5 NOWRAP>성명</TD>
					  			<TD CLASS=TD6 NOWRAP><INPUT NAME="txtName1" MAXLENGTH=30 SIZE=20 ALT ="성명" tag="14X"></TD>
					  		</TR>
					    </TABLE>
                        </FIELDSET>
					</TD>			  
			    </TR>
			    <TR>
			        <TD <%=LR_SPACE_TYPE_03%> WIDTH=100% COLSPAN=2></TD>
			    </TR>
   	            <TR>
   	                <TD WIDTH=10% VALIGN="TOP" HEIGHT="*" BGCOLOR=#eeeeec>
                        <TABLE <%=LR_SPACE_TYPE_60%>> 
					    	<TR><TD><img src="../../../CShared/image/default_picture.jpg" name="imgPhoto" WIDTH=100 HEIGHT=150></TD></TR>
					    	<TR><TD><INPUT style="WIDTH: 99px; HEIGHT: 22px" type=button size=48 value="사진등록" id=button10 name=button10 ONCLICK="VBSCRIPT:PgmJump1(BIZ_PGM_JUMP_ID10)" tag=2></TD></TR>
					    	<TR><TD><INPUT style="WIDTH: 99px; HEIGHT: 22px" type=button size=48 value="가족사항" id=button0 name=button0 ONCLICK="VBSCRIPT:PgmJump1(BIZ_PGM_JUMP_ID)" tag=2></TD></TR>
					    	<TR><TD><INPUT style="WIDTH: 99px; HEIGHT: 22px" type=button size=48 value="학    력" id=button1 name=button1 ONCLICK="vbscript:PgmJump1(BIZ_PGM_JUMP_ID1)" tag=2></TD></TR>
					    	<TR><TD><INPUT style="WIDTH: 99px; HEIGHT: 22px" type=button size=48 value="경    력" id=button2 name=button2 ONCLICK="vbscript:PgmJump1(BIZ_PGM_JUMP_ID2)" tag=2></TD></TR>
					    	<TR><TD><INPUT style="WIDTH: 99px; HEIGHT: 22px" type=button size=48 value="보 증 인" id=button3 name=button3 ONCLICK="vbscript:PgmJump1(BIZ_PGM_JUMP_ID3)" tag=2></TD></TR>
					    	<TR><TD><INPUT style="WIDTH: 99px; HEIGHT: 22px" type=button size=48 value="자격면허" id=button4 name=button4 ONCLICK="vbscript:PgmJump1(BIZ_PGM_JUMP_ID4)" tag=2></TD></TR>
					    	<TR><TD><INPUT style="WIDTH: 99px; HEIGHT: 22px" type=button size=48 value="인사변동" id=button5 name=button5 ONCLICK="vbscript:PgmJump1(BIZ_PGM_JUMP_ID5)" tag=2></TD></TR>
					    	<TR><TD><INPUT style="WIDTH: 99px; HEIGHT: 22px" type=button size=48 value="교    육" id=button6 name=button6 ONCLICK="vbscript:PgmJump1(BIZ_PGM_JUMP_ID6)" tag=2></TD></TR>
					    	<TR><TD><INPUT style="WIDTH: 99px; HEIGHT: 22px" type=button size=48 value="상    벌" id=button7 name=button7 ONCLICK="vbscript:PgmJump1(BIZ_PGM_JUMP_ID7)" tag=2></TD></TR>
					    	<TR><TD><INPUT style="WIDTH: 99px; HEIGHT: 22px" type=button size=48 value="고    과" id=button8 name=button8 ONCLICK="vbscript:PgmJump1(BIZ_PGM_JUMP_ID8)" tag=2></TD></TR>
					    	<TR><TD><INPUT style="WIDTH: 99px; HEIGHT: 22px" type=button size=48 value="외 국 어" id=button9 name=button9 ONCLICK="vbscript:PgmJump1(BIZ_PGM_JUMP_ID9)" tag=2></TD></TR>
                        </TABLE> 
   	                </TD>
   	                <TD WIDTH=90% VALIGN="TOP" HEIGHT="*">
					    <DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
                        <TABLE <%=LR_SPACE_TYPE_60%>>
					    	<TR>
					    	    <TD WIDTH=100%>
					    			<TABLE WIDTH=100% HEIGHT=100% CELLSPACING=0 CELLPADDING=0>
                                        <TR><TD CLASS="TD6" HEIGHT=5 WIDTH=100% colspan=4></TD></TR>
                                        <TR>
					    					<TD CLASS="TD5" NOWRAP>사번</TD>   
					    					<TD CLASS="TD6"><INPUT NAME="txtEmp_no" ALT="사번" TYPE="Text" MAXLENGTH=13 SiZE=13 tag="23XXXU"></TD>
					    					<TD CLASS="TD5" NOWRAP>한자성명</TD>
					    					<TD CLASS="TD6"><INPUT NAME="txtHanJa_name" ALT="한자성명" TYPE="Text" MAXLENGTH=40 SiZE=30 tag="21" ></TD>
                                        </TR>
                                        <TR>
					    					<TD CLASS="TD5" NOWRAP>성명</TD>
					    					<TD CLASS="TD6"><INPUT NAME="txtName" ALT="성명" TYPE="Text" MAXLENGTH=30 SiZE=20 tag="22XXX"></TD>
					    					<TD CLASS="TD5" NOWRAP>영문명</TD>
					    					<TD CLASS="TD6"><INPUT NAME="txtEng_name" ALT="영문명" TYPE="Text" MAXLENGTH=50 SiZE=30 tag="21" ONCHANGEONCLICK="vbscript:txtEng_name_Change()"></TD>
                                        </TR>
                                        <TR>
					    					<TD CLASS="TD5" NOWRAP>국적</TD>
					    					<TD CLASS="TD6"><INPUT NAME="txtNat_cd" ALT="국적" TYPE="Text" MAXLENGTH=2 SiZE=5 tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSItemDC frm1.txtName.value,102">&nbsp;<INPUT NAME="txtNat_cd_nm" TYPE="Text" MAXLENGTH=20 SIZE=20 tag="24"></TD>
					    					<TD CLASS="TD5" NOWRAP>출신도</TD>
					    					<TD CLASS="TD6"><INPUT NAME="txtNatv_State" ALT="출신도" TYPE="Text" MAXLENGTH=2 SiZE=5 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSItemDC frm1.txtName.value,27">&nbsp;<INPUT NAME="txtNatv_state_nm" TYPE="Text" MAXLENGTH=20 SIZE=20 tag="24"></TD>
                                        </TR>
                                        <TR>
					    					<TD CLASS="TD5" NOWRAP>주민번호</TD>
					    					<TD CLASS="TD6"><INPUT NAME="txtRes_no" ALT="주민번호" TYPE="Text" MAXLENGTH=18 SiZE=14 tag="22XXXU">
                                            	            <INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtSex_cd" tag="22X" ID="txtSex_cd1" VALUE="1"><LABEL FOR="txtSex_cd1">남</LABEL>
					    					                <INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtSex_cd" tag="22X" ID="txtSex_cd2" VALUE="2"><LABEL FOR="txtSex_cd2">여</LABEL>

					    					<TD CLASS="TD5" NOWRAP>생년월일</TD>
					    					<TD CLASS="TD6">
					    						<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtBirt CLASSID=<%=gCLSIDFPDT%> tag="22X1" ALT="생년월일" VIEWASTEXT id=txtBirt> </OBJECT>');</SCRIPT>&nbsp;
                                            	<INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtso_lu_cd" tag="22X" CHECKED ID="txtso_lu_cd1" VALUE="1"><LABEL FOR="txtso_lu_cd1">양력</LABEL>
					    					    <INPUT TYPE="RADIO" CLASS="RADIO" NAME="txtso_lu_cd" tag="22X" ID="txtso_lu_cd2" VALUE="2"><LABEL FOR="txtso_lu_cd2">음력</LABEL>
                                            </TD>
                                        </TR>
                                        <TR>
                                            <TD CLASS="TD5" NOWRAP>입사구분</TD>
                                            <TD CLASS="TD6"><INPUT NAME="txtEntr_cd" ALT="입사구분" TYPE="Text" MAXLENGTH=1 SiZE=5 tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSItemDC frm1.txtName.value,16">&nbsp;<INPUT NAME="txtEntr_cd_nm" TYPE="Text" MAXLENGTH=20 SIZE=20 tag="24"></TD>
                                            </TD>
					    					<TD CLASS="TD5" NOWRAP>채용구분</TD>
					    					<TD CLASS="TD6"><INPUT NAME="txtApp_cd" ALT="채용구분" TYPE="Text" MAXLENGTH=1 SiZE=5 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSItemDC frm1.txtName.value,17">&nbsp;<INPUT NAME="txtApp_cd_nm" TYPE="Text" MAXLENGTH=20 SIZE=20 tag="24"></TD>
                                        </TR>
                                        <TR>
					    					<TD CLASS="TD5" NOWRAP>직종</TD>
					    					<TD CLASS="TD6"><INPUT NAME="txtOcpt_type" ALT="직종" TYPE="Text" MAXLENGTH=2 SiZE=5 tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSItemDC frm1.txtName.value,3">&nbsp;<INPUT NAME="txtOcpt_type_nm" TYPE="Text" MAXLENGTH=20 SIZE=20 tag="24"></TD>
					    					<TD CLASS="TD5" NOWRAP></TD>
					    					<TD CLASS="TD6"></TD>
                                        </TR>
                                        <TR>
					    					<TD CLASS=TD5 NOWRAP>최종학력</TD>
					    					<TD CLASS="TD6">
					    					    <SELECT NAME="txtSch_ship" ALT="최종학력" CLASS ="cbonormal" TAG="21"><OPTION VALUE=""></OPTION></SELECT>
					    					</TD>
					    					<TD CLASS="TD5" NOWRAP></TD>
					    					<TD CLASS="TD6"></TD>
                                        </TR>
                                        <TR>
					    					<TD CLASS="TD5" NOWRAP>결혼여부</TD>
					    					<TD CLASS="TD6">
					    					    <SELECT NAME="txtMarry_Cd" ALT="결혼여부" CLASS ="cbonormal" TAG="21"><OPTION VALUE=""></OPTION></SELECT>
					    					</TD>
					    					<TD CLASS="TD5" NOWRAP>기념일구분</TD>
					    					<TD CLASS="TD6">
					    					    <INPUT NAME="txtMemo_cd" ALT="기념일구분" TYPE="Text" MAXLENGTH=1 SiZE=5 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSItemDC frm1.txtName.value,28">&nbsp;<INPUT NAME="txtMemo_cd_nm" TYPE="Text" MAXLENGTH=20 SIZE=20 tag="24">
					    					</TD>
                                        </TR>
                                        <TR>
					    					<TD CLASS="TD5" NOWRAP>주거구분</TD>
					    					<TD CLASS="TD6"><INPUT NAME="txtHouse_Cd" ALT="주거구분" TYPE="Text" MAXLENGTH=1 SiZE=5 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSItemDC frm1.txtName.value,15">&nbsp;<INPUT NAME="txtHouse_Cd_nm" TYPE="Text" MAXLENGTH=20 SIZE=20 tag="24"></TD>
                                            <TD CLASS="TD5" NOWRAP>기념일</TD>
					    					<TD CLASS="TD6">
					    					    <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtMemo_dt id=txtMemo_dt CLASSID=<%=gCLSIDFPDT%> tag="21X1" ALT="기념일" VIEWASTEXT> </OBJECT>');</SCRIPT>
					    					</TD>
                                        </TR>
                                        <TR>
					    					<TD CLASS="TD5" NOWRAP>직간접구분</TD>
					    					<TD CLASS="TD6"><INPUT NAME="txtDir_indir" ALT="직간접구분" TYPE="Text" MAXLENGTH=1 SiZE=5 tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSItemDC frm1.txtName.value,71">&nbsp;<INPUT NAME="txtDir_indir_nm" TYPE="Text" MAXLENGTH=20 SIZE=20 tag="24"></TD>
					    				  	<TD CLASS="TD5" NOWRAP></TD>
                                            <TD CLASS="TD6"></TD>
                                        </TR>
                                        <% Call SubFillRemBodyTD5656(8) %>
					    			</TABLE>  
					    		</TD>
					    	</TR>
					    </TABLE>
					    </DIV>
					    <DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
                        <TABLE <%=LR_SPACE_TYPE_60%>>
					    	<TR>
					    	    <TD WIDTH=100%>
					    			<TABLE WIDTH=100% HEIGHT=100% CELLSPACING=0 CELLPADDING=0>
                                        <TR><TD CLASS="TD6" HEIGHT=5 WIDTH=100% colspan=4></TD></TR>
                                        <TR>
					    					<TD CLASS="TD5" NOWRAP>법인코드</TD>
					    					<TD CLASS="TD6"><INPUT NAME="txtComp_cd" ALT="법인코드" TYPE="Text" MAXLENGTH=12 SiZE=10 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSItemDC frm1.txtName.value,103">&nbsp;<INPUT NAME="txtComp_cd_nm" TYPE="Text" MAXLENGTH=20 SIZE=20 tag="24"></TD>
					    					<TD CLASS="TD5"></TD>
					    					<TD CLASS="TD6"></TD>
                                        </TR>
                                        <TR>
					    					<TD CLASS="TD5" NOWRAP>신고사업장</TD>
					    					<TD CLASS="TD6"><SELECT NAME="txtYear_area_cd" ALT="신고사업장" CLASS ="cbonormal" TAG="22"></SELECT></TD>
					    					<TD CLASS="TD5" NOWRAP></TD>
					    					<TD CLASS="TD6"></TD>
                                        </TR>
                                        <TR>
					    					<TD CLASS="TD5" NOWRAP>근무구역</TD>
					    					<TD CLASS="TD6"><INPUT NAME="txtSect_cd" ALT="근무구역" TYPE="Text" MAXLENGTH=2 SiZE=5 tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSItemDC frm1.txtName.value,35">&nbsp;<INPUT NAME="txtSect_cd_nm" TYPE="Text" MAXLENGTH=20 SIZE=20 tag="24"></TD>
					    					<TD CLASS="TD5" NOWRAP>근무지</TD>
					    					<TD CLASS="TD6"><INPUT NAME="txtWk_Area_cd" ALT="근무지" TYPE="Text" MAXLENGTH=3 SiZE=5 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSItemDC frm1.txtName.value,36">&nbsp;<INPUT NAME="txtWk_Area_cd_nm" TYPE="Text" MAXLENGTH=20 SIZE=20 tag="24"></TD>
                                        </TR>
                                        <TR>
					    					<TD CLASS="TD5" NOWRAP>부서</TD>
					    					<TD CLASS="TD6"><INPUT NAME="txtDept_cd" ALT="부서" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnDept_cd" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenDept(0)">&nbsp;<INPUT NAME="txtDept_cd_nm" TYPE="Text" ALT="부서명" MAXLENGTH=40 SIZE=20 tag="24">
					    					</TD>
					    					<TD CLASS="TD5"></TD>
					    					<TD CLASS="TD6"></TD>
                                        </TR>
                                        <TR>
					    					<TD CLASS="TD5" NOWRAP>직위</TD>
					    					<TD CLASS="TD6"><INPUT NAME="txtRoll_pstn" ALT="직위" TYPE="Text" MAXLENGTH=2 SiZE=5 tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnRoll_pstn" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSItemDC frm1.txtRoll_pstn.value,2">&nbsp;<INPUT NAME="txtRoll_pstn_nm" TYPE="Text" MAXLENGTH=20 SIZE=20 tag="24"></TD>
					    					<TD CLASS="TD5" NOWRAP>직책</TD>
					    					<TD CLASS="TD6"><INPUT NAME="txtRole_cd" ALT="직책" TYPE="Text" MAXLENGTH=3 SiZE=5 tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSItemDC frm1.txtName.value,5">&nbsp;<INPUT NAME="txtRole_cd_nm" TYPE="Text" MAXLENGTH=20 SIZE=20 tag="24"></TD>
                                        </TR>
                                        <TR>
                                        	<TD CLASS="TD5" NOWRAP>직무</TD>
					    					<TD CLASS="TD6"><INPUT NAME="txtFunc_cd" ALT="직무" TYPE="Text" MAXLENGTH=2 SiZE=5 tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSItemDC frm1.txtName.value,4">&nbsp;<INPUT NAME="txtFunc_cd_nm" TYPE="Text" MAXLENGTH=20 SIZE=20 tag="24"></TD>
					    					<TD CLASS="TD5"></TD>
					    					<TD CLASS="TD6"></TD>
                                        </TR>
                                        <TR>
					    					<TD CLASS="TD5" NOWRAP>급호</TD>
					    					<TD CLASS="TD6"><INPUT NAME="txtPay_grd1" ALT="급호" TYPE="Text" MAXLENGTH=2 SiZE=5 tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSItemDC frm1.txtName.value,1">&nbsp;<INPUT NAME="txtPay_grd1_nm" TYPE="Text" MAXLENGTH=20 SIZE=20 tag="24"></TD>
					    					<TD CLASS="TD5" NOWRAP>호봉</TD>
					    					<TD CLASS="TD6"><INPUT NAME="txtPay_grd2" ALT="호봉" TYPE="Text" MAXLENGTH=3 SIZE=5 tag="22XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSItemDC frm1.txtPay_grd2.value,105"></TD>
                                        </TR>
                                        <TR>
                                            <TD CLASS="TD5" NOWRAP>그룹입사일</TD>
                                            <TD CLASS="TD6">
                                                <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtGroup_entr_dt CLASSID=<%=gCLSIDFPDT%> tag="21X1" ALT="그룹입사일" VIEWASTEXT id=txtGroup_entr_dt> </OBJECT>');</SCRIPT>
                                            </TD>
					    					<TD CLASS="TD5" NOWRAP>입사일</TD>
					    					<TD CLASS="TD6">
                                                <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtEntr_dt CLASSID=<%=gCLSIDFPDT%> tag="22" ALT="입사일" VIEWASTEXT id=txtEntr_dt> </OBJECT>');</SCRIPT>
                                            </TD>
                                        </TR>
                                        <TR>
                                            <TD CLASS="TD5" NOWRAP>수습만료일</TD>
                                            <TD CLASS="TD6">
                                                <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtIntern_dt CLASSID=<%=gCLSIDFPDT%> tag="21X1" ALT="수습만료일" VIEWASTEXT id=txtIntern_dt> </OBJECT>');</SCRIPT>										
                                            </TD>
					    					<TD CLASS="TD5"></TD>
					    					<TD CLASS="TD6"></TD>
                                        </TR>
                                        <TR>
					    					<TD CLASS="TD5" NOWRAP>인정경력</TD>
					    					<TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtCareer_mm name=txtCareer_mm CLASS=FPDS65 title=FPDOUBLESINGLE tag="21X9Z" ALT="인정경력"></OBJECT>');</SCRIPT>개월</TD>
					    					<TD CLASS="TD5"></TD>
					    					<TD CLASS="TD6"></TD>
                                        </TR>
                                        <TR>
					    					<TD CLASS="TD5" NOWRAP>최근승급일</TD>
					    					<TD CLASS="TD6">
                                                <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtResent_promote_dt CLASSID=<%=gCLSIDFPDT%> tag="21X1" ALT="최근승급일" VIEWASTEXT id=txtResent_promote_dt> </OBJECT>');</SCRIPT>
                                            </TD>
					    					<TD CLASS="TD5" NOWRAP>인사변동일</TD>
					    					<TD CLASS="TD6">
                                                <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtOrder_change_dt CLASSID=<%=gCLSIDFPDT%> tag="21X1" ALT="인사변동일" VIEWASTEXT id=txtOrder_change_dt> </OBJECT>');</SCRIPT>
                                            </TD>
                                        </TR>
                                        <TR>
					    					<TD CLASS="TD5" NOWRAP>퇴직일</TD>
					    					<TD CLASS="TD6">
                                                <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtRetire_dt CLASSID=<%=gCLSIDFPDT%> tag="21X1" ALT="퇴직일" VIEWASTEXT id=txtRetire_dt> </OBJECT>');</SCRIPT>
                                            </TD>
					    				  	<TD CLASS="TD5" NOWRAP>퇴직사유</TD>
					    					<TD CLASS="TD6">
					    					    <SELECT NAME="txtRetire_Resn" ALT="퇴직사유" CLASS ="cbonormal" TAG="21"><OPTION VALUE=""></OPTION></SELECT>
					    					</TD>
                                        </TR>
                                        <TR>
					    				  	<TD CLASS="TD5" NOWRAP>휴직개월</TD>
                                            <TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtRest_month name=txtRest_month CLASS=FPDS65 title=FPDOUBLESINGLE tag="21X9Z" ALT="휴직개월"></OBJECT>');</SCRIPT>개월</TD>
					    					<TD CLASS="TD5"></TD>
					    					<TD CLASS="TD6"></TD>
                                        </TR>
                                        <TR>
                                            <TD CLASS="TD5" NOWRAP></TD>
					    					<TD CLASS="TD6"><INPUT TYPE=HIDDEN TAG="21" NAME="txtTech_man" VALUE="Y"></TD>
					    					<TD CLASS="TD5"></TD>
					    					<TD CLASS="TD6"></TD>
                                        </TR>
                                        <% Call SubFillRemBodyTD5656(4) %>
					    			</TABLE>
					    		</TD>
					    	</TR>
					    </TABLE>
					    </DIV>
					    <DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
					        <TABLE <%=LR_SPACE_TYPE_60%>>
					    	    <TR>
					    	        <TD WIDTH=100%>
					    	    		<TABLE WIDTH=100% HEIGHT=100% CELLSPACING=0 CELLPADDING=0>
                                            <TR><TD CLASS="TD6" HEIGHT=5 WIDTH=100% colspan=2></TD>
                                            </TR>
                                            <TR>
					    	    				<TD WIDTH=100% CLASS=TDT>
                                                    <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=left>병역사항</LEGEND>
                                                    <TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
                                                        <TR>
					    	    				            <TD CLASS="TD5" NOWRAP>병역구분</TD>
					    	    				            <TD CLASS="TD6" colspan=3><INPUT NAME="txtMil_type" ALT="병역구분" TYPE="Text" MAXLENGTH=2 SiZE=5 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSItemDC frm1.txtName.value,19">&nbsp;<INPUT NAME="txtMil_type_nm" TYPE="Text" MAXLENGTH=10 SIZE=10 tag="24"></TD>
					    	    				            <TD CLASS="TD5" NOWRAP>병역군별</TD>
					    	    				            <TD CLASS="TD6" colspan=3><INPUT NAME="txtMil_kind" ALT="병역군별" TYPE="Text" MAXLENGTH=1 SiZE=5 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSItemDC frm1.txtName.value,20">&nbsp;<INPUT NAME="txtMil_kind_nm" TYPE="Text" MAXLENGTH=20 SIZE=20 tag="24"></TD>
                                                        </TR>
                                                        <TR>
					    	    				            <TD CLASS="TD5" NOWRAP>복무기간</TD>
					    	    				            <TD CLASS="TD6" colspan=3>
					    	    				                <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtMil_start CLASSID=<%=gCLSIDFPDT%> tag="21X1" ALT="복무기간1" VIEWASTEXT id=txtMil_start></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
                                                                <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT title=FPDATETIME CLASS=FPDTYYYYMMDD name=txtMil_End CLASSID=<%=gCLSIDFPDT%> tag="21X1" ALT="복무기간2" VIEWASTEXT id=txtMil_End></OBJECT>');</SCRIPT>
                                                            </TD>      
					    	    				            <TD CLASS="TD5" NOWRAP>병역등급</TD>
					    	    				            <TD CLASS="TD6" colspan=3><INPUT NAME="txtMil_grade" ALT="병역등급" TYPE="Text" MAXLENGTH=1 SiZE=5 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSItemDC frm1.txtName.value,21">&nbsp;<INPUT NAME="txtMil_grade_nm" TYPE="Text" MAXLENGTH=10 SIZE=10 tag="24"></TD>
					    	    				        </TR>
					    	    				        <TR>
					    	    				            <TD CLASS="TD5" NOWRAP>병역병과</TD>
					    	    				            <TD CLASS="TD6" colspan=3><INPUT NAME="txtMil_branch" ALT="병역병과" TYPE="Text" MAXLENGTH=2 SiZE=5 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSItemDC frm1.txtName.value,22">&nbsp;<INPUT NAME="txtMil_branch_nm" TYPE="Text" MAXLENGTH=10 SIZE=10 tag="24"></TD>
					    	    				            <TD CLASS="TD5" NOWRAP>군번</TD>
					    	    				            <TD CLASS="TD6" colspan=3><INPUT NAME="txtMil_no" ALT="군번" TYPE="Text" MAXLENGTH=10 SiZE=12 tag="21XXXU"></TD>
                                                        </TR>
                                                        <TR HEIGHT=3>
                                                            <TD CLASS="TD5"></TD>
                                                            <TD CLASS="TD6"></TD>
                                                        </TR>
                                                    </TABLE>
                                                    </FIELDSET>
					    	    				</TD>
					    	    				</TR>
					    	    				<TR>
					    	    				<TD WIDTH=100% CLASS=TDT>
                                                    <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=left>기타사항</LEGEND>
                                                    <TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
                                                        <TR>
					    	    				            <TD CLASS="TD5" NOWRAP>보훈구분</TD>
					    	    				            <TD CLASS="TD6"><SELECT NAME="txtRelief_cd" ALT="보훈구분" CLASS ="cbonormal" TAG="21"><OPTION VALUE=""></OPTION></SELECT></TD>
					    	    				            <TD CLASS="TD5" NOWRAP>보훈등급</TD>
					    	    				            <TD CLASS="TD6"><INPUT NAME="txtRelief_grade" ALT="보훈등급" TYPE="Text" MAXLENGTH=2 SiZE=5 tag=21XXXU></TD>                                                
                                                        </TR>
                                                        <TR>
					    	    				            <TD CLASS="TD5" NOWRAP>장애구분</TD>
					    	    				            <TD CLASS="TD6"><SELECT NAME="txtParia_cd" ALT="장애구분" CLASS ="cbonormal" TAG="21"><OPTION VALUE=""></OPTION></SELECT></TD>
					    	    				            <TD CLASS="TD5" NOWRAP>장애등급</TD>
					    	    				            <TD CLASS="TD6"><INPUT NAME="txtParia_grade" ALT="장애등급" TYPE="Text" MAXLENGTH=2 SiZE=5 tag=21XXXU></TD>                                                
                                                        </TR>
                                                        <TR>
					    	    				            <TD CLASS="TD5" NOWRAP>특기</TD>
					    	    				            <TD CLASS="TD6"><INPUT NAME="txtTalent" ALT="특기" TYPE="Text" MAXLENGTH=20 SiZE=20 tag="21XXXU"></TD>
					    	    				            <TD CLASS="TD5" NOWRAP>종교</TD>
					    	    				            <TD CLASS="TD6"><INPUT NAME="txtRelig_cd" ALT="종교" TYPE="Text" MAXLENGTH=1 SiZE=5 tag="21XXXU"><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnSItemDC" align=top TYPE="BUTTON" ONCLICK="vbscript:OpenSItemDC frm1.txtName.value,18">&nbsp;<INPUT NAME="txtRelig_cd_nm" TYPE="Text" MAXLENGTH=10 SIZE=10 tag="24"></TD>
                                                        </TR>
                                                    </TABLE>
                                                    </FIELDSET>
					    	    				</TD>   
                                            </TR>
                                            <TR>
					    	    				<TD WIDTH=100% CLASS=TDT>
                                                    <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=left>추천인사항</LEGEND>
                                                    <TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
                                                        <TR>
					    	    				            <TD CLASS="TD5" NOWRAP>성명</TD>
					    	    				            <TD CLASS="TD6"><INPUT NAME="txtNomit_name" ALT="추천인성명" TYPE="Text" MAXLENGTH=30 SiZE=20 tag="21"></TD>
					    	    				            <TD CLASS="TD5" NOWRAP>관계</TD>
					    	    				            <TD CLASS="TD6"><INPUT NAME="txtNomit_rel" ALT="추천인관계" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="21XXXU"></TD>                                                
                                                        </TR>
                                                        <TR>
					    	    				            <TD CLASS="TD5" NOWRAP>근무지</TD>
					    	    				            <TD CLASS="TD6">
                                                                <INPUT NAME="txtNomit_comp_nm" ALT="근무지" TYPE="Text" MAXLENGTH=30 SiZE=30 tag="21XXXU"></TD>
                                                            </TD>      
					    	    				            <TD CLASS="TD5" NOWRAP>직위</TD>
					    	    				            <TD CLASS="TD6">
                                                                <INPUT NAME="txtNomit_roll_pstn" ALT="직위" TYPE="Text" MAXLENGTH=10 SiZE=10 tag="21XXXU"></TD>
                                                        </TR>
                                                    </TABLE>
                                                    </FIELDSET>
					    	    				</TD>
					    	    				</TR>
					    	    				<TR>
					    	    				<TD WIDTH=100% CLASS=TDT>
                                                    <FIELDSET CLASS="CLSFLD"><LEGEND ALIGN=left>신체사항</LEGEND>
                                                    <TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
                                                        <TR>
					    	    				            <TD CLASS="TD5" NOWRAP>신장</TD>
					    	    				            <TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtHgt name=txtHgt CLASS=FPDS65 title=FPDOUBLESINGLE tag="21X70" ALT="신장"></OBJECT>');</SCRIPT>&nbsp;Cm</TD>
					    	    				            <TD CLASS="TD5" NOWRAP>체중</TD>
					    	    				            <TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtWgt name=txtWgt CLASS=FPDS65 title=FPDOUBLESINGLE tag="21X70" ALT="체중"></OBJECT>');</SCRIPT>&nbsp;Kg</TD>
                                                        </TR>
                                                        <TR>
					    	    				            <TD CLASS="TD5" NOWRAP>혈액형</TD>
					    	    				            <TD CLASS="TD6"><SELECT NAME="txtBlood_type1" ALT="혈액형1" CLASS ="cbonormal" TAG="21"><OPTION VALUE=""></OPTION></SELECT>&nbsp;형
					    	    				                            <SELECT NAME="txtBlood_type2" ALT="혈액형2" CLASS ="cbonormal" TAG="21"><OPTION VALUE=""></OPTION></SELECT></TD>
					    	    				            <TD CLASS="TD5" NOWRAP>색맹</TD>
					    	    				            <TD CLASS="TD6"><INPUT TYPE=CHECKBOX STYLE="BORDER-BOTTOM:0px solid; BORDER-LEFT:0px solid; BORDER-RIGHT:0px solid; BORDER-TOP:0px solid" TAG="21" NAME="txtDalt_type"></TD>
                                                        </TR>

                                                        <TR>
					    	    				            <TD CLASS="TD5" NOWRAP>시력</TD>
					    	    				            <TD CLASS="TD6"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtEyesgt_left name=txtEyesgt_left CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X8" ALT="시력(좌)"></OBJECT>');</SCRIPT>&nbsp;좌
                                                                            <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDS%> id=txtEyesgt_right name=txtEyesgt_right CLASS=FPDS40 title=FPDOUBLESINGLE tag="21X8" ALT="시력(우)"></OBJECT>');</SCRIPT>&nbsp;우</TD>
                                                            <TD CLASS="TD5"></TD>
                                                            <TD CLASS="TD6"></TD>
                                                        </TR>
                                                    </TABLE>
                                                    </FIELDSET>
					    	    				</TD>
                                            </TR>
                                            <TR><TD CLASS="TDT">&nbsp;</TD></TR>
                                            <TR><TD CLASS="TDT">&nbsp;</TD></TR>
                                            <TR><TD CLASS="TDT">&nbsp;</TD></TR>
                                            <TR><TD CLASS="TDT">&nbsp;</TD></TR>
                                            <TR><TD CLASS="TDT">&nbsp;</TD></TR>
                                        </TABLE>
					    	        </TD>
					    	    </TR>
					    	</TABLE>
                        </DIV>
					    <DIV ID="TabDiv" STYLE="FLOAT: left; HEIGHT: 100%; WIDTH: 100%;" SCROLL="NO">
					        <TABLE <%=LR_SPACE_TYPE_60%>>
					    	    <TR>
					    	        <TD WIDTH=100%>
					    	    		<TABLE WIDTH=100% HEIGHT=100% CELLSPACING=0 CELLPADDING=0>  
                                            <TR>
					    	    				<TD>
                                                    <TABLE CLASS="BasicTB" CELLSPACING=0 CELLPADDING=0>
                                                        <TR>
					    	    				            <TD CLASS="TD5" NOWRAP>본적</TD>
					    	    				            <TD CLASS="TD656">
                                                                <INPUT NAME="txtDomi" ALT="본적" TYPE="Text" MAXLENGTH=128 SiZE=80 tag=21XXXU></TD>
                                                            </TD>
                                                        </TR>
                                                        <TR>
					    	    				            <TD CLASS="TD5" NOWRAP>주민등록지</TD>
					    	    				            <TD CLASS="TD656">
                                                                <INPUT NAME="txtZip_cd" ALT="주민등록지우편번호" TYPE="Text" MAXLENGTH=12 SiZE=12 tag=21XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="VBScript:Call OpenZipCode(frm1.txtZip_cd.value, 0)"></TD>
                                                            </TD>      
                                                        </TR>
                                                        <TR>
					    	    				            <TD CLASS="TD5" NOWRAP></TD>
					    	    				            <TD CLASS="TD656">
                                                                <INPUT NAME="txtAddr" ALT="주민등록지" TYPE="Text" MAXLENGTH=128 SiZE=80 tag=21XXXU></TD>
                                                            </TD>      
                                                        </TR>
                                                        <TR>
					    	    				            <TD CLASS="TD5" NOWRAP>현주소</TD>
					    	    				            <TD CLASS="TD656">
                                                                <INPUT NAME="txtCurr_zip_cd" ALT="현주소우편번호" TYPE="Text" MAXLENGTH=12 SiZE=12 tag=21XXXU><IMG SRC="../../../CShared/image/btnPopup.gif" NAME="btnCalType" align=top TYPE="BUTTON" onclick="VBScript:Call OpenZipCode(frm1.txtCurr_zip_cd.value, 1)"></TD>
                                                        </TR>
                                                        <TR>
					    	    				            <TD CLASS="TD5" NOWRAP></TD>
					    	    				            <TD CLASS="TD656">
                                                                <INPUT NAME="txtCurr_addr" ALT="현주소" TYPE="Text" MAXLENGTH=128 SiZE=80 tag=21XXXU></TD>
                                                            </TD>      
                                                        </TR>
                                                        <TR>
					    	    				            <TD CLASS="TD5" NOWRAP>전화번호</TD>
					    	    				            <TD CLASS="TD656"><INPUT NAME="txtTel_no" ALT="전화번호" TYPE="Text" MAXLENGTH=20 SiZE=25 tag=21XXXU></TD>
                                                        </TR>
                                                        <TR>
					    	    				            <TD CLASS="TD5" NOWRAP>비상연락번호</TD>
					    	    				            <TD CLASS="TD656"><INPUT NAME="txtEm_tel_no" ALT="비상연락번호" TYPE="Text" MAXLENGTH=20 SiZE=25 tag=21XXXU></TD>
                                                        </TR>
                                                        <TR>
					    	    				            <TD CLASS="TD5" NOWRAP>E-Mail</TD>
					    	    				            <TD CLASS="TD656">
                                                                <INPUT NAME="txtEmail_addr" ALT="E-Mail" TYPE="Text" MAXLENGTH=30 SiZE=35 tag=21></TD>
                                                            </TD>      
                                                        </TR>
                                                        <TR>
					    	    				            <TD CLASS="TD5" NOWRAP>핸드폰</TD>
					    	    				            <TD CLASS="TD656">
                                                                <INPUT NAME="txtHand_tel_no" ALT="핸드폰" TYPE="Text" MAXLENGTH=20 SiZE=25 tag=21XXXU></TD>
                                                            </TD>      
                                                        </TR>
                                                        <% Call SubFillRemBodyTD56(10) %>
                                                    </TABLE>
					    	    				</TD>
                                            </TR>
                                        </TABLE>
					    	        </TD>
					    	    </TR>
					    	</TABLE>
                        </DIV>
                    </TD>
                </TR>
            </TABLE>
		</TD>
	</TR>
	<TR HEIGHT=20>
	    <TD>
	        <TABLE <%=LR_SPACE_TYPE_30%>>
	            <TR>
	                <TD WIDTH=10>&nbsp;</TD>
	         		<TD WIDTH="*" ALIGN=RIGHT><a href = "VBSCRIPT:PgmJump(BIZ_PGM_JUMP_ID11)" onClick="VBSCRIPT:CookiePage 1">조직도</a></TD>
	                <TD WIDTH=10>&nbsp;</TD>
	            </TR>
	        </TABLE>
	    </TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
	</TR>
</TABLE>
<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<INPUT TYPE=HIDDEN NAME="temp_flg_chk"   TAG="24">

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
</DIV>
</BODY>
</HTML>



