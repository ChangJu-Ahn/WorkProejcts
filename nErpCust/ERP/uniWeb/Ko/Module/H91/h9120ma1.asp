<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : 연말정산 
*  2. Function Name        : 통합연말정산신고(퇴직자미포함)
*  3. Program ID           : H9120ma1
*  4. Program Name         : 통합연말정산신고 
*  5. Program Desc         : 통합연말정산신고(퇴직자미포함)
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2003/06/13
*  9. Modifier (First)     : Hwang Jeong Won
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/IncEB.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>

<Script Language="VBScript">
Option Explicit

Const BIZ_PGM_ID      = "h9120mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID2     = "h9120mb2.asp"                                 '☆: File Creation Asp Name
Const C_SHEETMAXROWS    = 10                                      '☜: Visble row
Const C_SHEETMAXROWS1    = 10                                      '☜: Visble row
Const C_SHEETMAXROWS2    = 10                                      '☜: Visble row
Const C_SHEETMAXROWS3    = 10                                      '☜: Visble row

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop
Dim lgStrComDateType		                                            'Company Date Type을 저장(년월 Mask에 사용함.)
Dim lgStrPrevKey1,lgStrPrevKey2,lgStrPrevKey3
Dim topleftOK

'  Constants for SpreadSheet #1
Dim C_RECORD_TYPE
Dim C_DATA_TYPE
Dim C_TAX     
Dim C_PROV_DT 
Dim C_P_TYPE  
Dim C_MAG_NO 
Dim C_HOMETAX_ID	'2004
Dim C_TAX_CODE		'2004
Dim C_OWN_RGST_NO
Dim C_CUST_NM_FULL
Dim C_WORKER_DEPT	'2004
Dim C_WORKER_NM		'2004
Dim C_WORKER_TEL	'2004 
Dim C_B_COUNT      
Dim C_KR_CODE      
Dim C_TERM_CODE    
Dim C_EMPTY   
'  Constants for SpreadSheet #2
Dim C_RECORD_TYPE1
Dim C_DATA_TYPE1  
Dim C_TAX1        
Dim C_NO1         
Dim C_OWN_RGST_NO1
Dim C_CUST_NM_FULL1 
Dim C_REPRE_NM1     
Dim C_BCA010T_REPRE_RGST_NO1
Dim C_COM_NO1               
Dim C_OLD_COM_NO1           
Dim C_TOT_PROV_AMT1         
Dim C_DECI_INCOME_TAX1      
Dim C_TOT_TAX1              
Dim C_DECI_RES_TAX1         
Dim C_DECI_FARM_TAX1        
Dim C_DECI_SUM1             
Dim C_EMPTY1  

'  Constants for SpreadSheet #3
Dim C_RECORD_TYPE2          
Dim C_DATA_TYPE2            
Dim C_TAX2                  
Dim C_NO2                   
Dim C_OWN_RGST_NO2          
Dim C_OLD_COM_NO2           
Dim C_HDF020T_RES_FLAG2  
Dim C_HAA010T_NAT_CD2   '2002
Dim C_FOREIN_TAXRATE   '2004
Dim C_HAA010T_ENTR_DT2      
Dim C_HAA010T_RETIRE_DT2    
Dim C_HAA010T_NAME2         
Dim C_FOR_TYPE2             
Dim C_RES_NO2               
Dim C_START_DT2             
Dim C_END_DT2               
Dim C_HFA050T_NEW_PAY_TOT2  
Dim C_HFA050T_NEW_BONUS_TOT2
Dim C_HFA030T_AFTER_BONUS_AMT2
Dim C_NEW_TOT2                
Dim C_HFA050T_NON_TAX52       
Dim C_HFA050T_NON_TAX12       
Dim C_NON_TAX2                
Dim C_NON_TAX_SUM2            
Dim C_HFA050T_INCOME_TOT_AMT2
Dim C_HFA050T_INCOME_SUB_AMT2
Dim C_HFA050T_INCOME_AMT2    
Dim C_HFA050T_PER_SUB_AMT2   
Dim C_HFA050T_SPOUSE_SUB_AMT2
Dim C_HFA050T_SUPP_CNT2 
Dim C_HFA050T_SUPP_SUB_AMT2 
Dim C_HFA050T_OLD_CNT2      
Dim C_HFA050T_OLD_SUB_AMT2  
Dim C_HFA050T_PARIA_CNT2    
Dim C_HFA050T_PARIA_SUB_AMT2
Dim C_HFA050T_LADY_SUB_AMT2 
Dim C_HFA050T_CHL_REAR2     
Dim C_HFA050T_CHL_REAR_SUB_AMT2
Dim C_HFA050T_SMALL_SUB_AMT2   
Dim C_HFA050T_INSUR_SUB_AMT2   
Dim C_HFA050T_MED_SUB_AMT2     
Dim C_HFA050T_EDU_SUB_AMT2     
Dim C_HFA050T_HOUSE_FUND_AMT2  
Dim C_HFA050T_CONTR_SUB_AMT2 
Dim C_HFA050T_CEREMONTY_AMT
Dim C_HFA050T_STD_SUB_TOT_AMT2 
Dim C_HFA050T_STD_SUB_AMT2     
Dim C_HFA050T_NAT_PEN_SUB_AMT2 
Dim C_HFA050T_SUB_INCOME_AMT2  
Dim C_HFA050T_INDIV_ANU_AMT2   
Dim C_HFA050T_INDIV_ANU2_AMT2  
Dim C_HFA050T_INVEST_SUB_SUM_AMT2
Dim C_HFA050T_CARD_SUB_SUM_AMT2  
Dim C_HFA050T_OUR_STOCK_AMT2 
Dim C_SPECIAL_tAX_SUM
Dim C_HFA050T_TAX_STD_AMT2       
Dim C_HFA050T_CALU_TAX_AMT2      
Dim C_HFA050T_INCOME_TAX_SUB_AMT2
Dim C_EMPTY4                     
Dim C_HFA050T_HOUSE_REPAY_AMT2   
Dim C_HFA050T_FORE_PAY_AMT2   
Dim C_HFA050T_POLI_TAX_SUB  
Dim C_EMPTY_32                   
Dim C_HFA050T_TAX_SUB_AMT2       
Dim C_HFA050T_INCOME_REDU_AMT2   
Dim C_HFA050T_TAXES_REDU_AMT2    
Dim C_EMPTY2                     
Dim C_HFA050T_REDU_SUM_AMT2      
Dim C_HFA050T_DEC_INCOME_TAX_AMT2
Dim C_HFA050T_DEC_RES_TAX_AMT2   
Dim C_HFA050T_DEC_FARM_TAX_AMT2  
Dim C_DEC_TOT2                   
Dim C_HFA050T_OLD_INCOME_TAX_AMT2
Dim C_HFA050T_OLD_RES_TAX_AMT2   
Dim C_HFA050T_OLD_FARM_TAX_AMT2  
Dim C_OLD_TOT2                   
Dim C_HFA050T_NEW_INCOME_TAX_AMT2
Dim C_HFA050T_NEW_RES_TAX_AMT2   
Dim C_HFA050T_NEW_FARM_TAX_AMT2  
Dim C_NEW_SUM_TOT2               
Dim C_EMPTY_2                    
Dim C_EMP_NO_2                      

'  Constants for SpreadSheet #4
Dim C_RECORD_TYPE3         
Dim C_DATA_TYPE3           
Dim C_TAX3                 
Dim C_NO3                  
Dim C_OWN_RGST_NO3         
Dim C_EMPTY3               
Dim C_RES_NO_V3            
Dim C_HFA040T_A_COMP_NM3   
Dim C_HFA040T_A_COMP_NO3   
Dim C_HFA040T_A_PAY_TOT3   
Dim C_HFA040T_A_BONUS_TOT3 
Dim C_HFA040T_A_AFTER_BONUS_AMT3
Dim C_PAY_TOT3                  
Dim C_WORK_NO3                 
Dim C_EMPTY_3                     

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize Column  value
'========================================================================================================
Sub initSpreadPosVariables(spd) 
	if spd="A" or spd="ALL" then
		C_RECORD_TYPE     = 1 										'☆: Spread Sheet 의 Columns 인덱스 
		C_DATA_TYPE       = 2
		C_TAX             = 3
		C_PROV_DT         = 4
		C_P_TYPE          = 5
		C_MAG_NO          = 6
		C_HOMETAX_ID      = 7	
		C_TAX_CODE       = 8
		C_OWN_RGST_NO     = 9
		C_CUST_NM_FULL    = 10
		C_WORKER_DEPT     = 11
		C_WORKER_NM		  = 12
		C_WORKER_TEL	  = 13
		C_B_COUNT         = 14
		C_KR_CODE         = 15
		C_TERM_CODE       = 16
		C_EMPTY        = 17
		
	end if
	
	if spd="B" or spd="ALL" then	
		C_RECORD_TYPE1          = 1								'☆: Spread Sheet 의 Columns 인덱스 
		C_DATA_TYPE1            = 2
		C_TAX1                  = 3
		C_NO1                   = 4
		C_OWN_RGST_NO1          = 5
		C_CUST_NM_FULL1         = 6
		C_REPRE_NM1             = 7
		C_BCA010T_REPRE_RGST_NO1 = 8
		C_COM_NO1               = 9
		C_OLD_COM_NO1           = 10
		C_TOT_PROV_AMT1         = 11
		C_DECI_INCOME_TAX1      = 12
		C_TOT_TAX1              = 13
		C_DECI_RES_TAX1         = 14
		C_DECI_FARM_TAX1        = 15
		C_DECI_SUM1             = 16
		C_EMPTY1               = 17
	end if	
	if spd="C" or spd="ALL" then	
		C_RECORD_TYPE2               = 1							'☆: SPREAD SHEET 의 COLUMNS 인덱스 
		C_DATA_TYPE2                 = 2
		C_TAX2                       = 3
		C_NO2                        = 4
		C_OWN_RGST_NO2               = 5
		C_OLD_COM_NO2                = 6
		C_HDF020T_RES_FLAG2          = 7
		C_HAA010T_NAT_CD2            = 8   '2002 거주지국 
		C_FOREIN_TAXRATE             = 9   '2004 외국인단일세율 
		C_HAA010T_ENTR_DT2           = 10
		C_HAA010T_RETIRE_DT2         = 11
		C_HAA010T_NAME2              = 12
		C_FOR_TYPE2                  = 13
		C_RES_NO2                    = 14
		C_START_DT2                  = 15
		C_END_DT2                    = 16
		C_HFA050T_NEW_PAY_TOT2       = 17
		C_HFA050T_NEW_BONUS_TOT2     = 18
		C_HFA030T_AFTER_BONUS_AMT2   = 19
		C_NEW_TOT2                   = 20
		C_HFA050T_NON_TAX52          = 21
		C_HFA050T_NON_TAX12          = 22
		C_NON_TAX2                   = 23
		C_NON_TAX_SUM2               = 24
		C_HFA050T_INCOME_TOT_AMT2    = 25
		C_HFA050T_INCOME_SUB_AMT2    = 26
		C_HFA050T_INCOME_AMT2        = 27
		C_HFA050T_PER_SUB_AMT2       = 28
		C_HFA050T_SPOUSE_SUB_AMT2    = 29
		C_HFA050T_SUPP_CNT2          = 30
		C_HFA050T_SUPP_SUB_AMT2      = 31
		C_HFA050T_OLD_CNT2           = 32
		C_HFA050T_OLD_SUB_AMT2       = 33
		C_HFA050T_PARIA_CNT2         = 34
		C_HFA050T_PARIA_SUB_AMT2     = 35
		C_HFA050T_LADY_SUB_AMT2      = 36
		C_HFA050T_CHL_REAR2          = 37
		C_HFA050T_CHL_REAR_SUB_AMT2  = 38
		C_HFA050T_SMALL_SUB_AMT2     = 39
		C_HFA050T_NAT_PEN_SUB_AMT2   = 40 '2002 순서변경 
		C_HFA050T_INSUR_SUB_AMT2     = 41
		C_HFA050T_MED_SUB_AMT2       = 42
		C_HFA050T_EDU_SUB_AMT2       = 43
		C_HFA050T_HOUSE_FUND_AMT2    = 44
		C_HFA050T_CONTR_SUB_AMT2     = 45
		C_HFA050T_CEREMONTY_AMT      = 46
	
		C_HFA050T_STD_SUB_TOT_AMT2   = 47
		C_HFA050T_STD_SUB_AMT2       = 48
		C_HFA050T_SUB_INCOME_AMT2    = 49
		C_HFA050T_INDIV_ANU_AMT2     = 50
		C_HFA050T_INDIV_ANU2_AMT2    = 51
		C_HFA050T_INVEST_SUB_SUM_AMT2= 52
		C_HFA050T_CARD_SUB_SUM_AMT2  = 53 
		C_HFA050T_OUR_STOCK_AMT2     = 54 '2002 우리사주조합출연금 
		C_SPECIAL_tAX_SUM			 = 55
		C_HFA050T_TAX_STD_AMT2       = 56
		C_HFA050T_CALU_TAX_AMT2      = 57
		C_HFA050T_INCOME_REDU_AMT2   = 58 '2002 순서변경 : 세액감면 
		C_HFA050T_TAXES_REDU_AMT2    = 59
		C_EMPTY2                     = 60
		C_HFA050T_REDU_SUM_AMT2      = 61
		C_HFA050T_INCOME_TAX_SUB_AMT2= 62 '2002 순서변경 : 세액공제 
		C_EMPTY4                     = 63
		C_HFA050T_HOUSE_REPAY_AMT2   = 64
		C_HFA050T_FORE_PAY_AMT2      = 65
		C_HFA050T_POLI_TAX_SUB		 = 66
		C_EMPTY_32                   = 67
		C_HFA050T_TAX_SUB_AMT2       = 68
		C_HFA050T_DEC_INCOME_TAX_AMT2= 69
		C_HFA050T_DEC_RES_TAX_AMT2   = 70
		C_HFA050T_DEC_FARM_TAX_AMT2  = 71
		C_DEC_TOT2                   = 72
		C_HFA050T_OLD_INCOME_TAX_AMT2= 73
		C_HFA050T_OLD_RES_TAX_AMT2   = 74
		C_HFA050T_OLD_FARM_TAX_AMT2  = 75
		C_OLD_TOT2                   = 76
		C_HFA050T_NEW_INCOME_TAX_AMT2= 77
		C_HFA050T_NEW_RES_TAX_AMT2   = 78
		C_HFA050T_NEW_FARM_TAX_AMT2  = 79
		C_NEW_SUM_TOT2               = 80
		C_EMPTY_2                  = 81
		C_EMP_NO_2                  = 82
	
	end if	
	if spd="D" or spd="ALL" then	
		C_RECORD_TYPE3                = 1	    						'☆: SPREAD SHEET 의 COLUMNS 인덱스 
		C_DATA_TYPE3                  = 2
		C_TAX3                        = 3
		C_NO3                         = 4
		C_OWN_RGST_NO3                = 5
		C_EMPTY3                      = 6
		C_RES_NO_V3                   = 7
		C_HFA040T_A_COMP_NM3          = 8
		C_HFA040T_A_COMP_NO3          = 9
		C_HFA040T_A_PAY_TOT3          = 10
		C_HFA040T_A_BONUS_TOT3        = 11
		C_HFA040T_A_AFTER_BONUS_AMT3  = 12
		C_PAY_TOT3                    = 13
		C_WORK_NO3                    = 14
		C_EMPTY_3                   = 15
	end if	
End Sub
'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode       = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue   = False								    '⊙: Indicates that no value changed
	lgIntGrpCount      = 0										'⊙: Initializes Group View Size
    lgStrPrevKey       = ""                                     '⊙: initializes Previous Key
    lgStrPrevKey1       = ""                                     '⊙: initializes Previous Key
    lgStrPrevKey2       = ""                                     '⊙: initializes Previous Key
    lgStrPrevKey3       = ""                                     '⊙: initializes Previous Key            
    lgSortKey          = 1                                      '⊙: initializes sort direction		
End Sub
'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================	
Sub SetDefaultVal()
'    frm1.txtDt.Text     = UniConvDateAToB("<%=GetSvrDate%>",parent.gServerDateFormat,parent.gDateFormat)
 '   frm1.txtBas_dt.Text = frm1.txtDt.Text
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
        lgKeyStream       = Trim(Frm1.txtGubun.value) & parent.gColSep       'You Must append one character(parent.gColSep)
        lgKeyStream       = lgKeyStream & Trim(frm1.txtGigan.value) & parent.gColSep
        lgKeyStream       = lgKeyStream & Trim(frm1.txtDt.text) & parent.gColSep
        lgKeyStream       = lgKeyStream & Trim(frm1.txtSer.value) & parent.gColSep
        lgKeyStream       = lgKeyStream & Trim(frm1.txtFile.value) & parent.gColSep
        lgKeyStream       = lgKeyStream & Trim(frm1.txtBas_dt.text) & parent.gColSep
        lgKeyStream       = lgKeyStream & Trim(frm1.txtComp_cd.value) & parent.gColSep        
End Sub
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iNameArr , iNameArr1 , iNameArr2
    Dim iCodeArr , iCodeArr1 , iCodeArr2         
    '제출자 구분 
    Call CommonQueryRs("MINOR_NM,MINOR_CD","B_MINOR","MAJOR_CD = " & FilterVar("H0118", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iNameArr = lgF0
    iCodeArr = lgF1   
    Call SetCombo2(frm1.txtGubun,iCodeArr,iNameArr,Chr(11))     
    frm1.txtGubun.value = 2    
    '대상기간 
    Call CommonQueryRs("MINOR_NM,MINOR_CD","B_MINOR","MAJOR_CD = " & FilterVar("H0119", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iNameArr1 = lgF0
    iCodeArr1 = lgF1       
    Call SetCombo2(frm1.txtGigan,iCodeArr1,iNameArr1,Chr(11))            ''''''''DB에서 불러 condition에서        
    '신고사업장 
    Call CommonQueryRs("YEAR_AREA_NM,YEAR_AREA_CD","HFA100T","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iNameArr2 = lgF0
    iCodeArr2 = lgF1   
    Call SetCombo2(frm1.txtComp_cd,iCodeArr2,iNameArr2,Chr(11))     
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
Sub InitSpreadSheet(strSPD)
	Dim strMaskYM
	If parent.gComDateType = "/" Then 
		lgStrComDateType = "/" & parent.gComDateType
	Else
		lgStrComDateType = parent.gComDateType
	End If
	strMaskYM = "9999" & lgStrComDateType & "99"
	
	call InitSpreadPosVariables(strSPD )

    ' Set SpreadSheet #1
	if (strSPD = "A" or strSPD = "ALL") then	     
		With Frm1.vspdData
		    ggoSpread.Source = Frm1.vspdData
			ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread    		    
		   .ReDraw = false			
		   .MaxCols = C_EMPTY + 1                                                   '☜: Add 1 to Maxcols
		   .Col = .MaxCols                                                          '☜: Hide maxcols
		   .ColHidden = True                                                        '☜:    

		   .MaxRows = 0

			Call GetSpreadColumnPos("A")  

				ggoSpread.SSSetEdit      C_RECORD_TYPE,     "레코드구분",               12
				ggoSpread.SSSetEdit      C_DATA_TYPE,       "자료구분",                 10
				ggoSpread.SSSetEdit      C_TAX,             "세무서",                   8
				ggoSpread.SSSetEdit      C_PROV_DT,         "제출연월일",               12
				ggoSpread.SSSetEdit      C_P_TYPE,          "제출자(대리인구분)",       20
				ggoSpread.SSSetEdit      C_MAG_NO,          "세무대리인관리번호",       20
				ggoSpread.SSSetEdit      C_HOMETAX_ID,		"홈텍스ID",					20	'2004 
				ggoSpread.SSSetEdit      C_TAX_CODE,		"세무프로그램코드",			45	'2004 				
				ggoSpread.SSSetEdit      C_OWN_RGST_NO,     "사업자등록번호",           16
				ggoSpread.SSSetEdit      C_CUST_NM_FULL,    "법인명(상호)",             14
				ggoSpread.SSSetEdit      C_WORKER_DEPT,		"담당자부서",				30	'2004 	
				ggoSpread.SSSetEdit      C_WORKER_NM,		"담당자성명",				30	'2004 	
				ggoSpread.SSSetEdit      C_WORKER_TEL,		"담당자전화번호",			15	'2004 	
				ggoSpread.SSSetEdit      C_B_COUNT,         "신고의무자(B레코드) 수",   20
				ggoSpread.SSSetEdit      C_KR_CODE,         "한글코드종류",             14
				ggoSpread.SSSetEdit      C_TERM_CODE,       "제출대상기간코드",         18
				ggoSpread.SSSetEdit      C_EMPTY,           "공란",                     8

		   .ReDraw = true
	
		   lgActiveSpd = "A"
		   Call SetSpreadLock 
    
		End With
    End if		

    ' Set SpreadSheet #2
    
   	if (strSPD = "B" or strSPD = "ALL") then
   		With Frm1.vspdData1
	
		    ggoSpread.Source = Frm1.vspdData1
			ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread    
		   .ReDraw = false
		   .MaxCols = C_EMPTY1 + 1                                                   '☜: Add 1 to Maxcols
		   .Col = .MaxCols                                                           '☜: Hide maxcols
		   .ColHidden = True                                                         '☜:

		   .MaxRows = 0
	
			Call GetSpreadColumnPos("B") 

				ggoSpread.SSSetEdit      C_RECORD_TYPE1,            "레코드구분",                   12
				ggoSpread.SSSetEdit      C_DATA_TYPE1,              "자료구분",                     10
				ggoSpread.SSSetEdit      C_TAX1,                    "세무서",                       8
				ggoSpread.SSSetEdit      C_NO1,                     "일련번호",                     12
				ggoSpread.SSSetEdit      C_OWN_RGST_NO1,            "사업자등록번호",               16
				ggoSpread.SSSetEdit      C_CUST_NM_FULL1,           "법인명(상호)",                 14
				ggoSpread.SSSetEdit      C_REPRE_NM1,               "대표자(성명)",                 13
				ggoSpread.SSSetEdit      C_BCA010T_REPRE_RGST_NO1,  "주민(법인)등록번호",           20
				ggoSpread.SSSetEdit      C_COM_NO1,                 "주(현)제출건수(C레코드수)",   24
				ggoSpread.SSSetEdit      C_OLD_COM_NO1,             "종(전)레코드수(D레코드수)",   24
				ggoSpread.SSSetEdit      C_TOT_PROV_AMT1,           "소득금액총계",                14
				ggoSpread.SSSetEdit      C_DECI_INCOME_TAX1,        "소득세결정세액총계",          20
				ggoSpread.SSSetEdit      C_TOT_TAX1,                "법인세결정세액총계",          20
				ggoSpread.SSSetEdit      C_DECI_RES_TAX1,           "주민세결정세액총계",          20
				ggoSpread.SSSetEdit      C_DECI_FARM_TAX1,          "농특세결정세액총계",          20
				ggoSpread.SSSetEdit      C_DECI_SUM1,               "결정세액총계",                14
				ggoSpread.SSSetEdit      C_EMPTY1,                  "공란",                         8

		   .ReDraw = true
	
		   lgActiveSpd = "B"
		   Call SetSpreadLock 
    
		End With
    End if			

    ' Set SpreadSheet #3
   	if (strSPD = "C" or strSPD = "ALL") then    
	    With Frm1.vspdData2
	        ggoSpread.Source = Frm1.vspdData2
			ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread    
		   .ReDraw = false
	       .MaxCols = C_EMP_NO_2 + 1                                                   '☜: Add 1 to Maxcols
		   .Col = .MaxCols                                                            '☜: Hide maxcols
	       .ColHidden = True                                                          '☜:
	    
	       .MaxRows = 0		   
			Call GetSpreadColumnPos("C")
	 
			ggoSpread.SSSetEdit     C_RECORD_TYPE2,                 "레코드구분",               12
			ggoSpread.SSSetEdit     C_DATA_TYPE2,                   "자료구분",                 10
			ggoSpread.SSSetEdit     C_TAX2,                         "세무서",                   8
			ggoSpread.SSSetEdit     C_NO2,                          "일련번호",                 10
			ggoSpread.SSSetEdit     C_OWN_RGST_NO2,                 "사업자등록번호",           16
			ggoSpread.SSSetEdit     C_OLD_COM_NO2,                  "종(전)근무처수",           16
			ggoSpread.SSSetEdit     C_HDF020T_RES_FLAG2,            "거주자구분코드",           16
            ggoSpread.SSSetEdit     C_HAA010T_NAT_CD2,              "거주지국코드",             12  '2002
            ggoSpread.SSSetEdit     C_FOREIN_TAXRATE,              "외국인단일세율적용",             12  '2004
			ggoSpread.SSSetEdit     C_HAA010T_ENTR_DT2,             "귀속년도시작연월일",       20
			ggoSpread.SSSetEdit     C_HAA010T_RETIRE_DT2 ,          "귀속연도종료연월일",       20
			ggoSpread.SSSetEdit     C_HAA010T_NAME2,                "성명",                     8
			ggoSpread.SSSetEdit     C_FOR_TYPE2,                    "내국/외국인구분코드",      21
			ggoSpread.SSSetEdit     C_RES_NO2,                      "주민등록번호",             14
			ggoSpread.SSSetEdit     C_START_DT2,                    "감면기간시작연월일",       20
			ggoSpread.SSSetEdit     C_END_DT2,                      "감면기간종료연월일",       20
			ggoSpread.SSSetEdit     C_HFA050T_NEW_PAY_TOT2,         "급여총액",                 10
			ggoSpread.SSSetEdit     C_HFA050T_NEW_BONUS_TOT2,       "상여총액",                 10
			ggoSpread.SSSetEdit     C_HFA030T_AFTER_BONUS_AMT2,     "인정상여",                 10
			ggoSpread.SSSetEdit     C_NEW_TOT2,                     "계",                       6
			ggoSpread.SSSetEdit     C_HFA050T_NON_TAX52,            "국외근로",                 10
			ggoSpread.SSSetEdit     C_HFA050T_NON_TAX12,            "야간근로수당등",           16
			ggoSpread.SSSetEdit     C_NON_TAX2,                     "기타비과세",               12
			ggoSpread.SSSetEdit     C_NON_TAX_SUM2,                 "계",                       6
			ggoSpread.SSSetEdit     C_HFA050T_INCOME_TOT_AMT2,      "총급여",                   11
			ggoSpread.SSSetEdit     C_HFA050T_INCOME_SUB_AMT2,      "근로소득공제",             14
			ggoSpread.SSSetEdit     C_HFA050T_INCOME_AMT2,          "과세대상근로소득금액",     22
			ggoSpread.SSSetEdit     C_HFA050T_PER_SUB_AMT2,         "본인공제금액",             14
			ggoSpread.SSSetEdit     C_HFA050T_SPOUSE_SUB_AMT2,      "배우자공제금액",           16
			ggoSpread.SSSetEdit     C_HFA050T_SUPP_CNT2,            "부양가족공제인원",         18
			ggoSpread.SSSetEdit     C_HFA050T_SUPP_SUB_AMT2,        "부양가족공제금액",         18
			ggoSpread.SSSetEdit     C_HFA050T_OLD_CNT2,             "경로우대공제인원",         18
			ggoSpread.SSSetEdit     C_HFA050T_OLD_SUB_AMT2,         "경로우대공제금액",         18
			ggoSpread.SSSetEdit     C_HFA050T_PARIA_CNT2,           "장애자공제인원",           16
			ggoSpread.SSSetEdit     C_HFA050T_PARIA_SUB_AMT2,       "장애자공제금액",           16
			ggoSpread.SSSetEdit     C_HFA050T_LADY_SUB_AMT2,        "부녀자공제금액",           16
			ggoSpread.SSSetEdit     C_HFA050T_CHL_REAR2,            "자녀양육비공제인원",       20
			ggoSpread.SSSetEdit     C_HFA050T_CHL_REAR_SUB_AMT2,    "자녀양육비공제금액",       20
			ggoSpread.SSSetEdit     C_HFA050T_SMALL_SUB_AMT2,       "소수공제자추가공제",       20
			ggoSpread.SSSetEdit     C_HFA050T_NAT_PEN_SUB_AMT2,     "연금보험료",               12
			ggoSpread.SSSetEdit     C_HFA050T_INSUR_SUB_AMT2,       "보험료",                   8
			ggoSpread.SSSetEdit     C_HFA050T_MED_SUB_AMT2,         "의료비",                   8
			ggoSpread.SSSetEdit     C_HFA050T_EDU_SUB_AMT2,         "교육비",                   8
			ggoSpread.SSSetEdit     C_HFA050T_HOUSE_FUND_AMT2,      "주택자금",                 10
			ggoSpread.SSSetEdit     C_HFA050T_CONTR_SUB_AMT2,       "기부금",                   10
			ggoSpread.SSSetEdit     C_HFA050T_CEREMONTY_AMT,        "혼인/이사/장례비",			10	'2004			
			ggoSpread.SSSetEdit     C_HFA050T_STD_SUB_TOT_AMT2,     "계(특별공제)",             14
			ggoSpread.SSSetEdit     C_HFA050T_STD_SUB_AMT2,         "표준공제",                 10
			ggoSpread.SSSetEdit     C_HFA050T_SUB_INCOME_AMT2,      "차감소득금액",             14
			ggoSpread.SSSetEdit     C_HFA050T_INDIV_ANU_AMT2,       "개인연금저축",             14
			ggoSpread.SSSetEdit     C_HFA050T_INDIV_ANU2_AMT2,      "연금저축",                 10
			ggoSpread.SSSetEdit     C_HFA050T_INVEST_SUB_SUM_AMT2,  "투자조합출자등소득공제",   24
			ggoSpread.SSSetEdit     C_HFA050T_CARD_SUB_SUM_AMT2,    "신용카드소득공제",         18
            ggoSpread.SSSetEdit     C_HFA050T_OUR_STOCK_AMT2,       "우리사주조합출연금",       18 '2002
            ggoSpread.SSSetEdit     C_SPECIAL_tAX_SUM,				"조특소득공제계",			18	'2004	
            ggoSpread.SSSetEdit     C_HFA050T_TAX_STD_AMT2,         "종합소득과세표준",         18
            ggoSpread.SSSetEdit     C_HFA050T_CALU_TAX_AMT2,        "산출세액",                 10
            ggoSpread.SSSetEdit     C_HFA050T_INCOME_REDU_AMT2,     "소득세법",                 10 '2002 세액감면 
            ggoSpread.SSSetEdit     C_HFA050T_TAXES_REDU_AMT2,      "조특법",                   8
            ggoSpread.SSSetEdit     C_EMPTY2,                       "공란",                     8
            ggoSpread.SSSetEdit     C_HFA050T_REDU_SUM_AMT2,        "계",                       6
            ggoSpread.SSSetEdit     C_HFA050T_INCOME_TAX_SUB_AMT2,  "근로소득세액공제",         18 '2002 세액공제 
            ggoSpread.SSSetEdit     C_EMPTY4,                       "납세조합공제",             14
            ggoSpread.SSSetEdit     C_HFA050T_HOUSE_REPAY_AMT2,     "주택차입금세액공제",       20
            ggoSpread.SSSetEdit     C_HFA050T_FORE_PAY_AMT2,        "외국납부세액공제",         18
            ggoSpread.SSSetEdit     C_HFA050T_POLI_TAX_SUB,		"기부정치자금",				18	'2004
            ggoSpread.SSSetEdit     C_EMPTY_32,                     "공란",                     8
            ggoSpread.SSSetEdit     C_HFA050T_TAX_SUB_AMT2,         "계",                       6
			ggoSpread.SSSetEdit     C_HFA050T_INCOME_REDU_AMT2,     "소득세법",                 10
			ggoSpread.SSSetEdit     C_HFA050T_TAXES_REDU_AMT2,      "조특법",                   8
			ggoSpread.SSSetEdit     C_EMPTY2,                       "공란",                     8
			ggoSpread.SSSetEdit     C_HFA050T_REDU_SUM_AMT2,        "계",                       6
			ggoSpread.SSSetEdit     C_HFA050T_DEC_INCOME_TAX_AMT2,  "소득세",                   8
			ggoSpread.SSSetEdit     C_HFA050T_DEC_RES_TAX_AMT2,     "주민세",                   8
			ggoSpread.SSSetEdit     C_HFA050T_DEC_FARM_TAX_AMT2,    "농어촌특별세",             14
			ggoSpread.SSSetEdit     C_DEC_TOT2,                     "계",                       6
			ggoSpread.SSSetEdit     C_HFA050T_OLD_INCOME_TAX_AMT2,  "소득세",                   8
			ggoSpread.SSSetEdit     C_HFA050T_OLD_RES_TAX_AMT2,     "주민세",                   8
			ggoSpread.SSSetEdit     C_HFA050T_OLD_FARM_TAX_AMT2,    "농어촌특별세",             14
			ggoSpread.SSSetEdit     C_OLD_TOT2,                     "계",                       6
			ggoSpread.SSSetEdit     C_HFA050T_NEW_INCOME_TAX_AMT2,  "소득세",                   8
			ggoSpread.SSSetEdit     C_HFA050T_NEW_RES_TAX_AMT2,     "주민세",                   8
			ggoSpread.SSSetEdit     C_HFA050T_NEW_FARM_TAX_AMT2,    "농어촌특별세",             14
			ggoSpread.SSSetEdit     C_NEW_SUM_TOT2,                 "계",                       6
			ggoSpread.SSSetEdit     C_EMPTY_2,                      "공란",                     8
 			ggoSpread.SSSetEdit     C_EMP_NO_2,                     "사원번호",                 10
 			
		   .ReDraw = true
		    Call ggoSpread.SSSetColHidden(C_EMP_NO_2,C_EMP_NO_2,True)		
		    call ggoSpread.SSSetColHidden(C_NO2,C_NO2,true)
	       lgActiveSpd = "C"
	       Call SetSpreadLock 
	    
	    End With
    End if	 	    

    ' Set SpreadSheet #4
   	if (strSPD = "D" or strSPD = "ALL") then        
		With Frm1.vspdData3
	
		    ggoSpread.Source = Frm1.vspdData3
			ggoSpread.Spreadinit "V20021127",,parent.gAllowDragDropSpread    
		   .ReDraw = false
		   .MaxCols = C_EMPTY_3 + 1                                                   '☜: Add 1 to Maxcols
		   .Col = .MaxCols                                                            '☜: Hide maxcols
		   .ColHidden = True                                                          '☜:
    
		   .MaxRows = 0	
			Call GetSpreadColumnPos("D")
				ggoSpread.SSSetEdit      C_RECORD_TYPE3,                "레코드구분",           12
				ggoSpread.SSSetEdit      C_DATA_TYPE3,                  "자료구분",             10
				ggoSpread.SSSetEdit      C_TAX3,                        "세무서",               8
				ggoSpread.SSSetEdit      C_NO3,                         "일련번호",             10
				ggoSpread.SSSetEdit      C_OWN_RGST_NO3,                "사업자등록번호",       16
				ggoSpread.SSSetEdit      C_EMPTY3,                      "공란",                 8
				ggoSpread.SSSetEdit      C_RES_NO_V3,                   "소득자주민등록번호",   20
				ggoSpread.SSSetEdit      C_HFA040T_A_COMP_NM3,          "법인명(상호)",         13
				ggoSpread.SSSetEdit      C_HFA040T_A_COMP_NO3,          "사업자등록번호",       16
				ggoSpread.SSSetEdit      C_HFA040T_A_PAY_TOT3,          "급여총액",             10
				ggoSpread.SSSetEdit      C_HFA040T_A_BONUS_TOT3,        "상여총액",             10
				ggoSpread.SSSetEdit      C_HFA040T_A_AFTER_BONUS_AMT3,  "인정상여",             10
				ggoSpread.SSSetEdit      C_PAY_TOT3,                    "계",                   6
				ggoSpread.SSSetEdit      C_WORK_NO3,                    "종(전)근무처일련번호", 21
				ggoSpread.SSSetEdit      C_EMPTY_3,                     "공란",                 8
		     call ggoSpread.SSSetColHidden(C_NO3,C_NO3,true)	
		   .ReDraw = true
	
		   lgActiveSpd = "D"
		   Call SetSpreadLock 
    
		End With
    End if		
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

							C_RECORD_TYPE     = iCurColumnPos(1)										'☆: Spread Sheet 의 Columns 인덱스 
				C_DATA_TYPE       = iCurColumnPos(2)
				C_TAX             = iCurColumnPos(3)
				C_PROV_DT         = iCurColumnPos(4)
				C_P_TYPE          = iCurColumnPos(5)
				C_MAG_NO          = iCurColumnPos(6)
				C_HOMETAX_ID	  = iCurColumnPos(7)
				C_TAX_CODE	      = iCurColumnPos(8)						
				C_OWN_RGST_NO     = iCurColumnPos(9)
				C_CUST_NM_FULL    = iCurColumnPos(10)
				C_WORKER_DEPT     = iCurColumnPos(11)
				C_WORKER_NM       = iCurColumnPos(12)
				C_WORKER_TEL      = iCurColumnPos(13)
				C_B_COUNT         = iCurColumnPos(14)
				C_KR_CODE         = iCurColumnPos(15)
				C_TERM_CODE       = iCurColumnPos(16)
				C_EMPTY			  = iCurColumnPos(17)         
       Case "B"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
				C_RECORD_TYPE1          = iCurColumnPos(1)								'☆: Spread Sheet 의 Columns 인덱스 
				C_DATA_TYPE1            = iCurColumnPos(2)
				C_TAX1                  = iCurColumnPos(3)
				C_NO1                   = iCurColumnPos(4)
				C_OWN_RGST_NO1          = iCurColumnPos(5)
				C_CUST_NM_FULL1         = iCurColumnPos(6)
				C_REPRE_NM1             = iCurColumnPos(7)
				C_BCA010T_REPRE_RGST_NO1 = iCurColumnPos(8)
				C_COM_NO1               = iCurColumnPos(9)
				C_OLD_COM_NO1           = iCurColumnPos(10)
				C_TOT_PROV_AMT1         = iCurColumnPos(11)
				C_DECI_INCOME_TAX1      = iCurColumnPos(12)
				C_TOT_TAX1              = iCurColumnPos(13)
				C_DECI_RES_TAX1         = iCurColumnPos(14)
				C_DECI_FARM_TAX1        = iCurColumnPos(15)
				C_DECI_SUM1             = iCurColumnPos(16)
				C_EMPTY1               = iCurColumnPos(17)     

       Case "C"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
			C_RECORD_TYPE2               = iCurColumnPos(1)							'☆: SPREAD SHEET 의 COLUMNS 인덱스 
			C_DATA_TYPE2                 = iCurColumnPos(2)
			C_TAX2                       = iCurColumnPos(3)
			C_NO2                        = iCurColumnPos(4)
			C_OWN_RGST_NO2               = iCurColumnPos(5)
			C_OLD_COM_NO2                = iCurColumnPos(6)
			C_HDF020T_RES_FLAG2          = iCurColumnPos(7)
			C_HAA010T_NAT_CD2            = iCurColumnPos(8)
			C_FOREIN_TAXRATE             = iCurColumnPos(9)
			C_HAA010T_ENTR_DT2           = iCurColumnPos(10)
			C_HAA010T_RETIRE_DT2         = iCurColumnPos(11)
			C_HAA010T_NAME2              = iCurColumnPos(12)
			C_FOR_TYPE2                  = iCurColumnPos(13)
			C_RES_NO2                    = iCurColumnPos(14)
			C_START_DT2                  = iCurColumnPos(15)
			C_END_DT2                    = iCurColumnPos(16)
			C_HFA050T_NEW_PAY_TOT2       = iCurColumnPos(17)
			C_HFA050T_NEW_BONUS_TOT2     = iCurColumnPos(18)
			C_HFA030T_AFTER_BONUS_AMT2   = iCurColumnPos(19)
			C_NEW_TOT2                   = iCurColumnPos(20)
			C_HFA050T_NON_TAX52          = iCurColumnPos(21)
			C_HFA050T_NON_TAX12          = iCurColumnPos(22)
			C_NON_TAX2                   = iCurColumnPos(23)
			C_NON_TAX_SUM2               = iCurColumnPos(24)
			C_HFA050T_INCOME_TOT_AMT2    = iCurColumnPos(25)
			C_HFA050T_INCOME_SUB_AMT2    = iCurColumnPos(26)
			C_HFA050T_INCOME_AMT2        = iCurColumnPos(27)
			C_HFA050T_PER_SUB_AMT2       = iCurColumnPos(28)
			C_HFA050T_SPOUSE_SUB_AMT2    = iCurColumnPos(29)
			C_HFA050T_SUPP_CNT2          = iCurColumnPos(30)
			C_HFA050T_SUPP_SUB_AMT2      = iCurColumnPos(31)
			C_HFA050T_OLD_CNT2           = iCurColumnPos(32)
			C_HFA050T_OLD_SUB_AMT2       = iCurColumnPos(33)
			C_HFA050T_PARIA_CNT2         = iCurColumnPos(34)
			C_HFA050T_PARIA_SUB_AMT2     = iCurColumnPos(35)
			C_HFA050T_LADY_SUB_AMT2      = iCurColumnPos(36)
			C_HFA050T_CHL_REAR2          = iCurColumnPos(37)
			C_HFA050T_CHL_REAR_SUB_AMT2  = iCurColumnPos(38)
			C_HFA050T_SMALL_SUB_AMT2     = iCurColumnPos(39)
			C_HFA050T_NAT_PEN_SUB_AMT2   = iCurColumnPos(40)
			C_HFA050T_INSUR_SUB_AMT2     = iCurColumnPos(41)
			C_HFA050T_MED_SUB_AMT2       = iCurColumnPos(42)
			C_HFA050T_EDU_SUB_AMT2       = iCurColumnPos(43)
			C_HFA050T_HOUSE_FUND_AMT2    = iCurColumnPos(44)
			C_HFA050T_CONTR_SUB_AMT2     = iCurColumnPos(45)
			C_HFA050T_CEREMONTY_AMT		 = iCurColumnPos(46)
			C_HFA050T_STD_SUB_TOT_AMT2   = iCurColumnPos(47)
			C_HFA050T_STD_SUB_AMT2       = iCurColumnPos(48)
			C_HFA050T_SUB_INCOME_AMT2    = iCurColumnPos(49)
			C_HFA050T_INDIV_ANU_AMT2     = iCurColumnPos(50)
			C_HFA050T_INDIV_ANU2_AMT2    = iCurColumnPos(51)
			C_HFA050T_INVEST_SUB_SUM_AMT2= iCurColumnPos(52)
			C_HFA050T_CARD_SUB_SUM_AMT2  = iCurColumnPos(53)
			C_HFA050T_OUR_STOCK_AMT2     = iCurColumnPos(54)
			C_SPECIAL_tAX_SUM			 = iCurColumnPos(55)
			C_HFA050T_TAX_STD_AMT2       = iCurColumnPos(56)
			C_HFA050T_CALU_TAX_AMT2      = iCurColumnPos(57)
			C_HFA050T_INCOME_REDU_AMT2   = iCurColumnPos(58)
			C_HFA050T_TAXES_REDU_AMT2    = iCurColumnPos(59)
			C_EMPTY2                     = iCurColumnPos(60)
			C_HFA050T_REDU_SUM_AMT2      = iCurColumnPos(61)
			C_HFA050T_INCOME_TAX_SUB_AMT2= iCurColumnPos(62)
			C_EMPTY4                     = iCurColumnPos(63)
			C_HFA050T_HOUSE_REPAY_AMT2   = iCurColumnPos(64)
			C_HFA050T_FORE_PAY_AMT2      = iCurColumnPos(65)
			C_HFA050T_POLI_TAX_SUB	     = iCurColumnPos(66)
			C_EMPTY_32                   = iCurColumnPos(67)
			C_HFA050T_TAX_SUB_AMT2       = iCurColumnPos(68)
			C_HFA050T_DEC_INCOME_TAX_AMT2= iCurColumnPos(69)
			C_HFA050T_DEC_RES_TAX_AMT2   = iCurColumnPos(70)
			C_HFA050T_DEC_FARM_TAX_AMT2  = iCurColumnPos(71)
			C_DEC_TOT2                   = iCurColumnPos(72)
			C_HFA050T_OLD_INCOME_TAX_AMT2= iCurColumnPos(73)
			C_HFA050T_OLD_RES_TAX_AMT2   = iCurColumnPos(74)
			C_HFA050T_OLD_FARM_TAX_AMT2  = iCurColumnPos(75)
			C_OLD_TOT2                   = iCurColumnPos(76)
			C_HFA050T_NEW_INCOME_TAX_AMT2= iCurColumnPos(77)
			C_HFA050T_NEW_RES_TAX_AMT2   = iCurColumnPos(78)
			C_HFA050T_NEW_FARM_TAX_AMT2  = iCurColumnPos(79)
			C_NEW_SUM_TOT2               = iCurColumnPos(80)
			C_EMPTY_2                  = iCurColumnPos(81)
			C_EMP_NO_2                  = iCurColumnPos(82)

       Case "D"
            ggoSpread.Source = frm1.vspdData3
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
				C_RECORD_TYPE3                = iCurColumnPos(1)	    						'☆: SPREAD SHEET 의 COLUMNS 인덱스 
				C_DATA_TYPE3                  = iCurColumnPos(2)
				C_TAX3                        = iCurColumnPos(3)
				C_NO3                         = iCurColumnPos(4)
				C_OWN_RGST_NO3                = iCurColumnPos(5)
				C_EMPTY3                      = iCurColumnPos(6)
				C_RES_NO_V3                   = iCurColumnPos(7)
				C_HFA040T_A_COMP_NM3          = iCurColumnPos(8)
				C_HFA040T_A_COMP_NO3          = iCurColumnPos(9)
				C_HFA040T_A_PAY_TOT3          = iCurColumnPos(10)
				C_HFA040T_A_BONUS_TOT3        = iCurColumnPos(11)
				C_HFA040T_A_AFTER_BONUS_AMT3  = iCurColumnPos(12)
				C_PAY_TOT3                    = iCurColumnPos(13)
				C_WORK_NO3                    = iCurColumnPos(14)
				C_EMPTY_3                   = iCurColumnPos(15)
    End Select  
       
End Sub
'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock()
    If Trim(lgActiveSpd) = "" Then
       lgActiveSpd = "A"
    End If

    Select Case UCase(Trim(lgActiveSpd))
        Case  "A"
            With frm1 
            .vspdData.ReDraw = False
                ggoSpread.SpreadLock      -1,-1,-1
                ggoSpread.SSSetProtected  .vspdData.MaxCols   , -1, -1
            .vspdData.ReDraw = True
           End With
        Case  "B"
            With frm1
            .vspdData1.ReDraw = False
               ggoSpread.SpreadLock      -1,-1,-1
               ggoSpread.SSSetProtected  .vspdData1.MaxCols   , -1, -1
            .vspdData1.ReDraw = True
            End With
        Case  "C"
            With frm1    
              .vspdData2.ReDraw = False
                ggoSpread.SpreadLock    -1,-1,-1
                ggoSpread.SSSetProtected  .vspdData2.MaxCols   , -1, -1
              .vspdData2.ReDraw = True
            End With
        Case  "D"
            With frm1
              .vspdData3.ReDraw = False
                ggoSpread.SpreadLock     -1,-1,-1
                ggoSpread.SSSetProtected  .vspdData3.MaxCols   , -1, -1
              .vspdData3.ReDraw = True
            End With
    End Select                    
End Sub
'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)

End Sub
'======================================================================================================
' Function Name : SubSetErrPos
' Function Desc : This method set focus to pos of err
'======================================================================================================
Sub SubSetErrPos(iPosArr)
    Dim iDx
    Dim iRow
    iPosArr = Split(iPosArr,parent.gColSep)
    If IsNumeric(iPosArr(0)) Then
       iRow = CInt(iPosArr(0))
       For iDx = 1 To  frm1.vspdData1.MaxCols - 1
           Frm1.vspdData1.Col = iDx
           Frm1.vspdData1.Row = iRow
           If Frm1.vspdData1.ColHidden <> True And Frm1.vspdData1.BackColor <> parent.UC_PROTECTED Then
              Frm1.vspdData1.Col    = iDx
              Frm1.vspdData1.Row    = iRow
              Frm1.vspdData1.Action = 0 ' go to 
              Exit For
           End If           
       Next          
    End If   
End Sub
'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format		

    Call ggoOper.FormatField(Document, "A",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field

    Call InitSpreadSheet("ALL")                                                             'Setup the Spread sheet

    Call InitVariables                                                              'Initializes local global variables

	frm1.txtDt.focus 											'⊙: Set ToolBar    
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
    
    FncQuery = False                                                            '☜: Processing is NG    
    Err.Clear                                                                   '☜: Protect system from crashing

    ggoSpread.Source = Frm1.vspdData
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgbox("900013", parent.VB_YES_NO,"X","X")			        '☜: "Will you destory previous data"		
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If    
	
    Call ggoOper.ClearField(Document, "2")										'⊙: Clear Contents  Field
    Call InitVariables															'⊙: Initializes local global variables
    
    If Not chkField(Document, "1") Then									        '⊙: This function check indispensable field
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
       
    FncQuery = True																'☜: Processing is OK
   
End Function	
'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
    Dim IntRetCD 
    
    FncNew = False																 '☜: Processing is NG
    Err.Clear                                                                    '☜: Clear err status
    
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
    
    FncDelete = True                                                             '☜: Processing is OK
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
    Dim IntRetCD 
    Dim strWhere    
    FncSave = False                                                              '☜: Processing is NG    
    Err.Clear                                                                    '☜: Clear err status

    If DbSave = False Then
		Exit Function
	End If
    FncSave = True                                                                   '☜: Processing is OK    
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
	Call Parent.FncPrint()                                                       '☜: Protect system from crashing
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
	ggoSpread.Source = Frm1.vspdData1
    If lgBlnFlgChgValue = True Or ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgbox("900016", parent.VB_YES_NO,"X","X")			 '⊙: Data is changed.  Do you want to exit? 
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If
    FncExit = True

End Function
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
    select case gActiveSpdSheet.id
		case "vaSpread"
			Call InitSpreadSheet("A")      
		case "vaSpread1"
			Call InitSpreadSheet("B")      		
		case "vaSpread2"
			Call InitSpreadSheet("C")      		
		case "vaSpread3"
			Call InitSpreadSheet("D")      		
	end select      
    
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    Dim strVal
    Dim strEmpno
    Dim strNo
    Dim i
    Err.Clear                                                                        '☜: Clear err status

    DbQuery = False                                                                  '☜: Processing is NG
    
    Call LayerShowHide(1)

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                         '☜: Query
    strVal = strVal     & "&lgCurrentSpd="       & lgCurrentSpd                      '☜: Next key tag
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key

    If lgCurrentSpd = "A" Then
	    strVal = strVal     & "&lgStrPrevKey="       & lgStrPrevKey
    elseIf lgCurrentSpd = "B" Then
	    strVal = strVal     & "&lgStrPrevKey1="       & lgStrPrevKey1
    elseIf lgCurrentSpd = "C" Then
	    strVal = strVal     & "&lgStrPrevKey2="       & lgStrPrevKey2
    elseIf lgCurrentSpd = "D" Then        
	    strVal = strVal     & "&lgStrPrevKey3="       & lgStrPrevKey3
          For i = 1 to frm1.vspdData2.MaxRows
            Frm1.vspdData2.Col = C_OLD_COM_NO2
            Frm1.vspdData2.Row = i
            If Cdbl(Frm1.vspdData2.Value) > 0 Then
                Frm1.vspdData2.Col = C_EMP_NO_2
                 strEmpno = strEmpno & Frm1.vspdData2.Value & parent.gColSep
                Frm1.vspdData2.Col = C_NO2
                strNo = strNo & Frm1.vspdData2.Value & parent.gColSep                
            End If
        Next
        strVal = strVal & "&C_EMP_NO_2=" & strEmpno
        strVal = strVal & "&C_NO2=" & strNo
    End If
    Call RunMyBizASP(MyBizASP, strVal)                                               '☜:  Run biz logic
	
    DbQuery = True                                                                   '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbSave
' Desc : This function is called by FncSave
'========================================================================================================
Function DbSave()
    Dim lRow        
    Dim lGrpCnt     
    Dim retVal      
    Dim boolCheck   
    Dim lStartRow   
    Dim lEndRow     
    Dim lRestGrpCnt 
	Dim strVal, strDel
	
    Err.Clear                                                                    '☜: Clear err status		
	DbSave = False														         '☜: Processing is NG
		
	Call LayerShowHide(1)		
	Call ExecMyBizASP(frm1, BIZ_PGM_ID)
    DbSave  = True                                                               '☜: Processing is NG
End Function
'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
    Err.Clear                                                                    '☜: Clear err status
		
	DbDelete = False			                                                 '☜: Processing is NG
		
	DbDelete = True                                                              '☜: Processing is OK
	
End Function
'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
	Dim i
    Err.Clear                                                                    '☜: Clear err status

    If lgCurrentSpd = "D" And (frm1.vspdData.MaxRows <= 0 And frm1.vspdData1.MaxRows<= 0 And frm1.vspdData2.MaxRows <= 0 And frm1.vspdData3.MaxRows <= 0) Then
		Call DisplayMsgbox("900014", "X","X","X")			                            '☜: 조회를 먼저하세요		
    End If	
    Call SetToolbar("1100000000011111")
	Call ggoOper.LockField(Document, "Q")
    If lgCurrentSpd = "A" then
		frm1.vspdData.focus
	ElseIf lgCurrentSpd = "B" then
		frm1.vspdData1.focus	
	ElseIf lgCurrentSpd = "C" then
		frm1.vspdData2.focus	
	ElseIf lgCurrentSpd = "D" then
		frm1.vspdData3.focus	
	End if
End Function
	
'========================================================================================================
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()
	Call InitVariables
End Function	
'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()

End Function

'======================================================================================================
'	Name : OpenCode()
'	Description : Code PopUp at vspdData
'=======================================================================================================
Function OpenCode(Byval strCode, Byval iWhere, ByVal Row)
	Dim arrRet
	Dim arrParam(5), arrField(6), arrHeader(6)

	If IsOpenPop = True Then Exit Function

	IsOpenPop = True

	Select Case iWhere
	    Case 0
	        arrParam(0) = "수당코드 팝업"			        ' 팝업 명칭 
	        arrParam(1) = "HDA010T"				 		        ' TABLE 명칭 
	        arrParam(2) = ""                		            ' Code Condition
	        arrParam(3) = strCode						        ' Name Cindition
	        arrParam(4) = " PAY_CD = " & FilterVar("*", "''", "S") & "  AND CODE_TYPE = " & FilterVar("1", "''", "S") & "  "  ' Where Condition
	        arrParam(5) = "수당코드"			            ' TextBox 명칭 
	
            arrField(0) = "ALLOW_CD"					        ' Field명(0)
            arrField(1) = "ALLOW_NM"				            ' Field명(1)
    
            arrHeader(0) = "수당코드"				        ' Header명(0)
            arrHeader(1) = "수당코드명"
	    Case 1
	        arrParam(0) = "상여구분코드 팝업"			    ' 팝업 명칭 
	    	arrParam(1) = "b_minor"	    						' TABLE 명칭 
	    	arrParam(2) = ""                  		        	' Code Condition
	        arrParam(3) = strCode						        ' Name Cindition
	    	arrParam(4) = " major_cd=" & FilterVar("h0040", "''", "S") & " "                  ' Where Condition
	    	arrParam(5) = "상여코드" 			            ' TextBox 명칭 
	
	    	arrField(0) = "minor_cd"					    	' Field명(0)
	    	arrField(1) = "minor_nm"    				    	' Field명(1)
    
	    	arrHeader(0) = "상여코드"	   		    	    ' Header명(0)
	    	arrHeader(1) = "상여코드명"	   		            ' Header명(1)
	End Select
    
	arrRet = window.showModalDialog("../../comasp/CommonPopup.asp", Array(arrParam, arrField, arrHeader), _
		"dialogWidth=420px; dialogHeight=450px; center: Yes; help: No; resizable: No; status: No;")
	
	IsOpenPop = False
	
	If arrRet(0) = "" Then
		Exit Function
	Else
		Call SetCode(arrRet, iWhere)
 	    If iWhere = 0 Then
           	ggoSpread.Source = frm1.vspdData
            ggoSpread.UpdateRow Row
        Else
           	ggoSpread.Source = frm1.vspdData1
            ggoSpread.UpdateRow Row
        End If
	End If	

End Function

'======================================================================================================
'	Name : SetCode()
'	Description : Code PopUp에서 Return되는 값 setting
'=======================================================================================================
Function SetCode(Byval arrRet, Byval iWhere)
	    With frm1
	    	Select Case iWhere
	    	    Case 0
	    	        .vspdData.Col = C_ALLOW_CD
	    	    	.vspdData.text = arrRet(0) 
	    	    	.vspdData.Col = C_ALLOW_NM
	    	    	.vspdData.text = arrRet(1)   
            End Select
	    	Select Case iWhere
	    	    Case 1
	    	        .vspdData1.Col = C_BONUS_TYPE
	    	    	.vspdData1.text = arrRet(0) 
	    	    	.vspdData1.Col = C_BONUS_TYPE_NM
	    	    	.vspdData1.text = arrRet(1)   
            End Select
	    End With

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

Sub vspdData1_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111")

    gMouseClickStatus = "SP1C"   
    Set gActiveSpdSheet = frm1.vspdData1
    
    If frm1.vspdData1.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
     If Row = 0 Then
        ggoSpread.Source = frm1.vspdData1
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If
    End If
	frm1.vspdData1.Row = Row 
End Sub
Sub vspdData2_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111")

    gMouseClickStatus = "SP2C"   
    Set gActiveSpdSheet = frm1.vspdData2
    
    If frm1.vspdData2.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
     If Row = 0 Then
        ggoSpread.Source = frm1.vspdData2
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If
    End If
	frm1.vspdData2.Row = Row 
End Sub

Sub vspdData3_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111")

    gMouseClickStatus = "SP3C"   
    Set gActiveSpdSheet = frm1.vspdData3
    
    If frm1.vspdData3.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
     If Row = 0 Then
        ggoSpread.Source = frm1.vspdData3
        If lgSortKey = 1 Then
            ggoSpread.SSSort
            lgSortKey = 2
        Else
            ggoSpread.SSSort ,lgSortKey
            lgSortKey = 1
        End If
    End If
	frm1.vspdData3.Row = Row 
End Sub
'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub
Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub
Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SSSetColWidth(pvCol1,pvCol2)

End Sub
Sub vspdData3_ColWidthChange(ByVal pvCol1, ByVal pvCol2)		
    ggoSpread.Source = frm1.vspdData3
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
			
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If  
End Sub
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	If frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
		If lgStrPrevKey1 <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			topleftOK = true	
			lgCurrentSpd = "B"		

			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If  
End Sub
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	If frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then
		If lgStrPrevKey2 <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			topleftOK = true	
			lgCurrentSpd = "C"		
			
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
		End If
	End If  
End Sub
Sub vspdData3_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )
   	If OldLeft <> NewLeft Then
		Exit Sub
	End If
	If frm1.vspdData3.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData3,NewTop) Then
		If lgStrPrevKey3 <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	
			
			Call DisableToolBar(Parent.TBC_QUERY)
			lgCurrentSpd = "D"					
			If DBQuery = False Then
				Call RestoreToolBar()
				Exit Sub
			End If
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
Sub vspdData1_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
		Exit Sub
	End if
	
	If frm1.vspdData1.MaxRows = 0 then
		Exit Sub
	End if
	
End Sub
Sub vspdData2_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
		Exit Sub
	End if
	
	If frm1.vspdData2.MaxRows = 0 then
		Exit Sub
	End if
	
End Sub
Sub vspdData3_DblClick(ByVal Col, ByVal Row)				
    Dim iColumnName
    
    If Row <= 0 Then
		Exit Sub
	End if
	
	If frm1.vspdData3.MaxRows = 0 then
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
Sub vspdData1_GotFocus()
    ggoSpread.Source = Frm1.vspdData1
End Sub
Sub vspdData2_GotFocus()
    ggoSpread.Source = Frm1.vspdData2
End Sub
Sub vspdData3_GotFocus()
    ggoSpread.Source = Frm1.vspdData3
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
Sub vspdData1_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    
    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("B")
End Sub
Sub vspdData2_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    
    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("C")
End Sub
Sub vspdData3_ScriptDragDropBlock( Col ,  Row,  Col2,  Row2,  NewCol,  NewRow,  NewCol2,  NewRow2,  Overwrite , Action , DataOnly , Cancel )
    
    ggoSpread.Source = frm1.vspdData3
    Call ggoSpread.SpreadDragDropBlock(Col , Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite , Action , DataOnly , Cancel )    
    Call GetSpreadColumnPos("D")
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
Sub vspdData1_MouseDown(Button , Shift , x , y)

       If Button = 2 And gMouseClickStatus = "SP1C" Then
          gMouseClickStatus = "SP1CR"
        End If
End Sub    
Sub vspdData2_MouseDown(Button , Shift , x , y)

       If Button = 2 And gMouseClickStatus = "SP2C" Then
          gMouseClickStatus = "SP2CR"
        End If
End Sub    
Sub vspdData3_MouseDown(Button , Shift , x , y)

       If Button = 2 And gMouseClickStatus = "SP3C" Then
          gMouseClickStatus = "SP3CR"
        End If
End Sub    

'========================================================================================================
'   Event Name : vspdData_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	    With Frm1.vspdData
	    	ggoSpread.Source = Frm1.vspdData
	    	If Row > 0 Then
	    		Select Case Col
	    		       Case C_ALLOW_NM_POP
	    		        	.Col = Col - 1
	    		        	.Row = Row
	    		        	Call OpenCode(.Text,0,Row)
	    		End Select
	    	End If
    
	    End With
End Sub
'========================================================================================================
'   Event Name : vspdData1_ButtonClicked
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_ButtonClicked(ByVal Col, ByVal Row, Byval ButtonDown)
	    With Frm1.vspdData1
	    	ggoSpread.Source = Frm1.vspdData1
	    	If Row > 0 Then
	    		Select Case Col
	    		       Case C_BONUS_TYPE_NM_POP
	    		        	.Col = Col - 1
	    		        	.Row = Row
	    		        	Call OpenCode(.Text,1,Row)
	    		End Select
	    	End If
    
	    End With
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
'========================================================================================================
'   Event Name : vspdData1_Change
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_Change( Col ,  Row)

    Dim iDx
    Dim IntRetCD
    Frm1.vspdData1.Row = Row
    Frm1.vspdData1.Col = Col

    If Frm1.vspdData1.CellType = parent.SS_CELL_TYPE_FLOAT Then
       If UNICDbl(Frm1.vspdData1.text) < CDbl(Frm1.vspdData1.TypeFloatMin) Then
          Frm1.vspdData1.text = Frm1.vspdData1.TypeFloatMin
       End If
    End If
	
    ggoSpread.Source = frm1.vspdData1
    ggoSpread.UpdateRow Row

End Sub
'=======================================================================================================
'   Event Name : txtDt_DblClick(Button)
'   Event Desc : 달력을 호출한다.
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
'   Event Desc : 달력을 호출한다.
'=======================================================================================================
Sub txtBas_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtBas_dt.Action = 7
        frm1.txtBas_dt.focus
    End If
End Sub
'======================================================================================================
' Function Name : btnCb_print2_onClick
' Function Desc : 플로피디스켓, 신고 공문 출력 
'=======================================================================================================
Sub btnCb_print2_onClick()
Dim RetFlag

    If frm1.vspdData.MaxRows <= 0 And frm1.vspdData1.MaxRows<= 0 And frm1.vspdData2.MaxRows <= 0 And frm1.vspdData3.MaxRows <= 0 Then
		Call DisplayMsgbox("800167", "X","X","X")			                            '☜: 조회를 먼저하세요 
		Exit Sub
    End If
    
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Sub
    End If
    
    RetFlag = DisplayMsgbox("900018", parent.VB_YES_NO,"x","x")                                '☜ 작업을 계속하시겠습니까?
	If RetFlag = VBNO Then
		Exit Sub
    Else
        Call FloppyDiskLabelForm()      '플로피디스켓 라벨양식 
        Call ReportOfDocuments()        '신고 공문 
	End IF
        

    
End Sub
'======================================================================================================
' Function Name : btnCb_print_onClick
' Function Desc : 집계표 출력 
'=======================================================================================================
Sub btnCb_print_onClick()
Dim RetFlag

    If frm1.vspdData.MaxRows <= 0 And frm1.vspdData1.MaxRows<= 0 And frm1.vspdData2.MaxRows <= 0 And frm1.vspdData3.MaxRows <= 0 Then
		Call DisplayMsgbox("800167", "X","X","X")			                            '☜: 조회를 먼저하세요 
		Exit Sub
    End If
    	
    If Not chkField(Document, "1") Then									'⊙: This function check indispensable field
       Exit Sub
    End If
    
    RetFlag = DisplayMsgbox("900018", parent.VB_YES_NO,"x","x")                                '☜ 작업을 계속하시겠습니까?
	If RetFlag = VBNO Then
		Exit Sub
	End IF
    
    Call FncBtnPrint() 
End Sub
'======================================================================================================
' Function Name : FncBtnPrint
' Function Desc : 집계표 출력 
'=======================================================================================================
Function FncBtnPrint() 
	Dim strUrl
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile    
	Dim objName

	dim bas_dt, bas_yy, biz_area_cd, present_dt
	
	StrEbrFile = "h9120oa1_1"
	bas_dt = UniConvDateAToB(frm1.txtbas_dt.text,parent.gDateFormat, parent.gServerDateFormat)
	bas_dt = replace(bas_dt,parent.gServerDateFormat,"")
	bas_yy = frm1.txtBas_dt.year
	biz_area_cd = frm1.txtComp_cd.value
	present_dt = UniConvDateAToB(frm1.txtDt.text,parent.gDateFormat, parent.gServerDateFormat)
	present_dt = replace(present_dt,parent.gServerDateFormat,"")

	strUrl = strUrl & "bas_dt|" & bas_dt
	strUrl = strUrl & "|bas_yy|" & bas_yy 
	strUrl = strUrl & "|year_area_cd|" & biz_area_cd
	strUrl = strUrl & "|present_dt|" & present_dt
	
	objname = AskEBDocumentName(StrEbrFile,"EBR")
	Call FncEBRPrint(EBAction,objname,strUrl)
End Function
'======================================================================================================
' Function Name : FloppyDiskLabelForm
' Function Desc : 플로피디스켓 라벨양식 
'=======================================================================================================
Function FloppyDiskLabelForm()
	Dim strUrl	
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile
	Dim objName

	dim bas_dt, bas_yy, biz_area_cd
	
	StrEbrFile = "h9120oa1_2"	
	
	bas_dt = UniConvDateAToB(frm1.txtbas_dt.text,parent.gDateFormat, parent.gServerDateFormat)
	bas_dt = replace(bas_dt,parent.gServerDateFormat,"")
	bas_yy = frm1.txtBas_dt.year
	biz_area_cd = frm1.txtComp_cd.value	

	strUrl = strUrl & "bas_dt|" & bas_dt
	strUrl = strUrl & "|bas_yy|" & bas_yy 
	strUrl = strUrl & "|biz_area_cd|" & biz_area_cd		
	
	objname = AskEBDocumentName(StrEbrFile,"EBR")
	Call FncEBRPrint(EBAction,objname,strUrl)
	
End Function
'======================================================================================================
' Function Name : ReportOfDocuments
' Function Desc : 신고 공문 
'=======================================================================================================
Function ReportOfDocuments()
	Dim strUrl
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile
	Dim objName     	

	dim bas_dt, bas_yy, biz_area_cd, present_dt
	
	StrEbrFile = "h9120oa1_3"
	
	bas_dt = UniConvDateAToB(frm1.txtbas_dt.text,parent.gDateFormat, parent.gServerDateFormat)
	bas_dt = replace(bas_dt,parent.gServerDateFormat,"")
	bas_yy = Year(frm1.txtBas_dt.Text)
	biz_area_cd = frm1.txtComp_cd.value
	present_dt = UniConvDateAToB(frm1.txtDt.text, parent.gDateFormat, parent.gServerDateFormat)
	present_dt = replace(present_dt,parent.gServerDateFormat,"")

	strUrl = strUrl & "bas_dt|" & bas_dt
	strUrl = strUrl & "|bas_yy|" & bas_yy 
	strUrl = strUrl & "|biz_area_cd|" & biz_area_cd	
	strUrl = strUrl & "|present_dt|" & present_dt
	
	objname = AskEBDocumentName(StrEbrFile,"EBR")
	Call FncEBRPrint(EBAction,objname,strUrl)
End Function
'==========================================================================================
'   Event Name : btnCb_creation_OnClick
'   Event Desc : 파일생성(Server)
'==========================================================================================
Function btnCb_creation_OnClick()
Dim RetFlag
Dim strVal
Dim intRetCD

    Err.Clear                                                                           '☜: Clear err status
    
    If Not chkField(Document, "1") Then                                                 'Required로 표시된 Element들의 입력 [유/무]를 Check 한다.
       Exit Function                            
    End If
    
    If frm1.vspdData.MaxRows <= 0 And frm1.vspdData1.MaxRows<= 0 And frm1.vspdData2.MaxRows <= 0 And frm1.vspdData3.MaxRows <= 0 Then
		Call DisplayMsgbox("800167", "X","X","X")			                            '☜: 조회를 먼저하세요 
		Exit Function		
    End If
 
	RetFlag = DisplayMsgbox("900018", parent.VB_YES_NO,"x","x")                                '☜ 작업을 계속하시겠습니까?
	If RetFlag = VBNO Then
		Exit Function
	End IF

    With frm1
        Call LayerShowHide(1)					 
        lgCurrentSpd = "A"		
	    Call MakeKeyStream(lgCurrentSpd)    
	    strVal = BIZ_PGM_ID2    & "?txtMode="           & parent.UID_M0001						'☜: 비지니스 처리 ASP의 상태 	    	    		    
        strVal = strVal         & "&lgCurrentSpd="      & lgCurrentSpd                  '☜: Mulit의 종류 
        strVal = strVal         & "&txtKeyStream="      & lgKeyStream                   '☜: Query Key	
	   
		Call RunMyBizASP(MyBizASP, strVal)
	
    End With    
End Function
'==========================================================================================
'   Event Name : subVatDiskOK
'   Event Desc : 파일생성(Client)
'==========================================================================================
Function subVatDiskOK(ByVal pFileName) 
Dim strVal
    Err.Clear                                                                           '☜: server에 만들어진 file이름 
    If Trim(pFileName) <> "" Then
	    strVal = BIZ_PGM_ID2 & "?txtMode=" & parent.UID_M0002							        '☜: 비지니스 처리 ASP의 상태 
	    strVal = strVal & "&txtFileName=" & pFileName							        '☜: 조회 조건 데이타	
	    Call RunMyBizASP(MyBizASP, strVal)										        '☜: 비지니스 ASP 를 가동 
    End If
End Function


'=======================================================================================================
'   Event Name : txtDt_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
'=======================================================================================================
Sub txtDt_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub


'=======================================================================================================
'   Event Name : txtBas_dt_Keypress(Key)
'   Event Desc : enter key down시에 조회한다.
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>통합연말정산신고</font></td>
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
								<TD CLASS="TD5" NOWRAP>제출자구분</TD>
								<TD CLASS="TD6" NOWRAP><SELECT NAME="txtGubun" ALT="제출자구분" STYLE="WIDTH: 100px" TAG="12N"></SELECT></TD>
								<TD CLASS="TD5" NOWRAP>대상기간</TD>
								<TD CLASS="TD6" NOWRAP><SELECT NAME="txtGigan" ALT="대상기간" STYLE="WIDTH: 170px" TAG="12N"></SELECT></TD>
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>신고사업장</TD>
								<TD CLASS="TD6" NOWRAP><SELECT NAME="txtComp_cd" ALT="신고사업장" STYLE="WIDTH: 150px" TAG="12N"></SELECT></TD>								
								<TD CLASS=TD5  NOWRAP>제출연월일</TD>
								<TD CLASS=TD6  NOWRAP><script language =javascript src='./js/h9120ma1_fpDateTime1_txtDt.js'></script></TD>
							</TR>	
				            <TR>
								<TD CLASS=TD5  NOWRAP>세무대리인관리번호</TD>
								<TD CLASS=TD6  NOWRAP><INPUT TYPE=TEXT ID="txtBizAreaCD" MAXLENGTH=6 NAME="txtSer" SIZE=15 tag="11XXX" ALT="세무대리인관리번호"></TD>								
								<TD CLASS=TD5  NOWRAP>기준연월일</TD>
								<TD CLASS=TD6  NOWRAP><script language =javascript src='./js/h9120ma1_fpDateTime2_txtBas_dt.js'></script></TD>
							</TR>							
								<INPUT TYPE=HIDDEN ID="txtFile" NAME="txtFile" SIZE=15 tag="14XXXU" ALT="저장파일경로">								
						</TABLE>
						</FIELDSET>
					</TD>
				</TR>
				<TR><TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD></TR>
				<TR >
					<TD WIDTH="100%" HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
            			    <TR HEIGHT="25%">
            					<TD WIDTH="50%" HEIGHT=*>
	            					<TABLE WIDTH="100%" HEIGHT="100%">
	            						<TR>
	            							<TD HEIGHT="100%"><script language =javascript src='./js/h9120ma1_vaSpread_vspdData.js'></script></TD>
		               					</TR>
					            	</TABLE>
		            			</TD>
            					<TD WIDTH="50%" HEIGHT=*>
	            					<TABLE WIDTH="100%" HEIGHT="100%">
	            						<TR>
	            							<TD HEIGHT="100%"><script language =javascript src='./js/h9120ma1_vaSpread1_vspdData1.js'></script></TD>
		               					</TR>
					            	</TABLE>
		            			</TD>
                            </TR>  
                            <TR HEIGHT="40%">
				            	<TD WIDTH="100%" HEIGHT=* VALIGN=TOP COLSPAN=3>
				            		<TABLE WIDTH="100%" HEIGHT="100%">
				            			<TR>
				            				<TD HEIGHT="100%"><script language =javascript src='./js/h9120ma1_vaSpread2_vspdData2.js'></script></TD>
					            		</TR>
					            	</TABLE>
					            </TD>
			                </TR>
                            <TR HEIGHT="35%">
				            	<TD WIDTH="100%" HEIGHT=* VALIGN=TOP COLSPAN=3>
				            		<TABLE WIDTH="100%" HEIGHT="100%">
				            			<TR>
				            				<TD HEIGHT="100%"><script language =javascript src='./js/h9120ma1_vaSpread3_vspdData3.js'></script></TD>
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
	                <TD><BUTTON NAME="btnCb_print2" CLASS="CLSMBTN">공문및표지출력</BUTTON>&nbsp;
	                    <BUTTON NAME="btnCb_print" CLASS="CLSMBTN">집계표출력</BUTTON>&nbsp;
	                    <BUTTON NAME="btnCb_creation" CLASS="CLSMBTN">파일생성</BUTTON>&nbsp;일련번호나 공란은 파일을 생성할때 제대로 보여집니다</TD>
	                <TD WIDTH=* ALIGN="right"></TD>
	                <TD WIDTH=10>&nbsp;</TD>
	            </TR>
	        </TABLE>
	    </TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP1" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
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

	
