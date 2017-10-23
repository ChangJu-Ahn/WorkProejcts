<%@ LANGUAGE="VBSCRIPT" %>
<!--
======================================================================================================
*  1. Module Name          : Human Resources
*  2. Function Name        : Multi Sample
*  3. Program ID           : H1a02ma1
*  4. Program Name         : H1a02ma1
*  5. Program Desc         : Multi Sample
*  6. Comproxy List        :
*  7. Modified date(First) : 2001/04/18
*  8. Modified date(Last)  : 2003/06/16
*  9. Modifier (First)     :
* 10. Modifier (Last)      : Lee SiNa
* 11. Comment              :
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
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliRdsQuery.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"> </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incEB.vbs"> </SCRIPT>
<Script Language="VBScript">
Option Explicit 
'========================================================================================================
'=                       4.2 Constant variables
'========================================================================================================
Const CookieSplit = 1233
Const BIZ_PGM_ID    = "ha110mb1.asp"                                      'Biz Logic ASP
Const BIZ_PGM_ID2   = "ha110mb2.asp"                                 '��: File Creation Asp Name
Const C_SHEETMAXROWS            = 21	                                                                '�� ȭ�鿡 �������� �ִ밹��*1.5%>

'========================================================================================================
'=                       4.3 Common variables
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim lsConcd
Dim IsOpenPop
Dim lgStrPrevKey1,lgStrPrevKey2
Dim topleftOK

Dim C_RECORD_TYPE  
Dim C_DATA_TYPE    
Dim C_TAX          
Dim C_PROV_DT      
Dim C_P_TYPE       
Dim C_MAG_NO  
Dim C_HOMETAX_ID
Dim C_TAX_CODE   
Dim C_OWN_RGST_NO  
Dim C_CUST_NM_FULL 
Dim C_WORKER_DEPT      
Dim C_WORKER_NM     
Dim C_WORKER_TEL       
Dim C_B_COUNT      
Dim C_KR_CODE      
Dim C_TERM_CODE    
Dim C_EMPTY        

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

Dim C_RECORD_TYPE2   
Dim C_DATA_TYPE2     
Dim C_TAX2           
Dim C_NO2            
Dim C_OWN_RGST_NO2   
Dim C_OLD_COM_NO2    
Dim C_HDF020T_RES_FLAG2       
Dim C_HAA010T_NAT_CD2          '2002 ���������ڵ� �߰� 
Dim C_HAA010T_ENTR_DT2     
Dim C_HAA010T_RETIRE_DT2   
Dim C_HAA010T_NAME2        
Dim C_FOR_TYPE2            
Dim C_HGA070T_RETIRE_AMT2  
Dim C_HGA070T_HONOR_AMT2   
Dim C_HGA070T_CORP_INSUR2  
Dim C_HGA070T_TOT_PROV_AMT2
Dim C_ENTR_DT2             
Dim C_RETIRE_DT2              
Dim C_HGA070T_TOT_DUTY_MM2 
Dim C_OLD_ENTR_DT2  
Dim C_OLD_RETIRE_DT2
Dim C_OLD_DUTY2     
Dim C_D_DUTY2       
Dim C_HGA070T_DUTY_CNT2

Dim C_H_ENTR_DT2
Dim C_H_RETIRE_DT2
Dim C_H_HGA070T_TOT_DUTY_MM2
Dim C_H_OLD_ENTR_DT2
Dim C_H_OLD_RETIRE_DT2
Dim C_H_OLD_DUTY2
Dim C_H_D_DUTY2
Dim C_H_HGA070T_DUTY_CNT2

Dim C_RETIRE_TOT_PROV_AMT2    
Dim C_HGA070T_INCOME_SUB2     
Dim C_HGA070T_TAX_STD2        
Dim C_HGA070T_AVR_TAX_STD2    
Dim C_HGA070T_AVR_CALC_TAX2   
Dim C_HGA070T_CALC_TAX2       
Dim C_RETIRE_SUB2   
Dim C_DECI_TAX

Dim C_H_RETIRE_TOT_PROV_AMT2
Dim C_H_HGA070T_INCOME_SUB2
Dim C_H_HGA070T_TAX_STD2
Dim C_H_HGA070T_AVR_TAX_STD2
Dim C_H_HGA070T_AVR_CALC_TAX2
Dim C_H_HGA070T_CALC_TAX2
Dim C_H_RETIRE_SUB2
Dim C_H_DECI_TAX

Dim C_T_RETIRE_TOT_PROV_AMT2    
Dim C_T_HGA070T_INCOME_SUB2     
Dim C_T_HGA070T_TAX_STD2        
Dim C_T_HGA070T_AVR_TAX_STD2    
Dim C_T_HGA070T_AVR_CALC_TAX2   
Dim C_T_HGA070T_CALC_TAX2       
Dim C_T_RETIRE_SUB2   
Dim C_T_DECI_TAX
          
Dim C_HGA070T_DECI_INCOME_TAX2
Dim C_HGA070T_DECI_RES_TAX2   
Dim C_DECI_FARM_TAX2          
Dim C_DECI_SUM2               
Dim C_HFA050T_OLD_INCOME_TAX2 
Dim C_HFA050T_OLD_RES_TAX2    
Dim C_HFA050T_OLD_FARM_TAX2   
Dim C_OLD_SUM2  

'========================================================================================================
' Name : initSpreadPosVariables()	
' Desc : Initialize value
'========================================================================================================
Sub initSpreadPosVariables(ByVal pvSpdNo)  

    If pvSpdNo = "A" Then
        C_RECORD_TYPE		= 1
        C_DATA_TYPE			= 2
        C_TAX				= 3
        C_PROV_DT			= 4
        C_P_TYPE			= 5
        C_MAG_NO			= 6
        C_HOMETAX_ID		= 7
        C_TAX_CODE			= 8
        C_OWN_RGST_NO		= 9
        C_CUST_NM_FULL		= 10
        C_WORKER_DEPT		= 11
        C_WORKER_NM			= 12
        C_WORKER_TEL		= 13
        C_B_COUNT 			= 14
        C_KR_CODE 			= 15
        C_TERM_CODE  		= 16
        C_EMPTY  			= 17  

    ElseIf pvSpdNo = "B" Then
        C_RECORD_TYPE1            = 1
        C_DATA_TYPE1              = 2
        C_TAX1                    = 3
        C_NO1                     = 4
        C_OWN_RGST_NO1            = 5
        C_CUST_NM_FULL1           = 6
        C_REPRE_NM1               = 7
        C_BCA010T_REPRE_RGST_NO1  = 8
        C_COM_NO1                 = 9
        C_OLD_COM_NO1             = 10
        C_TOT_PROV_AMT1           = 11
        C_DECI_INCOME_TAX1        = 12
        C_TOT_TAX1                = 13
        C_DECI_RES_TAX1           = 14
        C_DECI_FARM_TAX1          = 15
        C_DECI_SUM1               = 16
        C_EMPTY1                  = 17
 
    ElseIf pvSpdNo = "C" Then
        C_RECORD_TYPE2            = 1
        C_DATA_TYPE2              = 2
        C_TAX2                    = 3
        C_NO2                     = 4
        C_OWN_RGST_NO2            = 5
        C_OLD_COM_NO2             = 6
        C_HDF020T_RES_FLAG2       = 7
        C_HAA010T_NAT_CD2         = 8  '2002 ���������ڵ� �߰� 
        C_HAA010T_ENTR_DT2        = 9
        C_HAA010T_RETIRE_DT2      = 10
        C_HAA010T_NAME2           = 11
        C_FOR_TYPE2               = 12
        C_HGA070T_RETIRE_AMT2     = 13
        C_HGA070T_HONOR_AMT2      = 14
        C_HGA070T_CORP_INSUR2     = 15
        C_HGA070T_TOT_PROV_AMT2   = 16
        C_ENTR_DT2                = 17
        C_RETIRE_DT2              = 18
        C_HGA070T_TOT_DUTY_MM2    = 19
        C_OLD_ENTR_DT2            = 20
        C_OLD_RETIRE_DT2          = 21
        C_OLD_DUTY2               = 22
        C_D_DUTY2                 = 23
        C_HGA070T_DUTY_CNT2       = 24
    
        C_H_ENTR_DT2				= 25
        C_H_RETIRE_DT2				= 26
        C_H_HGA070T_TOT_DUTY_MM2	= 27
        C_H_OLD_ENTR_DT2			= 28
        C_H_OLD_RETIRE_DT2			= 29
        C_H_OLD_DUTY2				= 30
        C_H_D_DUTY2					= 31
        C_H_HGA070T_DUTY_CNT2		= 32

        C_RETIRE_TOT_PROV_AMT2		= 33
        C_HGA070T_INCOME_SUB2		= 34
        C_HGA070T_TAX_STD2			= 35
        C_HGA070T_AVR_TAX_STD2		= 36
        C_HGA070T_AVR_CALC_TAX2		= 37
        C_HGA070T_CALC_TAX2			= 38
        C_RETIRE_SUB2				= 39
		C_DECI_TAX					= 40

        C_H_RETIRE_TOT_PROV_AMT2	= 41
        C_H_HGA070T_INCOME_SUB2		= 42
        C_H_HGA070T_TAX_STD2		= 43
        C_H_HGA070T_AVR_TAX_STD2	= 44
        C_H_HGA070T_AVR_CALC_TAX2	= 45
        C_H_HGA070T_CALC_TAX2		= 46
        C_H_RETIRE_SUB2				= 47
        C_H_DECI_TAX        		= 48
     
        C_T_RETIRE_TOT_PROV_AMT2	= 49
        C_T_HGA070T_INCOME_SUB2		= 50
        C_T_HGA070T_TAX_STD2		= 51
        C_T_HGA070T_AVR_TAX_STD2	= 52
        C_T_HGA070T_AVR_CALC_TAX2	= 53
        C_T_HGA070T_CALC_TAX2		= 54
        C_T_RETIRE_SUB2				= 55
		C_T_DECI_TAX				= 56

        C_HGA070T_DECI_INCOME_TAX2	= 57
        C_HGA070T_DECI_RES_TAX2		= 58
        C_DECI_FARM_TAX2			= 59
        C_DECI_SUM2					= 60
        C_HFA050T_OLD_INCOME_TAX2	= 61
        C_HFA050T_OLD_RES_TAX2		= 62
        C_HFA050T_OLD_FARM_TAX2		= 63
        C_OLD_SUM2					= 64
    End If

End Sub

'========================================================================================================
' Name : InitVariables()
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode      = parent.OPMD_CMODE						        '��: Indicates that current mode is Create mode
	lgBlnFlgChgValue  = False								    '��: Indicates that no value changed
	lgIntGrpCount     = 0										'��: Initializes Group View Size
    lgStrPrevKey      = ""                                      '��: initializes Previous Key
    lgStrPrevKey1       = ""                                     '��: initializes Previous Key
    lgStrPrevKey2       = ""         
    lgSortKey         = 1                                       '��: initializes sort direction
End Sub

'========================================================================================================
' Name : SetDefaultVal()
' Desc : Set default value
'========================================================================================================

Sub SetDefaultVal()
    frm1.txtDt.Text =  UniConvDateAToB("<%=GetSvrDate%>" ,parent.gServerDateFormat,parent.gDateFormat)
  '  frm1.txtStrt_dt.Text = UniConvDateAToB("<%=GetSvrDate%>",parent.gServerDateFormat,parent.gDateFormat)
	frm1.txtBas_dt.Text =  UniConvDateAToB("<%=GetSvrDate%>" ,parent.gServerDateFormat,parent.gDateFormat)
	Dim strYear,strMonth,strDay
    Call ExtractDateFrom("<%=GetsvrDate%>",parent.gServerDateFormat , parent.gServerDateType ,strYear,strMonth,strDay)
 
    frm1.txtStrt_dt.year = strYear
    frm1.txtStrt_dt.month = "1"
    frm1.txtStrt_dt.day = "1" 	
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
Sub MakeKeyStream(pRow)

   Dim strDt, strStrt_dt, strBas_dt
   
   strDt      = UniConvDateToYYYYMMDD(frm1.txtDt, parent.gDateFormat, "")
   strStrt_dt = UNIConvDateCompanyToDB(frm1.txtStrt_dt, parent.gDateFormat)
   strBas_dt  = UNIConvDateCompanyToDB(frm1.txtBas_dt, parent.gDateFormat)

   lgKeyStream       = Trim(Frm1.txtGubun.value) & parent.gColSep						'0 
   lgKeyStream       = lgKeyStream & Trim(frm1.txtGigan.value) & parent.gColSep		'1
   lgKeyStream       = lgKeyStream & Trim(strDt) & parent.gColSep			            '2
   lgKeyStream       = lgKeyStream & Trim(frm1.txtSer.value) & parent.gColSep			'3
   lgKeyStream       = lgKeyStream & Trim(frm1.txtFile.value) & parent.gColSep		    '4
   lgKeyStream       = lgKeyStream & Trim(strStrt_dt) & parent.gColSep		'5
   lgKeyStream       = lgKeyStream & Trim(strBas_dt) & parent.gColSep		'6
   lgKeyStream       = lgKeyStream & Trim(frm1.txtComp_cd.value) & parent.gColSep		'7
   lgKeyStream       = lgKeyStream & Trim(frm1.txtStrt_dt.Year) & parent.gColSep		'8
   lgKeyStream       = lgKeyStream & Trim(frm1.txtBas_dt.Year) & parent.gColSep		'9

End Sub
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()

    Dim iNameArr , iNameArr1 , iNameArr2
    Dim iCodeArr , iCodeArr1 , iCodeArr2
'������ ���� 
    Call CommonQueryRs("MINOR_NM,MINOR_CD","B_MINOR","MAJOR_CD = " & FilterVar("H0118", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iNameArr = lgF0
    iCodeArr = lgF1
    Call SetCombo2(frm1.txtGubun,iCodeArr,iNameArr,Chr(11))
    frm1.txtGubun.value = 2      
    '���Ⱓ 
    Call CommonQueryRs("MINOR_NM,MINOR_CD","B_MINOR","MAJOR_CD = " & FilterVar("H0119", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iNameArr1 = lgF0
    iCodeArr1 = lgF1
    Call SetCombo2(frm1.txtGigan,iCodeArr1,iNameArr1,Chr(11))            ''''''''DB���� �ҷ� condition���� 
    '�Ű����� 
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
Sub InitSpreadSheet(ByVal pvSpdNo)

    If pvSpdNo = "" OR pvSpdNo = "A" Then

        Call initSpreadPosVariables("A")   'sbk 

    	With frm1.vspdData
                ggoSpread.Source = frm1.vspdData
                ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk

    	       .ReDraw = false

               .MaxCols = C_EMPTY + 1										'��: �ִ� Columns�� �׻� 1�� ������Ŵ %>	   
    	       .Col = .MaxCols                                                              ' ��:��: Hide maxcols
               .ColHidden = True                                                            ' ��:��:

               .MaxRows = 0
                ggoSpread.ClearSpreadData

                Call AppendNumberPlace("6","2","0")

                Call GetSpreadColumnPos("A") 'sbk

                ggoSpread.SSSetEdit     C_RECORD_TYPE,      "���ڵ屸��",           12
                ggoSpread.SSSetEdit     C_DATA_TYPE,        "�ڷᱸ��",             10
                ggoSpread.SSSetEdit     C_TAX,              "������",               8
                ggoSpread.SSSetEdit     C_PROV_DT,          "���⿬����",           12
                ggoSpread.SSSetEdit     C_P_TYPE,           "������(�븮�α���)",   19
                ggoSpread.SSSetEdit     C_MAG_NO,           "�����븮�ΰ�����ȣ",   20
				ggoSpread.SSSetEdit     C_HOMETAX_ID,		"Ȩ�ؽ�ID",					20	'2004 
				ggoSpread.SSSetEdit     C_TAX_CODE,			"�������α׷��ڵ�",			45	'2004 	                
                ggoSpread.SSSetEdit     C_OWN_RGST_NO,      "����ڵ�Ϲ�ȣ",       16
                ggoSpread.SSSetEdit     C_CUST_NM_FULL,     "���θ�(��ȣ)",         13
				ggoSpread.SSSetEdit     C_WORKER_DEPT,		"����ںμ�",				30	'2004 	
				ggoSpread.SSSetEdit     C_WORKER_NM,		"����ڼ���",				30	'2004 	
				ggoSpread.SSSetEdit     C_WORKER_TEL,		"�������ȭ��ȣ",			15	'2004 
                ggoSpread.SSSetEdit     C_B_COUNT,          "�Ű��ǹ���(B���ڵ�)��",22
                ggoSpread.SSSetEdit     C_KR_CODE,          "�ѱ��ڵ�����",         14
                ggoSpread.SSSetEdit     C_TERM_CODE,        "������Ⱓ�ڵ�",     18
                ggoSpread.SSSetEdit     C_EMPTY,            "����",                 8

    	    	.ReDraw = true	
        End With
	    
        Call SetSpreadLock("A")
            
    End If
    
    If pvSpdNo = "" OR pvSpdNo = "B" Then

        Call initSpreadPosVariables("B")   'sbk 

    	With frm1.vspdData1
                ggoSpread.Source = frm1.vspdData1
                ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk

    	       .ReDraw = false

               .MaxCols = C_EMPTY1 + 1										'��: �ִ� Columns�� �׻� 1�� ������Ŵ %>	   
    	       .Col = .MaxCols                                                              ' ��:��: Hide maxcols
               .ColHidden = True                                                            ' ��:��:

               .MaxRows = 0
                ggoSpread.ClearSpreadData

                Call AppendNumberPlace("6","2","0")

                Call GetSpreadColumnPos("B") 'sbk

                ggoSpread.SSSetEdit     C_RECORD_TYPE1,             "���ڵ屸��",               12
                ggoSpread.SSSetEdit     C_DATA_TYPE1,               "�ڷᱸ��",                 10
                ggoSpread.SSSetEdit     C_TAX1,                     "������",                   8
                ggoSpread.SSSetEdit     C_NO1,                      "�Ϸù�ȣ",                 10
                ggoSpread.SSSetEdit     C_OWN_RGST_NO1,             "����ڵ�Ϲ�ȣ",           16
                ggoSpread.SSSetEdit     C_CUST_NM_FULL1,            "���θ�(��ȣ)",             13
                ggoSpread.SSSetEdit     C_REPRE_NM1,                "��ǥ��(����)",             13
                ggoSpread.SSSetEdit     C_BCA010T_REPRE_RGST_NO1,   "�ֹ�(����)��Ϲ�ȣ",      19
                ggoSpread.SSSetEdit     C_COM_NO1,                  "��(��)����Ǽ�(C���ڵ��)",25
                ggoSpread.SSSetEdit     C_OLD_COM_NO1,              "��(��)����Ǽ�(D���ڵ��)",25
                ggoSpread.SSSetEdit     C_TOT_PROV_AMT1,            "�ҵ�ݾ��Ѱ�",             14
                ggoSpread.SSSetEdit     C_DECI_INCOME_TAX1,         "�ҵ���������Ѱ�",         18
                ggoSpread.SSSetEdit     C_TOT_TAX1,                 "���ΰ��������Ѱ�",         18
                ggoSpread.SSSetEdit     C_DECI_RES_TAX1,            "�ֹΰ��������Ѱ�",         18
                ggoSpread.SSSetEdit     C_DECI_FARM_TAX1,           "��Ư�����������Ѱ�",       20
                ggoSpread.SSSetEdit     C_DECI_SUM1,                "���������Ѱ�",             14
                ggoSpread.SSSetEdit     C_EMPTY1,                   "����",                     8

    	    	.ReDraw = true	
        End With
	    
        Call SetSpreadLock("B")
            
    End If
    
    If pvSpdNo = "" OR pvSpdNo = "C" Then
 
        Call initSpreadPosVariables("C")   'sbk 
 
    	With frm1.vspdData2
                ggoSpread.Source = frm1.vspdData2
                ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk

    	       .ReDraw = false

               .MaxCols = C_OLD_SUM2 + 1										'��: �ִ� Columns�� �׻� 1�� ������Ŵ %>	   
    	       .Col = .MaxCols                                                  ' ��:��: Hide maxcols
               .ColHidden = True                                                ' ��:��:

               .MaxRows = 0
                ggoSpread.ClearSpreadData

                Call AppendNumberPlace("6","2","0")

                Call GetSpreadColumnPos("C") 'sbk

                ggoSpread.SSSetEdit     C_RECORD_TYPE2,             "���ڵ屸��",               12
                ggoSpread.SSSetEdit     C_DATA_TYPE2,               "�ڷᱸ��",                 10
                ggoSpread.SSSetEdit     C_TAX2,                     "������",                   8
                ggoSpread.SSSetEdit     C_NO2,                      "�Ϸù�ȣ",                 10
                ggoSpread.SSSetEdit     C_OWN_RGST_NO2,             "����ڵ�Ϲ�ȣ",           16
                ggoSpread.SSSetEdit     C_OLD_COM_NO2,              "��(��)�ٹ�ó��",           15
                ggoSpread.SSSetEdit     C_HDF020T_RES_FLAG2,        "�����ڱ����ڵ�",           16  
                ggoSpread.SSSetEdit     C_HAA010T_NAT_CD2 ,         "���������ڵ�",             12
                ggoSpread.SSSetEdit     C_HAA010T_ENTR_DT2,         "�ͼӿ������ۿ�����",       18
                ggoSpread.SSSetEdit     C_HAA010T_RETIRE_DT2,       "�ͼӿ������Ῥ����",       18
                ggoSpread.SSSetEdit     C_HAA010T_NAME2,            "����",                     8
                ggoSpread.SSSetEdit     C_FOR_TYPE2,                "��/�ܱ��α����ڵ�",        17
                ggoSpread.SSSetEdit     C_HGA070T_RETIRE_AMT2,      "�����޿�",                 10
                ggoSpread.SSSetEdit     C_HGA070T_HONOR_AMT2,       "������Ǵ��߰������޿�", 20
                ggoSpread.SSSetEdit     C_HGA070T_CORP_INSUR2,      "��ü���������",           14
                ggoSpread.SSSetEdit     C_HGA070T_TOT_PROV_AMT2,    "��",                       12

                ggoSpread.SSSetEdit     C_ENTR_DT2,                 "��(��)�ٹ����Ի翬����",   22
                ggoSpread.SSSetEdit     C_RETIRE_DT2,               "��(��)�ٹ�����翬����",   22
                ggoSpread.SSSetEdit     C_HGA070T_TOT_DUTY_MM2,     "��(��)�ٹ����ټӿ���",     20
                ggoSpread.SSSetEdit     C_OLD_ENTR_DT2,             "��(��)�ٹ����Ի翬����",   22
                ggoSpread.SSSetEdit     C_OLD_RETIRE_DT2,           "��(��)�ٹ�����翬����",   22
                ggoSpread.SSSetEdit     C_OLD_DUTY2,                "��(��)�ٹ����ټӿ���",     20
                ggoSpread.SSSetEdit     C_D_DUTY2,                  "�ߺ�����",                 10
                ggoSpread.SSSetEdit     C_HGA070T_DUTY_CNT2,        "�ټӿ���",                 10

                ggoSpread.SSSetEdit     C_H_ENTR_DT2,                "��(��)�ٹ����Ի翬����",   22
                ggoSpread.SSSetEdit     C_H_RETIRE_DT2,              "��(��)�ٹ�����翬����",   22
                ggoSpread.SSSetEdit     C_H_HGA070T_TOT_DUTY_MM2,    "��(��)�ٹ����ټӿ���",     20
                ggoSpread.SSSetEdit     C_H_OLD_ENTR_DT2,            "��(��)�ٹ����Ի翬����",   22
                ggoSpread.SSSetEdit     C_H_OLD_RETIRE_DT2,          "��(��)�ٹ�����翬����",   22
                ggoSpread.SSSetEdit     C_H_OLD_DUTY2,               "��(��)�ٹ����ټӿ���",     20
                ggoSpread.SSSetEdit     C_H_D_DUTY2,                 "�ߺ�����",                 10
                ggoSpread.SSSetEdit     C_H_HGA070T_DUTY_CNT2,       "�ټӿ���",                 10
 
                ggoSpread.SSSetEdit     C_RETIRE_TOT_PROV_AMT2,     "�����޿���",               12
                ggoSpread.SSSetEdit     C_HGA070T_INCOME_SUB2,      "�����ҵ������",           16
                ggoSpread.SSSetEdit     C_HGA070T_TAX_STD2,         "�����ҵ����ǥ��",         18
                ggoSpread.SSSetEdit     C_HGA070T_AVR_TAX_STD2,     "����հ���ǥ��",           16
                ggoSpread.SSSetEdit     C_HGA070T_AVR_CALC_TAX2,    "����ջ��⼼��",           16
                ggoSpread.SSSetEdit     C_HGA070T_CALC_TAX2,        "���⼼��",                 10
                ggoSpread.SSSetEdit     C_RETIRE_SUB2,              "�����ҵ漼�װ���",         18
                ggoSpread.SSSetEdit     C_DECI_TAX,					"��������",					18

                ggoSpread.SSSetEdit     C_H_RETIRE_TOT_PROV_AMT2,    "�����޿���",               12
                ggoSpread.SSSetEdit     C_H_HGA070T_INCOME_SUB2,     "�����ҵ������",           16
                ggoSpread.SSSetEdit     C_H_HGA070T_TAX_STD2,        "�����ҵ����ǥ��",         18
                ggoSpread.SSSetEdit     C_H_HGA070T_AVR_TAX_STD2,    "����հ���ǥ��",           16
                ggoSpread.SSSetEdit     C_H_HGA070T_AVR_CALC_TAX2,   "����ջ��⼼��",           16
                ggoSpread.SSSetEdit     C_H_HGA070T_CALC_TAX2,       "���⼼��",                 10
                ggoSpread.SSSetEdit     C_H_RETIRE_SUB2,             "�����ҵ漼�װ���",         18
                ggoSpread.SSSetEdit     C_H_DECI_TAX,				 "��������",				 18
 
                ggoSpread.SSSetEdit     C_T_RETIRE_TOT_PROV_AMT2,     "�����޿���",               12
                ggoSpread.SSSetEdit     C_T_HGA070T_INCOME_SUB2,      "�����ҵ������",           16
                ggoSpread.SSSetEdit     C_T_HGA070T_TAX_STD2,         "�����ҵ����ǥ��",         18
                ggoSpread.SSSetEdit     C_T_HGA070T_AVR_TAX_STD2,     "����հ���ǥ��",           16
                ggoSpread.SSSetEdit     C_T_HGA070T_AVR_CALC_TAX2,    "����ջ��⼼��",           16
                ggoSpread.SSSetEdit     C_T_HGA070T_CALC_TAX2,        "���⼼��",                 10
                ggoSpread.SSSetEdit     C_T_RETIRE_SUB2,              "�����ҵ漼�װ���",         18
                ggoSpread.SSSetEdit     C_T_DECI_TAX,				  "��������",				  18
 
                ggoSpread.SSSetEdit     C_HGA070T_DECI_INCOME_TAX2, "�ҵ漼",                   8
                ggoSpread.SSSetEdit     C_HGA070T_DECI_RES_TAX2,    "�ֹμ�",                   8
                ggoSpread.SSSetEdit     C_DECI_FARM_TAX2,           "�����Ư����",             14
                ggoSpread.SSSetEdit     C_DECI_SUM2,                "��",                       12
                ggoSpread.SSSetEdit     C_HFA050T_OLD_INCOME_TAX2,  "�ҵ漼",                   8
                ggoSpread.SSSetEdit     C_HFA050T_OLD_RES_TAX2,     "�ֹμ�",                   8
                ggoSpread.SSSetEdit     C_HFA050T_OLD_FARM_TAX2,    "�����Ư����",             14
                ggoSpread.SSSetEdit     C_OLD_SUM2,                 "��",                       12

    	    	.ReDraw = true	
        End With
 	    
        Call SetSpreadLock("C")
            
    End If

End Sub

'======================================================================================================
' Function Name : SetSpreadLock
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadLock(ByVal pvSpdNo)

    Select Case pvSpdNo
        Case  "A"
            With frm1
            ggoSpread.Source = frm1.vspdData
            .vspdData.ReDraw = False
            ggoSpread.SpreadLock C_RECORD_TYPE, -1,  C_RECORD_TYPE, -1
            ggoSpread.SpreadLock C_DATA_TYPE,   -1,  C_DATA_TYPE,   -1  
            ggoSpread.SpreadLock C_TAX,         -1,  C_TAX,         -1
            ggoSpread.SpreadLock C_PROV_DT,     -1,  C_PROV_DT,     -1
            ggoSpread.SpreadLock C_P_TYPE,      -1,  C_P_TYPE,      -1
            ggoSpread.SpreadLock C_MAG_NO,      -1,  C_MAG_NO,      -1
            ggoSpread.SpreadLock C_HOMETAX_ID,	-1,  C_HOMETAX_ID,      -1
            ggoSpread.SpreadLock C_TAX_CODE,	-1,  C_TAX_CODE,      -1                        
            ggoSpread.SpreadLock C_OWN_RGST_NO, -1,  C_OWN_RGST_NO, -1
            ggoSpread.SpreadLock C_CUST_NM_FULL,-1,  C_CUST_NM_FULL,-1
            ggoSpread.SpreadLock C_WORKER_DEPT,	-1,  C_WORKER_DEPT,     -1
            ggoSpread.SpreadLock C_WORKER_NM,	-1,  C_WORKER_NM,    -1
            ggoSpread.SpreadLock C_WORKER_TEL,	-1,  C_WORKER_TEL,      -1
            ggoSpread.SpreadLock C_B_COUNT,     -1,  C_B_COUNT,     -1
            ggoSpread.SpreadLock C_KR_CODE,     -1,  C_KR_CODE,     -1
            ggoSpread.SpreadLock C_TERM_CODE,   -1,  C_TERM_CODE,   -1
            ggoSpread.SpreadLock C_EMPTY,       -1,  C_EMPTY,       -1
            ggoSpread.SSSetProtected  .vspdData.MaxCols   , -1, -1
            .vspdData.ReDraw = True
            End With

        Case  "B"
            With frm1
            ggoSpread.Source = frm1.vspdData1
            .vspdData1.ReDraw = False
            ggoSpread.SpreadLock C_RECORD_TYPE1          ,-1,   C_RECORD_TYPE1           ,-1
            ggoSpread.SpreadLock C_DATA_TYPE1            ,-1,   C_DATA_TYPE1             ,-1
            ggoSpread.SpreadLock C_TAX1                  ,-1,   C_TAX1                   ,-1
            ggoSpread.SpreadLock C_NO1                   ,-1,   C_NO1                    ,-1
            ggoSpread.SpreadLock C_OWN_RGST_NO1          ,-1,   C_OWN_RGST_NO1           ,-1
            ggoSpread.SpreadLock C_CUST_NM_FULL1         ,-1,   C_CUST_NM_FULL1          ,-1
            ggoSpread.SpreadLock C_REPRE_NM1             ,-1,   C_REPRE_NM1              ,-1
            ggoSpread.SpreadLock C_BCA010T_REPRE_RGST_NO1,-1,   C_BCA010T_REPRE_RGST_NO1 ,-1
            ggoSpread.SpreadLock C_COM_NO1               ,-1,   C_COM_NO1                ,-1
            ggoSpread.SpreadLock C_OLD_COM_NO1           ,-1,   C_OLD_COM_NO1            ,-1
            ggoSpread.SpreadLock C_TOT_PROV_AMT1         ,-1,   C_TOT_PROV_AMT1          ,-1
            ggoSpread.SpreadLock C_DECI_INCOME_TAX1      ,-1,   C_DECI_INCOME_TAX1       ,-1
            ggoSpread.SpreadLock C_TOT_TAX1              ,-1,   C_TOT_TAX1               ,-1
            ggoSpread.SpreadLock C_DECI_RES_TAX1         ,-1,   C_DECI_RES_TAX1          ,-1
            ggoSpread.SpreadLock C_DECI_FARM_TAX1        ,-1,   C_DECI_FARM_TAX1         ,-1
            ggoSpread.SpreadLock C_DECI_SUM1             ,-1,   C_DECI_SUM1              ,-1
            ggoSpread.SpreadLock C_EMPTY1                ,-1,   C_EMPTY1                 ,-1    
            ggoSpread.SSSetProtected  .vspdData1.MaxCols   , -1, -1
            .vspdData1.ReDraw = True
            End With

        Case  "C"
            With frm1
            ggoSpread.Source = frm1.vspdData2
            .vspdData2.ReDraw = False
            ggoSpread.SpreadLock C_RECORD_TYPE2            ,-1,  C_RECORD_TYPE2            ,-1 
            ggoSpread.SpreadLock C_DATA_TYPE2              ,-1,  C_DATA_TYPE2              ,-1 
            ggoSpread.SpreadLock C_TAX2                    ,-1,  C_TAX2                    ,-1 
            ggoSpread.SpreadLock C_NO2                     ,-1,  C_NO2                     ,-1 
            ggoSpread.SpreadLock C_OWN_RGST_NO2            ,-1,  C_OWN_RGST_NO2            ,-1 
            ggoSpread.SpreadLock C_OLD_COM_NO2             ,-1,  C_OLD_COM_NO2             ,-1 
            ggoSpread.SpreadLock C_HDF020T_RES_FLAG2       ,-1,  C_HDF020T_RES_FLAG2       ,-1 
            ggoSpread.SpreadLock C_HAA010T_NAT_CD2         ,-1,  C_HAA010T_NAT_CD2         ,-1 
            ggoSpread.SpreadLock C_HAA010T_ENTR_DT2        ,-1,  C_HAA010T_ENTR_DT2        ,-1 
            ggoSpread.SpreadLock C_HAA010T_RETIRE_DT2      ,-1,  C_HAA010T_RETIRE_DT2      ,-1 
            ggoSpread.SpreadLock C_HAA010T_NAME2           ,-1,  C_HAA010T_NAME2           ,-1 
            ggoSpread.SpreadLock C_FOR_TYPE2               ,-1,  C_FOR_TYPE2               ,-1 
            ggoSpread.SpreadLock C_HGA070T_RETIRE_AMT2     ,-1,  C_HGA070T_RETIRE_AMT2     ,-1 
            ggoSpread.SpreadLock C_HGA070T_HONOR_AMT2      ,-1,  C_HGA070T_HONOR_AMT2      ,-1 
            ggoSpread.SpreadLock C_HGA070T_CORP_INSUR2     ,-1,  C_HGA070T_CORP_INSUR2     ,-1 
            ggoSpread.SpreadLock C_HGA070T_TOT_PROV_AMT2   ,-1,  C_HGA070T_TOT_PROV_AMT2   ,-1 
            
            ggoSpread.SpreadLock C_ENTR_DT2                ,-1,  C_ENTR_DT2                ,-1 
            ggoSpread.SpreadLock C_RETIRE_DT2              ,-1,  C_RETIRE_DT2              ,-1 
            ggoSpread.SpreadLock C_HGA070T_TOT_DUTY_MM2    ,-1,  C_HGA070T_TOT_DUTY_MM2    ,-1
            ggoSpread.SpreadLock C_OLD_ENTR_DT2            ,-1,  C_OLD_ENTR_DT2            ,-1
            ggoSpread.SpreadLock C_OLD_RETIRE_DT2          ,-1,  C_OLD_RETIRE_DT2          ,-1
            ggoSpread.SpreadLock C_OLD_DUTY2               ,-1,  C_OLD_DUTY2               ,-1
            ggoSpread.SpreadLock C_D_DUTY2                 ,-1,  C_D_DUTY2                 ,-1
            ggoSpread.SpreadLock C_HGA070T_DUTY_CNT2       ,-1,  C_HGA070T_DUTY_CNT2       ,-1
 
            ggoSpread.SpreadLock C_H_ENTR_DT2                ,-1,  C_H_ENTR_DT2                ,-1 
            ggoSpread.SpreadLock C_H_RETIRE_DT2              ,-1,  C_H_RETIRE_DT2              ,-1 
            ggoSpread.SpreadLock C_H_HGA070T_TOT_DUTY_MM2    ,-1,  C_H_HGA070T_TOT_DUTY_MM2    ,-1
            ggoSpread.SpreadLock C_H_OLD_ENTR_DT2            ,-1,  C_H_OLD_ENTR_DT2            ,-1
            ggoSpread.SpreadLock C_H_OLD_RETIRE_DT2          ,-1,  C_H_OLD_RETIRE_DT2          ,-1
            ggoSpread.SpreadLock C_H_OLD_DUTY2               ,-1,  C_H_OLD_DUTY2               ,-1
            ggoSpread.SpreadLock C_H_D_DUTY2                 ,-1,  C_H_D_DUTY2                 ,-1
            ggoSpread.SpreadLock C_H_HGA070T_DUTY_CNT2       ,-1,  C_H_HGA070T_DUTY_CNT2       ,-1
 
            ggoSpread.SpreadLock C_RETIRE_TOT_PROV_AMT2    ,-1,  C_RETIRE_TOT_PROV_AMT2    ,-1
            ggoSpread.SpreadLock C_HGA070T_INCOME_SUB2     ,-1,  C_HGA070T_INCOME_SUB2     ,-1
            ggoSpread.SpreadLock C_HGA070T_TAX_STD2        ,-1,  C_HGA070T_TAX_STD2        ,-1
            ggoSpread.SpreadLock C_HGA070T_AVR_TAX_STD2    ,-1,  C_HGA070T_AVR_TAX_STD2    ,-1
            ggoSpread.SpreadLock C_HGA070T_AVR_CALC_TAX2   ,-1,  C_HGA070T_AVR_CALC_TAX2   ,-1
            ggoSpread.SpreadLock C_HGA070T_CALC_TAX2       ,-1,  C_HGA070T_CALC_TAX2       ,-1
            ggoSpread.SpreadLock C_RETIRE_SUB2             ,-1,  C_RETIRE_SUB2             ,-1
            ggoSpread.SpreadLock C_DECI_TAX				   ,-1,  C_DECI_TAX				   ,-1
 
            ggoSpread.SpreadLock C_H_RETIRE_TOT_PROV_AMT2    ,-1,  C_H_RETIRE_TOT_PROV_AMT2    ,-1
            ggoSpread.SpreadLock C_H_HGA070T_INCOME_SUB2     ,-1,  C_H_HGA070T_INCOME_SUB2     ,-1
            ggoSpread.SpreadLock C_H_HGA070T_TAX_STD2        ,-1,  C_H_HGA070T_TAX_STD2        ,-1
            ggoSpread.SpreadLock C_H_HGA070T_AVR_TAX_STD2    ,-1,  C_H_HGA070T_AVR_TAX_STD2    ,-1
            ggoSpread.SpreadLock C_H_HGA070T_AVR_CALC_TAX2   ,-1,  C_H_HGA070T_AVR_CALC_TAX2   ,-1
            ggoSpread.SpreadLock C_H_HGA070T_CALC_TAX2       ,-1,  C_H_HGA070T_CALC_TAX2       ,-1
            ggoSpread.SpreadLock C_H_RETIRE_SUB2             ,-1,  C_H_RETIRE_SUB2             ,-1
            ggoSpread.SpreadLock C_H_DECI_TAX				   ,-1,  C_H_DECI_TAX				,-1

            ggoSpread.SpreadLock C_T_RETIRE_TOT_PROV_AMT2    ,-1,  C_T_RETIRE_TOT_PROV_AMT2    ,-1
            ggoSpread.SpreadLock C_T_HGA070T_INCOME_SUB2     ,-1,  C_T_HGA070T_INCOME_SUB2     ,-1
            ggoSpread.SpreadLock C_T_HGA070T_TAX_STD2        ,-1,  C_T_HGA070T_TAX_STD2        ,-1
            ggoSpread.SpreadLock C_T_HGA070T_AVR_TAX_STD2    ,-1,  C_T_HGA070T_AVR_TAX_STD2    ,-1
            ggoSpread.SpreadLock C_T_HGA070T_AVR_CALC_TAX2   ,-1,  C_T_HGA070T_AVR_CALC_TAX2   ,-1
            ggoSpread.SpreadLock C_T_HGA070T_CALC_TAX2       ,-1,  C_T_HGA070T_CALC_TAX2       ,-1
            ggoSpread.SpreadLock C_T_RETIRE_SUB2             ,-1,  C_T_RETIRE_SUB2             ,-1
            ggoSpread.SpreadLock C_T_DECI_TAX				 ,-1,  C_T_DECI_TAX				   ,-1
            
            ggoSpread.SpreadLock C_HGA070T_DECI_INCOME_TAX2,-1,  C_HGA070T_DECI_INCOME_TAX2,-1
            ggoSpread.SpreadLock C_HGA070T_DECI_RES_TAX2   ,-1,  C_HGA070T_DECI_RES_TAX2   ,-1
            ggoSpread.SpreadLock C_DECI_FARM_TAX2          ,-1,  C_DECI_FARM_TAX2          ,-1
            ggoSpread.SpreadLock C_DECI_SUM2               ,-1,  C_DECI_SUM2               ,-1
            ggoSpread.SpreadLock C_HFA050T_OLD_INCOME_TAX2 ,-1,  C_HFA050T_OLD_INCOME_TAX2 ,-1
            ggoSpread.SpreadLock C_HFA050T_OLD_RES_TAX2    ,-1,  C_HFA050T_OLD_RES_TAX2    ,-1 
            ggoSpread.SpreadLock C_HFA050T_OLD_FARM_TAX2   ,-1,  C_HFA050T_OLD_FARM_TAX2   ,-1 
            ggoSpread.SpreadLock C_OLD_SUM2                ,-1,  C_OLD_SUM2                ,-1 
    
            ggoSpread.SSSetProtected  .vspdData2.MaxCols   , -1, -1
            .vspdData2.ReDraw = True
            End With
    End Select
End Sub

'======================================================================================================
' Function Name : SetSpreadColor
' Function Desc : This method set color and protect in spread sheet celles
'======================================================================================================
Sub SetSpreadColor(ByVal pvStartRow,ByVal pvEndRow)
    With frm1

    End With

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
       For iDx = 1 To  frm1.vspdData.MaxCols - 1
           Frm1.vspdData.Col = iDx
           Frm1.vspdData.Row = iRow
           If Frm1.vspdData.ColHidden <> True And Frm1.vspdData.BackColor <> parent.UC_PROTECTED Then
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

            C_RECORD_TYPE             = iCurColumnPos(1)
            C_DATA_TYPE               = iCurColumnPos(2)
            C_TAX                     = iCurColumnPos(3)
            C_PROV_DT                 = iCurColumnPos(4)
            C_P_TYPE                  = iCurColumnPos(5)
            C_MAG_NO                  = iCurColumnPos(6)
            C_HOMETAX_ID			= iCurColumnPos(7)
            C_TAX_CODE				= iCurColumnPos(8)
            C_OWN_RGST_NO			= iCurColumnPos(9)
            C_CUST_NM_FULL			= iCurColumnPos(10)
            C_WORKER_DEPT			= iCurColumnPos(11)
            C_WORKER_NM				= iCurColumnPos(12)
            C_WORKER_TEL			= iCurColumnPos(13)
            C_B_COUNT               = iCurColumnPos(14)
            C_KR_CODE				= iCurColumnPos(15)
            C_TERM_CODE				= iCurColumnPos(16)
            C_EMPTY					= iCurColumnPos(17)

       Case "B"
            ggoSpread.Source = frm1.vspdData1
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)

            C_RECORD_TYPE1            = iCurColumnPos(1)
            C_DATA_TYPE1              = iCurColumnPos(2)
            C_TAX1                    = iCurColumnPos(3)
            C_NO1                     = iCurColumnPos(4)
            C_OWN_RGST_NO1            = iCurColumnPos(5)
            C_CUST_NM_FULL1           = iCurColumnPos(6)
            C_REPRE_NM1               = iCurColumnPos(7)
            C_BCA010T_REPRE_RGST_NO1  = iCurColumnPos(8)
            C_COM_NO1                 = iCurColumnPos(9)
            C_OLD_COM_NO1             = iCurColumnPos(10)
            C_TOT_PROV_AMT1           = iCurColumnPos(11)
            C_DECI_INCOME_TAX1        = iCurColumnPos(12)
            C_TOT_TAX1                = iCurColumnPos(13)
            C_DECI_RES_TAX1           = iCurColumnPos(14)
            C_DECI_FARM_TAX1          = iCurColumnPos(15)
            C_DECI_SUM1               = iCurColumnPos(16)
            C_EMPTY1                  = iCurColumnPos(17)
    
       Case "C"
            ggoSpread.Source = frm1.vspdData2
            Call ggoSpread.GetSpreadColumnPos(iCurColumnPos)
    
            C_RECORD_TYPE2            = iCurColumnPos(1)
            C_DATA_TYPE2              = iCurColumnPos(2)
            C_TAX2                    = iCurColumnPos(3)
            C_NO2                     = iCurColumnPos(4)
            C_OWN_RGST_NO2            = iCurColumnPos(5)
            C_OLD_COM_NO2             = iCurColumnPos(6)
            C_HDF020T_RES_FLAG2       = iCurColumnPos(7)
            C_HAA010T_NAT_CD2         = iCurColumnPos(8)   '2002 ���������ڵ� �߰� 
            C_HAA010T_ENTR_DT2        = iCurColumnPos(9)
            C_HAA010T_RETIRE_DT2      = iCurColumnPos(10)
            C_HAA010T_NAME2           = iCurColumnPos(11)
            C_FOR_TYPE2               = iCurColumnPos(12)
            C_HGA070T_RETIRE_AMT2     = iCurColumnPos(13)
            C_HGA070T_HONOR_AMT2      = iCurColumnPos(14)
            C_HGA070T_CORP_INSUR2     = iCurColumnPos(15)
            C_HGA070T_TOT_PROV_AMT2   = iCurColumnPos(16)
            C_ENTR_DT2                = iCurColumnPos(17)
            C_RETIRE_DT2              = iCurColumnPos(18)
            C_HGA070T_TOT_DUTY_MM2    = iCurColumnPos(19)
            C_OLD_ENTR_DT2            = iCurColumnPos(20)
            C_OLD_RETIRE_DT2          = iCurColumnPos(21)
            C_OLD_DUTY2               = iCurColumnPos(22)
            C_D_DUTY2                 = iCurColumnPos(23)
            C_HGA070T_DUTY_CNT2       = iCurColumnPos(24)
            
            C_H_ENTR_DT2				= iCurColumnPos(25)
            C_H_RETIRE_DT2				= iCurColumnPos(26)
            C_H_HGA070T_TOT_DUTY_MM2	= iCurColumnPos(27)
            C_H_OLD_ENTR_DT2			= iCurColumnPos(28)
            C_H_OLD_RETIRE_DT2			= iCurColumnPos(29)
            C_H_OLD_DUTY2				= iCurColumnPos(30)
            C_H_D_DUTY2					= iCurColumnPos(31)
            C_H_HGA070T_DUTY_CNT2		= iCurColumnPos(32)

            C_RETIRE_TOT_PROV_AMT2		= iCurColumnPos(33)
            C_HGA070T_INCOME_SUB2		= iCurColumnPos(34)
            C_HGA070T_TAX_STD2			= iCurColumnPos(35)
            C_HGA070T_AVR_TAX_STD2		= iCurColumnPos(36)
            C_HGA070T_AVR_CALC_TAX2		= iCurColumnPos(37)
            C_HGA070T_CALC_TAX2			= iCurColumnPos(38)
            C_RETIRE_SUB2				= iCurColumnPos(39)
			C_DECI_TAX					= iCurColumnPos(40)
 
            C_H_RETIRE_TOT_PROV_AMT2	= iCurColumnPos(41)
            C_H_HGA070T_INCOME_SUB2		= iCurColumnPos(42)
            C_H_HGA070T_TAX_STD2		= iCurColumnPos(43)
            C_H_HGA070T_AVR_TAX_STD2	= iCurColumnPos(44)
            C_H_HGA070T_AVR_CALC_TAX2	= iCurColumnPos(45)
            C_H_HGA070T_CALC_TAX2		= iCurColumnPos(46)
            C_H_RETIRE_SUB2				= iCurColumnPos(47)
            C_H_DECI_TAX        		= iCurColumnPos(48)          
 
            C_T_RETIRE_TOT_PROV_AMT2	= iCurColumnPos(49)
            C_T_HGA070T_INCOME_SUB2		= iCurColumnPos(50)
            C_T_HGA070T_TAX_STD2		= iCurColumnPos(51)
            C_T_HGA070T_AVR_TAX_STD2	= iCurColumnPos(52)
            C_T_HGA070T_AVR_CALC_TAX2	= iCurColumnPos(53)
            C_T_HGA070T_CALC_TAX2		= iCurColumnPos(54)
            C_T_RETIRE_SUB2				= iCurColumnPos(55)
			C_T_DECI_TAX				= iCurColumnPos(56)
            
            C_HGA070T_DECI_INCOME_TAX2= iCurColumnPos(57)
            C_HGA070T_DECI_RES_TAX2   = iCurColumnPos(58)
            C_DECI_FARM_TAX2          = iCurColumnPos(59)
            C_DECI_SUM2               = iCurColumnPos(60)
            C_HFA050T_OLD_INCOME_TAX2 = iCurColumnPos(61)
            C_HFA050T_OLD_RES_TAX2    = iCurColumnPos(62)
            C_HFA050T_OLD_FARM_TAX2   = iCurColumnPos(63)
            C_OLD_SUM2                = iCurColumnPos(64)
    End Select    
End Sub

'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '��: Clear err status
	Call LoadInfTB19029                                                             '��: Load table , B_numeric_format
 
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'��: Lock Field
 
    Call InitSpreadSheet("")                                                            'Setup the Spread sheet
    Call InitVariables                                                              'Initializes local global variables
 
    Call SetDefaultVal
  
    Call InitComboBox
    Call SetToolbar("1100000000001111")										        '��ư ���� ���� 
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

    FncQuery = False															 '��: Processing is NG
    Err.Clear                                                                    '��: Clear err status

    ggoSpread.Source = Frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgbox("900013", parent.VB_YES_NO,"x","x")					 '��: Data is changed.  Do you want to display it?
		If IntRetCD = vbNo Then
			Exit Function
		End If
    End If

    Call ggoOper.ClearField(Document, "2")										 '��: Clear Contents  Field

    If Not chkField(Document, "1") Then									         '��: This function check required field
       Exit Function
    End If

    If CompareDateByFormat(frm1.txtStrt_dt.Text,frm1.txtBas_dt.Text,frm1.txtStrt_dt.Alt,frm1.txtBas_dt.Alt,"970023",frm1.txtStrt_dt.UserDefinedFormat,parent.gComDateType,True) = False Then
        frm1.txtBas_dt.focus()
        Set gActiveElement = document.activeElement
        Exit Function
    End If

    If CompareDateByFormat(frm1.txtBas_dt.Text,frm1.txtDt.Text,frm1.txtBas_dt.Alt,frm1.txtDt.Alt,"970023",frm1.txtBas_dt.UserDefinedFormat,parent.gComDateType,True) = False Then
        frm1.txtDt.focus()
        Set gActiveElement = document.activeElement
        Exit Function
    End If

    Call InitVariables                                                           '��: Initializes local global variables

    lgCurrentSpd = "A"
	topleftOK = false
    Call MakeKeyStream(lgCurrentSpd)
    If DbQuery = False Then
		Exit Function
	End If                                                                 '��: Query db data

    FncQuery = True                                                              '��: Processing is OK

End Function

'========================================================================================================
' Name : FncNew
' Desc : developer describe this line Called by MainNew in Common.vbs
'========================================================================================================
Function FncNew()
End Function

'========================================================================================================
' Name : FncDelete
' Desc : developer describe this line Called by MainDelete in Common.vbs
'========================================================================================================
Function FncDelete()
End Function

'========================================================================================================
' Name : FncSave
' Desc : developer describe this line Called by MainSave in Common.vbs
'========================================================================================================
Function FncSave()
End Function

'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()
End Function

'========================================================================================================
' Function Name : FncCancel
' Function Desc : This function is related to Cancel Button of Main ToolBar
'========================================================================================================
Function FncCancel()
    ggoSpread.Source = Frm1.vspdData
    ggoSpread.EditUndo
End Function

'========================================================================================================
' Function Name : FncInsertRow
' Function Desc : This function is related to InsertRow Button of Main ToolBar
'========================================================================================================
Function FncInsertRow()
End Function

'========================================================================================================
' Function Name : FncDeleteRow
' Function Desc : This function is related to DeleteRow Button of Main ToolBar
'========================================================================================================
Function FncDeleteRow()
End Function

'========================================================================================================
' Function Name : FncPrint
' Function Desc : This function is related to Print Button of Main ToolBar
'========================================================================================================
Function FncPrint()
    Call parent.FncPrint()
End Function

'========================================================================================================
' Function Name : FncExcel
' Function Desc : This function is related to Excel
'========================================================================================================
Function FncExcel()
    Call parent.FncExport(parent.C_MULTI)                                         '��: ȭ�� ���� 
End Function

'========================================================================================================
' Function Name : FncFind
' Function Desc :
'========================================================================================================
Function FncFind()
    Call parent.FncFind(parent.C_MULTI, False)                                    '��:ȭ�� ����, Tab ���� 
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
    select case gActiveSpdSheet.id
		case "vaSpread"
			Call InitSpreadSheet("A")      
		case "vaSpread1"
			Call InitSpreadSheet("B")      		
		case "vaSpread2"
			Call InitSpreadSheet("C")      		
	end select 
    
	Call ggoSpread.ReOrderingSpreadData()
	Call InitData()
End Sub

'========================================================================================================
' Function Name : FncExit
' Function Desc :
'========================================================================================================
Function FncExit()
    Dim IntRetCD

	FncExit = False

    ggoSpread.Source = frm1.vspdData
    If ggoSpread.SSCheckChange = True Then
		IntRetCD = DisplayMsgbox("900016", parent.VB_YES_NO,"x","x")			'��: Data is changed.  Do you want to exit?
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

    DbQuery = False

    Err.Clear                                                                        '��: Clear err status

    If LayerShowHide(1) = false Then
        Exit Function
    End If

	Dim strVal

    With Frm1
		strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001
		strVal = strVal     & "&lgCurrentSpd="       & lgCurrentSpd                      '��: Next key tag
        strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '��: Query Key
		If lgCurrentSpd = "A" Then
		    strVal = strVal     & "&lgStrPrevKey="       &  lgStrPrevKey
		elseIf lgCurrentSpd = "B" Then
		    strVal = strVal     & "&lgStrPrevKey1="       &  lgStrPrevKey1
		elseIf lgCurrentSpd = "C" Then
		    strVal = strVal     & "&lgStrPrevKey2="       &  lgStrPrevKey2
		end if        
    End With

	Call RunMyBizASP(MyBizASP, strVal)                                               '��: Run Biz Logic

    DbQuery = True

End Function

'========================================================================================================
' Name : DbSave
' Desc : This function is data query and display
'========================================================================================================
Function DbSave()
End Function

'========================================================================================================
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()

    If lgCurrentSpd = "C" And (frm1.vspdData.MaxRows <= 0 And frm1.vspdData1.MaxRows<= 0 And frm1.vspdData2.MaxRows <= 0) Then
		Call DisplayMsgbox("900014", "X","X","X")			                            '��: ��ȸ�� �����ϼ��� 
    End If
    lgIntFlgMode = parent.OPMD_UMODE
    Call ggoOper.LockField(Document, "Q")										'��: Lock field
    Call InitData()
	Call SetToolbar("1100000000011111")
	If lgCurrentSpd = "A" then
		frm1.vspdData.focus
	elseIf lgCurrentSpd = "B" then
		frm1.vspdData1.focus	
	elseIf lgCurrentSpd = "C" then		
		frm1.vspdData2.focus	
	end if
End Function

'========================================================================================================
'   Event Name : vspdData_Change
'   Event Desc :
'========================================================================================================
Sub vspdData_Change(ByVal Col , ByVal Row )
    Dim iDx

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
'   Event Name : vspdData_Click
'   Event Desc : �÷��� Ŭ���� ��� �߻� 
'========================================================================================================
Sub vspdData_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111")

    gMouseClickStatus = "SPC" 

    Set gActiveSpdSheet = frm1.vspdData
   
    If frm1.vspdData1.MaxRows = 0 Then                                                    'If there is no data.
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

Sub vspdData1_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111")

    gMouseClickStatus = "SP1C" 

    Set gActiveSpdSheet = frm1.vspdData1
   
    If frm1.vspdData1.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData1
       
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

Sub vspdData2_Click(ByVal Col, ByVal Row)
	Call SetPopupMenuItemInf("0000111111")

    gMouseClickStatus = "SP2C" 

    Set gActiveSpdSheet = frm1.vspdData2
   
    If frm1.vspdData1.MaxRows = 0 Then                                                    'If there is no data.
       Exit Sub
   	End If
    
    If Row <= 0 Then
       ggoSpread.Source = frm1.vspdData2
       
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

'-----------------------------------------
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


'========================================================================================================
'   Event Name : vspdData_ColWidthChange
'   Event Desc : 
'========================================================================================================
Sub vspdData_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData
    call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData1_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData1
    call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub

Sub vspdData2_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData2
    call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
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

Sub vspdData1_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData1
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("B")
End Sub 

Sub vspdData2_ScriptDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)

    ggoSpread.Source = frm1.vspdData2
    Call ggoSpread.SpreadDragDropBlock(Col, Row, Col2, Row2, NewCol, NewRow, NewCol2, NewRow2, Overwrite, Action, DataOnly, Cancel)    
    Call GetSpreadColumnPos("C")
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

Sub vspdData1_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName
    if Row <= 0 then
		exit sub
	end if
	if Frm1.vspdData1.MaxRows = 0 then
		exit sub
	end if
End Sub

Sub vspdData2_DblClick(ByVal Col, ByVal Row)
    Dim iColumnName
    if Row <= 0 then
		exit sub
	end if
	if Frm1.vspdData2.MaxRows = 0 then
		exit sub
	end if
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
'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData1_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    if frm1.vspdData1.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData1,NewTop) Then
    	If lgStrPrevKey1 <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	

      	   Call DisableToolBar(parent.TBC_QUERY)
			topleftOK = true	
			lgCurrentSpd = "B"

      	   If DBQuery = false Then
      	    Call RestoreToolBar()
      	    Exit Sub
      	   End If
    	End If
    End if
End Sub

'========================================================================================================
'   Event Name : vspdData_TopLeftChange
'   Event Desc : This function is data query with spread sheet scrolling
'========================================================================================================
Sub vspdData2_TopLeftChange(ByVal OldLeft , ByVal OldTop , ByVal NewLeft , ByVal NewTop )

    If OldLeft <> NewLeft Then
        Exit Sub
    End If

    if frm1.vspdData2.MaxRows < NewTop + VisibleRowCnt(frm1.vspdData2,NewTop) Then
    	If lgStrPrevKey2 <> "" Then
			If CheckRunningBizProcess = True Then
				Exit Sub
			End If	

      	   Call DisableToolBar(parent.TBC_QUERY)
			lgCurrentSpd = "C"

      	   If DBQuery = false Then
      	    Call RestoreToolBar()
      	    Exit Sub
      	   End If
    	End If
    End if
End Sub

'======================================================================================================
' Function Name : btnCb_print2_onClick
' Function Desc : �÷��ǵ���, �Ű� ���� ��� 
'=======================================================================================================
Sub btnCb_print2_onClick()
Dim RetFlag

    If frm1.vspdData.MaxRows <= 0 And frm1.vspdData1.MaxRows<= 0 And frm1.vspdData2.MaxRows <= 0 Then
		Call DisplayMsgbox("800167", "X","X","X")			                            '��: ��ȸ�� �����ϼ��� 
		Exit Sub
    End If

    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Sub
    End If

    If CompareDateByFormat(frm1.txtStrt_dt.Text,frm1.txtBas_dt.Text,frm1.txtStrt_dt.Alt,frm1.txtBas_dt.Alt,"970023",frm1.txtStrt_dt.UserDefinedFormat,parent.gComDateType,True) = False Then
        frm1.txtBas_dt.focus
        Set gActiveElement = document.activeElement
        Exit Sub
    End If

    RetFlag = DisplayMsgbox("900018", parent.VB_YES_NO,"x","x")                                '�� �۾��� ����Ͻðڽ��ϱ�?
	If RetFlag = VBNO Then
		Exit Sub
    Else
        Call FloppyDiskLabelForm()      '�÷��ǵ��� �󺧾�� 
        Call ReportOfDocuments()        '�Ű� ���� 
	End IF



End Sub
'======================================================================================================
' Function Name : btnCb_print_onClick
' Function Desc : ����ǥ ��� 
'=======================================================================================================
Sub btnCb_print_onClick()
Dim RetFlag

    If frm1.vspdData.MaxRows <= 0 And frm1.vspdData1.MaxRows<= 0 And frm1.vspdData2.MaxRows <= 0 Then
		Call DisplayMsgbox("800167", "X","X","X")			                            '��: ��ȸ�� �����ϼ��� 
		Exit Sub
    End If

    If Not chkField(Document, "1") Then									'��: This function check indispensable field
       Exit Sub
    End If

    If CompareDateByFormat(frm1.txtStrt_dt.Text,frm1.txtBas_dt.Text,frm1.txtStrt_dt.Alt,frm1.txtBas_dt.Alt,"970023",frm1.txtStrt_dt.UserDefinedFormat,parent.gComDateType,True) = False Then
        frm1.txtBas_dt.focus
        Set gActiveElement = document.activeElement
        Exit Sub
    End If

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
	Dim condvar	
    Dim StrEbrFile
	Dim ObjName
	
	dim biz_area_cd, end_dt, end_yy, present_dt, start_dt, start_yy

	StrEbrFile = "ha108oa1_1"

	biz_area_cd = frm1.txtComp_cd.value
	
	end_dt = UniConvDateToYYYYMMDD(frm1.txtBas_dt.Text,parent.gDateFormat,parent.gServerDateType)
	end_yy = frm1.txtBas_dt.Year
	
	present_dt = UniConvDateToYYYYMMDD(frm1.txtDt.Text,parent.gDateFormat,parent.gServerDateType)

	start_dt = UniConvDateToYYYYMMDD(frm1.txtStrt_dt.Text,parent.gDateFormat,parent.gServerDateType)
	start_yy = frm1.txtStrt_dt.Year

	condvar = "biz_area_cd|" & biz_area_cd
	condvar = condvar & "|end_dt|" & end_dt
	condvar = condvar & "|end_yy|" & end_yy
	condvar = condvar & "|present_dt|" & present_dt
	condvar = condvar & "|start_dt|" & start_dt
	condvar = condvar & "|start_yy|" & start_yy

	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
 	call FncEBRPrint(EBAction , ObjName , condvar)
End Function
'======================================================================================================
' Function Name : FloppyDiskLabelForm
' Function Desc : �÷��ǵ��� �󺧾�� 
'=======================================================================================================
Function FloppyDiskLabelForm()
	Dim condvar
	Dim StrEbrFile
    Dim ObjName

	dim biz_area_cd, end_dt, end_yy, present_dt, start_dt, start_yy

	StrEbrFile = "ha108oa1_2"

	biz_area_cd = frm1.txtComp_cd.value

	end_dt = UniConvDateToYYYYMMDD(frm1.txtBas_dt.Text,parent.gDateFormat,parent.gServerDateType)
	end_yy = frm1.txtBas_dt.Year
	
	start_dt = UniConvDateToYYYYMMDD(frm1.txtStrt_dt.Text,parent.gDateFormat,parent.gServerDateType)
	start_yy = frm1.txtStrt_dt.Year
	
	condvar = "biz_area_cd|" & biz_area_cd
	condvar = condvar & "|end_dt|" & end_dt
	condvar = condvar & "|end_yy|" & end_yy
	condvar = condvar & "|start_dt|" & start_dt
	condvar = condvar & "|start_yy|" & start_yy

	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
 	call FncEBRPrint(EBAction , ObjName , condvar)
End Function
'======================================================================================================
' Function Name : ReportOfDocuments
' Function Desc : �Ű� ���� 
'=======================================================================================================
Function ReportOfDocuments()
	Dim condvar
	Dim StrEbrFile
    Dim ObjName

	dim biz_area_cd, end_dt, end_yy, present_dt, start_dt, start_yy

	StrEbrFile = "ha108oa1_3"

	biz_area_cd = frm1.txtComp_cd.value

	end_dt = UniConvDateToYYYYMMDD(frm1.txtBas_dt.Text,parent.gDateFormat,parent.gServerDateType)
	end_yy = frm1.txtBas_dt.Year
	
	present_dt = UniConvDateToYYYYMMDD(frm1.txtDt.Text,parent.gDateFormat,parent.gServerDateType)

	start_dt = UniConvDateToYYYYMMDD(frm1.txtStrt_dt.Text,parent.gDateFormat,parent.gServerDateType)
	start_yy = frm1.txtStrt_dt.Year

	condvar = "biz_area_cd|" & biz_area_cd
	condvar = condvar & "|end_dt|" & end_dt
	condvar = condvar & "|end_yy|" & end_yy
	condvar = condvar & "|present_dt|" & present_dt
	condvar = condvar & "|start_dt|" & start_dt
	condvar = condvar & "|start_yy|" & start_yy
	
	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
 	call FncEBRPrint(EBAction , ObjName , condvar)
End Function
'==========================================================================================
'   Event Name : btnCb_autoisrt_OnClick()
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

    If CompareDateByFormat(frm1.txtStrt_dt.Text,frm1.txtBas_dt.Text,frm1.txtStrt_dt.Alt,frm1.txtBas_dt.Alt,"970023",frm1.txtStrt_dt.UserDefinedFormat,parent.gComDateType,True) = False Then
        frm1.txtBas_dt.focus
        Set gActiveElement = document.activeElement
        Exit Function
    End If

    If frm1.vspdData.MaxRows <= 0 And frm1.vspdData1.MaxRows<= 0 And frm1.vspdData2.MaxRows <= 0 Then
		Call DisplayMsgbox("800167", "X","X","X")			                            '��: ��ȸ�� �����ϼ��� 
		Exit Function
    End If

	RetFlag = DisplayMsgbox("900018", parent.VB_YES_NO,"x","x")                                '�� �۾��� ����Ͻðڽ��ϱ�?
	If RetFlag = VBNO Then
		Exit Function
	End IF

    With frm1
    If LayerShowHide(1) = false Then
        Exit Function
    End If
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


'=======================================
'   Event Name : txtDt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================

Sub txtDt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtDt.Action = 7
        frm1.txtDt.focus
    End If
End Sub

Sub txtStrt_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtStrt_dt.Action = 7
        frm1.txtStrt_dt.focus
    End If
End Sub

Sub txtBas_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtBas_dt.Action = 7
        frm1.txtBas_dt.focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtDt_Keypress(Key)
'   Event Desc : enter key down�ÿ� ��ȸ�Ѵ�.
'=======================================================================================================
Sub txtDt_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub

'=======================================================================================================
'   Event Name : txtStrt_dt_Keypress(Key)
'   Event Desc : enter key down�ÿ� ��ȸ�Ѵ�.
'=======================================================================================================
Sub txtStrt_dt_Keypress(Key)
    If Key = 13 Then
        MainQuery()
    End If
End Sub

'=======================================================================================================
'   Event Name : txtBas_dt_Keypress(Key)
'   Event Desc : enter key down�ÿ� ��ȸ�Ѵ�.
'=======================================================================================================
Sub txtBas_dt_Keypress(Key)
    If Key = 13 Then
        MainQuery()
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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>����庰��������Ű�(������)</font></td>
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
		<TD CLASS="Tab11">
			<TABLE <%=LR_SPACE_TYPE_20%>>
			    <TR><TD <%=HEIGHT_TYPE_02%>></TD></TR>
				<TR>
					<TD HEIGHT=20>
						<FIELDSET CLASS="CLSFLD">
						<TABLE <%=LR_SPACE_TYPE_40%>>
			            	<TR>
								<TD CLASS="TD5" NOWRAP>�����ڱ���</TD>
								<TD CLASS="TD6" NOWRAP><SELECT NAME="txtGubun" ALT="�����ڱ���" STYLE="WIDTH: 100px" TAG="12N"></SELECT></TD>
								<TD CLASS="TD5" NOWRAP>���Ⱓ</TD>
								<TD CLASS="TD6" NOWRAP><SELECT NAME="txtGigan" ALT="���Ⱓ" STYLE="WIDTH: 170px" TAG="12N"></SELECT></TD>
							</TR>
						    <TR>
						        <TD CLASS="TD5" NOWRAP>�Ű�����</TD>
								<TD CLASS="TD6" NOWRAP><SELECT NAME="txtComp_cd" ALT="�Ű�����" STYLE="WIDTH: 150px" TAG="12N"></SELECT></TD>
						        <TD CLASS=TD5  NOWRAP>��������</TD>
								<TD CLASS=TD6  NOWRAP><script language =javascript src='./js/ha110ma1_fpDateTime1_txtDt.js'></script></TD>
							</TR>
							<TR>
							    <TD CLASS=TD5  NOWRAP>�����븮�ΰ�����ȣ</TD>
								<TD CLASS=TD6  NOWRAP><INPUT TYPE=TEXT ID="txtBizAreaCD" MAXLENGTH=6 NAME="txtSer" SIZE=15 tag="11XXX" ALT="�����븮�ΰ�����ȣ"></TD>
								<TD CLASS=TD5  NOWRAP>�����Ⱓ</TD>
								<TD CLASS=TD6  NOWRAP><script language =javascript src='./js/ha110ma1_fpDateTime2_txtStrt_dt.js'></script>&nbsp;~&nbsp;
								                      <script language =javascript src='./js/ha110ma1_fpDateTime3_txtBas_dt.js'></script></TD>
							</TR>
							    <INPUT TYPE=HIDDEN ID="txtFile" NAME="txtFile" SIZE=15 tag="14XXXU" ALT="�������ϰ��">
						</TABLE>
						</FIELDSET>
					</TD>
				</TR>

				<TR><TD <%=HEIGHT_TYPE_03%>></TD></TR>
				<TR >
					<TD WIDTH="100%" HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
            			    <TR >
            					<TD WIDTH="50%" HEIGHT=*>
	            					<TABLE WIDTH="100%" HEIGHT="100%">
	            						<TR>
	            							<TD HEIGHT="100%"><script language =javascript src='./js/ha110ma1_vaSpread_vspdData.js'></script></TD>
		               					</TR>
					            	</TABLE>
		            			</TD>
            					<TD WIDTH="50%" HEIGHT=*>
	            					<TABLE WIDTH="100%" HEIGHT="100%">
	            						<TR>
	            							<TD HEIGHT="100%"><script language =javascript src='./js/ha110ma1_vaSpread1_vspdData1.js'></script></TD>
		               					</TR>
					            	</TABLE>
		            			</TD>
                            </TR>
                            <TR HEIGHT="70%">
				            	<TD WIDTH="100%" HEIGHT=* VALIGN=TOP COLSPAN=3>
				            		<TABLE WIDTH="100%" HEIGHT="100%">
				            			<TR>
				            				<TD HEIGHT="100%"><script language =javascript src='./js/ha110ma1_vaSpread2_vspdData2.js'></script></TD>
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

	<TR HEIGHT=20>
	    <TD>
	        <TABLE <%=LR_SPACE_TYUPE_30%>>
	            <TR>
	                <TD WIDTH=10>&nbsp;</TD>
	                <TD>
	                <!--<BUTTON NAME="btnCb_print2"		CLASS="CLSMBTN" Flag=1>������ǥ�����</BUTTON>&nbsp; -->
	                 <!--<BUTTON NAME="btnCb_print"		CLASS="CLSMBTN" Flag=1>����ǥ���</BUTTON>&nbsp; -->
	                    <BUTTON NAME="btnCb_creation"	CLASS="CLSMBTN" Flag=1>���ϻ���</BUTTON></TD>
	                <TD WIDTH=* ALIGN="right"></TD>
	                <TD WIDTH=10>&nbsp;</TD>
	            </TR>
	        </TABLE>
	    </TD>
	</TR>

	<TR>
		<TD width=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
		<TD WIDTH=100% HEIGHT=<%=BizSize%>><IFRAME NAME="MyBizASP1" SRC="../../blank.htm" WIDTH=100% HEIGHT=<%=BizSize%> FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>

</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' ����ó��ASP�� �ѱ�� ���� ������ ��� �ִ� Tag�� %>
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24"><INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24"><INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="hPayCd" tag="24">
</FORM>
<DIV ID="MousePT" NAME="MousePT">
<iframe name="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 width=220 height=41 src="../../inc/cursor.htm"></iframe>
</DIV>
<FORM NAME="EBAction" TARGET = "MyBizASP" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">
</FORM>
<FORM NAME="EBAction1" TARGET = "MyBizASP1" METHOD="POST">
	<INPUT TYPE="HIDDEN" NAME="uname">
	<INPUT TYPE="HIDDEN" NAME="dbname">
	<INPUT TYPE="HIDDEN" NAME="filename">
	<INPUT TYPE="HIDDEN" NAME="condvar">
	<INPUT TYPE="HIDDEN" NAME="date">
</FORM>
</BODY>
</HTML>

