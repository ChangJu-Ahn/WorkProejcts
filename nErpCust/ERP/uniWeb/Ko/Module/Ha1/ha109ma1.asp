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
Const BIZ_PGM_ID    = "Ha109mb1.asp"                                      'Biz Logic ASP 
Const BIZ_PGM_ID2   = "ha109mb2.asp"                                 '��: File Creation Asp Name
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
Dim StartDate
Dim lgStrPrevKey1,lgStrPrevKey2
Dim topleftOK

StartDate	= "<%=GetSvrDate%>"

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
    frm1.txtsubmit_Dt.Text =  UniConvDateAToB(StartDate ,parent.gServerDateFormat,parent.gDateFormat)
    'frm1.txtretire_dt1.Text = UniConvDateAToB(StartDate,parent.gServerDateFormat,parent.gDateFormat)
	frm1.txtretire_dt2.Text =  UniConvDateAToB(StartDate ,parent.gServerDateFormat,parent.gDateFormat)

    Dim strYear,strMonth,strDay
    Call ExtractDateFrom("<%=GetsvrDate%>",parent.gServerDateFormat , parent.gServerDateType ,strYear,strMonth,strDay)
 
    frm1.txtretire_dt1.year = strYear

    frm1.txtretire_dt1.month = "1"
    frm1.txtretire_dt1.day = "1" 	
    
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
' Function Name : MakeKeyStream
' Function Desc : This method set focus to pos of err
'========================================================================================================
Sub MakeKeyStream(pRow)
	  
   Dim strSubmitDt, strretire_dt1, strretire_dt2
   
  	dim bas_dt
  	dim send_dt
  	
   
   strSubmitDt      = frm1.txtsubmit_Dt.Year & right("0" & frm1.txtsubmit_Dt.Month,2) & right("0" & frm1.txtsubmit_Dt.Day,2)
   strretire_dt1 = frm1.txtretire_dt1.Year & right("0" & frm1.txtretire_dt1.Month,2) & right("0" & frm1.txtretire_dt1.Day,2)
   strretire_dt2  = frm1.txtretire_dt2.Year & right("0" & frm1.txtretire_dt2.Month,2) & right("0" & frm1.txtretire_dt2.Day,2)

  	bas_dt = frm1.txtretire_dt2.year & right("00" & frm1.txtretire_dt2.month,2) & right("00" & frm1.txtretire_dt2.Day,2) 
	send_dt  = frm1.txtsubmit_Dt.year & right("00" & frm1.txtsubmit_Dt.month,2) & right("00" & frm1.txtsubmit_Dt.Day,2)

  lgKeyStream=""
  	lgKeyStream       = Trim(frm1.txtFile.value) & parent.gColSep					'���ϸ� 
	lgKeyStream       = lgKeyStream & Trim(frm1.txtComp_cd.value) & parent.gColSep	'�Ű����� 
	lgKeyStream       = lgKeyStream & Trim(strSubmitDt) & parent.gColSep				'���⿬���� 
	lgKeyStream       = lgKeyStream & Trim(strretire_dt1) & parent.gColSep					'�Ⱓ1
	lgKeyStream       = lgKeyStream & Trim(strretire_dt2) & parent.gColSep					'�Ⱓ1

	IF (frm1.txtComp_type1.checked = True) Then '�����Ű��̸� ���õ� ����� �ڵ�� 
		lgKeyStream       = lgKeyStream & Trim(frm1.txtComp_cd.value) & parent.gColSep
	Else
		lgKeyStream       = lgKeyStream & "Y"  & parent.gColSep           '���սŰ��̸� ��ü "%" �� 
	End If	  
	

	lgKeyStream       = lgKeyStream & Trim(Frm1.txtGubun.value) & parent.gColSep	'�����ڱ��� 
	lgKeyStream       = lgKeyStream & Trim(frm1.txtGigan.value) & parent.gColSep	'���Ⱓ 
	lgKeyStream       = lgKeyStream & Trim(frm1.txtSer.value) & parent.gColSep		'�����븮�ΰ�����ȣ 
	

	 
	 
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
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet(ByVal pvSpdNo)

    If pvSpdNo = "" OR pvSpdNo = "A" Then

    	With frm1.vspdData
                ggoSpread.Source = frm1.vspdData
                ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk

    	       .ReDraw = false

               .MaxCols = 17 + 1										'��: �ִ� Columns�� �׻� 1�� ������Ŵ %>	   
    	       .Col = .MaxCols                                                              ' ��:��: Hide maxcols
               .ColHidden = True                                                            ' ��:��:

               .MaxRows = 0
                ggoSpread.ClearSpreadData

                Call AppendNumberPlace("6","2","0")

                ggoSpread.SSSetEdit     1,		"���ڵ屸��",           12
                ggoSpread.SSSetEdit     2,		"�ڷᱸ��",             10
                ggoSpread.SSSetEdit     3,		"������",               8
                ggoSpread.SSSetEdit     4,		"���⿬����",           12
                ggoSpread.SSSetEdit     5,		"������(�븮�α���)",   19
                ggoSpread.SSSetEdit     6,		"�����븮�ΰ�����ȣ",   20
				ggoSpread.SSSetEdit     7,		"Ȩ�ؽ�ID",					20	'2004 
				ggoSpread.SSSetEdit     8,		"�������α׷��ڵ�",			45	'2004 	                
                ggoSpread.SSSetEdit     9,      "����ڵ�Ϲ�ȣ",       16
                ggoSpread.SSSetEdit     10,     "���θ�(��ȣ)",         13
				ggoSpread.SSSetEdit     11,		"����ںμ�",				30	'2004 	
				ggoSpread.SSSetEdit     12,		"����ڼ���",				30	'2004 	
				ggoSpread.SSSetEdit     13,		"�������ȭ��ȣ",			15	'2004 
                ggoSpread.SSSetEdit     14,		"�Ű��ǹ���(B���ڵ�)��",22
                ggoSpread.SSSetEdit     15,		"�ѱ��ڵ�����",         14
                ggoSpread.SSSetEdit     16,		"������Ⱓ�ڵ�",     18
                ggoSpread.SSSetEdit     17,		"����",                 8

    	    	.ReDraw = true	
        End With
	    
        Call SetSpreadLock("A")
            
    End If
    
    If pvSpdNo = "" OR pvSpdNo = "B" Then

    	With frm1.vspdData1
                ggoSpread.Source = frm1.vspdData1
                ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk

    	       .ReDraw = false

               .MaxCols = 17 + 1										'��: �ִ� Columns�� �׻� 1�� ������Ŵ %>	   
    	       .Col = .MaxCols                                                              ' ��:��: Hide maxcols
               .ColHidden = True                                                            ' ��:��:

               .MaxRows = 0
                ggoSpread.ClearSpreadData

                Call AppendNumberPlace("6","2","0")

                ggoSpread.SSSetEdit     1,	"���ڵ屸��",               12
                ggoSpread.SSSetEdit     2,	"�ڷᱸ��",                 10
                ggoSpread.SSSetEdit     3,	"������",                   8
                ggoSpread.SSSetEdit     4,	"�Ϸù�ȣ",                 10
                ggoSpread.SSSetEdit     5,	"����ڵ�Ϲ�ȣ",           16
                ggoSpread.SSSetEdit     6,	"���θ�(��ȣ)",             13
                ggoSpread.SSSetEdit     7,	"��ǥ��(����)",             13
                ggoSpread.SSSetEdit     8,	"�ֹ�(����)��Ϲ�ȣ",		19
                ggoSpread.SSSetEdit     9,	"��(��)����Ǽ�(C���ڵ��)",25
                ggoSpread.SSSetEdit     10,	"��(��)����Ǽ�(D���ڵ��)",25
                ggoSpread.SSSetEdit     11,	"�����޿����Ѱ�",             14
                ggoSpread.SSSetEdit     12,	"�ҵ���������Ѱ�",         18
                ggoSpread.SSSetEdit     13,	"���ΰ��������Ѱ�",         18
                ggoSpread.SSSetEdit     14,	"�ֹΰ��������Ѱ�",         18
                ggoSpread.SSSetEdit     15,	"��Ư�����������Ѱ�",       20
                ggoSpread.SSSetEdit     16,	"���������Ѱ�",             14
                ggoSpread.SSSetEdit     17,	"����",                     8

    	    	.ReDraw = true	
        End With
	    
        Call SetSpreadLock("B")
            
    End If
    
    If pvSpdNo = "" OR pvSpdNo = "C" Then

    	With frm1.vspdData2
                ggoSpread.Source = frm1.vspdData2
                ggoSpread.Spreadinit "V20021121",,parent.gAllowDragDropSpread    'sbk

    	       .ReDraw = false

               .MaxCols = 87 + 1										'��: �ִ� Columns�� �׻� 1�� ������Ŵ %>	   
    	       .Col = .MaxCols                                                  ' ��:��: Hide maxcols
               .ColHidden = True                                                ' ��:��:

               .MaxRows = 0
                ggoSpread.ClearSpreadData

                Call AppendNumberPlace("6","2","0")

                ggoSpread.SSSetEdit     1,	"���ڵ屸��",               12
                ggoSpread.SSSetEdit     2,	"�ڷᱸ��",                 10
                ggoSpread.SSSetEdit     3,	"������",                   8
                ggoSpread.SSSetEdit     4,	"�Ϸù�ȣ",                 10
                ggoSpread.SSSetEdit     5,	"����ڵ�Ϲ�ȣ",           16
                ggoSpread.SSSetEdit     6,	"��(��)�ٹ�ó��",           15
                ggoSpread.SSSetEdit     7,	"�����ڱ����ڵ�",           16  
                ggoSpread.SSSetEdit     8 ,	"���������ڵ�",             12
                ggoSpread.SSSetEdit     9,	"�ͼӿ������ۿ�����",       18
                ggoSpread.SSSetEdit     10,	"�ͼӿ������Ῥ����",       18
                ggoSpread.SSSetEdit     11,	"����",                     8
                ggoSpread.SSSetEdit     12,	"��/�ܱ��α����ڵ�",        17
                ggoSpread.SSSetEdit     13,	"�ֹε�Ϲ�ȣ",        15
                ggoSpread.SSSetEdit     14,	"�����޿�",                 10
                ggoSpread.SSSetEdit     15,	"������(�߰�������)", 24
                ggoSpread.SSSetEdit     16,	"���������Ͻñ�",           24
                ggoSpread.SSSetEdit     17,	"��",                       12
                ggoSpread.SSSetEdit     18,	"������ҵ�",               12
                
                ggoSpread.SSSetEdit     19,	"�Ѽ��ɾ�",               12
                ggoSpread.SSSetEdit     20,	"�������հ��",               12
                ggoSpread.SSSetEdit     21,	"�ҵ��ں��Ծ�",               12
                ggoSpread.SSSetEdit     22,	"�������ݼҵ������",               20
                ggoSpread.SSSetEdit     23,	"���������Ͻñ�",               12
                '-------------����ȯ���--------------------------------------------------
                ggoSpread.SSSetEdit     24,	"���������Ͻñ����޿����",               12
                ggoSpread.SSSetEdit     25,	"���Ͻñ�",               12
                ggoSpread.SSSetEdit     26,	"���ɰ��������޿���",               22
                ggoSpread.SSSetEdit     27,	"ȯ�������ҵ����",               22
                ggoSpread.SSSetEdit     28,	"ȯ�������ҵ��������ǥ��",               25
                ggoSpread.SSSetEdit     29,	"ȯ�꿬��հ���ǥ��",               20
                ggoSpread.SSSetEdit     30,	"ȯ�꿬��ջ��⼼��",               20
                
                
                ggoSpread.SSSetEdit     31,	"���������Ͻñ����޿����",               12
                ggoSpread.SSSetEdit     32,	"���Ͻñ�",               12
                ggoSpread.SSSetEdit     33,	"���ɰ��������޿���",               22
                ggoSpread.SSSetEdit     34,	"ȯ�������ҵ����",               22
                ggoSpread.SSSetEdit     35,	"ȯ�������ҵ��������ǥ��",               25
                ggoSpread.SSSetEdit     36,	"ȯ�꿬��հ���ǥ��",               20
                ggoSpread.SSSetEdit     37,	"ȯ�꿬��ջ��⼼��",               20
                '---------------�ټӿ���------------------------------------------------
                ggoSpread.SSSetEdit     38,	"��(��)�ٹ����Ի翬����",               20
                ggoSpread.SSSetEdit     39,	"��(��)�ٹ�����翬����",               20
                ggoSpread.SSSetEdit     40,	"��(��)�ٹ����ټӿ���",               20
                ggoSpread.SSSetEdit     41,	"��(��)�ٹ������ܿ���",               20
                ggoSpread.SSSetEdit     42,	"��(��)�ٹ����Ի翬����",               20
                ggoSpread.SSSetEdit     43,	"��(��)��翬����",               20
                ggoSpread.SSSetEdit     44,	"��(��)�ټӿ���",               20
                ggoSpread.SSSetEdit     45,	"��(��)���ܿ���",               20
                ggoSpread.SSSetEdit     46,	"�ߺ�����",                 10
                ggoSpread.SSSetEdit     47,	"�ټӿ���",                 10
                
                ggoSpread.SSSetEdit     48,	"��(��)�ٹ����Ի翬����",               20
                ggoSpread.SSSetEdit     49,	"��(��)�ٹ�����翬����",               20
                ggoSpread.SSSetEdit     50,	"��(��)�ٹ����ټӿ���",               20
                ggoSpread.SSSetEdit     51,	"��(��)�ٹ������ܿ���",               20
                ggoSpread.SSSetEdit     52,	"��(��)�ٹ����Ի翬����",               20
                ggoSpread.SSSetEdit     53,	"��(��)��翬����",               20
                ggoSpread.SSSetEdit     54,	"��(��)�ټӿ���",               20
                ggoSpread.SSSetEdit     55,	"��(��)���ܿ���",               20
                ggoSpread.SSSetEdit     56,	"�ߺ�����",                 10
                ggoSpread.SSSetEdit     57,	"�ټӿ���",                 10
                                        
                '---------------�����  ------------------------------------------------
                
                
                ggoSpread.SSSetEdit     58,	"�����޿���",               12
                ggoSpread.SSSetEdit     59,	"�����ҵ����",           16
                ggoSpread.SSSetEdit     60,	"�����ҵ����ǥ��",         18
                ggoSpread.SSSetEdit     61,	"����հ���ǥ��",           16
                ggoSpread.SSSetEdit     62,	"����ջ��⼼��",           16
                ggoSpread.SSSetEdit     63,	"���⼼��",                 10
                ggoSpread.SSSetEdit     64,	"��������",                 10
                                          
                ggoSpread.SSSetEdit     65,	"�����޿���",               12
                ggoSpread.SSSetEdit     66,	"�����ҵ����",           16
                ggoSpread.SSSetEdit     67,	"�����ҵ����ǥ��",         18
                ggoSpread.SSSetEdit     68,	"����հ���ǥ��",           16
                ggoSpread.SSSetEdit     69,	"����ջ��⼼��",           16
                ggoSpread.SSSetEdit     70,	"���⼼��",                 10
                ggoSpread.SSSetEdit     71,	"��������",                 10
                                        
                '---------------��------  ------------------------------------------
                
                
                ggoSpread.SSSetEdit     72,	"�����޿���",               12
                ggoSpread.SSSetEdit     73,	"�����ҵ����",           16
                ggoSpread.SSSetEdit     74,	"�����ҵ����ǥ��",         18
                ggoSpread.SSSetEdit     75,	"����հ���ǥ��",           16
                ggoSpread.SSSetEdit     76,	"����ջ��⼼��",           16
                ggoSpread.SSSetEdit     77,	"���⼼��",                 10
                ggoSpread.SSSetEdit     78,	"��������",                 10
                
                 '---------------��������------------------------------------------------
                 
                ggoSpread.SSSetEdit     79, "�ҵ漼",                   8
                ggoSpread.SSSetEdit     80,	"�ֹμ�",                   8
                ggoSpread.SSSetEdit     81,	"��Ư��",             		14
                ggoSpread.SSSetEdit     82,	"��",                       12
                                          
                '---------------�ⳳ�μ��� ------------------------------------------------
                
                ggoSpread.SSSetEdit     83,	"�ҵ漼",                   8
                ggoSpread.SSSetEdit     84,	"�ֹμ�",                   8
                ggoSpread.SSSetEdit     85,	"��Ư��",             		14
                ggoSpread.SSSetEdit     86,	"��",                       12
                                          
                ggoSpread.SSSetEdit     87,	"����",                     12


    	    	.ReDraw = true	
        End With
        
        Call SetSpreadLock("C")
            
    End If
    
    
      If pvSpdNo = "" OR pvSpdNo = "D" Then

    	With frm1.vspdData3
                ggoSpread.Source = frm1.vspdData3
                ggoSpread.Spreadinit "V20070121",,parent.gAllowDragDropSpread    'sbk

    	       .ReDraw = false

               .MaxCols = 21 + 1										'��: �ִ� Columns�� �׻� 1�� ������Ŵ %>	   
    	       .Col = .MaxCols                                                              ' ��:��: Hide maxcols
               .ColHidden = True                                                            ' ��:��:

               .MaxRows = 0
                ggoSpread.ClearSpreadData

                Call AppendNumberPlace("6","2","0")

                ggoSpread.SSSetEdit     1,	"���ڵ屸��",               12
                ggoSpread.SSSetEdit     2,	"�ڷᱸ��",                 10
                ggoSpread.SSSetEdit     3,	"������",                   8
                ggoSpread.SSSetEdit     4,	"�Ϸù�ȣ",                 10
                ggoSpread.SSSetEdit     5,	"����ڵ�Ϲ�ȣ",           16
                ggoSpread.SSSetEdit     6,	"����",						13
                ggoSpread.SSSetEdit     7,	"�ҵ����ֹι�ȣ",           13
                ggoSpread.SSSetEdit     8,	"�ٹ�ó��",					19
                ggoSpread.SSSetEdit     9,	"����ڵ�Ϲ�ȣ",			25
                ggoSpread.SSSetEdit     10,	"�����޿�",					10
                ggoSpread.SSSetEdit     11,	"�������޿�(�߰������޿�)",30
                ggoSpread.SSSetEdit     12,	"���������Ͻñ�",			18
                ggoSpread.SSSetEdit     13,	"��",						15
                ggoSpread.SSSetEdit     14,	"������ҵ�",				15
                ggoSpread.SSSetEdit     15,	"�Ѽ��ɾ�",					15
                ggoSpread.SSSetEdit     16,	"�������հ��",             14
                ggoSpread.SSSetEdit     17,	"�ҵ��ں��Ծ�",             18
                ggoSpread.SSSetEdit     18,	"�������ݼҵ������",       18
                ggoSpread.SSSetEdit     19,	"���������Ͻñ�",           18
                ggoSpread.SSSetEdit     20,	"��(��)�ٹ�ó�Ϸù�ȣ",     18
                ggoSpread.SSSetEdit     21,	"����",						18

    	    	.ReDraw = true	
        End With
	    
        Call SetSpreadLock("D")
            
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
                ggoSpread.SpreadLock      -1,-1,-1
                ggoSpread.SSSetProtected  .vspdData2.MaxCols   , -1, -1
            .vspdData2.ReDraw = True
           End With
       
          Case  "D"
            With frm1 
            .vspdData2.ReDraw = False
                ggoSpread.SpreadLock      -1,-1,-1
                ggoSpread.SSSetProtected  .vspdData3.MaxCols   , -1, -1
            .vspdData2.ReDraw = True
           End With
               
    End Select
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

    Call InitVariables                                                           '��: Initializes local global variables

    If CompareDateByFormat(frm1.txtretire_dt1.Text,frm1.txtretire_dt2.Text,frm1.txtretire_dt1.Alt,frm1.txtretire_dt2.Alt,"970023",frm1.txtretire_dt1.UserDefinedFormat,parent.gComDateType,True) = False Then
        frm1.txtretire_dt2.focus()
        Set gActiveElement = document.activeElement
        Exit Function
    End If

    If CompareDateByFormat(frm1.txtretire_dt2.Text,frm1.txtsubmit_Dt.Text,frm1.txtretire_dt2.Alt,frm1.txtsubmit_Dt.Alt,"970023",frm1.txtretire_dt2.UserDefinedFormat,parent.gComDateType,True) = False Then
        frm1.txtsubmit_Dt.focus()
        Set gActiveElement = document.activeElement
        Exit Function
    End If
    
    
    

    if frm1.txtretire_dt1.year <> frm1.txtretire_dt2.year then
		call DisplayMsgBox("971012","X", "�����Ⱓ","X")
		frm1.txtretire_dt1.focus
		exit function
    end if
    
    
    
    lgCurrentSpd = "A"
	topleftOK = false    
    Call MakeKeyStream(lgCurrentSpd)

    If DbQuery = False Then  
		Exit Function
	End If                                                               '��: Query db data
       
    FncQuery = True                                                              '��: Processing is OK
    
End Function


'========================================================================================================
' Function Name : FncCopy
' Function Desc : This function is related to Copy Button of Main ToolBar
'========================================================================================================
Function FncCopy()

    If Frm1.vspdData.MaxRows < 1 Then
       Exit Function
    End If
    
    ggoSpread.Source = Frm1.vspdData
	With Frm1.VspdData
         .ReDraw = False
		 If .ActiveRow > 0 Then
            ggoSpread.CopyRow
			SetSpreadColor .ActiveRow
    
            .ReDraw = True
		    .Focus
		 End If
	End With
	
    Set gActiveElement = document.ActiveElement   

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
		case "vaSpread3"
			Call InitSpreadSheet("D") 
				    		
	end select 
    
	Call ggoSpread.ReOrderingSpreadData()

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
' Name : DbDelete
' Desc : This function is called by FncDelete
'========================================================================================================
Function DbDelete()
    Dim IntRetCd
    
    FncDelete = False                                                      '��: Processing is NG
    
    If lgIntFlgMode <> parent.OPMD_UMODE Then                                      'Check if there is retrived data
        Call DisplayMsgbox("900002","X","X","X")                                '��:
        Exit Function
    End If
    
    IntRetCD = DisplayMsgbox("900003", parent.VB_YES_NO,"X","X")		            '��: "Will you destory previous data"
	If IntRetCD = vbNo Then													'------ Delete function call area ------ 
		Exit Function	
	End If
    
    Call DbDelete															'��: Delete db data
    
    FncDelete = True                                                        '��: Processing is OK

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
' Function Name : DbSaveOk
' Function Desc : Called by MB Area when save operation is successful
'========================================================================================================
Function DbSaveOk()

    Call ggoOper.ClearField(Document, "2")										'��: Clear Contents  Field
    
    Call InitVariables															'��: Initializes local global variables
	Call DisableToolBar(parent.TBC_QUERY)
	If DBQuery = False Then
		Call RestoreToolBar()
		Exit Function
	End If
End Function

'========================================================================================================
' Function Name : DbDeleteOk
' Function Desc : Called by MB Area when delete operation is successful
'========================================================================================================
Function DbDeleteOk()

End Function

'===============================================================================
' Function Name : ValidDateCheckThisForm
' Function Desc : Valid Date Check Function
'===============================================================================

Function ValidDateCheckThisForm(ThisObjFromDt, ThisObjToDt)

	ValidDateCheckThisForm = False

	If Len(Trim(ThisObjToDt.Text)) And Len(Trim(ThisObjFromDt.Text)) Then
		If ValidDateCheck(ThisObjFromDt,ThisObjToDt) =False Then
			Exit Function
		End If
	End If

	ValidDateCheckThisForm = True

End Function

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
    
    If Not(ValidDateCheckThisForm(frm1.txtretire_dt1, frm1.txtretire_dt2)) Then        
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
    
    If Not(ValidDateCheckThisForm(frm1.txtretire_dt1, frm1.txtretire_dt2)) Then        
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
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile    
	Dim ObjName
	
	dim biz_area_cd, end_dt, end_yy, present_dt, start_dt, start_yy
	
	StrEbrFile = "ha109oa1_1"
	
	biz_area_cd = frm1.txtComp_cd.value
	end_dt		= UniConvDateToYYYYMMDD(frm1.txtretire_dt2.Text,parent.gDateFormat,parent.gServerDateType)
	end_yy		= frm1.txtretire_dt2.Year
	present_dt	= UniConvDateToYYYYMMDD(frm1.txtsubmit_Dt.Text,parent.gDateFormat,parent.gServerDateType)
	start_dt	= UniConvDateToYYYYMMDD(frm1.txtretire_dt1.Text,parent.gDateFormat,parent.gServerDateType)		
	start_yy	= frm1.txtretire_dt1.Year	
	
	condvar = "year_area_cd|" & biz_area_cd
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
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile
    Dim ObjName

	dim end_dt, end_yy, present_dt, start_dt, start_yy, year_area_cd
	
	StrEbrFile = "ha109oa1_2"		
	
	end_dt		= UniConvDateToYYYYMMDD(frm1.txtretire_dt2.Text,parent.gDateFormat,parent.gServerDateType)
	end_yy		= frm1.txtretire_dt2.Year	
	start_dt	= UniConvDateToYYYYMMDD(frm1.txtretire_dt1.Text,parent.gDateFormat,parent.gServerDateType)		
	start_yy	= frm1.txtretire_dt1.Year	
	year_area_cd = frm1.txtComp_cd.value
	
	condvar = "end_dt|" & end_dt 
	condvar = condvar & "|end_yy|" & end_yy 	
	condvar = condvar & "|start_dt|" & start_dt 
	condvar = condvar & "|start_yy|" & start_yy 	
	condvar = condvar & "|year_area_cd|" & year_area_cd	

	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
 	call FncEBRPrint(EBAction , ObjName , condvar)
	
End Function
'======================================================================================================
' Function Name : ReportOfDocuments
' Function Desc : �Ű� ���� 
'=======================================================================================================
Function ReportOfDocuments()
	Dim condvar
	Dim lngPos
	Dim intCnt
    Dim StrEbrFile
    Dim ObjName    	

	dim end_dt, end_yy, present_dt, start_dt, start_yy, year_area_cd
	
	StrEbrFile = "ha109oa1_3"	
	
	end_dt		= UniConvDateToYYYYMMDD(frm1.txtretire_dt2.Text,parent.gDateFormat,parent.gServerDateType)
	end_yy		= frm1.txtretire_dt2.Year
	present_dt	= UniConvDateToYYYYMMDD(frm1.txtsubmit_Dt.Text,parent.gDateFormat,parent.gServerDateType)
	start_dt	= UniConvDateToYYYYMMDD(frm1.txtretire_dt1.Text,parent.gDateFormat,parent.gServerDateType)		
	start_yy	= frm1.txtretire_dt1.Year	
	year_area_cd = frm1.txtComp_cd.value

	condvar = "end_dt|" & end_dt  
	condvar = condvar & "|end_yy|" & end_yy 
	condvar = condvar & "|present_dt|" & present_dt
	condvar = condvar & "|start_dt|" & start_dt 
	condvar = condvar & "|start_yy|" & start_yy 	
	condvar = condvar & "|year_area_cd|" & year_area_cd
	
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
    
    If Not(ValidDateCheckThisForm(frm1.txtretire_dt1, frm1.txtretire_dt2)) Then        
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
		strVal = strVal & "&txtFileName=" & frm1.txtFile.value 
		
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
'   Event Name : txtsubmit_Dt_DblClick(Button)
'   Event Desc : �޷��� ȣ���Ѵ�.
'=======================================

Sub txtsubmit_Dt_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtsubmit_Dt.Action = 7
        frm1.txtsubmit_Dt.focus
    End If
End Sub

Sub txtretire_dt1_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtretire_dt1.Action = 7
        frm1.txtretire_dt1.focus
    End If
End Sub

Sub txtretire_dt2_DblClick(Button) 
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtretire_dt2.Action = 7
        frm1.txtretire_dt2.focus
    End If
End Sub

'=======================================================================================================
'   Event Name : txtsubmit_Dt_Keypress(Key)
'   Event Desc : enter key down�ÿ� ��ȸ�Ѵ�.
'=======================================================================================================
Sub txtsubmit_Dt_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

'=======================================================================================================
'   Event Name : txtretire_dt1_Keypress(Key)
'   Event Desc : enter key down�ÿ� ��ȸ�Ѵ�.
'=======================================================================================================
Sub txtretire_dt1_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

'=======================================================================================================
'   Event Name : txtretire_dt2_Keypress(Key)
'   Event Desc : enter key down�ÿ� ��ȸ�Ѵ�.
'=======================================================================================================
Sub txtretire_dt2_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

Sub txtSer_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
End Sub

</SCRIPT>
<!-- #Include file="../../inc/uni2kcm.inc" -->	
<SCRIPT LANGUAGE="JavaScript">
<!-- Hide script from old browsers

function setCookie(name, value, expire)
{
	document.cookie = name + "=" + escape(value)
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/bin"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
	document.cookie = name + "=" + escape(value)
		+ "; path=/EasyBaseWeb/lib"
		+ ((expire == null) ? "" : ("; expires=" + expire.toGMTString()))
}

setCookie("client", "-1", null)
setCookie("owner", "admin", null)
setCookie("identity", "admin", null)
-->
</SCRIPT>

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
								<td background="../../../CShared/image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>��������Ű�</font></td>
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
								<TD CLASS="TD6" NOWRAP><SELECT NAME="txtComp_cd" ALT="�Ű�����" STYLE="WIDTH: 150px" TAG="12N"></SELECT><br>
														<INPUT TYPE="RADIO" CLASS="RADIO" ID="txtComp_type1" NAME="txtComp_type" TAG="21X" VALUE="Y" CHECKED><LABEL FOR="txtComp_type1">����尳���Ű�</LABEL>
														<INPUT TYPE="RADIO" CLASS="RADIO" ID="txtComp_type2" NAME="txtComp_type" TAG="21X" VALUE="N"><LABEL FOR="txtComp_type2">��������սŰ�</LABEL>								
								</TD>
						        <TD CLASS=TD5  NOWRAP>��������</TD>
								<TD CLASS=TD6  NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtsubmit_Dt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="��������" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>															
							</TR>
							<TR>
							    <TD CLASS=TD5  NOWRAP>�����븮�ΰ�����ȣ</TD>
								<TD CLASS=TD6  NOWRAP><INPUT TYPE=TEXT ID="txtSer" MAXLENGTH=6 NAME="txtSer" SIZE=15 tag="11XXX" ALT="�����븮�ΰ�����ȣ">
								                      <INPUT TYPE=HIDDEN ID="txtFile" NAME="txtFile" SIZE=15 tag="14XXXU" ALT="�������ϰ��"></TD>
								<TD CLASS=TD5  NOWRAP>�����Ⱓ</TD>
								<TD CLASS=TD6  NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtretire_dt1 CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="���������Ⱓ" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT>&nbsp;~&nbsp;
								                      <SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime3 name=txtretire_dt2 CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="���������Ⱓ" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>
							
														
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
	            							<TD HEIGHT="100%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT></TD>
		               					</TR>
					            	</TABLE>
		            			</TD>
            					<TD WIDTH="50%" HEIGHT=*>
	            					<TABLE WIDTH="100%" HEIGHT="100%">
	            						<TR>
	            							<TD HEIGHT="100%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData1 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread1> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT></TD>
		               					</TR>
					            	</TABLE>
		            			</TD>
                            </TR>  
                            <TR HEIGHT="70%">
				            	<TD WIDTH="100%" HEIGHT=* VALIGN=TOP COLSPAN=3>
				            		<TABLE WIDTH="100%" HEIGHT="100%">
				            			<TR>
				            				<TD HEIGHT="50%">
				            				<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread2> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
				            				
				            				</TD>
					            		</TR>
					            			<TR>
				            				<TD HEIGHT="50%">
				            				<SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData3 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread3> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT>
				            				
				            				</TD>
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
	                <TD><!--BUTTON NAME="btnCb_print2"		CLASS="CLSMBTN"	Flag=1>������ǥ�����</BUTTON>&nbsp;
	                    <BUTTON NAME="btnCb_print"		CLASS="CLSMBTN"	Flag=1>����ǥ���</BUTTON>&nbsp; -->
	                    <BUTTON NAME="btnCb_creation"	CLASS="CLSMBTN"	Flag=1>���ϻ���</BUTTON></TD>
	                <TD WIDTH=* ALIGN="right"></TD>
	                <TD WIDTH=10>&nbsp;</TD>
	            </TR>
	        </TABLE>
	    </TD>
	</TR>	
	
	<TR>
		<TD width=100% HEIGHT=0><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=10 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
		<TD WIDTH=100% HEIGHT=0><IFRAME NAME="MyBizASP1" SRC="../../blank.htm" WIDTH=100% HEIGHT=10 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</TD>
	</TR>
	

	
</TABLE>
<TEXTAREA CLASS="hidden" NAME="txtSpread" tag="24"></TEXTAREA>
<INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24">
<INPUT TYPE=HIDDEN NAME="txtMode" tag="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId" tag="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" tag="24">
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

