<%@ LANGUAGE=VBSCript %>
<!--
======================================================================================================
*  1. Module Name          : 연말정산 
*  2. Function Name        : 사업장연말정산신고(퇴직자 포함)
*  3. Program ID           : H9117ma1
*  4. Program Name         : 사업장연말정산신고 
*  5. Program Desc         : 사업장연말정산신고(퇴직자 포함)
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
<!-- #Include file="../../inc/IncSvrHTML.inc" -->

<LINK REL="stylesheet" TYPE="Text/css" HREF="../../inc/SheetStyle.css">		

<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAMain.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAEvent.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliVariables.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incCliMAOperation.vbs"></SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/IncCliRdsQuery.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/IncEB.vbs">   </SCRIPT>
<SCRIPT LANGUAGE="VBScript"   SRC="../../inc/incHRQuery.vbs"></SCRIPT>
<Script Language="VBScript">
Option Explicit 
'========================================================================================================
'=                       4.1 External ASP File
'========================================================================================================

Const BIZ_PGM_ID      = "h9117mb1.asp"						           '☆: Biz Logic ASP Name
Const BIZ_PGM_ID2     = "h9117mb2.asp"                                 '☆: File Creation Asp Name
Const C_SHEETMAXROWS    = 10                                      '☜: Visble row
Const C_SHEETMAXROWS1    = 10
Const C_SHEETMAXROWS2    = 10	                                      '☜: Visble row
Const C_SHEETMAXROWS3    = 10	                                      '☜: Visble row
Const C_SHEETMAXROWS4    = 10	 

'========================================================================================================
'=                       4.3 Common variables 
'========================================================================================================
<!-- #Include file="../../inc/lgvariables.inc" -->	
'========================================================================================================
'=                       4.4 User-defind Variables
'========================================================================================================
Dim IsOpenPop
Dim lgStrComDateType		                                            'Company Date Type을 저장(년월 Mask에 사용함.)

'========================================================================================================
' Name : InitVariables()	
' Desc : Initialize value
'========================================================================================================
Sub InitVariables()
	lgIntFlgMode       = parent.OPMD_CMODE						        '⊙: Indicates that current mode is Create mode
	lgBlnFlgChgValue   = False								    '⊙: Indicates that no value changed
	lgIntGrpCount      = 0										'⊙: Initializes Group View Size
    lgSortKey          = 1                                      '⊙: initializes sort direction		
End Sub
'========================================================================================================
' Name : SetDefaultVal()	
' Desc : Set default value
'========================================================================================================	
Sub SetDefaultVal()
  '  frm1.txtDt.Text     = UniConvDateAToB("<%=GetSvrDate%>",parent.gServerDateFormat,parent.gDateFormat)
   ' frm1.txtBas_dt.Text = frm1.txtDt.Text
        
    Dim strYear,strMonth,strDay
    Call ExtractDateFrom("<%=GetsvrDate%>",parent.gServerDateFormat , parent.gServerDateType ,strYear,strMonth,strDay)
' strYear ="2005"
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
	<% Call loadInfTB19029A("Q", "H","NOCOOKIE","MA") %>
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
Sub MakeKeyStream()   
	Dim bas_dt, send_dt
	bas_dt = frm1.txtBas_dt.year & right("00" & frm1.txtBas_dt.month,2) & right("00" & frm1.txtBas_dt.Day,2) 
	send_dt  = frm1.txtDt.year & right("00" & frm1.txtDt.month,2) & right("00" & frm1.txtDt.Day,2)

	lgKeyStream       = Trim(frm1.txtFile.value) & parent.gColSep					'파일명 
	lgKeyStream       = lgKeyStream & Trim(frm1.txtComp_cd.value) & parent.gColSep	'신고사업장 
	lgKeyStream       = lgKeyStream & Trim(send_dt) & parent.gColSep				'제출연월일 
	lgKeyStream       = lgKeyStream & Trim(bas_dt) & parent.gColSep					'기준연월일 

	lgKeyStream       = lgKeyStream & Trim(frm1.txtAllYn.value) & parent.gColSep	'통합신고여부 
	lgKeyStream       = lgKeyStream & Trim(frm1.txtRetireYn.value) & parent.gColSep	'퇴직포함여부 
				
	lgKeyStream       = lgKeyStream & Trim(Frm1.txtGubun.value) & parent.gColSep	'제출자구분 
	lgKeyStream       = lgKeyStream & Trim(frm1.txtGigan.value) & parent.gColSep	'대상기간 
	lgKeyStream       = lgKeyStream & Trim(frm1.txtSer.value) & parent.gColSep		'세무대리인관리번호 
	
End Sub
'========================================================================================================
' Name : InitComboBox()
' Desc : Set ComboBox
'========================================================================================================
Sub InitComboBox()
    Dim iNameArr , iCodeArr     
   '제출자 구분 
    Call CommonQueryRs("MINOR_NM,MINOR_CD","B_MINOR","MAJOR_CD = " & FilterVar("H0118", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iNameArr = lgF0
    iCodeArr = lgF1   
    Call SetCombo2(frm1.txtGubun,iCodeArr,iNameArr,Chr(11))     
    frm1.txtGubun.value = 2
    '대상기간 
    Call CommonQueryRs("MINOR_NM,MINOR_CD","B_MINOR","MAJOR_CD = " & FilterVar("H0119", "''", "S") & "",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    iNameArr = lgF0
    iCodeArr = lgF1       
    Call SetCombo2(frm1.txtGigan,iCodeArr,iNameArr,Chr(11))            ''''''''DB에서 불러 condition에서        
    '신고사업장 
    Call CommonQueryRs("YEAR_AREA_NM,YEAR_AREA_CD","HFA100T","",lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    iNameArr = lgF0
    iCodeArr = lgF1   
    Call SetCombo2(frm1.txtComp_cd,iCodeArr,iNameArr,Chr(11))  
    
    iCodeArr = "Y" & Chr(11) & "N" & Chr(11)
    Call SetCombo2(frm1.txtAllYn,iCodeArr,iCodeArr,Chr(11))
    Call SetCombo2(frm1.txtRetireYn,iCodeArr,iCodeArr,Chr(11))    
End Sub

'========================================================================================================
' Function Name : InitSpreadSheet
' Function Desc : This method initializes spread sheet column property
'========================================================================================================
Sub InitSpreadSheet(strSPD)
	Dim strMaskYM, i 
	
	If parent.gComDateType = "/" Then 
		lgStrComDateType = "/" & parent.gComDateType
	Else
		lgStrComDateType = parent.gComDateType
	End If
	strMaskYM = "9999" & lgStrComDateType & "99"

    ' Set SpreadSheet #1
	if (strSPD = "A" or strSPD = "ALL") then
		With Frm1.vspdData
			ggoSpread.Source = Frm1.vspdData
			ggoSpread.Spreadinit "V20021128",, parent.gAllowDragDropSpread	
		.ReDraw = false
		.MaxCols = 17 + 1                                                   '☜: Add 1 to Maxcols
		.Col = .MaxCols                                                          '☜: Hide maxcols
		.ColHidden = True                                                        '☜:    

		.MaxRows = 0

		ggoSpread.SSSetEdit      1,		"레코드구분",               10
		ggoSpread.SSSetEdit      2,		"자료구분",                 9
		ggoSpread.SSSetEdit      3,		"세무서",                   8
		ggoSpread.SSSetEdit      4,		"제출연월일",               8
		ggoSpread.SSSetEdit      5,		"제출자(대리인구분)",       20
		ggoSpread.SSSetEdit      6,		"세무대리인관리번호",       20
		ggoSpread.SSSetEdit      7,		"홈텍스ID",					20	'2004 
		ggoSpread.SSSetEdit      8,		"세무프로그램코드",			45	'2004 				
		ggoSpread.SSSetEdit      9,		"사업자등록번호",           16
		ggoSpread.SSSetEdit      10,	"법인명(상호)",             14
		ggoSpread.SSSetEdit      11,	"담당자부서",				30	'2004 	
		ggoSpread.SSSetEdit      12,	"담당자성명",				30	'2004 	
		ggoSpread.SSSetEdit      13,	"담당자전화번호",			15	'2004 	
		ggoSpread.SSSetEdit      14,	"신고의무자(B레코드) 수",   20
		ggoSpread.SSSetEdit      15,	"한글코드종류",             14
		ggoSpread.SSSetEdit      16,	"제출대상기간코드",         18
		ggoSpread.SSSetEdit      17,	"공란",                     6

		.ReDraw = true
		
		lgActiveSpd = "A"
		Call SetSpreadLock 
	    
		End With
    End if
    ' Set SpreadSheet #2
   	if (strSPD = "B" or strSPD = "ALL") then
   		With Frm1.vspdData1
	   	
			ggoSpread.Source = Frm1.vspdData1
			ggoSpread.Spreadinit "V20021128",, parent.gAllowDragDropSpread
		.ReDraw = false
		.MaxCols = 17 + 1                                                   '☜: Add 1 to Maxcols
		.Col = .MaxCols                                                           '☜: Hide maxcols
		.ColHidden = True                                                         '☜:

		.MaxRows = 0
	      
		ggoSpread.SSSetEdit      1,		"레코드구분",                   12
		ggoSpread.SSSetEdit      2,		"자료구분",                     10
		ggoSpread.SSSetEdit      3,		"세무서",                       8
		ggoSpread.SSSetEdit      4,		"일련번호",                     12
		ggoSpread.SSSetEdit      5,		"사업자등록번호",               16
		ggoSpread.SSSetEdit      6,		"법인명(상호)",                 14
		ggoSpread.SSSetEdit      7,		"대표자(성명)",                 13
		ggoSpread.SSSetEdit      8,		"주민(법인)등록번호",           20
		ggoSpread.SSSetEdit      9,		"주(현)제출건수(C레코드수)",   24
		ggoSpread.SSSetEdit      10,	"종(전)레코드수(D레코드수)",   24
		ggoSpread.SSSetEdit      11,	"소득금액총계",                14
		ggoSpread.SSSetEdit      12,	"소득세결정세액총계",          20
		ggoSpread.SSSetEdit      13,	"법인세결정세액총계",          20
		ggoSpread.SSSetEdit      14,	"주민세결정세액총계",          20
		ggoSpread.SSSetEdit      15,	"농특세결정세액총계",          20
		ggoSpread.SSSetEdit      16,	"결정세액총계",                14
		ggoSpread.SSSetEdit      17,	"공란",                         6

		.ReDraw = true
		
		lgActiveSpd = "B"
		Call SetSpreadLock 
	    
		End With
    End if
    ' Set SpreadSheet #3
  
    if (strSPD = "C" or strSPD = "ALL") then
		With Frm1.vspdData2
			
			ggoSpread.Source = Frm1.vspdData2
			ggoSpread.Spreadinit "V20021128",, parent.gAllowDragDropSpread
			.ReDraw = false
		.MaxCols = 85 + 1                                                   '☜: Add 1 to Maxcols
		.Col = .MaxCols                                                            '☜: Hide maxcols
		.ColHidden = True                                                          '☜:
	      
		.MaxRows = 0
		
			ggoSpread.SSSetEdit     1,		"레코드구분",               10
			ggoSpread.SSSetEdit     2,		"자료구분",                 8
			ggoSpread.SSSetEdit     3,		"세무서",                   8
			ggoSpread.SSSetEdit     4,		"일련번호",                 10
			ggoSpread.SSSetEdit     5,		"사업자등록번호",           14
			ggoSpread.SSSetEdit     6,		"종(전)근무처수",           14
			ggoSpread.SSSetEdit     7,		"거주자구분코드",           14
            ggoSpread.SSSetEdit     8,		"거주지국코드",             12  '2002
            ggoSpread.SSSetEdit     9,		"외국인단일세율적용",		12  '2004
			ggoSpread.SSSetEdit     10,		"귀속년도시작연월일",       16
			ggoSpread.SSSetEdit     11 ,	"귀속연도종료연월일",       16
			ggoSpread.SSSetEdit     12,		"성명",                     8
			ggoSpread.SSSetEdit     13,		"내국/외국인구분코드",      16
			ggoSpread.SSSetEdit     14,		"주민등록번호",             12
			ggoSpread.SSSetEdit     15,		"감면기간시작연월일",       16
			ggoSpread.SSSetEdit     16,		"감면기간종료연월일",       16
			ggoSpread.SSSetEdit     17,		"급여총액",                 10
			ggoSpread.SSSetEdit     18,		"상여총액",                 10
			ggoSpread.SSSetEdit     19,     "인정상여",                 10
			ggoSpread.SSSetEdit     20,		"계",                       10
			ggoSpread.SSSetEdit     21,		"국외근로",                 10
			ggoSpread.SSSetEdit     22,		"야간근로수당등",           10
			ggoSpread.SSSetEdit     23,		"기타비과세",               10
			ggoSpread.SSSetEdit     24,		"계",                       10
			ggoSpread.SSSetEdit     25,		"총급여",                   10
			ggoSpread.SSSetEdit     26,		"근로소득공제",             10
			ggoSpread.SSSetEdit     27,		"과세대상근로소득금액",     10
			ggoSpread.SSSetEdit     28,		"본인공제금액",             10
			ggoSpread.SSSetEdit     29,		"배우자공제금액",           10
			ggoSpread.SSSetEdit     30,		"부양가족공제인원",         10
			ggoSpread.SSSetEdit     31,		"부양가족공제금액",         10
			ggoSpread.SSSetEdit     32,		"경로우대공제인원",         10
			ggoSpread.SSSetEdit     33,		"경로우대공제금액",         10
			ggoSpread.SSSetEdit     34,		"장애자공제인원",           10
			ggoSpread.SSSetEdit     35,		"장애자공제금액",           10
			ggoSpread.SSSetEdit     36,		"부녀자공제금액",           10
			ggoSpread.SSSetEdit     37,		"자녀양육비공제인원",       10
			ggoSpread.SSSetEdit     38,		"자녀양육비공제금액",       10
			ggoSpread.SSSetEdit     39,		"소수공제자추가공제",       10
			ggoSpread.SSSetEdit     40,		"연금보험료",               10
			ggoSpread.SSSetEdit     41,		"보험료",                   10
			ggoSpread.SSSetEdit     42,		"의료비",                   10
			ggoSpread.SSSetEdit     43,		"교육비",                   10
			ggoSpread.SSSetEdit     44,		"주택자금",                 10
			ggoSpread.SSSetEdit     45,		"기부금",                   10
			
			ggoSpread.SSSetEdit     46,		"혼인/이사/장례비",			10	'2004	
			ggoSpread.SSSetEdit     47,		"공란",                     6		
			ggoSpread.SSSetEdit     48,		"계(특별공제)",             10
			ggoSpread.SSSetEdit     49,		"표준공제",                 10
			ggoSpread.SSSetEdit     50,		"차감소득금액",             10
			ggoSpread.SSSetEdit     51,		"개인연금저축",             10
			ggoSpread.SSSetEdit     52,		"연금저축",                 10
			ggoSpread.SSSetEdit     53,		"투자조합출자등소득공제",   10
			ggoSpread.SSSetEdit     54,		"신용카드소득공제",         10
            ggoSpread.SSSetEdit     55,		"우리사주조합소득공제",     10 '2002
            ggoSpread.SSSetEdit     56,		"퇴직연금소득공제",			10 '2005
            ggoSpread.SSSetEdit     57,		"기타소득공제계",			10	'2004	
            ggoSpread.SSSetEdit     58,		"종합소득과세표준",         18
            ggoSpread.SSSetEdit     59,		"산출세액",                 10
            ggoSpread.SSSetEdit     60,     "소득세법",                 10 '2002 세액감면 
            ggoSpread.SSSetEdit     61,		"조특법",                   10
            ggoSpread.SSSetEdit     62,		"공란",                     6
            ggoSpread.SSSetEdit     63,		"계",                       10
            ggoSpread.SSSetEdit     64,		"근로소득세액공제",         10 '2002 세액공제 
            ggoSpread.SSSetEdit     65,		"납세조합공제",             10
            ggoSpread.SSSetEdit     66,		"주택차입금세액공제",       10

            ggoSpread.SSSetEdit     67,		"기부정치자금",				10'2004            
            ggoSpread.SSSetEdit     68,		"외국납부세액공제",         10

            ggoSpread.SSSetEdit     69,		"공란",                     6
            ggoSpread.SSSetEdit     70,		"공란",                     6
            ggoSpread.SSSetEdit     71,		"세액공제계",               10
			ggoSpread.SSSetEdit     72,		"소득세",                   10
			ggoSpread.SSSetEdit     73,		"주민세",                   10
			ggoSpread.SSSetEdit     74,		"농특세",					10
			ggoSpread.SSSetEdit     75,		"계",                       10
			ggoSpread.SSSetEdit     76,		"소득세",                   10
			ggoSpread.SSSetEdit     77,     "주민세",                   10
			ggoSpread.SSSetEdit     78,		"농특세",					10
			ggoSpread.SSSetEdit     79,		"계",                       10
			ggoSpread.SSSetEdit     80,		"소득세",                   10
			ggoSpread.SSSetEdit     81,     "주민세",                   10
			ggoSpread.SSSetEdit     82,		"농특세",					10
			ggoSpread.SSSetEdit     83,		"계",                       10
			ggoSpread.SSSetEdit     84,		"현금영수증사용액",			10			
			ggoSpread.SSSetEdit     85,		"공란",                     6
			
		.ReDraw = true
 
		lgActiveSpd = "C"
		Call SetSpreadLock 
	    
		End With
    end if
  
    ' Set SpreadSheet #4
    if (strSPD = "D" or strSPD = "ALL") then
		With Frm1.vspdData3
	    
			ggoSpread.Source = Frm1.vspdData3
			ggoSpread.Spreadinit "V20021128",, parent.gAllowDragDropSpread	
		.ReDraw = false
		.MaxCols = 15 + 1                                                   '☜: Add 1 to Maxcols
		.Col = .MaxCols                                                            '☜: Hide maxcols
		.ColHidden = True                                                          '☜:
	    
		.MaxRows = 0
		
		ggoSpread.SSSetEdit      1,		"레코드구분",           12
		ggoSpread.SSSetEdit      2,		"자료구분",             10
		ggoSpread.SSSetEdit      3,		"세무서",               8
		ggoSpread.SSSetEdit      4,		"일련번호",             10
		ggoSpread.SSSetEdit      5,		"사업자등록번호",       16
		ggoSpread.SSSetEdit      6,		"공란",                 6
		ggoSpread.SSSetEdit      7,		"소득자주민등록번호",   20
		ggoSpread.SSSetEdit      8,		"법인명(상호)",         13
		ggoSpread.SSSetEdit      9,		"사업자등록번호",       16
		ggoSpread.SSSetEdit      10,	"급여총액",             10
		ggoSpread.SSSetEdit      11,	"상여총액",             10
		ggoSpread.SSSetEdit      12,	"인정상여",             10
		ggoSpread.SSSetEdit      13,	"계",                   6
		ggoSpread.SSSetEdit      14,	"종(전)근무처일련번호", 21
		ggoSpread.SSSetEdit      15,	"공란",                 6
				
			
		.ReDraw = true
		
		lgActiveSpd = "D"
		Call SetSpreadLock 
	    
		End With
	end if
	
 
	    ' Set SpreadSheet #4
    if (strSPD = "E" or strSPD = "ALL") then
		With Frm1.vspdData4
	    
			ggoSpread.Source = Frm1.vspdData4
			ggoSpread.Spreadinit "V20021128",, parent.gAllowDragDropSpread	
		.ReDraw = false
		.MaxCols = 91 + 1														'☜: Add 1 to Maxcols
		.Col = .MaxCols                                                            '☜: Hide maxcols
		.ColHidden = True                                                          '☜:
	    
		.MaxRows = 0
		
		ggoSpread.SSSetEdit      1,		"레코드구분",           12
		ggoSpread.SSSetEdit      2,		"자료구분",             10
		ggoSpread.SSSetEdit      3,		"세무서",               8
		ggoSpread.SSSetEdit      4,		"일련번호",             10
		ggoSpread.SSSetEdit      5,		"사업자등록번호",       16
		ggoSpread.SSSetEdit      6,		"소득자주민등록번호",   20

		For i = 1 To  5

			ggoSpread.SSSetEdit      17*i-10,	"관계"&i,				6
			ggoSpread.SSSetEdit      17*i-10+1 ,	"내외국인코드"&i,				15
			ggoSpread.SSSetEdit      17*i-10+2 ,	"성명"&i,			14
			ggoSpread.SSSetEdit      17*i-10+3 ,	"주민등록번호"&i,       12
			ggoSpread.SSSetEdit      17*i-10+4 ,	"기본공제"&i,           10
			ggoSpread.SSSetEdit      17*i-10+5 ,	"장애인"&i,             10
			ggoSpread.SSSetEdit      17*i-10+6 ,	"자녀양육비"&i,			10
			ggoSpread.SSSetEdit      17*i-10+7 ,	"보험료"&i,             10
			ggoSpread.SSSetEdit      17*i-10+8 ,	"의료비"&i,             10
			ggoSpread.SSSetEdit      17*i-10+9 ,	"교육비"&i,             10
			ggoSpread.SSSetEdit      17*i-10+10 ,	"신용카드등"&i,         10
			ggoSpread.SSSetEdit      17*i-10+11 ,	"현금영수증등"&i,       10
			
			ggoSpread.SSSetEdit      17*i-10+12 ,	"보험료"&i & " 외",            14
			ggoSpread.SSSetEdit      17*i-10+13 ,	"의료비"&i& " 외",             15
			ggoSpread.SSSetEdit      17*i-10+14 ,	"교육비"&i& " 외",             14
			ggoSpread.SSSetEdit      17*i-10+15,	"신용카드등"&i& " 외",         14
			ggoSpread.SSSetEdit      17*i-10+16 ,	"기부금"&i,       14
			
				          
		Next
 
		.ReDraw = true
		
		lgActiveSpd = "E"
		Call SetSpreadLock 
	    
		End With
	end if	
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
                ggoSpread.SpreadLock      -1,-1,-1
                ggoSpread.SSSetProtected  .vspdData2.MaxCols   , -1, -1
              .vspdData2.ReDraw = True
            End With
        Case  "D"
            With frm1
              .vspdData3.ReDraw = False
                ggoSpread.SpreadLock      -1,-1,-1
                ggoSpread.SSSetProtected  .vspdData3.MaxCols   , -1, -1
              .vspdData3.ReDraw = True
            End With
        Case  "E"
            With frm1
              .vspdData4.ReDraw = False
                ggoSpread.SpreadLock      -1,-1,-1
                ggoSpread.SSSetProtected  .vspdData4.MaxCols   , -1, -1
              .vspdData4.ReDraw = True
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
Sub vspdData3_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData3
    call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub
Sub vspdData4_ColWidthChange(ByVal pvCol1, ByVal pvCol2)
    ggoSpread.Source = frm1.vspdData4
    call ggoSpread.SSSetColWidth(pvCol1, pvCol2)
End Sub
 
'========================================================================================================
' Name : Form_Load
' Desc : developer describe this line Called by Window_OnLoad() evnt
'========================================================================================================
Sub Form_Load()

    Err.Clear                                                                       '☜: Clear err status
	Call LoadInfTB19029                                                             '☜: Load table , B_numeric_format		
    Call ggoOper.FormatField(Document, "1",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
    Call ggoOper.FormatField(Document, "2",ggStrIntegeralPart, ggStrDeciPointPart,parent.gDateFormat,parent.gComNum1000,parent.gComNumDec)
	Call ggoOper.LockField(Document, "N")											'⊙: Lock Field
 
    Call InitSpreadSheet("ALL")                                                           'Setup the Spread sheet
 
    Call InitVariables  
	ProtectTag(frm1.txtSer) 
    frm1.txtDt.focus 
    Call SetDefaultVal
    
	Call SetToolbar("1100000000001111")												'⊙: Set ToolBar    	
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
    Dim send_dt
	Dim strFrom, strWhere
	Dim strEmp_no
	
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
    dim txtRetireYn
    txtRetireYn = frm1.txtRetireYn.value 
'-------데이터 존재하는지 체크 
    
	send_dt  = frm1.txtDt.year & right("00" & frm1.txtDt.month,2) & right("00" & frm1.txtDt.Day,2)

    strFrom = "hfa050t 	left outer join  haa010t on hfa050t.emp_no = haa010t.emp_no "& chr(13)
    strFrom = strFrom & "left outer join  hdf020t on hfa050t.emp_no = hdf020t.emp_no "& chr(13)
 
    strWhere = "hdf020t.res_flag = 'Y' "& chr(13)
    strWhere = strWhere & " AND hdf020t.year_mon_give = 'Y' "& chr(13)
    
	
	strWhere = strWhere & " 	AND (hdf020t.retire_dt IS NULL" & chr(13)
	strWhere = strWhere & " 							OR CONVERT(VARCHAR(4), DATEPART(year, hdf020t.retire_dt)) > " &  FilterVar(frm1.txtBas_dt.Year, "''", "S")& chr(13)
	strWhere = strWhere & " 							OR (CONVERT(VARCHAR(4), DATEPART(year, hdf020t.retire_dt)) =   " &  FilterVar(frm1.txtBas_dt.Year, "''", "S")& chr(13)
	strWhere = strWhere & " 								AND haa010t.retire_resn  IN (	SELECT DISTINCT CASE WHEN '"&txtRetireYn&"' ='Y' THEN MINOR_CD ELSE '6' END "& chr(13)
	strWhere = strWhere & " 												FROM  B_MINOR "& chr(13)
	strWhere = strWhere & " 												WHERE MAJOR_CD ='H0025'"& chr(13)
	strWhere = strWhere & " 											)"& chr(13)
	strWhere = strWhere & " 								)"& chr(13)
	strWhere = strWhere & " 						     )"& chr(13)
					     
					     
					     
    IF frm1.txtAllYn.value = "N" Then
		strWhere = strWhere & " 	AND haa010t.year_area_cd = " & FilterVar(frm1.txtComp_cd.value, "''", "S")
	End If	    

	    
    IntRetCD = CommonQueryRs(" hfa050t.emp_no ", strFrom, strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)

    If IntRetCD = False Then
		Call DisplayMsgbox("900014", "X","X","X")	'조회된 데이터가 없습니다.
		Exit Function
	End If
	
'-------배우자 체크 

    strWhere = "emp_no not in ( select emp_no"
    strWhere = strWhere & " 		from hfa150t "
    strWhere = strWhere & " 		where family_rel ='3' and base_yn='Y' ) "
    strWhere = strWhere & " 	and  hfa050t.year_yy= " & FilterVar(frm1.txtBas_dt.Year, "''", "S")
    strWhere = strWhere & " 	and spouse='Y'"

    strWhere = strWhere & " and emp_no  in ( select emp_no from haa010t where "
	strWhere = strWhere & " (haa010t.retire_dt IS NULL "
	strWhere = strWhere & "		OR CONVERT(VARCHAR(4), DATEPART(year, haa010t.retire_dt)) >  " & FilterVar(frm1.txtBas_dt.Year, "''", "S")
	strWhere = strWhere & "		OR (CONVERT(VARCHAR(4), DATEPART(year, haa010t.retire_dt)) =  " & FilterVar(frm1.txtBas_dt.Year, "''", "S")
    strWhere = strWhere & "     OR haa010t.retire_resn = '6' ))"
	strWhere = strWhere & " and haa010t.entr_dt < " & FilterVar(send_dt, "''", "S")	
    
    IF frm1.txtAllYn.value = "N" Then
		strWhere = strWhere & " 	AND haa010t.year_area_cd = " & FilterVar(frm1.txtComp_cd.value, "''", "S")
	End If	  
		
	strWhere = strWhere & " ) "

    IntRetCD = CommonQueryRs(" emp_no ", "hfa050t", strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    
    strEmp_no = Trim(Replace(lgF0,Chr(11),","))
    
    If IntRetCD = True Then
		Call DisplayMsgbox("971012", "X","부양가족공제자 등록화면에" & left(strEmp_no,len(strEmp_no)-1) & "사원의 배우자가 등록되어 있고,부양자 체크가 되어있는지","X")	
		Exit Function
	End If

'-------경로우대 체크 
	strFrom	= ""
	strFrom = strFrom & " ( "
 	strFrom = strFrom & " SELECT a.emp_no "		'경로우대(65세이상) 체크 
	strFrom = strFrom & " FROM hfa050t a left outer join "& chr(13)
    strFrom = strFrom & "	(	 select year_yy, emp_no, count(family_res_no) cnt "& chr(13)
    strFrom = strFrom & "		from hfa150t "& chr(13)
    strFrom = strFrom & "		where "& frm1.txtBas_dt.Year-1900 &"-left(family_res_no,2) >=65 and "& frm1.txtBas_dt.Year-1900 &"-left(family_res_no,2) <70 "& chr(13)
    strFrom = strFrom & "			and substring(replace(family_res_no,'-',''),7,1) in ('1','2') "& chr(13)
    strFrom = strFrom & "			and base_yn='Y' "& chr(13)
    strFrom = strFrom & "		group by year_yy,emp_no"& chr(13)
    strFrom = strFrom & "	) b on a.year_yy = b.year_yy and a.emp_no = b.emp_no"& chr(13)
    strFrom = strFrom  & "WHERE  a.year_yy= " & FilterVar(frm1.txtBas_dt.Year, "''", "S")& chr(13)
    strFrom = strFrom & "		and a.old_cnt>0"& chr(13)
    strFrom = strFrom & "		and a.old_cnt <> isnull(b.cnt,0) "& chr(13)
    strFrom = strFrom & "  and a.emp_no  in ( select emp_no from haa010t where "& chr(13)
	strFrom = strFrom & " (haa010t.retire_dt IS NULL "& chr(13)
	strFrom = strFrom & "		OR CONVERT(VARCHAR(4), DATEPART(year, haa010t.retire_dt)) >  " & FilterVar(frm1.txtBas_dt.Year, "''", "S")& chr(13)
    strFrom = strFrom & "		OR haa010t.retire_resn = '6' )"& chr(13)
	strFrom = strFrom & " and haa010t.entr_dt < " & FilterVar(send_dt, "''", "S")	& chr(13)
   
    IF frm1.txtAllYn.value = "N" Then
		strFrom = strFrom & " 	AND haa010t.year_area_cd = " & FilterVar(frm1.txtComp_cd.value, "''", "S")& chr(13)
	End If		
	
	strFrom = strFrom & " ) "


    strFrom = strFrom & "UNION ALL "

	strFrom = strFrom & " SELECT a.emp_no "	& chr(13)	'경로우대(70세이상) 체크 
	strFrom = strFrom & " FROM hfa050t a left outer join "& chr(13)
    strFrom = strFrom & "	(	 select year_yy, emp_no, count(family_res_no) cnt "& chr(13)
    strFrom = strFrom & "		from hfa150t "& chr(13)
    strFrom = strFrom & "		where "& frm1.txtBas_dt.Year-1900 &"-left(family_res_no,2) >=70 "& chr(13)
    strFrom = strFrom & "			and substring(replace(family_res_no,'-',''),7,1) in ('1','2') "& chr(13)
    strFrom = strFrom & "			and base_yn='Y'"& chr(13)
    strFrom = strFrom & "		group by year_yy,emp_no"& chr(13)
    strFrom = strFrom & "	) b on a.year_yy = b.year_yy and a.emp_no = b.emp_no"& chr(13)

    strFrom = strFrom & " WHERE a.year_yy= " & FilterVar(frm1.txtBas_dt.Year, "''", "S")& chr(13)
    strFrom = strFrom & "		and a.old_cnt2>0"& chr(13)
    strFrom = strFrom & "		and a.old_cnt2 <> isnull(b.cnt,0) "& chr(13)
    strFrom = strFrom & "  and a.emp_no  in ( select emp_no from haa010t where "& chr(13)
	strFrom = strFrom & " (haa010t.retire_dt IS NULL "& chr(13)
	strFrom = strFrom & "		OR CONVERT(VARCHAR(4), DATEPART(year, haa010t.retire_dt)) >  " & FilterVar(frm1.txtBas_dt.Year, "''", "S")& chr(13)
    strFrom = strFrom & "		OR haa010t.retire_resn = '6' )"& chr(13)
	strFrom = strFrom & " and haa010t.entr_dt < " & FilterVar(send_dt, "''", "S")	& chr(13)
    
    IF frm1.txtAllYn.value = "N" Then
		strFrom = strFrom & " 	AND haa010t.year_area_cd = " & FilterVar(frm1.txtComp_cd.value, "''", "S")
	End If
		
	strFrom = strFrom & " ) "    
	strFrom = strFrom & " ) T "

    IntRetCD = CommonQueryRs(" T.emp_no ", strFrom, "1=1" ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    strEmp_no = Trim(Replace(lgF0,Chr(11),","))

    If IntRetCD = True Then
		Call DisplayMsgbox("971012", "X",left(strEmp_no,len(strEmp_no)-1) & ":급여 마스터에 입력된 경로자 인원수와 부양가족공제자에 등록된 경로자 명수가 일치 하지 않습니다. 부양가족공제자에 주민등록번호 나이와 부양자 체크","X")	
		Exit Function
	End If
	
	
	'-------자녀양육 체크 

	strFrom = " hfa050t a left outer join "& chr(13)
    strFrom = strFrom & "	(	 select year_yy, emp_no, count(family_res_no) cnt "& chr(13)
    strFrom = strFrom & "		from hfa150t "& chr(13)
    strFrom = strFrom & "		where ((("& frm1.txtBas_dt.Year-2000 &"-left(family_res_no,2) <=6"& chr(13)
    strFrom = strFrom & "			and substring(replace(family_res_no,'-',''),7,1) in ('3','4')) or "& chr(13)
    strFrom = strFrom & "		    ( "& frm1.txtBas_dt.Year-1900 &"-left(family_res_no,2) <=6 "& chr(13)
    strFrom = strFrom & "			and substring(replace(family_res_no,'-',''),7,1) in ('1','2') ))"    & chr(13)
    strFrom = strFrom & "			and child_yn='Y'  and nat_flag='1') "& chr(13)
    strFrom = strFrom & "			or (child_yn='Y'  and  nat_flag='9' )"& chr(13)
    strFrom = strFrom & "		group by year_yy,emp_no"& chr(13)
    strFrom = strFrom & "	) b on a.year_yy = b.year_yy and a.emp_no = b.emp_no"& chr(13)

    strWhere = " a.year_yy= " & FilterVar(frm1.txtBas_dt.Year, "''", "S")& chr(13)
    strWhere = strWhere & " and a.chl_rear>0"& chr(13)
    strWhere = strWhere & "	and a.chl_rear <> isnull(b.cnt,0) "& chr(13)
    
    strWhere = strWhere & " and a.emp_no  in ( select emp_no from haa010t where "& chr(13)
	strWhere = strWhere & " (haa010t.retire_dt IS NULL "& chr(13)
	strWhere = strWhere & "		OR CONVERT(VARCHAR(4), DATEPART(year, haa010t.retire_dt)) >  " & FilterVar(frm1.txtBas_dt.Year, "''", "S")& chr(13)
    strWhere = strWhere & "		OR haa010t.retire_resn = '6' )"& chr(13)
	strWhere = strWhere & " and haa010t.entr_dt < " & FilterVar(send_dt, "''", "S")	& chr(13)
    
    IF frm1.txtAllYn.value = "N" Then
		strWhere = strWhere & " 	AND haa010t.year_area_cd = " & FilterVar(frm1.txtComp_cd.value, "''", "S")& chr(13)
	End If	
	
	strWhere = strWhere & " ) "& chr(13)
	
	IntRetCD = False
	
    IntRetCD = CommonQueryRs(" a.emp_no ", strFrom, strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    strEmp_no = Trim(Replace(lgF0,Chr(11),","))
    
    If IntRetCD = True Then
		Call DisplayMsgbox("971012", "X",left(strEmp_no,len(strEmp_no)-1) & ":급여 마스터에 입력된 자녀양육 인원수와 부양가족공제자에 등록된 자녀양육 명수가 일치 하지 않습니다. 부양가족공제자에 주민등록번호 나이와 자녀양육 체크","X")	    
		Exit Function
	End If
		
	'-------부양자 체크 

	strFrom = " hfa050t a left outer join "
    strFrom = strFrom & "	(	 select year_yy, emp_no, count(family_res_no) cnt "
    strFrom = strFrom & "		from hfa150t "
    strFrom = strFrom & "		where family_rel not in ('0','3') "
    strFrom = strFrom & "		and base_yn='Y'"
    strFrom = strFrom & "		group by year_yy,emp_no"
    strFrom = strFrom & "	) b on a.year_yy = b.year_yy and a.emp_no = b.emp_no"

    strWhere = " a.year_yy= " & FilterVar(frm1.txtBas_dt.Year, "''", "S")
    strWhere = strWhere & " and a.supp_cnt>0"
    strWhere = strWhere & "	and a.supp_cnt <> isnull(b.cnt,0) "
    strWhere = strWhere & " and a.emp_no  in ( select emp_no from haa010t where "
	strWhere = strWhere & " (haa010t.retire_dt IS NULL "
	strWhere = strWhere & "		OR CONVERT(VARCHAR(4), DATEPART(year, haa010t.retire_dt)) >  " & FilterVar(frm1.txtBas_dt.Year, "''", "S")
    strWhere = strWhere & "		OR haa010t.retire_resn = '6' )"
	strWhere = strWhere & " and haa010t.entr_dt < " & FilterVar(send_dt, "''", "S")	
    
    IF frm1.txtAllYn.value = "N" Then
		strWhere = strWhere & " 	AND haa010t.year_area_cd = " & FilterVar(frm1.txtComp_cd.value, "''", "S")
	End If	
		
	strWhere = strWhere & " ) "
	
	IntRetCD = False
	
    IntRetCD = CommonQueryRs(" a.emp_no ", strFrom, strWhere ,lgF0,lgF1,lgF2,lgF3,lgF4,lgF5,lgF6)
    strEmp_no = Trim(Replace(lgF0,Chr(11),","))
    
    If IntRetCD = True Then
   		Call DisplayMsgbox("971012", "X",left(strEmp_no,len(strEmp_no)-1) & ":급여 마스터에 입력된 부양자 인원수와 부양가족공제자에 등록된 부양자 명수가 일치 하지 않습니다. 부양가족공제자에 부양자 체크","X")	    
		Exit Function
	End If
	'----------		
    Call MakeKeyStream()

    If DbQuery = False Then  
		Exit Function
	End If
       
    FncQuery = True																'☜: Processing is OK
   
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
'========================================================================================================
' Name : DbQuery
' Desc : This function is called by FncQuery
'========================================================================================================
Function DbQuery()
    Dim strVal
    Dim i

    Err.Clear                                                                        '☜: Clear err status

    DbQuery = False                                                                  '☜: Processing is NG
    
    Call LayerShowHide(1)

    strVal = BIZ_PGM_ID & "?txtMode="            & parent.UID_M0001                         '☜: Query
    strVal = strVal     & "&txtKeyStream="       & lgKeyStream                       '☜: Query Key

    Call RunMyBizASP(MyBizASP, strVal)                                               '☜:  Run biz logic
	
    DbQuery = True                                                                   '☜: Processing is NG
End Function

'========================================================================================================
' Function Name : DbQueryOk
' Function Desc : Called by MB Area when query operation is successful
'========================================================================================================
Function DbQueryOk()
	Dim i
    Err.Clear                                                                    '☜: Clear err status
    If (frm1.vspdData.MaxRows <= 0 And frm1.vspdData1.MaxRows<= 0 And frm1.vspdData2.MaxRows <= 0 And frm1.vspdData3.MaxRows <= 0) Then
		Call DisplayMsgbox("900014", "X","X","X")			                            '☜: 조회를 먼저하세요		
    End If	

    Call SetToolbar("1100000000011111")
	
    Set gActiveElement = document.activeElement	

End Function
 

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
	Dim StrEbrFile    
	Dim objName

	dim bas_dt, bas_yy, biz_area_cd, present_dt
	
	StrEbrFile = "h9117oa1_1"
	bas_dt = UniConvDateAToB(frm1.txtBas_dt.Text, parent.gDateFormat, parent.gServerDateFormat)
	bas_dt = replace(bas_dt,parent.gServerDateFormat,"")
	bas_yy = frm1.txtBas_dt.Year
	biz_area_cd = frm1.txtComp_cd.value
	present_dt = UniConvDateAToB(frm1.txtdt.text, parent.gDateFormat, parent.gServerDateFormat)
	present_dt = replace(present_dt,parent.gServerDateFormat,"")


	strUrl = strUrl & "bas_dt|" & bas_dt
	strUrl = strUrl & "|bas_yy|" & bas_yy 
	strUrl = strUrl & "|biz_area_cd|" & biz_area_cd
	strUrl = strUrl & "|present_dt|" & present_dt
	strUrl = strUrl & "|yn|" & frm1.txtAllYn.value
	strUrl = strUrl & "|retire_yn|" & frm1.txtRetireYn.value
	
'    objname = AskEBDocumentName(StrEbrFile,"EBR")
'	Call FncEBRPrint(EBAction,objname,strUrl)
   	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	call FncEBRPreview(ObjName , strUrl)
End Function
'======================================================================================================
' Function Name : FloppyDiskLabelForm
' Function Desc : 플로피디스켓 라벨양식 
'=======================================================================================================
Function FloppyDiskLabelForm()
	Dim strUrl	
    Dim StrEbrFile
	Dim objName
	
	dim bas_dt, bas_yy, biz_area_cd
	
	StrEbrFile = "h9117oa1_2"	
	bas_dt = UniConvDateAToB(frm1.txtbas_dt.text, parent.gDateFormat, parent.gServerDateFormat)
	bas_dt = replace(bas_dt,parent.gServerDateFormat,"")
	bas_yy = frm1.txtBas_dt.Year
	biz_area_cd = frm1.txtComp_cd.value	

	strUrl = strUrl & "bas_dt|" & bas_dt
	strUrl = strUrl & "|bas_yy|" & bas_yy 
	strUrl = strUrl & "|biz_area_cd|" & biz_area_cd		
	strUrl = strUrl & "|yn|" & frm1.txtAllYn.value
	strUrl = strUrl & "|retire_yn|" & frm1.txtRetireYn.value

'	objname = AskEBDocumentName(StrEbrFile,"EBR")
'	Call FncEBRPrint(EBAction,objname,strUrl)

 '  	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
 '	call FncEBRPrint(EBAction , ObjName , strUrl)
   	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	call FncEBRPreview(ObjName , strUrl) 
	
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
        	
	dim bas_dt, bas_yy, biz_area_cd, present_dt
	
	StrEbrFile = "h9117oa1_3"
	
	bas_dt = UniConvDateAToB(frm1.txtbas_dt.text, parent.gDateFormat, parent.gServerDateFormat)
	bas_dt = replace(bas_dt,parent.gServerDateFormat,"")
	bas_yy = frm1.txtBas_dt.Year
	biz_area_cd = frm1.txtComp_cd.value
	present_dt = UniConvDateAToB(frm1.txtdt.text, parent.gDateFormat, parent.gServerDateFormat)
	present_dt = replace(present_dt,parent.gServerDateFormat,"")

	strUrl = strUrl & "bas_dt|" & bas_dt
	strUrl = strUrl & "|bas_yy|" & bas_yy 
	strUrl = strUrl & "|biz_area_cd|" & biz_area_cd	
	strUrl = strUrl & "|present_dt|" & present_dt
	strUrl = strUrl & "|yn|" & frm1.txtAllYn.value
	strUrl = strUrl & "|retire_yn|" & frm1.txtRetireYn.value
		
'    objname = AskEBDocumentName(StrEbrFile,"EBR")
	'Call FncEBRPrint(EBAction,objname,strUrl)

   	ObjName = AskEBDocumentName(StrEbrFile,"ebr")
	call FncEBRPreview(ObjName , strUrl)	
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

	    Call MakeKeyStream()    
	    strVal = BIZ_PGM_ID2    & "?txtMode="           & parent.UID_M0001						'☜: 비지니스 처리 ASP의 상태 	    	    		    
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

Sub txtBas_dt_Keypress(Key)
    If Key = 13 Then
        FncQuery()
    End If
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
Sub txtBas_dt_DblClick(Button)
    If Button = 1 Then
		Call SetFocusToDocument("M")    
        frm1.txtBas_dt.Action = 7
        frm1.txtBas_dt.focus
    End If
End Sub
'========================================================================================================
'   Event Name : txtEmp_no_change            
'========================================================================================================
Function txtGubun_Onchange()
    Dim IntRetCd
  
	If  frm1.txtGubun.value <> "1" Then
		ProtectTag(frm1.txtSer) 
	Else
		ReleaseTag(frm1.txtSer) 
	End If	

End Function 

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
								<td background="../../../Cshared/Image/table/seltab_up_bg.gif"><IMG src="../../../Cshared/Image/table/seltab_up_left.gif" width="9" height="23" ></td>
								<td background="../../../Cshared/Image/table/seltab_up_bg.gif" align="center" CLASS="CLSMTABP"><font color=white>연말정산신고</font></td>
								<td background="../../../Cshared/Image/table/seltab_up_bg.gif" align="right"><IMG src="../../../Cshared/Image/table/seltab_up_right.gif" width="10" height="23" ></td>
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
								<TD CLASS=TD5  NOWRAP>세무대리인관리번호</TD>
								<TD CLASS=TD6  NOWRAP><INPUT TYPE=TEXT ID="txtBizAreaCD" MAXLENGTH=6 NAME="txtSer" SIZE=15 tag="11XXX" ALT="세무대리인관리번호"></TD>								
							</TR>
							<TR>
								<TD CLASS="TD5" NOWRAP>대상기간</TD>
								<TD CLASS="TD6" NOWRAP><SELECT NAME="txtGigan" ALT="대상기간" STYLE="WIDTH: 170px" TAG="12N"></SELECT></TD>							
								<TD CLASS=TD5  NOWRAP>제출연월일</TD>
								<TD CLASS=TD6  NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime1 name=txtDt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="제출년월일" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>	
				            <TR>
								<TD CLASS="TD5" NOWRAP>신고사업장</TD>
								<TD CLASS="TD6" NOWRAP><SELECT NAME="txtComp_cd" ALT="신고사업장" STYLE="WIDTH: 150px" TAG="12N"></SELECT></TD>								
								<TD CLASS=TD5  NOWRAP>기준연월일</TD>
								<TD CLASS=TD6  NOWRAP><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPDT%> id=fpDateTime2 name=txtBas_dt CLASS=FPDTYYYYMMDD title=FPDATETIME ALT="기준연월일" tag="12X1" VIEWASTEXT></OBJECT>');</SCRIPT></TD>
							</TR>
				            <TR>
								<TD CLASS=TD5  NOWRAP>통합신고여부</TD>
								<TD CLASS=TD6  NOWRAP><SELECT NAME="txtAllYn" ALT="통합신고여부" STYLE="WIDTH: 100px" TAG="12N"></SELECT></TD>								
								<TD CLASS=TD5  NOWRAP>퇴직포함여부</TD>
								<TD CLASS=TD6  NOWRAP><SELECT NAME="txtRetireYn" ALT="퇴직포함여부" STYLE="WIDTH: 100px" TAG="12N"></SELECT></TD>
							</TR>							
							<TR>								
						    <INPUT TYPE=HIDDEN ID="txtFile" NAME="txtFile" SIZE=15 tag="14XXXU" ALT="저장파일경로">
						</TABLE>
						</FIELDSET>
					</TD>
				</TR>
	
				<TR><TD <%=HEIGHT_TYPE_03%> WIDTH=100%></TD></TR>
				<TR >
					<TD WIDTH="100%" HEIGHT=* VALIGN=TOP>
						<TABLE <%=LR_SPACE_TYPE_60%>>
            			    <TR HEIGHT="20%">
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
                            <TR HEIGHT="30%">
				            	<TD WIDTH="100%" HEIGHT=* VALIGN=TOP COLSPAN=3>
				            		<TABLE WIDTH="100%" HEIGHT="100%">
				            			<TR>
				            				<TD HEIGHT="100%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData2 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread2> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT></TD>
					            		</TR>
					            	</TABLE>
					            </TD>
			                </TR>
                            <TR HEIGHT="20%">
				            	<TD WIDTH="100%" HEIGHT=* VALIGN=TOP COLSPAN=3>
				            		<TABLE WIDTH="100%" HEIGHT="100%">
				            			<TR>
				            				<TD HEIGHT="100%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData3 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread3> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT></TD>
					            		</TR>
					            	</TABLE>
					            </TD>
			                </TR>
                            <TR HEIGHT="30%">
				            	<TD WIDTH="100%" HEIGHT=* VALIGN=TOP COLSPAN=3>
				            		<TABLE WIDTH="100%" HEIGHT="100%">
				            			<TR>
				            				<TD HEIGHT="100%"><SCRIPT LANGUAGE=JavaScript> ExternalWrite('<OBJECT CLASSID=<%=gCLSIDFPSPD%> NAME=vspdData4 WIDTH=100% HEIGHT=100% tag="23" TITLE="SPREAD" id=vaSpread4> <PARAM NAME="MaxCols" VALUE="0"><PARAM NAME="MaxRows" VALUE="0"></OBJECT>');</SCRIPT></TD>
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
	                    <BUTTON NAME="btnCb_creation" CLASS="CLSMBTN">파일생성</BUTTON>&nbsp;</TD>
	                <TD WIDTH=* ALIGN="right"></TD>
	                <TD WIDTH=10>&nbsp;</TD>
	            </TR>
	        </TABLE>
	    </TD>
	</TR>
	<TR>
		<TD WIDTH=100% HEIGHT=0><IFRAME NAME="MyBizASP" SRC="../../blank.htm" WIDTH=100% HEIGHT=10 FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0></IFRAME>
		</Td>
		
	</TR>
</TABLE>

<INPUT TYPE=HIDDEN NAME="txtMode"        TAG="24">
<INPUT TYPE=HIDDEN NAME="txtKeyStream"   TAG="24">
<INPUT TYPE=HIDDEN NAME="txtUpdtUserId"  TAG="24">
<INPUT TYPE=HIDDEN NAME="txtInsrtUserId" TAG="24">
<INPUT TYPE=HIDDEN NAME="txtFlgMode"     TAG="24">
<INPUT TYPE=HIDDEN NAME="txtPrevNext"    TAG="24">
<TEXTAREA CLASS="hidden" NAME="txtSpread" TAG="24"></TEXTAREA><INPUT TYPE=HIDDEN NAME="txtMaxRows" tag="24"><%' 업무처리ASP로 넘기기 위한 변수를 담고 있는 Tag들 %>

</FORM>
<DIV ID="MousePT" NAME="MousePT">
	<IFRAME NAME="MouseWindow" FRAMEBORDER=0 SCROLLING=NO noresize framespacing=0 WIDTH=220 HEIGHT=41 SRC="../../inc/cursor.htm"></IFRAME>
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


