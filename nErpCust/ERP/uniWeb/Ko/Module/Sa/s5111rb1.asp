<%'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : NEGO관리 
'*  3. Program ID           : S5111RB1
'*  4. Program Name         : 매출채권 참조(NEGO등록)
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/08/07
'*  8. Modified date(Last)  : 
'*  9. Modifier (First)     : Hwangseongbae
'* 10. Modifier (Last)      : 
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%

On Error Resume Next
                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'☜ : DBAgent Parameter 선언 
Dim lgStrData                                                 '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
Dim lgMaxCount                                                '☜ : Spread sheet 의 visible row 수 
Dim lgTailList                                                '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "RB")
    Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgMaxCount     = CInt(30)             '☜ : 한번에 가져올수 있는 데이타 건수 
	lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
	lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
	lgDataExist      = "No"
	    
    Call FixUNISQLData()									 '☜ : DB-Agent로 보낼 parameter 데이타 set
    Call QueryData()										 '☜ : DB-Agent를 통한 ADO query
    
'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
    End If
    
    iLoopCount = -1
    
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < lgMaxCount Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        
        rs0.MoveNext
	Loop

    If iLoopCount < lgMaxCount Then                                 '☜: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF

End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Sub SetConditionData()
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
	Dim iStrWhere1

    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(0,2)

    UNISqlId(0) = "S5111RA101"
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    UNIValue(0,0) = Trim(lgSelectList)                                      '☜: Select list
    
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	iStrWhere1 = ""
	
    If Len(Trim(Request("txtSoldToParty"))) Then
		iStrWhere1 = iStrWhere1 & " AND BH.sold_to_party =  " & FilterVar(Request("txtSoldToParty"), "''", "S") & ""	'수입자 
	End If
	
    If Len(Trim(Request("txtCurrency"))) Then
	    iStrWhere1 = iStrWhere1 & " AND BH.cur =  " & FilterVar(Request("txtCurrency"), "''", "S") & ""					'화폐단위 
	End If

    If Len(Trim(Request("txtSalesGrpCd"))) Then
	    iStrWhere1 = iStrWhere1 & " AND BH.sales_grp =  " & FilterVar(Request("txtSalesGrpCd"), "''", "S") & ""			'영업그룹 
	End If

    If Len(Trim(Request("txtFromDt"))) Then
		iStrWhere1 = iStrWhere1 & " AND BH.bill_dt >=  " & FilterVar(UNIConvDate(Request("txtFromDt")), "''", "S") & ""						'시작일 
	End If

    If Len(Trim(Request("txtToDt"))) Then
	    iStrWhere1 = iStrWhere1 & " AND BH.bill_dt <=  " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & ""							'종료일 
	End If

	UNIValue(0,1) = iStrWhere1 
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
         
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        rs0.Close
        Set rs0 = Nothing
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        %>
		<Script Language=vbscript>
		parent.frm1.txtSoldToParty.focus
		</Script>	
        <%
    Else    
        Call  MakeSpreadSheetData()
        Call  SetConditionData()
    End If  
End Sub

%>

<Script Language=vbscript>
    With parent.frm1
		
		If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area
			' "1" means that this query is first and next data exists	
			<%If lgPageNo = "1" Then %>
			.txtHSoldToParty.value	= "<%=ConvSPChars(Request("txtSoldToParty"))%>"
			.txtHFromDt.value		= "<%=Request("txtFromDT")%>"
			.txtHToDt.value			= "<%=Request("txtToDT")%>"
			.txtHSalesGrp.value		= "<%=ConvSPChars(Request("txtSalesGrp"))%>"
			.txtHCurrency.value		= "<%=ConvSPChars(Request("txtCurrency"))%>"
			<%End If%>

			'Show multi spreadsheet data from this line
			parent.ggoSpread.Source		= .vspdData 
			.vspdData.Redraw = False
			parent.ggoSpread.SSShowDataByClip "<%=lgstrData%>","F"						'☜ : Display data
			Call parent.ReFormatSpreadCellByCellByCurrency(.vspdData,-1,-1,parent.GetKeyPos("A",5),parent.GetKeyPos("A",6),"A","I","X","X")

			parent.lgPageNo				=  "<%=lgPageNo%>"						  '☜: Next next data tag
			parent.DbQueryOk
			.vspdData.Redraw = True      
		End If
	End with
</Script>	
