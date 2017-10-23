<%'======================================================
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출채권관리 
'*  3. Program ID           : S5311QA1
'*  4. Program Name         : 세금계산서현황조회 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/05/10
'*  8. Modified date(Last)  : 2002/05/10
'*  9. Modifier (First)     : Hwangseongbae
'* 10. Modifier (Last)      : Hwangseongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'=======================================================
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
On Error Resume Next                                                                         
                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2, rs3			   '☜ : DBAgent Parameter 선언 
Dim lgStrData                                                 '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
Dim lgMaxCount                                                '☜ : Spread sheet 의 visible row 수 
Dim lgTailList                                                '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

Dim iFrPoint
iFrPoint=0

    Call HideStatusWnd 

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "QB")
	Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "QB")


	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgMaxCount     = 100							             '☜ : 한번에 가져올수 있는 데이타 건수 
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
       iFrPoint		= CLng(lgMaxCount) * CLng(lgPageNo)
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

    Redim UNISqlId(0)                                                     '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(0,8)

    UNISqlId(0) = "S5311QA101"
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    UNIValue(0,0) = Trim(lgSelectList)                                      '☜: Select list
    
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	If Len(Trim(Request("txtBillToParty"))) Then
		UNIValue(0,1) = " " & FilterVar(Request("txtBillToParty"), "''", "S") & ""		'발행처 
    Else
		UNIValue(0,1) = "NULL"
   	End If
	
    If Len(Trim(Request("txtFromDt"))) Then
		UNIValue(0,2) = " " & FilterVar(UNIConvDate(Request("txtFromDt")), "''", "S") & ""                           '시작일 
	Else
		UNIValue(0,2) = "Null"
	End If

    If Len(Trim(Request("txtToDt"))) Then
		UNIValue(0,3) = " " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & ""								'종료일 
	Else
		UNIValue(0,3) = "Null"
	End If

	If Len(Trim(Request("txtTaxBizArea"))) Then
		UNIValue(0,4) = " " & FilterVar(Request("txtTaxBizArea"), "''", "S") & ""			'세금신고사업장 
    Else
		UNIValue(0,4) = "NULL"
   	End If

	If Len(Trim(Request("txtSalesGrp"))) Then
		UNIValue(0,5) = " " & FilterVar(Request("txtSalesGrp"), "''", "S") & ""			'영업그룹 
    Else
		UNIValue(0,5) = "NULL"
   	End If

	If Len(Trim(Request("txtTaxRadio"))) Then
		UNIValue(0,6) = " " & FilterVar(Request("txtTaxRadio"), "''", "S") & ""			'계산서구분 
    Else
		UNIValue(0,6) = "NULL"
   	End If


	If Len(Trim(Request("txtIssueRadio"))) Then
		UNIValue(0,7) = " " & FilterVar(Request("txtIssueRadio"), "''", "S") & ""			'발행여부 
    Else
		UNIValue(0,7) = "NULL"
   	End If

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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)

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
		parent.frm1.txtBillToParty.focus
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

			'Show multi spreadsheet data from this line
			Parent.ggoSpread.Source		= .vspdData 
			.vspdData.Redraw = False
			Parent.ggoSpread.SSShowDataByClip "<%=lgstrData%>","F"						'☜ : Display data
			Call Parent.ReFormatSpreadCellByCellByCurrency(.vspdData,"<%=iFrPoint+1%>",.vspdData.MaxRows,Parent.GetKeyPos("A",7),Parent.GetKeyPos("A",6),"A","Q","X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency(.vspdData,"<%=iFrPoint+1%>",.vspdData.MaxRows,Parent.GetKeyPos("A",7),Parent.GetKeyPos("A",8),"A","Q","X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency(.vspdData,"<%=iFrPoint+1%>",.vspdData.MaxRows,Parent.GetKeyPos("A",7),Parent.GetKeyPos("A",9),"A","Q","X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency2(.vspdData,"<%=iFrPoint+1%>",.vspdData.MaxRows,Parent.parent.gCurrency,Parent.GetKeyPos("A",10),"A","Q","X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency2(.vspdData,"<%=iFrPoint+1%>",.vspdData.MaxRows,Parent.parent.gCurrency,Parent.GetKeyPos("A",11),"A","Q","X","X")
			Call Parent.ReFormatSpreadCellByCellByCurrency2(.vspdData,"<%=iFrPoint+1%>",.vspdData.MaxRows,Parent.parent.gCurrency,Parent.GetKeyPos("A",12),"A","Q","X","X")

			parent.lgPageNo				=  "<%=lgPageNo%>"							'☜: Next next data tag
			.txtHFromDt.value		= "<%=Request("txtFromDT")%>"
			.txtHToDt.value			= "<%=Request("txtToDT")%>"
			.txtHTaxBizArea.value	= "<%=ConvSPChars(Request("txtTaxBizArea"))%>"
			.txtHSalesGrp.value		= "<%=ConvSPChars(Request("txtSalesGrp"))%>"
			.txtHBillToParty.value	= "<%=ConvSPChars(Request("txtBillToParty"))%>"
			.txtHIssueRadio.value	= "<%=Request("txtIssueRadio")%>"
			.txtHTaxRadio.value		= "<%=Request("txtTaxRadio")%>"
			parent.DbQueryOk
			.vspdData.Redraw = True        
		End If
	End with
</Script>	
