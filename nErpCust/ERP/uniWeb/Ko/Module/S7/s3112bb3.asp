<%'======================================================
'*  1. Module Name          : 영업 
'*  2. Function Name        : 매출관리 
'*  3. Program ID           : s3112bb3
'*  4. Program Name         : 출하내역참조 
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
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "RB")

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgMaxCount     = 2000							             '☜ : 한번에 가져올수 있는 데이타 건수 
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
    Redim UNIValue(0,14)

    UNISqlId(0) = "S3112BA301"
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    UNIValue(0,0) = Trim(lgSelectList)                                      '☜: Select list
    
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    UNIValue(0,1) = " " & FilterVar(Request("txtSoldToParty"), "''", "S") & ""                    '주문처 
    UNIValue(0,2) = " " & FilterVar(Request("txtCurrency"), "''", "S") & ""                       '화폐단위 
    UNIValue(0,3) = " " & FilterVar(Request("txtSalesGrpCd"), "''", "S") & ""                     '영업그룹 
    
    If Len(Trim(Request("txtFromDt"))) Then
		UNIValue(0,4) = " " & FilterVar(UNIConvDate(Request("txtFromDt")), "''", "S") & ""                             '시작일 
	Else
		UNIValue(0,4) = "Null"
	End If

	UNIValue(0,5) = " " & FilterVar(UNIConvDate(Request("txtToDt")), "''", "S") & ""									'종료일 

    If Len(Trim(Request("txtSoNo"))) Then
		UNIValue(0,6) = " " & FilterVar(Request("txtSoNo"), "''", "S") & ""                      'S/O번호 
	Else
		UNIValue(0,6) = "Null"
	End If

	UNIValue(0,7) = " " & FilterVar(Request("txtVatCalcType"), "''", "S") & ""                   'VAT 적용기준 
	UNIValue(0,8) = " " & FilterVar(Request("txtVatType"), "''", "S") & ""                       'VAT 유형 
	UNIValue(0,9) = " " & FilterVar(Request("txtVatIncFlag"), "''", "S") & ""                    'VAT포함구분 
	UNIValue(0,10) = " " & FilterVar(Request("txtPayTermsCd"), "''", "S") & ""                   '결제방법 
	UNIValue(0,11) = UNIConvNum(Request("txtVatRate"), 0)

    If Len(Trim(Request("txtItemCd"))) Then
		UNIValue(0,12) = " " & FilterVar(Request("txtItemCd"), "''", "S") & ""                   '품목코드 
	Else
		UNIValue(0,12) = "Null"
	End If
	
	If Len(Trim(Request("txtTrackingNo"))) Then		
		UNIValue(0,13) = " " & FilterVar(Request("txtTrackingNo"), "''", "S") & ""    		
	Else
		UNIValue(0,13) = "Null"
	End If

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
		Call parent.SetFocusToDocument("P")
		parent.frm1.txtFromDt.focus
		</Script>	
        <%
    Else    
        Call  MakeSpreadSheetData()
        Call  SetConditionData()
    End If  
End Sub

If lgDataExist = "Yes" Then
%>
<Script Language=vbscript>
    With parent
		
		'Set condition data to hidden area
		<%IF CLng(lgPageNo) = 1 Then%>
		.frm1.txtHFromDt.value	= "<%=Request("txtFromDT")%>"
		.frm1.txtHToDt.value	= "<%=Request("txtToDT")%>"
		.frm1.txtHItemCd.value  = "<%=Request("txtItemCd")%>"
		.frm1.txtHTrackingNo.value	= "<%=Request("txtTrackingNo")%>"
		<%End If%>
		
		'Show multi spreadsheet data from this line
		.ggoSpread.Source		= .frm1.vspdData 
        .frm1.vspdData.Redraw = False
		.ggoSpread.SSShowDataByClip "<%=lgstrData%>","F"						'☜ : Display data
		Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspdData.MaxRows,"<%=Request("txtCurrency")%>",.GetKeyPos("A",7),"C","I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspdData.MaxRows,"<%=Request("txtCurrency")%>",.GetKeyPos("A",8),"A","I","X","X")
		.lgPageNo				=  "<%=lgPageNo%>"							  '☜: Next next data tag
		.DbQueryOk
        .frm1.vspdData.Redraw = True                		
	End with
</Script>	
<%End If%>
