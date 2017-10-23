<%'======================================================================================================
'*  1. Module Name          : Sales
'*  2. Function Name        : 
'*  3. Program ID           : S2111QB31
'*  4. Program Name         : 조직별 품목그룹판매계획실적조회 
'*  5. Program Desc         :
'*  6. Modified date(First) : 2000/11/01
'*  7. Modified date(Last)  : 2001/12/19
'*  8. Modifier (First)     : KimTaeHyun
'*  9. Modifier (Last)      : sonbumyeol
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 12. History              :
'=======================================================================================================
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Call LoadBasisGlobalInf()

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgStrPrevKey                                                           '☜ : 이전 값 
Dim lgMaxCount                                                             '☜ : 한번에 가져올수 있는 데이타 건수 
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
Dim strPoType	                                                           '⊙ : 발주형태 
Dim strPoFrDt	                                                           '⊙ : 발주일 
Dim strPoToDt	                                                           '⊙ :
Dim strSpplCd	                                                           '⊙ : 공급처 
Dim strPurGrpCd	                                                           '⊙ : 구매그룹 
Dim strItemCd	                                                           '⊙ : 품목 
Dim strTrackNo	                                                           '⊙ : Tracking No
Dim FalsechkFlg
Dim arrRsVal(7)
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 

    lgStrPrevKey   = Request("lgStrPrevKey_A")                               '☜ : Next key flag
    lgMaxCount     = CInt(100)                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value

    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iCnt
    Dim  iRCnt
    Dim  iStr

    iCnt = 0
    lgstrData = ""

    If Len(Trim(lgStrPrevKey)) Then                                        '☜ : Chnage Nextkey str into int value
       If Isnumeric(lgStrPrevKey) Then
          iCnt = CInt(lgStrPrevKey)
       End If   
    End If   

    For iRCnt = 1 to iCnt  *  lgMaxCount                                   '☜ : Discard previous data
        rs0.MoveNext
    Next

    iRCnt = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iRCnt =  iRCnt + 1
        iStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            Select Case  lgSelectListDT(ColCnt)
               Case "DD"   '날짜 
                           iStr = iStr & Chr(11) & UNIDateClientFormat(rs0(ColCnt))
               Case "F2"  ' 금액 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggAmtOfMoney.DecPoint, 0)
               Case "F3"  '수량 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggQty.DecPoint       , 0)
               Case "F4"  '단가 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggUnitCost.DecPoint  , 0)
               Case "F5"   '환율 
                           iStr = iStr & Chr(11) & UNINumClientFormat(rs0(ColCnt), ggExchRate.DecPoint  , 0)
               Case Else
                    iStr = iStr & Chr(11) & ConvSPChars(rs0(ColCnt)) 
            End Select
		Next
 
        If  iRCnt < lgMaxCount Then
            lgstrData      = lgstrData      & iStr & Chr(11) & Chr(12)
        Else
            iCnt = iCnt + 1
            lgStrPrevKey = CStr(iCnt)
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iRCnt < lgMaxCount Then                                            '☜: Check if next data exists
        lgStrPrevKey = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
    Set lgADF = Nothing                                                    '☜: ActiveX Data Factory Object Nothing
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
    Dim arrVal(2)
    Dim Majorcd(1)
    Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(3,3)

    UNISqlId(0) = "S2111QA301"    
    UNISqlId(1) = "B1254MA803"	' 기존 S0000QA006    
    UNISqlId(2) = "S0000QA000"
    UNISqlId(3) = "S0000QA000"

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	strVal = " "

	Select Case CStr(Request("txtConPlanTypeCd"))
	Case CStr(1)					'수주인경우 
		strVal = "SELECT BG.ITEM_GROUP_CD  ITEM_GROUP_CD, SUM(SD.NET_AMT_LOC) NET_AMT_LOC"
		strVal = strVal + " FROM S_SO_DTL SD, S_SO_HDR SH , B_ITEM_GROUP BG, B_ITEM BI"
		strVal = strVal + " WHERE SD.SO_NO = SH.SO_NO AND SD.ITEM_CD = BI.ITEM_CD"
		strVal = strVal + " AND BI.ITEM_GROUP_CD = BG.ITEM_GROUP_CD AND BG.LEAF_FLG = " & FilterVar("Y", "''", "S") & " "
		strVal = strVal + " AND SH.CFM_FLAG = " & FilterVar("Y", "''", "S") & " "

		If CStr(Request("txtConDealTypeCd")) = CStr(1) Then
			strVal = strVal + " AND SH.EXPORT_FLAG = " & FilterVar("N", "''", "S") & " "
		ElseIf CStr(Request("txtConDealTypeCd")) = CStr(2) Then
			strVal = strVal + " AND SH.EXPORT_FLAG = " & FilterVar("Y", "''", "S") & " "
		Else
			strVal = strVal + " AND SH.EXPORT_FLAG = " & FilterVar("Z", "''", "S") & " "
		End If

		If Len(Request("txtConSalesOrg")) Then
			strVal = strVal + " AND SH.SALES_ORG = " & FilterVar(UCase(Request("txtConSalesOrg")), "''", "S") & " "
		Else
			strVal = strVal + ""
		End If
	
		If Len(Request("txtConSpYear")) Then
			strVal = strVal + " AND YEAR(SH.SO_DT) = " & FilterVar(Request("txtConSpYear"), "''", "S") & ""
		Else
			strVal = strVal + ""
		End If

		strVal = strVal + " GROUP BY BG.ITEM_GROUP_CD"

	Case CStr(2)					'매출인경우 
		strVal = "SELECT BG.ITEM_GROUP_CD  ITEM_GROUP_CD, SUM(SD.BILL_AMT_LOC) NET_AMT_LOC"
		strVal = strVal + " FROM S_BILL_DTL SD, S_BILL_HDR SH , B_ITEM_GROUP BG, B_ITEM BI"
		strVal = strVal + " WHERE SD.BILL_NO = SH.BILL_NO AND SD.ITEM_CD = BI.ITEM_CD"
		strVal = strVal + " AND BI.ITEM_GROUP_CD = BG.ITEM_GROUP_CD AND BG.LEAF_FLG = " & FilterVar("Y", "''", "S") & " "

		If CStr(Request("txtConDealTypeCd")) = CStr(1) Then
			strVal = strVal + " AND SH.BL_FLAG = " & FilterVar("N", "''", "S") & " "
		ElseIf CStr(Request("txtConDealTypeCd")) = CStr(2) Then
			strVal = strVal + " AND SH.BL_FLAG = " & FilterVar("Y", "''", "S") & " "
		Else
			strVal = strVal + " AND SH.BL_FLAG = " & FilterVar("Z", "''", "S") & " "
		End If

		If Len(Request("txtConSalesOrg")) Then
			strVal = strVal + " AND SH.SALES_ORG = " & FilterVar(UCase(Request("txtConSalesOrg")), "''", "S") & " "
		Else
			strVal = strVal + ""
		End If
	
		If Len(Request("txtConSpYear")) Then
			strVal = strVal + " AND YEAR(SH.BILL_DT) = " & FilterVar(Request("txtConSpYear"), "''", "S") & ""
		Else
			strVal = strVal + ""
		End If

		strVal = strVal + " GROUP BY BG.ITEM_GROUP_CD"

	Case Else
		strVal = "SELECT BG.ITEM_GROUP_CD  ITEM_GROUP_CD, SUM(SD.NET_AMT_LOC) NET_AMT_LOC"
		strVal = strVal + " FROM S_SO_DTL SD, S_SO_HDR SH , B_ITEM_GROUP BG, B_ITEM BI"
		strVal = strVal + " WHERE SD.SO_NO = SH.SO_NO AND SD.ITEM_CD = BI.ITEM_CD"
		strVal = strVal + " AND BI.ITEM_GROUP_CD = BG.ITEM_GROUP_CD AND BG.LEAF_FLG = " & FilterVar("Y", "''", "S") & " "
		strVal = strVal + " AND SH.CFM_FLAG = " & FilterVar("Y", "''", "S") & " "
		strVal = strVal + " AND YEAR(SH.SO_DT) = " & FilterVar("9999", "''", "S") & ""
		strVal = strVal + " GROUP BY BG.ITEM_GROUP_CD"
	End Select	
	
    UNIValue(0,1) = strVal


	strVal = " "

	strVal = " AND A.ORG_GRP_FLAG=" & FilterVar("O", "''", "S") & " "

	If Len(Request("txtConSalesOrg")) Then
		strVal = strVal + " AND A.SALES_ORG = " & FilterVar(UCase(Request("txtConSalesOrg")), "''", "S") & " "
		arrVal(0) = Trim(Request("txtConSalesOrg"))
	Else
		strVal = strVal + ""
		arrVal(0) = ""
	End If
	
	If Len(Request("txtConSpYear")) Then
		strVal = strVal + " AND A.SP_YEAR = " & FilterVar(Request("txtConSpYear"), "''", "S") & ""
	Else
		strVal = strVal + ""
	End If
	
	If Len(Request("txtConPlanTypeCd")) Then
		strVal = strVal + " AND A.PLAN_FLAG = " & FilterVar(UCase(Request("txtConPlanTypeCd")), "''", "S") & " "
		arrVal(1) = Trim(Request("txtConPlanTypeCd"))
		Majorcd(0) = "S4089"
	Else
		strVal = strVal + ""
		arrVal(1) = ""
		Majorcd(0) = ""
	End If
	
	If Len(Request("txtConDealTypeCd")) Then
		strVal = strVal + " AND A.EXPORT_FLAG = " & FilterVar(UCase(Request("txtConDealTypeCd")), "''", "S") & " "
		arrVal(2) = Trim(Request("txtConDealTypeCd"))
		Majorcd(1) = "S4225"
	Else
		strVal = strVal + ""
		arrVal(2) = ""
		Majorcd(1) = ""
	End If

    UNIValue(0,2) = strVal   

    UNIValue(1,0) = " " & FilterVar(arrVal(0), "''", "S") & "  AND END_ORG_FLAG=" & FilterVar("Y", "''", "S") & "  AND USAGE_FLAG=" & FilterVar("Y", "''", "S") & " "
    UNIValue(2,0) = FilterVar(Majorcd(0), " " , "S")
    UNIValue(2,1) = FilterVar(arrVal(1), " " , "S")
    UNIValue(3,0) = FilterVar(Majorcd(1), " " , "S")
    UNIValue(3,1) = FilterVar(arrVal(2), " " , "S")
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
'==    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNIValue(0,UBound(UNIValue,2)) = " " + UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub


'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim iStr
   
    FalsechkFlg = False
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3)

    FalsechkFlg = False
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing

		If Len(Request("txtConSalesOrg")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "영업조직", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
				%>
				<Script Language=vbscript>
				    parent.frm1.txtConSalesOrg.focus
				</Script>	
				<%
			   FalsechkFlg = True			
		End If	
    Else    
		arrRsVal(0) = rs1(0)
		arrRsVal(1) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If

    If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing

		If Len(Request("txtConPlanTypeCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "계획구분", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
				%>
				<Script Language=vbscript>
				    parent.frm1.txtConPlanTypeCd.focus
				</Script>	
				<%
				FalsechkFlg = True	
		End If
    Else    
		arrRsVal(2) = rs2(0)
		arrRsVal(3) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If

    If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing

		If Len(Request("txtConDealTypeCd")) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "거래구분", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
				%>
				<Script Language=vbscript>
				    parent.frm1.txtConDealTypeCd.focus
				</Script>	
				<%
				FalsechkFlg = True		
		End If
    Else    
		arrRsVal(4) = rs3(0)
		arrRsVal(5) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If
    
    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    

    If  rs0.EOF And rs0.BOF And FalsechkFlg = False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
    Else    
        Call  MakeSpreadSheetData()
    End If
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
End Sub


Sub ClearTextBox()
	%>
	<Script Language=vbscript>
	    With parent
	  		.frm1.txtConSalesOrgNm.value		=  ""
	  		.frm1.txtConPlanTypeNm.value		=  ""
	  		.frm1.txtConDealTypeNm.value		=  ""
	  	End with
	</Script>	
	<%
End Sub

%>

<Script Language=vbscript>
    With parent
         .ggoSpread.Source    = .frm1.vspdData 
         .ggoSpread.SSShowDataByClip "<%=lgstrData%>"                            '☜: Display data 
         .lgStrPrevKey_A      =  "<%=lgStrPrevKey%>"                       '☜: set next data tag

  		.frm1.txtConSalesOrgNm.value		=  "<%=ConvSPChars(arrRsVal(1))%>" 
  		.frm1.txtConPlanTypeNm.value		=  "<%=ConvSPChars(arrRsVal(3))%>" 	
  		.frm1.txtConDealTypeNm.value		=  "<%=ConvSPChars(arrRsVal(5))%>"
        .frm1.txtHConSalesOrg.value			= "<%=ConvSPChars(Request("txtConSalesOrg"))%>"
        .frm1.txtHConSpYear.value			= "<%=ConvSPChars(Request("txtConSpYear"))%>"
		.frm1.txtHConPlanTypeCd.value		= "<%=ConvSPChars(Request("txtConPlanTypeCd"))%>"
		.frm1.txtHConDealTypeCd.value		= "<%=ConvSPChars(Request("txtConDealTypeCd"))%>"
		.frm1.txtHConCurr.value				= "<%=ConvSPChars(Request("txtConCurr"))%>"
        If "<%=FalsechkFlg%>" = False Then
        .DbQueryOk("A")
        End If
	End with
</Script>	
<%
    Response.End													'☜: 비지니스 로직 처리를 종료함 
%>

