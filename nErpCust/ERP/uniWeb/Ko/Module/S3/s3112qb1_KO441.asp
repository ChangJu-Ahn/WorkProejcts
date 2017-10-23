<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : s3112qb1
'*  4. Program Name         : 수주상세조회 
'*  5. Program Desc         : 수주상세조회 
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2001/12/18
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Kim Hyungsuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%                                                                         
Call LoadBasisGlobalInf()
Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "MB")
Call LoadBNumericFormatB("Q", "S", "NOCOOKIE", "MB")

On Error Resume Next

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0                              '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim lgPageNo
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
Dim arrRsVal(7)
Dim txtconBp_cd
Dim txtSoType
Dim txtSalesGroup
Dim txtItem_cd 
Dim txtSOFrDt
Dim txtSOToDt
Dim txtRadio
Dim txtTrackingNo
Dim txtConSo_no

Const C_SHEETMAXROWS_D  = 100                                   '☆: Server에서 한번에 fetch할 최대 데이타 건수 
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
  
    Call HideStatusWnd 

	If Request("txtFlgMode") = Request("OPMD_UMODE") Then
		txtconBp_cd = Request("HtxtconBp_cd")
		txtSoType = Request("HtxtSoType")
		txtSalesGroup = Request("HtxtSalesGroup")
		txtItem_cd = Request("HtxtItem_cd")
		txtSOFrDt = Request("HtxtSOFrDt")
		txtSOToDt = Request("HtxtSOToDt")
		txtRadio = Request("HtxtRadio")
		txtTrackingNo = Request("HtxtTrackingNo")
		txtConSo_no = Request("HtxtConSo_no")
	Else
		txtconBp_cd = Request("txtconBp_cd")
		txtSoType = Request("txtSoType")
		txtSalesGroup = Request("txtSalesGroup")
		txtItem_cd = Request("txtItem_cd")
		txtSOFrDt = Request("txtSOFrDt")
		txtSOToDt = Request("txtSOToDt")
		txtRadio = Request("txtRadio")
		txtTrackingNo = Request("txtTrackingNo")
		txtConSo_no = Request("txtConSo_no")
	End If
    
	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)                  '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
    'lgDataExist      = "No"

    Call TrimData()
    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    
    lgstrData      = ""
    
    If CInt(lgPageNo) > 0 Then
       rs0.Move  =  C_SHEETMAXROWS_D * CInt(lgPageNo)
    End If

    iLoopCount = -1
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < C_SHEETMAXROWS_D Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop

    If  iLoopCount < C_SHEETMAXROWS_D Then                             '☜: Check if next data exists
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
End Sub

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
    Dim arrVal(3)
    Redim UNISqlId(4)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(4,2)

    UNISqlId(0) = "S3112QA101"
    UNISqlId(1) = "S0000QA002"
    UNISqlId(2) = "S0000QA007"
    UNISqlId(3) = "S0000QA005"
    UNISqlId(4) = "S0000QA001"

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	strVal = " "
	
	If Len(txtconBp_cd) Then
		strVal = "AND A.SOLD_TO_PARTY = " & FilterVar(txtconBp_cd, "''", "S") & " "		
		arrVal(0) = Trim(txtconBp_cd)
	Else
		strVal = ""
	End If

	If Len(txtSoType) Then
		strVal = strVal & " AND A.SO_TYPE = " & FilterVar(txtSoType, "''", "S") & " "		
		arrVal(1) = Trim(txtSoType)
	End If		
		   
	If Len(txtSalesGroup) Then
		strVal = strVal & " AND A.SALES_GRP = " & FilterVar(txtSalesGroup, "''", "S") & " "		
		arrVal(2) = Trim(txtSalesGroup)
	End If		
    
 	If Len(txtItem_cd) Then
 		strVal = strVal & " AND B.ITEM_CD = " & FilterVar(txtItem_cd, "''", "S") & " "		
		arrVal(3) = Trim(txtItem_cd) 
	End If		
    
    If Len(txtSOFrDt) Then
		strVal = strVal & " AND A.SO_DT >='" & UNIConvDate(txtSOFrDt) & "'"		
	End If		
	
	If Len(txtSOToDt) Then
		strVal = strVal & " AND A.SO_DT <='" & UNIConvDate(txtSOToDt) & "'"		
	End If

	If Trim(txtRadio) = "Y" Or Trim(txtRadio) = "N" Then
		strVal = strVal & " AND A.CFM_FLAG ='" & Trim(txtRadio) & "'"		
	End If
	
	If Len(txtTrackingNo) Then
		strVal = strVal & " AND B.TRACKING_NO = " & FilterVar(txtTrackingNo, "''" , "S") & " "				
	End If		

	If Len(txtConSo_no) Then
		strVal = strVal & " AND A.SO_NO >= " & FilterVar(txtConSo_no, "''", "S") & " "		
	End If

	If Len(Request("gBizArea")) Then
		strVal = strVal & " AND A.BIZ_AREA = " & FilterVar(Request("gBizArea"), "''", "S") & " "			
	End If

	If Len(Request("gPlant")) Then
		strVal = strVal & " AND B.PLANT_CD = " & FilterVar(Request("gPlant"), "''", "S") & " "			
	End If

	If Len(Request("gSalesGrp")) Then
		strVal = strVal & " AND A.SALES_GRP = " & FilterVar(Request("gSalesGrp"), "''", "S") & " "			
	End If

	If Len(Request("gSalesOrg")) Then
		strVal = strVal & " AND A.SALES_ORG = " & FilterVar(Request("gSalesOrg"), "''", "S") & " "			
	End If	

    UNIValue(0,1) = strVal   '---발주일 
    UNIValue(1,0) = FilterVar(arrVal(0), "" , "S")
    UNIValue(2,0) = FilterVar(arrVal(1), "" , "S")
    UNIValue(3,0) = FilterVar(arrVal(2), "" , "S")
    UNIValue(4,0) = FilterVar(arrVal(3), "" , "S")
    
    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()
    Dim iStr
    Dim FalsechkFlg
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)
    
    FalsechkFlg = False
    
    If  rs1.EOF And rs1.BOF Then
        rs1.Close
        Set rs1 = Nothing

		If Len(txtconBp_cd) Then
		   Call DisplayMsgBox("970000", vbInformation, "주문처", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtconBp_cd.focus    
                </Script>
            <%	       
		End If	
    Else    
		arrRsVal(0) = rs1(0)
		arrRsVal(1) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    End If
    
    If  rs4.EOF And rs4.BOF Then
        rs4.Close
        Set rs4 = Nothing

		If Len(txtItem_cd) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "품목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtItem_cd.focus    
                </Script>
            <%	       
		End If	
    Else    
		arrRsVal(6) = rs4(0)
		arrRsVal(7) = rs4(1)
        rs4.Close
        Set rs4 = Nothing
    End If
    
    If  rs3.EOF And rs3.BOF Then
        rs3.Close
        Set rs3 = Nothing

		If Len(txtSalesGroup) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "영업그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtSalesGroup.focus    
                </Script>
            <%	       
		End If
    Else    
		arrRsVal(4) = rs3(0)
		arrRsVal(5) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    End If
    
    If  rs2.EOF And rs2.BOF Then
        rs2.Close
        Set rs2 = Nothing

		If Len(txtSoType) And FalsechkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "수주형태", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       FalsechkFlg = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtSoType.focus    
                </Script>
            <%	       
		End If	
    Else    
		arrRsVal(2) = rs2(0)
		arrRsVal(3) = rs2(1)
        rs2.Close
        Set rs2 = Nothing
    End If

    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        ' Modify Focus Events    
        %>
            <Script language=vbs>
            Parent.frm1.txtconBp_cd.focus    
            </Script>
        <%        
    Else    
        Call  MakeSpreadSheetData()
    End If
   
End Sub

'----------------------------------------------------------------------------------------------------------
' Set default value or preset value
'----------------------------------------------------------------------------------------------------------
Sub TrimData()
End Sub

%>

<Script Language=vbscript>

With parent

	.ggoSpread.Source    = .frm1.vspdData 
	.frm1.vspdData.Redraw = False
    .ggoSpread.SSShowDataByClip "<%=lgstrData%>", "F"                       '☜: Display data 
	
	Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,-1,-1,Parent.GetKeyPos("A",7),Parent.GetKeyPos("A",9),"C", "Q" ,"X","X")
    Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.Frm1.vspdData,-1,-1,Parent.GetKeyPos("A",7),Parent.GetKeyPos("A",10),"A", "Q" ,"X","X")
        
        
    .lgPageNo = "<%=lgPageNo%>"

    .frm1.txtconBp_Nm.value		= "<%=ConvSPChars(arrRsVal(1))%>" 
    .frm1.txtSOTypeNm.value		= "<%=ConvSPChars(arrRsVal(3))%>" 
    .frm1.txtSalesGroupNm.value = "<%=ConvSPChars(arrRsVal(5))%>" 
	.frm1.txtItem_Nm.value		= "<%=ConvSPChars(arrRsVal(7))%>"  

	.frm1.HtxtconBp_cd.value = "<%=ConvSPChars(txtconBp_cd)%>"
	.frm1.HtxtSoType.value	= "<%=ConvSPChars(txtSoType)%>"
	.frm1.HtxtSalesGroup.value	= "<%=ConvSPChars(txtSalesGroup)%>"
	.frm1.HtxtItem_cd.value	= "<%=ConvSPChars(txtItem_cd)%>"
	.frm1.HtxtSOFrDt.value	= "<%=txtSOFrDt%>"
	.frm1.HtxtSOToDt.value	= "<%=txtSOToDt%>"
	.frm1.HtxtRadio.value	= "<%=txtRadio%>"
	.frm1.HtxtTrackingNo.value = "<%=ConvSPChars(txtTrackingNo)%>"

    .DbQueryOk
    .frm1.vspdData.Redraw = True
End with	
</Script>	
<%
	Response.End													'☜: 비지니스 로직 처리를 종료함 
%>
