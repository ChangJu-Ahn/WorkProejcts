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
'''' On Error Resume Next                                                                         

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

	lgPageNo       = UNICInt(Trim(Request("txtHlgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgMaxCount     = 100							             '☜ : 한번에 가져올수 있는 데이타 건수 
	lgSelectList   = Request("txtHlgSelectList")                               '☜ : select 대상목록 
	lgTailList     = Request("txtHlgTailList")                                 '☜ : Orderby value
	lgSelectListDT = Split(Request("txtHlgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
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
	Dim strWhere

    Redim UNISqlId(2)                                                     '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(2,2)

    UNISqlId(0) = "S4512RA101"
	UNISqlId(1) = "S0000QA002"  ' 납품처 
	UNISqlId(2) = "S0000QA005"  ' 영업그룹 


	UNIValue(1,0) = FilterVar(Trim(Request("txtHPtnBpCd")), " ", "S")		'납품처코드 
	UNIValue(2,0) = FilterVar(Trim(Request("txtHSalesGrp")), " ", "S")

'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    UNIValue(0,0) = Trim(lgSelectList)                                      '☜: Select list

	
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    
'	2006-10-24 10:21:51 AM - 안준선:
'	s5\s4512rb1.asp에서 
'	========================================
'	strWhere = " WHERE SD.DN_REQ_QTY > 0 AND SD.DN_REQ_QTY > SD.REQ_QTY AND  DD.REQ_QTY >  ISNULL(DN.DN_REQ_QTY, 0) "
'	========================================
'	를 
'	========================================
	strWhere = " WHERE SD.DN_REQ_QTY > 0 "
	strWhere = strWhere & " AND SD.DN_REQ_QTY >  (SELECT ISNULL(SUM(GI_QTY), 0) FROM S_DN_DTL WHERE SO_NO = SD.SO_NO AND SO_SEQ = SD.SO_SEQ) "
	strWhere = strWhere & " AND DD.REQ_QTY >  ISNULL(DN.DN_REQ_QTY, 0) "
'	========================================
'	로 변경 

    If Len(Trim(Request("txtHFromDt"))) Then
		strWhere = strWhere & " AND DH.PROMISE_DT >= " & FilterVar(UNIConvDate(Request("txtHFromDt")), "''", "S") & ""                             '시작일 
	End If

    If Len(Trim(Request("txtHToDt"))) Then
		strWhere = strWhere & " AND DH.PROMISE_DT <= " & FilterVar(UNIConvDate(Request("txtHToDt")), "''", "S") & ""                             '시작일 
	End If

    If Len(Trim(Request("txtHSalesGrp"))) Then
		strWhere = strWhere & " AND DH.SALES_GRP = " & FilterVar(Request("txtHSalesGrp"), "''", "S") & ""                             '시작일 
	End If
    If Len(Trim(Request("txtHSoNo"))) Then
		strWhere = strWhere & " AND DD.SO_NO = " & FilterVar(Request("txtHSoNo"), "''", "S") & ""                             '시작일 
	End If
    If Len(Trim(Request("txtHPtnBpCd"))) Then
		strWhere = strWhere & " AND DH.SHIP_TO_PARTY = " & FilterVar(Request("txtHPtnBpCd"), "''", "S") & ""                             '시작일 
	End If
    If Len(Trim(Request("txtHDnReqNo"))) Then
		strWhere = strWhere & " AND DH.DN_REQ_NO >= " & FilterVar(Request("txtHDnReqNo"), "''", "S") & ""                             '시작일 
	End If
    If Len(Trim(Request("txtHItem"))) Then
		strWhere = strWhere & " AND DD.ITEM_CD = " & FilterVar(Request("txtHItem"), "''", "S") & ""                             '시작일 
	End If
    If Len(Trim(Request("txtHPlantCd"))) Then
		strWhere = strWhere & " AND DD.PLANT_CD= " & FilterVar(Request("txtHPlantCd"), "''", "S") & ""                             '시작일 
	End If

	UNIValue(0, 1) = strWhere                                      '☜: WHERE 절 
	UNIValue(0, 2) = UCase(Trim(lgTailList))
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

	' 납품처명    
	If Len(Request("txtHPtnBpCd")) Then
		If Not(rs1.EOF Or rs1.BOF) Then
			Response.Write "<Script language=VBScript> " & VbCr
			Response.Write " Parent.frm1.txtPtnBpNm.value = """ & ConvSPChars(rs1(1)) & """" &VbCr
			Response.Write "</Script> " & VbCr
		Else
			Call DisplayMsgBox("970000", vbInformation, "납품처", "", I_MKSCRIPT)
			Response.Write "<Script language=VBScript> " & VbCr
			Response.Write " Parent.frm1.txtPtnBpNm.value = """"" & VbCr
			Response.Write " Parent.frm1.txtPtnBpCd.Focus " & VbCr
			Response.Write "</Script> " & VbCr
			Exit Sub
		End If
	Else
		Response.Write "<Script language=VBScript> " & VbCr
		Response.Write " Parent.frm1.txtPtnBpNm.value = """"" & VbCr
		Response.Write "</Script> " & VbCr
	End If

	' 영업그룹명    
	If Len(Request("txtHSalesGrp")) Then
		If Not(rs2.EOF Or rs2.BOF) Then
			Response.Write "<Script language=VBScript> " & VbCr
			Response.Write " Parent.frm1.txtSalesGrpNm.value = """ & ConvSPChars(rs2(1)) & """" &VbCr
			Response.Write "</Script> " & VbCr
		Else
			Call DisplayMsgBox("970000", vbInformation, "영업그룹", "", I_MKSCRIPT)
			Response.Write "<Script language=VBScript> " & VbCr
			Response.Write " Parent.frm1.txtSalesGrpNm.value = """"" & VbCr
			Response.Write " Parent.frm1.txtSalesGrp.Focus " & VbCr
			Response.Write "</Script> " & VbCr
			Exit Sub
		End If
	Else
		Response.Write "<Script language=VBScript> " & VbCr
		Response.Write " Parent.frm1.txtSalesGrpNm.value = """"" & VbCr
		Response.Write "</Script> " & VbCr
	End If

	
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
		<%IF UNICInt(Trim(Request("txtHlgPageNo")),0) = 1 Then%>
		.frm1.txtHFromDt.value	= "<%=Request("txtHFromDT")%>"
		.frm1.txtHToDt.value	= "<%=Request("txtHToDT")%>"
		<%End If%>
		
		'Show multi spreadsheet data from this line
		.ggoSpread.Source		= .frm1.vspdData 
        .frm1.vspdData.Redraw = False
		.ggoSpread.SSShowDataByClip "<%=lgstrData%>","F"						'☜ : Display data
'		Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspdData.MaxRows,"<%=Request("txtCurrency")%>",.GetKeyPos("A",7),"C","I","X","X")
'		Call .ReFormatSpreadCellByCellByCurrency2(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspdData.MaxRows,"<%=Request("txtCurrency")%>",.GetKeyPos("A",8),"A","I","X","X")
		.lgPageNo				=  "<%=lgPageNo%>"							  '☜: Next next data tag
		.DbQueryOk
        .frm1.vspdData.Redraw = True                		
	End with
</Script>	
<%End If%>
