<%'===================================================================
'*  1. Module Name          : 영업 
'*  2. Function Name        : 출하관리 
'*  3. Program ID           : S4112QA1
'*  4. Program Name         : 출하상세조회 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2000/12/09
'*  8. Modified date(Last)  : 2001/12/19
'*  9. Modifier (First)     : Byun Jee Hyun
'* 10. Modifier (Last)      : Kim Hyungsuk
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/19	Date표준적용 
'*                            -2002/12/20 : Get방식 --> Post방식으로 변경 
'========================================

%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<%                                                                         '☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next

Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0  ,rs1, rs2, rs3, rs4, rs5                            '☜ : DBAgent Parameter 선언 
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
Dim arrRsVal(9)
Dim lgDataExist
Dim lgPageNo
Dim txtTrackingNo

MsgDisplayFlag = False															

	Call LoadBasisGlobalInf()
    Call HideStatusWnd 

if Request("txtFlgMode") = Request("OPMD_UMODE") Then
		txtMode= Request("txtMode")
		txtconBp_cd= Request("HtxtconBp_cd")
		txtDNType= Request("HtxtDNType")
		txtSalesGroup= Request("HtxtSalesGroup")
		txtPlant= Request("HtxtPlant")
		txtStoRo_cd= Request("HtxtStoRo_cd")
		txtDNFrDt= Request("HtxtDNFrDt")
		txtDNToDt= Request("HtxtDNToDt")
		txtRadio= Request("HtxtRadio")
		txtTrackingNo = Request("HtxtTrackingNo")
		lgStrPrevKey= Request("txt_lgStrPrevKey")
else
		txtMode= Request("txtMode")
		txtconBp_cd= Request("txtconBp_cd")
		txtDNType= Request("txtDNType")
		txtSalesGroup= Request("txtSalesGroup")
		txtPlant= Request("txtPlant")
		txtStoRo_cd= Request("txtStoRo_cd")
		txtDNFrDt= Request("txtDNFrDt")
		txtDNToDt= Request("txtDNToDt")
		txtRadio= Request("txtRadio")
		txtTrackingNo = Request("txtTrackingNo")
end if

    lgPageNo         = UNICInt(Trim(Request("txt_lgPageNo")),0)              '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount     = 100								                           '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList   = Request("txt_lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("txt_lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("txt_lgTailList")                                 '☜ : Orderby value
    lgDataExist      = "No"
'response.write "ccc"

    Call FixUNISQLData()
    Call QueryData()

'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------

Sub MakeSpreadSheetData()

    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(lgMaxCount) * CLng(lgPageNo)   'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
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

    If iLoopCount < lgMaxCount Then                                      '☜: Check if next data exists
       lgPageNo = ""
    End If
    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF

End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Function SetConditionData()
	
	On Error Resume Next
	SetConditionData = False	
	
	If Not(rs1.EOF Or rs1.BOF ) Then
		arrRsVal(0) = rs1(0)
		arrRsVal(1) = rs1(1)
        rs1.Close
        Set rs1 = Nothing
    Else    
        rs1.Close
        Set rs1 = Nothing

   		If Len(Trim(txtconBp_cd)) Then
		   Call DisplayMsgBox("970000", vbInformation, "납품처", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	       MsgDisplayFlag = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtconBp_cd.focus    
                </Script>
            <%       	       	
		End If
    End If
    
    
    If  Not(rs2.EOF Or rs2.BOF )Then		
		arrRsVal(2) = rs2(0)
		arrRsVal(3) = rs2(1)
        rs2.Close
        Set rs2 = Nothing        
    Else    
        rs2.Close
        Set rs2 = Nothing

		If MsgDisplayFlag = False Then
   			If Len(Trim(txtDNType)) Then
			   Call DisplayMsgBox("970000", vbInformation, "출하형태", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
			   MsgDisplayFlag = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtDNType.focus    
                </Script>
            <%       	       	
			   
			End If	
		End If
    End If
    

    If  Not(rs3.EOF Or rs3.BOF) Then
		arrRsVal(4) = rs3(0)
		arrRsVal(5) = rs3(1)
        rs3.Close
        Set rs3 = Nothing
    Else    
        rs3.Close
        Set rs3 = Nothing

   		If MsgDisplayFlag = False Then
			If Len(Trim(txtSalesGroup)) Then
			   Call DisplayMsgBox("970000", vbInformation, "영업그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
			   MsgDisplayFlag = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtSalesGroup.focus    
                </Script>
            <%       	       	
			   
			End If	
		End If    
    End If
   
    If  Not(rs4.EOF And rs4.BOF )Then
		arrRsVal(6) = rs4(0)
		arrRsVal(7) = rs4(1)
        rs4.Close
        Set rs4 = Nothing
    Else    
        rs4.Close
        Set rs4 = Nothing

   		If MsgDisplayFlag = False Then
			If Len(Trim(txtPlant)) Then
			   Call DisplayMsgBox("970000", vbInformation, "공장", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
			   MsgDisplayFlag = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtPlant.focus    
                </Script>
            <%       	       	
			   
			End If	
		End If	
    End If

	If  Not(rs5.EOF Or rs5.BOF ) Then
		arrRsVal(8) = rs5(0)
		arrRsVal(9) = rs5(1)
        rs5.Close
        Set rs5 = Nothing
    Else    

        rs5.Close
        Set rs5 = Nothing

		If MsgDisplayFlag = False Then
   			If Len(Trim(txtStoRo_cd)) Then
			   Call DisplayMsgBox("970000", vbInformation, "창고", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
			   MsgDisplayFlag = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtStoRo_cd.focus    
                </Script>
            <%       	       	
			   
			End If	
		End If
    End If    
	
	If MsgDisplayFlag = True Then Exit Function
	SetConditionData = True
	
End Function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
    Dim arrVal(4)
    Dim MajorCd
    Redim UNISqlId(5)                                                     '☜: SQL ID 저장을 위한 영역확보 
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

    Redim UNIValue(5,2)

    UNISqlId(0) = "S4112QA101"
    UNISqlId(1) = "s0000qa002"
    UNISqlId(2) = "s0000qa000"
    UNISqlId(3) = "s0000qa005"
    UNISqlId(4) = "s0000qa009"
    UNISqlId(5) = "s0000qa010"

    '--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    '--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	strVal = " "
		
	If Len(txtconBp_cd) Then
		strVal = "AND A.SHIP_TO_PARTY = " & FilterVar(txtconBp_cd, "''", "S") & " "		
	Else
		strVal = ""
	End If
	arrVal(0) = FilterVar(Trim(txtconBp_cd), " ", "S")

	If Len(txtDNType) Then
		strVal = strVal & " AND A.MOV_TYPE = " & FilterVar(txtDNType, "''", "S") & " "			
		Majorcd = "I0001"
	End If	
	arrVal(1) = FilterVar(Trim(txtDNType), " ", "S")	
		   
	If Len(txtSalesGroup) Then
		strVal = strVal & " AND A.SALES_GRP = " & FilterVar(txtSalesGroup, "''", "S") & " "	
	End If		
	arrVal(2) = FilterVar(Trim(txtSalesGroup), " ", "S")
    
    If Len(txtPlant) Then
		strVal = strVal & " AND B.PLANT_CD = " & FilterVar(txtPlant, "''", "S") & " "		
	End If		
	arrVal(3) = FilterVar(Trim(txtPlant), " ", "S")

	If Len(txtStoRo_cd) Then
		strVal = strVal & " AND B.SL_CD = " & FilterVar(txtStoRo_cd, "''", "S") & " "		
	End If		
    arrVal(4) = FilterVar(Trim(txtStoRo_cd), " ", "S")
    
    If Len(txtDNFrDt) Then
		strVal = strVal & " AND A.ACTUAL_GI_DT >= " & FilterVar(UNIConvDate(txtDNFrDt), "''", "S") & ""		
	End If		
	
	If Len(txtDNToDt) Then
		strVal = strVal & " AND A.ACTUAL_GI_DT <= " & FilterVar(UNIConvDate(txtDNToDt), "''", "S") & ""		
	End If

	If Trim(txtRadio) = "Y" Then
		strVal = strVal & " AND A.POST_FLAG = " & FilterVar(txtRadio, "''", "S") & ""		
	ElseIf Trim(txtRadio) = "N" Then
		strVal = strVal & " AND A.POST_FLAG <>" & FilterVar("Y", "''", "S") & " "	
	End If
	
	If Len(txtTrackingNo) Then
		strVal = strVal & " AND B.TRACKING_NO = " & FilterVar(Trim(txtTrackingNo), "''" , "S") & ""
	End If

    UNIValue(0,1) = strVal   '---발주일 
    UNIValue(1,0) = arrVal(0)
    UNIValue(2,0) = FilterVar(MajorCd, " ", "S")
    UNIValue(2,1) = arrVal(1)
    UNIValue(3,0) = arrVal(2)
    UNIValue(4,0) = arrVal(3)
    UNIValue(5,0) = arrVal(4)
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
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5)
    
    Set lgADF   = Nothing
    
'    FalsechkFlg = False
    

    iStr = Split(lgstrRetMsg,gColSep)
    
    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
    
    If SetConditionData = False Then Exit Sub 
    
    If  rs0.EOF And rs0.BOF  Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        MsgDisplayFlag = True
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



%>

<Script Language=vbscript>

With parent

    .frm1.txtconBp_nm.value		= "<%=ConvSPChars(arrRsVal(1))%>" 
    .frm1.txtDNTypeNm.value		= "<%=ConvSPChars(arrRsVal(3))%>" 
    .frm1.txtSalesGroupNm.value	= "<%=ConvSPChars(arrRsVal(5))%>" 
    .frm1.txtPlantNm.value		= "<%=ConvSPChars(arrRsVal(7))%>" 
    .frm1.txtStoRo_Nm.value		= "<%=ConvSPChars(arrRsVal(9))%>" 

    
	If "<%=lgDataExist%>" = "Yes" Then
       'Set condition data to hidden area

		If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			
			parent.frm1.HtxtconBp_cd.value = "<%=ConvSPChars(txtconBp_cd)%>"
			parent.frm1.HtxtDNType.value = "<%=ConvSPChars(txtDNType)%>"
			parent.frm1.HtxtSalesGroup.value = "<%=ConvSPChars(txtSalesGroup)%>"
			parent.frm1.HtxtPlant.value = "<%=ConvSPChars(txtPlant)%>"
			parent.frm1.HtxtStoRo_cd.value = "<%=ConvSPChars(txtStoRo_cd)%>"
			parent.frm1.HtxtDNFrDt.value = "<%=txtDNFrDt%>"
			parent.frm1.HtxtDNToDt.value = "<%=txtDNToDt%>"
			parent.frm1.HtxtRadio.value = "<%=txtRadio%>"
			parent.frm1.HtxtTrackingNo.value = "<%=ConvSPChars(txtTrackingNo)%>"
			
        End if 
    
       .frm1.vspdData.Redraw = False
        .ggoSpread.Source    = .frm1.vspdData 
        .ggoSpread.SSShowDataByClip "<%=lgstrData%>", "F"                            '☜: Display data 
        .lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
       .frm1.vspdData.Redraw = True
       .DbQueryOk
    End if

End with
</Script>	
<%
    Response.End													'☜: 비지니스 로직 처리를 종료함 
%>
