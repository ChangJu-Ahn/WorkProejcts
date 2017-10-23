<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 출하관리 
'*  3. Program ID           : s4511ab1
'*  4. Program Name         : 수주참조(출하등록)
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/04/12
'*  8. Modified date(Last)  : 2003/06/11
'*  9. Modifier (First)     : Cho inkuk
'* 10. Modifier (Last)      : Hwang Seongbae
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/18  Date 표준적용 
'*							  2002/04/12 ADO 변환 
'**********************************************************************************************
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<%
                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2, rs3, rs4, rs5			   '☜ : DBAgent Parameter 선언 
Dim lgArrData                                                 '☜ : Spread sheet에 보여줄 데이타를 위한 변수 

Dim lgMaxCount                                                '☜ : Spread sheet 의 visible row 수 
Dim lgTailList                                                '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgPageNo

	Call LoadBasisGlobalInf()
    Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgMaxCount     = 30							             '☜ : 한번에 가져올수 있는 데이타 건수 
	lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
	lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
	    
    Call FixUNISQLData()									 '☜ : DB-Agent로 보낼 parameter 데이타 set
    Call QueryData()										 '☜ : DB-Agent를 통한 ADO query

'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()

    Dim iArrRow
    Dim iRowCnt
    Dim iColCnt
	Dim iLngStartRow
    
    ReDim iArrRow(UBound(lgSelectListDT) - 1)
	
	iLngStartRow = CLng(lgMaxCount) * CLng(lgPageNo)
	
	' Scroll 조회시 Client로 보낼 첫 번재 Row로 이동한다.
    If CLng(lgPageNo) > 0 Then
       rs0.Move = iLngStartRow
    End If
    
    ' Client로 전송할 조회결과가 한 Page를 넘어갈 때 
    If rs0.RecordCount > CLng(lgMaxCount) * (CLng(lgPageNo) + 1) Then
        lgPageNo = lgPageNo + 1
	    Redim lgArrData(lgMaxCount - 1)

    ' Client로 전송할 조회결과가 한 Page를 넘지 않을 때, 즉 마지막 자료인 경우 
    Else
		Redim lgArrData(rs0.RecordCount - (iLngStartRow + 1))
		lgPageNo = ""
    End If

    For iRowCnt = 0 To UBound(lgArrData)
		For iColCnt = 0 To UBound(lgSelectListDT) - 1 
            iArrRow(iColCnt) = FormatRsString(lgSelectListDT(iColCnt),rs0(iColCnt))
		Next
		
		lgArrData(iRowCnt) = Chr(11) & Join(iArrRow, Chr(11))
		
        rs0.MoveNext
    Next

    rs0.Close                                                       '☜: Close recordset object
    Set rs0 = Nothing	                                            '☜: Release ADF
End Sub

' 조회조건 명 Display
'----------------------------------------------------------------------------------------------------------
Function SetConditionData()
    On Error Resume Next
    
    SetConditionData = False
    
	' 공장명 
	If Len(Request("txtConPlantCd")) Then
		If Not(rs1.EOF Or rs1.BOF) Then
			Call WriteConDesc("txtConPlantNm", rs1(1), rs1)
		Else
			Call ConNotFound("txtConPlantCd", rs1)
			Call DisplayMsgBox("970000", vbInformation, "공장", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
			Exit Function
		End If
	Else
		Call ClearConDesc("txtConPlantNm")
	End If

	' 납품처명    
	If Len(Request("txtConShipToParty")) Then
		If Not(rs2.EOF Or rs2.BOF) Then
			Call WriteConDesc("txtConShipToPartyNm", rs2(1), rs2)
		Else
			Call ConNotFound("txtConShipToParty", rs2)
			Call DisplayMsgBox("970000", vbInformation, "납품처", "", I_MKSCRIPT)
			Exit Function
		End If
	Else
		Call ClearConDesc("txtConShipToPartyNm")
	End If

	' 출하형태명 
	If Len(Request("txtConMovType")) Then
		If Not(rs3.EOF Or rs3.BOF) Then
			Call WriteConDesc("txtConMovTypeNm", rs3(1), rs3)
		Else
			Call ConNotFound("txtConMovType", rs3)
			Call DisplayMsgBox("970000", vbInformation, "출하형태", "", I_MKSCRIPT)
			Exit Function
		End If
	Else
		Call ClearConDesc("txtConMovTypeNm")
	End If

	' 수주형태명 
	If Len(Request("txtConSOType")) Then
		If Not(rs4.EOF Or rs4.BOF) Then
			Call WriteConDesc("txtConSOTypeNm", rs4(1), rs4)
		Else
			Call ConNotFound("txtConSOType", rs4)
			Call DisplayMsgBox("970000", vbInformation, "수주형태", "", I_MKSCRIPT)
			Exit Function
		End If
	Else
		Call ClearConDesc("txtConSOTypeNm")
	End If

	' 영업그룹명 
	If Len(Request("txtConSalesGrp")) Then
		If Not(rs5.EOF Or rs5.BOF) Then
			Call WriteConDesc("txtConSalesGrpNm", rs5(1), rs5)
		Else
			Call ConNotFound("txtConSalesGrp", rs5)
			Call DisplayMsgBox("970000", vbInformation, "영업그룹", "", I_MKSCRIPT)
			Exit Function
		End If
	Else
		Call ClearConDesc("txtConSalesGrpNm")
	End If
	
	SetConditionData = True
End Function

'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim iStrVal
    Redim UNISqlId(5)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(5,2)


	If Request("lgStrAllocInvFlag") = "N" Then
		UNISqlId(0) = "S4511AA101"
	Else
		UNISqlId(0) = "S4511AA102"
	End If
    UNIValue(0,0) = Trim(lgSelectList)                                      '☜: Select list
	    
	'========= 조회조건 
	' 공장 
    UNISqlId(1) = "S0000QA009"
    UNIValue(1,0) = FilterVar(Trim(Request("txtConPlantCd")), " ", "S")
	iStrVal = " AND	SD.PLANT_CD =  " & FilterVar(Trim(Request("txtConPlantCd")), "''", "S") & ""		' 공장 
	

	' 납기일 
	iStrVal = iStrVal & " AND SD.DLVY_DT >=  " & FilterVar(UNIConvDate(Request("txtConFrDlvyDt")), "''", "S") & ""		
	iStrVal = iStrVal & " AND SD.DLVY_DT <=  " & FilterVar(UNIConvDate(Request("txtConToDlvyDt")), "''", "S") & ""		


	' 납품처 
	If Len(Request("txtConShipToParty")) Then
	    UNISqlId(2) = "S0000QA002"
	    UNIValue(2,0) = FilterVar(Trim(Request("txtConShipToParty")), " ", "S")		'납품처코드 
		iStrVal = iStrVal & " AND SD.SHIP_TO_PARTY =  " & FilterVar(Trim(Request("txtConShipToParty")), "''", "S") & ""
	End If


	' 출하형태 
	If Len(Request("txtConMovType")) Then
		UNISqlId(3) = "S0000QA000"					'출하형태명 
    	UNIValue(3,0) = FilterVar("I0001", " ", "S")					    '출하형태코드 
		UNIValue(3,1) = FilterVar(Trim(Request("txtConMovType")), " ", "S")
		iStrVal = iStrVal & " AND STC.MOV_TYPE =  " & FilterVar(Trim(Request("txtConMovType")), "''", "S") & ""		
	End If

	' 수주형태 
 	If Len(Request("txtConSOType")) Then
	    UNISqlId(4) = "S0000QA007"					'수주형태명  
	    UNIValue(4,0) = FilterVar(Trim(Request("txtConSOType")), " ", "S")
		iStrVal = iStrVal & " AND SH.SO_TYPE =  " & FilterVar(Trim(Request("txtConSOType")), "''", "S") & ""		
	End If
	
	' 영업그룹 
 	If Len(Request("txtConSalesGrp")) Then
	    UNISqlId(5) = "S0000QA005"
	    UNIValue(5,0) = FilterVar(Trim(Request("txtConSalesGrp")), " ", "S")

		iStrVal = iStrVal & " AND SH.SALES_GRP =  " & FilterVar(Trim(Request("txtConSalesGrp")), "''", "S") & ""		
	End If	    

	If Len(Request("txtConSoNo")) Then
		iStrVal = iStrVal & " AND SD.SO_NO =  " & FilterVar(Request("txtConSoNo"), "''", "S") & ""	
	End If

    UNIValue(0,1) = iStrVal   
    
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)
	
	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
		Exit Sub
    End If 
         
	Call BeginScriptTag
	
	' 처음 조회시 조회조건 명을 Display한다.
	If lgPageNo = 0 Then
		If Not SetConditionData	Then Exit Sub
		Call SetHiddenQueryCon()
	End If
	
    If rs0.EOF And rs0.BOF Then
		Call ConNotFound("txtConPlantCd", rs0)
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
    Else    
        Call MakeSpreadSheetData()
        Call WriteResult()
		Call EndScriptTag()
    End If  
End Sub

'----------------------------------------------------------------------------------------------------------
' Write the Result
'----------------------------------------------------------------------------------------------------------
Sub BeginScriptTag()
	Response.Write "<Script language=VBScript> " & VbCr
End Sub

Sub EndScriptTag()
	Response.Write "</Script> " & VbCr
End Sub

' 조회조건이 존재하지 않는 경우 Focus 처리 
Sub ConNotFound(ByVal pvStrField, ByRef prObjRs)
	Response.Write " Parent.frm1." & pvStrField & ".focus " & VbCr
	prObjRs.Close
	Set prObjRs = Nothing
	Call EndScriptTag()
End Sub

' 조회조건이 입력되지 않은 관련 명을 clear 시킨다.
Sub ClearConDesc(ByVal pvStrField)
	Response.Write " Parent.frm1." & pvStrField & ".value = """"" & VbCr
End Sub

' 조회조건에 해당하는 명을 Display하는 Script 작성 
Sub WriteConDesc(ByVal pvStrField, ByVal pvStrFieldDesc, ByRef prObjRs)
	Response.Write " Parent.frm1." & pvStrField & ".value = """ & ConvSPChars(pvStrFieldDesc) & """" &VbCr
	prObjRs.Close
	Set prObjRs = Nothing
End Sub

' 다음(Scrollbar) 조회를 위한 조회조건을 Hidden 필드에 설정 
Sub SetHiddenQueryCon()
    Response.Write "With parent.frm1" & vbCr
    Response.Write " .HPlantCd.value		 = """ & ConvSPChars(Request("txtConPlantCd")) & """" & vbCr
    Response.Write " .HShipToParty.value = """ & ConvSPChars(Request("txtConShipToParty")) & """" & vbCr
    Response.Write " .HFrDlvyDt.value	 = """ & Request("txtConFrDlvyDt") & """" & vbCr
    Response.Write " .HToDlvyDt.value	 = """ & Request("txtConToDlvyDt") & """" & vbCr
    Response.Write " .HMovType.value		 = """ & ConvSPChars(Request("txtConMovType")) & """" & vbCr
    Response.Write " .HSOType.value		 = """ & ConvSPChars(Request("txtConSOType")) & """" & vbCr
    Response.Write " .HSalesGrp.value	 = """ & ConvSPChars(Request("txtConSalesGrp")) & """" & vbCr
    Response.Write " .HSoNo.value		 = """ & ConvSPChars(Request("txtConSoNo")) & """" & vbCr
    Response.Write "End with" & vbCr
End Sub

' 조회 결과를 Display하는 Script 작성 
Sub WriteResult()
	Response.Write " Parent.ggoSpread.Source  = Parent.frm1.vspdData " & vbCr
	Response.Write " Parent.frm1.vspdData.Redraw = False " & vbCr      	
	Response.Write " Parent.ggoSpread.SSShowDataByClip  """ & Join(lgArrData, Chr(11) & Chr(12)) & Chr(11) & Chr(12) & """ ,""F""" & vbCr
	Response.Write " parent.lgPageNo = """ & lgPageNo & """" & vbCr	
	Response.Write " Parent.DbQueryOk " & vbCr		
 	Response.Write " Parent.frm1.vspdData.Redraw = True " & vbCr      
End Sub
%>
