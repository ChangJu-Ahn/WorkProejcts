<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 출하관리 
'*  3. Program ID           : S4111MA8
'*  4. Program Name         : 출하현황조회 
'*  5. Program Desc         :
'*  6. Comproxy List        : S41118ListDnHdrSvr
'*  7. Modified date(First) : 2000/04/17
'*  8. Modified date(Last)  : 2002/04/11
'*  9. Modifier (First)     : Cho song hyon
'* 10. Modifier (Last)      : RYU KYUNG RAE(1)
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            -2000/04/17 : 화면 Layout & ASP Coding
'*                            -2001/12/19 : Date 표준적용 
'*                            -2002/04/11 : ADO 변환 
'**********************************************************************************************

'								'☜ : ASP가 캐쉬되지 않도록 한다.
'								'☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.

'☜ : 항상 서버 사이드 구문의 시작점인 좌꺽쇠(<)% 와 %우꺽쇠(>)는 New Line에 위치하여 
'	  서버 사이드 구문과 클라이언트 사이드 구문의 위치를 가늠할 수 있도록 한다.
'☜ : 아래 HTML 구문은 변경되어서는 안된다. 
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next
'============================================  2002-04-10 시작  =============================================
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4
															'☜ : DBAgent Parameter 선언 
Dim lgStrData                                               '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
Dim lgMaxCount                                              '☜ : Spread sheet 의 visible row 수 
Dim lgTailList
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo

'--------------- 개발자 coding part(변수선언,Start)--------------------------------------------------------
DIM strMovType		'출하형태 
DIM strSoNo			'수주번호 
DIM strBPartner		'납품처 
DIM strTranMeth		'운송방법 
DIM strPostFlag		'출고여부 
Dim strSalesGrp		'영업그룹 
DIM strReqFromDate
DIM strReqToxxDate
Dim arrRsVal(7)												'☜ : QueryData()실행시 레코드셋을 배열로 받을때 사용 
															'☜ : 받을 레코드셋의 갯수만큼 배열 크기 선언			
	MsgDisplayFlag = False															
'--------------- 개발자 coding part(변수선언,End)----------------------------------------------------------
   
	Call LoadBasisGlobalInf()
	Call HideStatusWnd
	
    lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount       = 100						                       '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList     = Request("lgSelectList")
    lgTailList       = Request("lgTailList")
    lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '☜ : 각 필드의 데이타 타입 
    lgDataExist      = "No"

    Call  FixUNISQLData()                                                '☜ : DB-Agent로 보낼 parameter 데이타 set
    call  QueryData()                                                    '☜ : DB-Agent를 통한 ADO query

'/////////////////////////////////////////////////////////////////////////////////////////////
'Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
'Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
'Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0				'☜ : DBAgent Parameter 선언 
'Dim rs1, rs2, rs3, rs4, rs5, rs6							'☜ : DBAgent Parameter 선언 
'Dim lgStrData                                               '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
'Dim lgStrPrevKey                                            '☜ : 이전 값 
'Dim lgMaxCount                                              '☜ : Spread sheet 의 visible row 수 
'Dim lgTailList
'Dim lgSelectList
'Dim lgSelectListDT
'============================================  2002-04-10 끝  ===============================================

' Make srpread sheet data
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
	
'	On Error Resume Next
 
	SetConditionData = False
    
    If Not(rs1.EOF Or rs1.BOF) Then
        strMovType =  rs1(1)
        Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Request("txtDn_Type")) And MsgDisplayFlag = False Then
			Call DisplayMsgBox("970000", vbInformation, "출하형태", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    MsgDisplayFlag = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtDn_Type.focus    
                </Script>
            <%        		    	
		End If
	End If   	    
     
	If Not(rs2.EOF Or rs2.BOF) Then
        strBPartner =  rs2(1)
        Set rs2 = Nothing
    Else
		Set rs2 = Nothing
		If Len(Request("txtShip_to_party")) And MsgDisplayFlag = False Then
			Call DisplayMsgBox("970000", vbInformation, "납품처", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    MsgDisplayFlag = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtShip_to_party.focus    
                </Script>
            <%        		    			    	
		End If			
    End If   	

    If Not(rs3.EOF Or rs3.BOF) Then
        strTranMeth =  rs3(1)
        Set rs3 = Nothing
    Else
		Set rs3 = Nothing
		If Len(Request("txtTrans_meth")) And MsgDisplayFlag = False Then
			Call DisplayMsgBox("970000", vbInformation, "운송방법", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    MsgDisplayFlag = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtTrans_meth.focus    
                </Script>
            <%        		    			    	
		End If				
    End If
    
    If Not(rs4.EOF Or rs4.BOF) Then
        strSalesGrp =  rs4(1)
        Set rs4 = Nothing
    Else
		Set rs4 = Nothing
		If Len(Request("txtSalesGrp")) And MsgDisplayFlag = False Then
			Call DisplayMsgBox("970000", vbInformation, "영업그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    MsgDisplayFlag = True	
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
	Dim arrVal(3)
	Redim UNISqlId(4)                                                     '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------	
	Redim UNIValue(4,2)

	UNISqlId(0) = "S4111MA801"
	UNISqlId(1) = "s0000qa000"    ' 출하형태    'I0001'
	UNISqlId(2) = "s0000qa002"    ' 납품처 
	UNISqlId(3) = "s0000qa000"    ' 운송방법	'B9009'
	UNISqlId(4) = "s0000qa005"    ' 영업그룹 
	
	'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

	UNIValue(0,0) = Trim(lgSelectList)		                              '☜: Select 절에서 Summary    필드 

	'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
	
	strVal = " "
		
    '---출하형태 
    If Len(Trim(Request("txtDn_Type"))) Then
    	strVal = strval & " AND B.MOV_TYPE =  " & FilterVar(Request("txtDn_Type"), "''", "S") & " "    			
    End If
    arrVal(0) = FilterVar(Trim(Request("txtDn_Type")), " " , "S")
	
	'---납품처 
	If Len(Trim(Request("txtShip_to_party"))) Then
    	strVal = strval & " AND B.SHIP_TO_PARTY =  " & FilterVar(Request("txtShip_to_party"), "''", "S") & " "    	
    End If
    arrVal(1) = FilterVar(Trim(Request("txtShip_to_party")), " " , "S")
    
	'---수주번호 
	If Len(Trim(Request("txtSo_no"))) Then
    	strVal = strval & " AND B.SO_NO =  " & FilterVar(Request("txtSo_no"), "''", "S") & " "
    End If    
    
    '---운송방법 
	If Len(Trim(Request("txtTrans_meth"))) Then
    	strVal = strval & " AND B.TRANS_METH =  " & FilterVar(Request("txtTrans_meth"), "''", "S") & " "    	
    End If
    arrVal(2) = FilterVar(Trim(Request("txtTrans_meth")), " " , "S")

	'---영업그룹 
	If Len(Trim(Request("txtSalesGrp"))) Then
    	strVal = strval & " AND B.SALES_GRP =  " & FilterVar(Request("txtSalesGrp"), "''", "S") & " "    	
    End If
    arrVal(3) = FilterVar(Trim(Request("txtSalesGrp")), "''", "S")
    
    '---출고여부 
	If Len(Trim(Request("txtPostGiFlag"))) Then
    	strVal = strval & " AND B.POST_FLAG =  " & FilterVar(Request("txtPostGiFlag"), "''", "S") & ""
    End If

     '---출고요청일 
    If Len(Trim(Request("txtReqGiDtFrom"))) Then
    	strVal = strval & " AND B.PROMISE_DT >=  " & FilterVar(uniConvDate(Trim(Request("txtReqGiDtFrom"))), "''", "S") & ""
    End If
    
    If Len(Trim(Request("txtReqGiDtTo"))) Then
    	strVal = strval & " AND B.PROMISE_DT <=  " & FilterVar(uniConvDate(Trim(Request("txtReqGiDtTo"))), "''", "S") & ""
    End If
		   
    UNIValue(0,1) = strVal   
    UNIValue(1,0) = FilterVar("I0001", "''", "S")    					'출하형태 
    UNIValue(1,1) = arrVal(0)					'출하형태 

    UNIValue(2,0) = arrVal(1)				    '납품처 
	UNIValue(3,0) = FilterVar("B9009", "''", "S") 					'운송방법 
	UNIValue(3,1) = arrVal(2)
	UNIValue(4,0) = arrVal(3)					'영업그룹	

	'--------------- 개발자 coding part(실행로직,End)----------------------------------------------------
	UNIValue(0,UBound(UNIValue,2)    ) = Trim(lgTailList)	'---Order By 조건 
	UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode

End Sub
'----------------------------------------------------------------------------------------------------------
' Query Data
' ADO의 Record Set이용하여 Query를 하고 Record Set을 넘겨서 MakeSpreadSheetData()으로 Spreadsheet에 데이터를 
' 뿌림 
' ADO 객체를 생성할때 prjPublic.dll파일을 이용한다.(상세내용은 vb로 작성된 prjPublic.dll 소스 참조)
'----------------------------------------------------------------------------------------------------------
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub QueryData()

	Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
	Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
	Dim iStr
	    
	Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
	lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)
	    
	Set lgADF   = Nothing

	iStr = Split(lgstrRetMsg,gColSep)
   
   
	If iStr(0) <> "0" Then
		Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
	End If    
	
	
   	If SetConditionData = False Then Exit Sub

   	
	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
		rs0.Close
		Set rs0 = Nothing
		MsgDisplayFlag = True
        ' Modify Focus Events    
        %>
            <Script language=vbs>
            Parent.frm1.txtDn_Type.focus    
            </Script>
        <%
	Else    
		Call  MakeSpreadSheetData()
	End If

'	Call  SetConditionData()
    
End Sub

%>
<Script Language=vbscript>

With Parent 	

	.frm1.txtDn_TypeNm.value			= "<%=ConvSPChars(strMovType)%>"
	.frm1.txtShip_to_partyNm.value		= "<%=ConvSPChars(strBPartner)%>"
	.frm1.txtTrans_meth_nm.value		= "<%=ConvSPChars(strTranMeth)%>"
	.frm1.txtSalesGrpNm.value			= "<%=ConvSPChars(strSalesGrp)%>"

	If "<%=lgDataExist%>" = "Yes" Then
       'Set condition data to hidden area

		If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			.frm1.txtHDn_Type.value			= "<%=ConvSPChars(Request("txtDn_Type"))%>"
			.frm1.txtHSo_no.value			= "<%=ConvSPChars(Request("txtSo_no"))%>"
			.frm1.txtHShip_to_party.value	= "<%=ConvSPChars(Request("txtShip_to_party"))%>"
			.frm1.txtHReqGiDtFrom.value		= "<%=ConvSPChars(Request("txtReqGiDtFrom"))%>"
			.frm1.txtHReqGiDtTo.value		= "<%=ConvSPChars(Request("txtReqGiDtTo"))%>"
			.frm1.txtHTrans_meth.value		= "<%=ConvSPChars(Request("txtTrans_meth"))%>"
			.frm1.txtHPostGiFlag.value		= "<%=ConvSPChars(Request("txtPostGiFlag"))%>"	    
			.frm1.txtHSalesGrp.value		= "<%=ConvSPChars(Request("txtSalesGrp"))%>"	    
		End If

       'Show multi spreadsheet data from this line
       .frm1.vspdData.Redraw = False
       .ggoSpread.Source  = .frm1.vspdData
       .ggoSpread.SSShowDataByClip "<%=lgstrData%>", "F"          '☜ : Display data
       .lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
       .frm1.vspdData.Redraw = True
       .DbQueryOk
    
    End If  
    
End With     
</Script>
