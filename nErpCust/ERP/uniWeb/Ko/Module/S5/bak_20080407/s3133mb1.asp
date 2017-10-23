<%'===================================================================
'*  1. Module Name          : 영업 
'*  2. Function Name        : 출하관리 
'*  3. Program ID           : S3133mb1
'*  4. Program Name         : 미출하생성현황조회 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2001/12/19
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Seo Jinkyung
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2002/09/18 Ado 표준적용 
'*                            -2002/12/20 : Get방식 --> Post방식으로 변경 
'========================================

                                                       '☜ : ASP가 캐쉬되지 않도록 한다.
                                                     '☜ : ASP가 버퍼에 저장되지 않고 바로 Client에 내려간다.
%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<%													'☜ : 여기서 부터 개발자 비지니스 로직을 처리하는 내용이 시작된다 

On Error Resume Next														

	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4	'☜ : DBAgent Parameter 선언 
	Dim lgStrData														'☜ : Spread sheet에 보여줄 데이타를 위한 변수 
	Dim lgMaxCount														'☜ : Spread sheet 의 visible row 수 
	Dim lgTailList
	Dim lgSelectList
	Dim lgSelectListDT
	Dim lgDataExist
	Dim lgPageNo

'--------------- 개발자 coding part(변수선언,Start)----------------------------------------------------
	Dim strSalesGrpNm					'영업그룹 
	Dim strItemCodeNm						'품목 
	Dim strSoldToPartyNm						'거래처 
	Dim strSoTypeNm						'수주형태 
	Dim MsgDisplayFlag
   
	MsgDisplayFlag = False
	Dim iFrPoint
    iFrPoint=0
'--------------- 개발자 coding part(변수선언,End)------------------------------------------------------

	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("Q", "S", "NOCOOKIE", "QB")
	Call LoadBNumericFormatB("Q","S","NOCOOKIE","QB")
	
    Call HideStatusWnd 
     
if Request("txtFlgMode") = Request("OPMD_UMODE") Then
		txtMode= Request("txtMode")
		txtSalesGrp= Request("HSalesGrp")
		txtSoDtFrom= Request("HSoDtFrom")
		txtSoDtTo= Request("HSoDtTo")
		txtDlvyDtFrom= Request("HDlvyDtFrom")
		txtDlvyDtTo= Request("HDlvyDtTo")
		txtSoldToParty= Request("HSoldToParty")
		txtItemCode= Request("HItemCode")
		txtSoType= Request("HSoType")
		txtTrackingNo = Request("HtxtTrackingNo")
		lgStrPrevKey= Request("txt_lgStrPrevKey")
else
		txtMode= Request("txtMode")
		txtSalesGrp= Request("txtSalesGrp")
		txtSoDtFrom= Request("txtSoDtFrom")
		txtSoDtTo= Request("txtSoDtTo")
		txtDlvyDtFrom= Request("txtDlvyDtFrom")
		txtDlvyDtTo= Request("txtDlvyDtTo")
		txtSoldToParty= Request("txtSoldToParty")
		txtItemCode= Request("txtItemCode")
		txtSoType= Request("txtSoType")
		txtTrackingNo = Request("txtTrackingNo")
		lgStrPrevKey= Request("txt_lgStrPrevKey")
end if

	lgStrPrevKey	 = Request("txt_lgStrPrevKey")
    lgPageNo         = UNICInt(Trim(Request("txt_lgPageNo")),0)              '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgMaxCount       = 100								                       '☜ : 한번에 가져올수 있는 데이타 건수 
    lgSelectList     = Request("txt_lgSelectList")
    lgTailList       = Request("txt_lgTailList")
    lgSelectListDT   = Split(Request("txt_lgSelectListDT"), gColSep)         '☜ : 각 필드의 데이타 타입 
    lgDataExist      = "No"

    Call  FixUNISQLData()                                                '☜ : DB-Agent로 보낼 parameter 데이타 set
    call  QueryData()                                                    '☜ : DB-Agent를 통한 ADO query
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move = CLng(lgMaxCount) * CLng(lgPageNo)                  'lgMaxCount:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
       iFrPoint	= CLng(lgMaxCount) * CLng(lgPageNo)
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
Function SetConditionData()

	SetConditionData = False
	
    On Error Resume Next
	
	If Not(rs1.EOF Or rs1.BOF) Then
       strSalesGrpNm =  rs1(1)
        rs1.Close
        Set rs1 = Nothing       
    Else
        rs1.Close
        Set rs1 = Nothing
            
		If Len(txtSalesGrp) And MsgDisplayFlag = False Then
			Call DisplayMsgBox("970000", vbInformation, "영업그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
		    MsgDisplayFlag = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtSalesGrp.focus    
                </Script>
            <%
		End If	
    End If   

    
	
	If Not(rs2.EOF Or rs2.BOF) Then
       strItemCodeNm =  rs2(1)
        rs2.Close
        Set rs2 = Nothing               
    Else
        rs2.Close
        Set rs2 = Nothing        
   		If Len(txtItemCode) And MsgDisplayFlag = False Then
		    Call DisplayMsgBox("970000", vbInformation, "품목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
		    MsgDisplayFlag = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtItemCode.focus    
                </Script>
            <%		    
		End If	
    End If   
    
    
    
	If Not(rs3.EOF Or rs3.BOF) Then
	    strSoldToPartyNm =  rs3(1)
		rs3.Close
		Set rs3 = Nothing       
    Else
        rs3.Close
        Set rs3 = Nothing    
   		If Len(txtSoldToParty) And MsgDisplayFlag = False Then
		    Call DisplayMsgBox("970000", vbInformation, "주문처", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
		    MsgDisplayFlag = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtSoldToParty.focus    
                </Script>
            <%		    
		End If	
    End If   
    
    
    
    If Not(rs4.EOF Or rs4.BOF) Then
       strSoTypeNm =  rs4(1)
        rs4.Close
        Set rs4 = Nothing       
    Else
        rs4.Close
        Set rs4 = Nothing    
   		If Len(txtSoType) And MsgDisplayFlag = False Then
    		Call DisplayMsgBox("970000", vbInformation, "수주형태", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code
	    	MsgDisplayFlag = True	
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtSoType.focus    
                </Script>
            <%	    	
		End If	
    End If   
    
    

	If MsgDisplayFlag = True Then Exit Function

	SetConditionData = True
	
End Function
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal															  '☜:UNISqlId(0)에 들어가는 입력변수 
																		  '아래에 보면 화면단에서 넣어 주는 query시 where조건임을 알 수 있다.	
    Dim arrVal(3)														  '☜: 화면에서 팝업하여 query
																		  '아래에 보면 UNISqlId(1),UNISqlId(2), UNISqlId(3)의 where조건임을 알 수 있다.
    Redim UNISqlId(4)                                                     '☜: SQL ID 저장을 위한 영역확보 
																		  '조회화면에서 필요한 query조건문들의 영역(Statements table에 있음)
    Redim UNIValue(4,2)													  '☜: 각각의 SQL ID와 입력될 where 조건의 쌍으로 된 2차원 배열 

    UNISqlId(0) = "S3133QA101"											  ' main query(spread sheet에 뿌려지는 query statement)
    UNISqlId(1) = "s0000qa005"											  ' main query(spread sheet에 뿌려지는 query statement)
    UNISqlId(2) = "s0000qa001"											  ' main query(spread sheet에 뿌려지는 query statement)
    UNISqlId(3) = "s0000qa002"											  ' main query(spread sheet에 뿌려지는 query statement)
    UNISqlId(4) = "s0000qa007"											  ' main query(spread sheet에 뿌려지는 query statement)
    
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
																		  '	UNISqlId(0)의 첫번째 ?에 입력됨				
	strVal = ""
     '---발생일자 
    If Len(Trim(txtSoDtFrom)) Then
    	strVal 	= strVal & "AND a.so_dt >=  " & FilterVar(uniConvDate(Trim(txtSoDtFrom)), "''", "S") & " "
    End If
    
    If Len(Trim(txtSoDtTo)) Then
    	strVal 	= strVal & "AND a.so_dt <= " & FilterVar(uniConvDate(Trim(txtSoDtTo)), "''", "S") & " "
    End If    
    
    If Len(Trim(txtDlvyDtFrom)) Then
    	strVal 	= strVal & "AND b.DLVY_DT >=  " & FilterVar(uniConvDate(Trim(txtDlvyDtFrom)), "''", "S") & " "
    End If
    
    If Len(Trim(txtDlvyDtTo)) Then
    	strVal 	= strVal & "AND b.DLVY_DT <= " & FilterVar(uniConvDate(Trim(txtDlvyDtTo)), "''", "S") & " "
    End If    

    '---영업그룹 
    If Len(Trim(txtSalesGrp)) Then
    	strVal 	  = strVal & "AND c.sales_grp =  " & FilterVar(txtSalesGrp, "''", "S") & "  "    	
    End If
    arrVal(0) = FilterVar(Trim(txtSalesGrp), " ", "S")

	'---품목 
	If Len(Trim(txtItemCode)) Then
    	strVal 	  = strVal & "AND f.item_cd  =  " & FilterVar(txtItemCode, "''", "S") & "  "      	
    End If
    arrVal(1) = FilterVar(Trim(txtItemCode), " ", "S")

	'---거래처 
	If Len(Trim(txtSoldToParty)) Then
    	strVal 	  = strVal & "AND d.bp_cd =  " & FilterVar(txtSoldToParty, "''", "S") & "  "       	
    End If
    arrVal(2) = FilterVar(Trim(txtSoldToParty), " ", "S")
    
    '---수주형태 
	If Len(Trim(txtSoType)) Then
    	strVal 	  = strVal & "AND g.so_type =  " & FilterVar(txtSoType, "''", "S") & "  "        
    End If
    arrVal(3) = FilterVar(Trim(txtSoType), " ", "S")
    
    If Len(txtTrackingNo) Then
		strVal = strVal & " AND B.TRACKING_NO = " & FilterVar(Trim(txtTrackingNo), "''" , "S") & ""
	End If

	UNIValue(0,1)  = UCase(Trim(strVal))	
	UNIValue(1,0)  = UCase(arrVal(0))	
	UNIValue(2,0)  = UCase(arrVal(1))	
	UNIValue(3,0)  = UCase(arrVal(2))	
	UNIValue(4,0)  = UCase(arrVal(3))	
	
    UNIValue(0,UBound(UNIValue,2)) = UCase(Trim(lgTailList))			  '	UNISqlId(0)의 마지막 ?에 입력됨	
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
 
End Sub
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
        Call ServerMesgBox(lgstrRetMsg, vbInformation, I_MKSCRIPT)
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
				Call parent.SetFocusToDocument("M")	
				parent.frm1.txtSoDtFrom.Focus
            </Script>
        <%	   
'       Exit Sub
    Else    
        Call  MakeSpreadSheetData()
    End If

End Sub

%>
<Script Language=vbscript>
    
    With parent
		
		 .frm1.txtSalesGrpNm.value	  = "<%=ConvSPChars(strSalesGrpNm)%>"
		 .frm1.txtItemCodeNm.value	  = "<%=ConvSPChars(strItemCodeNm)%>"
		 .frm1.txtSoldToPartyNm.value = "<%=ConvSPChars(strSoldToPartyNm)%>"
		 .frm1.txtSoTypeNm.value	  = "<%=ConvSPChars(strSoTypeNm)%>"
			
    	 
     If "<%=lgDataExist%>" = "Yes" Then
        'Set condition data to hidden area
        
		If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists
			.frm1.HSalesGrp.value	 = "<%=ConvSPChars(txtSalesGrp)%>"
			.frm1.HSoDtFrom.value	 = "<%=txtSoDtFrom%>"
			.frm1.HSoDtTo.value		 = "<%=txtSoDtTo%>"
			.frm1.HDlvyDtFrom.value	 = "<%=ConvSPChars(txtDlvyDtFrom)%>"
			.frm1.HDlvyDtTo.value	 = "<%=ConvSPChars(txtDlvyDtTo)%>"
			.frm1.HSoldToParty.value = "<%=ConvSPChars(txtSoldToParty)%>"
			.frm1.HItemCode.value	 = "<%=ConvSPChars(txtItemCode)%>"
			.frm1.HSoType.value		 = "<%=ConvSPChars(txtSoType)%>"
			.frm1.HtxtTrackingNo.value = "<%=ConvSPChars(txtTrackingNo)%>"    
		End If
		
		.ggoSpread.Source  = .frm1.vspdData
                
        .frm1.vspdData.Redraw = False
		.ggoSpread.SSShowDataByClip  "<%=lgstrData%>", "F"

		Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,Parent.GetKeyPos("A",2),Parent.GetKeyPos("A",3),"A", "Q" ,"X","X")		
		Call Parent.ReFormatSpreadCellByCellByCurrency(Parent.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,Parent.GetKeyPos("A",2),Parent.GetKeyPos("A",4),"A", "Q" ,"X","X")		
				
		.lgPageNo	  	   =  "<%=lgPageNo%>"  				  '☜: Next next data tag
       	.DbQueryOk
       	
       	.frm1.vspdData.Redraw = True
    End If
       
	End with	
</Script>	

