<% Option explicit%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../inc/incSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%

On Error Resume Next

	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2,rs3     '☜ : DBAgent Parameter 선언 
	Dim lgStrData												   '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
	Dim lgTailList
	Dim lgSelectList
	Dim lgSelectListDT
	Dim lgDataExist
	Dim lgPageNo
	Dim iTotstrData
		   
	Dim strPlantNm
	Dim strItemNm
	Dim strVatTypeNm   
   
	Dim iPrevEndRow
	Dim iEndRow
	
    Call HideStatusWnd 
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "RB")
     
    lgPageNo         = UNICInt(Trim(Request("lgPageNo")),0)              '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
    lgSelectList     = Request("lgSelectList")
    lgTailList       = Request("lgTailList")
    lgSelectListDT   = Split(Request("lgSelectListDT"), gColSep)         '☜ : 각 필드의 데이타 타입 
    lgDataExist      = "No"
    iPrevEndRow = 0
    iEndRow = 0
    
	SELECT CASE REQUEST("txtMode")								 '☜ : onChange 에서 호출할경우와 메인쿼리인경우 
	CASE "changeItemPlant"
		Call FixUNISQLData2()									 '☜ : DB-Agent로 보낼 parameter 데이타 set
		Call QueryData2()										 '☜ : DB-Agent를 통한 ADO query
	CASE ELSE
		Call FixUNISQLData()									 '☜ : DB-Agent로 보낼 parameter 데이타 set
		Call QueryData()										 '☜ : DB-Agent를 통한 ADO query
	END SELECT    

 
'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim  RecordCnt
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    Dim PvArr
    
    Const C_SHEETMAXROWS_D = 100    
    
    lgDataExist    = "Yes"
    lgstrData      = ""
    iPrevEndRow = 0
    If CInt(lgPageNo) > 0 Then
       iPrevEndRow = C_SHEETMAXROWS_D * CInt(lgPageNo)
       rs0.Move  = iPrevEndRow                 
    End If

    iLoopCount = -1
    ReDim PvArr(C_SHEETMAXROWS_D - 1)
    
    Do while Not (rs0.EOF Or rs0.BOF)
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
		For ColCnt = 0 To UBound(lgSelectListDT) - 1
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If  iLoopCount < C_SHEETMAXROWS_D Then
            lgstrData      = lgstrData      & iRowStr & Chr(11) & Chr(12)
            PvArr(iLoopCount) = lgstrData	
			lgstrData = ""
        Else
            lgPageNo = lgPageNo + 1
            Exit Do
        End If
        rs0.MoveNext
	Loop
	
	iTotstrData = Join(PvArr, "")
	
    If  iLoopCount < C_SHEETMAXROWS_D Then                                            '☜: Check if next data exists
        iEndRow = iPrevEndRow + iLoopCount + 1
        lgPageNo = ""                                                  '☜: 다음 데이타 없다.
    Else
        iEndRow = iPrevEndRow + iLoopCount
    End If
  	
	rs0.Close
    Set rs0 = Nothing 
End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Function SetConditionData()
    'On Error Resume Next
    
    SetConditionData = TRUE
    
    If Not(rs1.EOF Or rs1.BOF) Then
       strPlantNm =  rs1(1)
       Set rs1 = Nothing 
 	else
	    Set rs1 = Nothing
		If Len(Request("txtPlantCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "공장", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
			SetConditionData = FALSE
			EXIT FUNCTION
		End If   
    End If   

    If Not(rs2.EOF Or rs2.BOF) Then
       strItemNm =  rs2(1)
       Set rs2 = Nothing 
 	else
	    Set rs2 = Nothing
		If Len(Request("txtItemCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "품목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
			SetConditionData = FALSE
			EXIT FUNCTION
		End If   
    End If   

    If Not(rs3.EOF Or rs3.BOF) Then
       strVatTypeNm =  rs3(1)
       Set rs3 = Nothing 
	else
	    Set rs3 = Nothing
		If Len(Request("txtVatType")) Then
			Call DisplayMsgBox("970000", vbInformation, "VAT유형", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
			SetConditionData = FALSE
			EXIT FUNCTION
		End If    
    End If       

    

End Function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
	Dim arrVal(2)
    Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(3,2)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                          '    parameter의 수에 따라 변경함 
     UNISqlId(0) = "M3112QA002" 
     UNISqlId(1) = "s0000qa009"											'공장 
     UNISqlId(2) = "s0000qa016"											'품목 
     UNISqlId(3) = "s0000qa026"											'VAT명 

	'--- 2004-08-20 by Byun Jee Hyun for UNICODE
     UNIValue(0,0) = Trim(lgSelectList)		                              '☜: Select 절에서 Summary    필드 

	strVal = " "

	If Len(Request("txtPoNo")) Then
		strVal = " AND A.PO_NO = " & FilterVar(UCase(Request("txtPoNo")), "''", "S") & " "
	End If

	If Len(Request("txtItemCd")) Then
		strVal = strVal & " AND B.ITEM_CD = " & FilterVar(UCase(Request("txtItemCd")), "''", "S") & " "
	End If
	arrVal(1)=FilterVar(Trim(UCase(Request("txtItemCd"))), " " , "S")

	If Len(Request("txtPlantCd")) Then
		strVal = strVal & " AND B.PLANT_CD = " & FilterVar(UCase(Request("txtPlantCd")), "''", "S") & " "
	End If
	arrVal(0)=FilterVar(Trim(UCase(Request("txtPlantCd"))), " " , "S")

	If Len(Request("txtSupplier")) Then
		strVal = strVal & " AND A.BP_CD = " & FilterVar(UCase(Request("txtSupplier")), "''", "S") & " "
	End If

    If Len(Trim(Request("txtGroup"))) Then
		strVal = strVal & " AND A.PUR_GRP = " & FilterVar(UCase(Request("txtGroup")), "''", "S") & " "		
	End If		
	
	If Len(Trim(Request("txtIvType"))) Then
		strVal = strVal & " AND A.IV_TYPE = " & FilterVar(UCase(Request("txtIvType")), "''", "S") & " "		
	End If

	If Len(Request("txtPoCur")) Then
		strVal = strVal & " AND A.PO_CUR = " & FilterVar(UCase(Request("txtPoCur")), "''", "S") & " "
	End If

    If Len(Trim(Request("txtFrPoDt"))) Then
		strVal = strVal & " AND A.PO_DT >= " & FilterVar(UNIConvDate(Request("txtFrPoDt")), "''", "S") & ""		
	End If		
	
	If Len(Trim(Request("txtToPoDt"))) Then
		strVal = strVal & " AND A.PO_DT <= " & FilterVar(UNIConvDate(Request("txtToPoDt")), "''", "S") & ""		
	End If
	
	If Len(Trim(Request("txtVatType"))) Then
		strVal = strVal & " AND B.VAT_TYPE = " & FilterVar(UCase(Request("txtVatType")), "''", "S") & " "	
	End If	
	arrVal(2)=FilterVar(Trim(UCase(Request("txtVatType"))), " " , "S")
	
	'---2003.07 TrackingNo 추가 
    If Len(Request("txtTrackingNo")) Then
		strVal = strVal & " AND B.TRACKING_NO = " & FilterVar(UCase(Request("txtTrackingNo")), "''", "S") & "  "		
	End If	

     If Request("gPlant") <> "" Then
        strVal = strVal & " AND b.PLANT_CD=" & FilterVar(Request("gPlant"),"''","S")
     End If
     If Request("gPurGrp") <> "" Then
        strVal = strVal & " AND a.PUR_GRP=" & FilterVar(Request("gPurGrp"),"''","S")
     End If
     If Request("gPurOrg") <> "" Then
        strVal = strVal & " AND a.PUR_ORG=" & FilterVar(Request("gPurOrg"),"''","S")
     End If
     If Request("gBizArea") <> "" Then
        strVal = strVal & " AND a.PUR_BIZ_AREA=" & FilterVar(Request("gBizArea"),"''","S")
     End If   
        
    UNIValue(0,1) = strVal
    UNIValue(1,0) = arrVal(0)  	'공장 
    UNIValue(2,0) = arrVal(1)  	'품목 
    UNIValue(3,0) = arrVal(2)		                              '☜: Select 절에서 Summary    필드    
       
    
    UNIValue(0,UBound(UNIValue,2)) = " GROUP BY A.PO_NO,B.PO_SEQ_NO,B.PLANT_CD,C.PLANT_NM,B.ITEM_CD,D.ITEM_NM,D.SPEC,B.PO_QTY,(B.PO_QTY-B.IV_QTY-B.INSPECT_QTY-B.LC_QTY+ISNULL(K.IV_QTY,0)),B.PO_UNIT,B.PO_PRC,B.PO_DOC_AMT,A.PO_CUR,B.VAT_TYPE,R.MINOR_NM,B.VAT_RATE,A.PO_DT,B.DLVY_DT,B.SL_CD,E.SL_NM,B.RCPT_QTY,B.LC_QTY,B.TRACKING_NO,B.VAT_INC_FLAG,B.AMT_UPT_FLG,B.VAT_DOC_AMT,B.VAT_AMT_RVS_FLG,B.IV_QTY,A.RET_FLG ORDER BY A.PO_NO DESC,B.PO_SEQ_NO ASC"
    UNILock = DISCONNREAD :	UNIFlag = "1"                                 '☜: set ADO read mode
	
End Sub


Sub FixUNISQLData2()

    Dim strVal
	Dim arrVal(2)
    Redim UNISqlId(2)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(2,2)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                          '    parameter의 수에 따라 변경함 
     UNISqlId(0) = "s0000qa009"											'공장 
     UNISqlId(1) = "s0000qa016"											'품목 
     UNISqlId(2) = "s0000qa026"											'VAT명 

	'--- 2004-08-20 by Byun Jee Hyun for UNICODE
	'If Len(Request("txtItemCd")) Then
		arrVal(1)=FilterVar(Trim(UCase(Request("txtItemCd"))), " " , "S") 
	'End If

	'If Len(Request("txtPlantCd")) Then
		arrVal(0)=FilterVar(Trim(UCase(Request("txtPlantCd"))), " " , "S") 
	'End If

	'If Len(Trim(Request("txtVatType"))) Then
		arrVal(2)=FilterVar(Trim(UCase(Request("txtVatType"))), " " , "S") 
	'End If	
        
    UNIValue(0,0) = arrVal(0)  	'공장 
    UNIValue(1,0) = arrVal(1)  	'품목 
    UNIValue(2,0) = arrVal(2)	'VAT
       
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
    
    if SetConditionData = False Then Exit Sub

    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        

 
    If rs0.EOF And rs0.BOF Then
       Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
       rs0.Close
       Set rs0 = Nothing
       Exit Sub
    Else    
        Call  MakeSpreadSheetData()
        
    End If
    
End Sub


'----------------------------------------------------------------------------------------------------------
' Query Data2
'----------------------------------------------------------------------------------------------------------
Sub QueryData2()

    Dim lgstrRetMsg                                             '☜ : Record Set Return Message 변수선언 
    Dim lgADF                                                   '☜ : ActiveX Data Factory 지정 변수선언 
    Dim iStr
    
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag,rs1, rs2, rs3)
    
    Set lgADF   = Nothing
    
    Call  SetConditionData()
    
End Sub

%>
<Script Language=vbscript>
	parent.frm1.txtPlantNm.Value 		= "<%=ConvSPChars(strPlantNm)%>"
	parent.frm1.txtItemNm.Value 		= "<%=ConvSPChars(strItemNm)%>"
	parent.frm1.txtVatNm.Value 			= "<%=ConvSPChars(strVatTypeNm)%>"	

    If "<%=lgDataExist%>" = "Yes" Then
       If "<%=lgPageNo%>" = "1" Then   ' "1" means that this query is first and next data exists       
			parent.frm1.hdnPoNo.value	= "<%=ConvSPChars(Request("txtPoNo"))%>"
			parent.frm1.hdnFrPoDt.value	= "<%=ConvSPChars(Request("txtFrPoDt"))%>"
			parent.frm1.hdnToPoDt.value	= "<%=ConvSPChars(Request("txtToPoDt"))%>"
			parent.frm1.hdnItemCd.value	= "<%=ConvSPChars(Request("txtItemCd"))%>"
			parent.frm1.hdnPlantCd.value = "<%=ConvSPChars(Request("txtPlantCd"))%>"			
			parent.frm1.txtVatType.value = "<%=ConvSPChars(Request("txtVatType"))%>"				
       End If
       
       parent.ggoSpread.Source  = parent.frm1.vspdData
       Parent.frm1.vspdData.Redraw = False
       parent.ggoSpread.SSShowData "<%=iTotstrData%>", "F"          '☜ : Display data
       
       Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnPoCur.value, Parent.GetKeyPos("A",11),"C", "I" ,"X","X")
       Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnPoCur.value, Parent.GetKeyPos("A",12),"A", "I" ,"X","X")
       Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnPoCur.value, Parent.GetKeyPos("A",16),"D", "I" ,"X","X")
       Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnPoCur.value, Parent.GetKeyPos("A",25),"A", "I" ,"X","X")
       Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnPoCur.value, Parent.GetKeyPos("A",27),"A", "I" ,"X","X")
       Call Parent.ReFormatSpreadCellByCellByCurrency2(Parent.Frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,parent.frm1.hdnPoCur.value, Parent.GetKeyPos("A",28),"A", "I" ,"X","X")
       
       
       parent.lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
 
       parent.DbQueryOk
       Parent.frm1.vspdData.Redraw = True
    End If   
</Script>	
