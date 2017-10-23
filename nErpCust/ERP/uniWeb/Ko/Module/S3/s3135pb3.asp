<%
'**********************************************************************************************
'*  1. Module Name          : 영업 
'*  2. Function Name        : 
'*  3. Program ID           : s3135pa1
'*  4. Program Name         : Tracking No(수주진행별조회)
'*  5. Program Desc         : Tracking No(수주진행별조회)
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 2002/02/16
'*  8. Modified date(Last)  : 2002/04/12
'*  9. Modifier (First)     : Choinkuk		
'* 10. Modifier (Last)      : Choinkuk
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
On Error Resume Next
                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2, rs3, rs4			   
Dim lgStrData                                                 
Dim lgTailList                                                
Dim lgSelectList
Dim lgSelectListDT

Dim lgPageNo

Dim strPtnBpNm												  ' 주문처명 
Dim strSalesGrpNm											  ' 영억그룹명 
Dim strPlantNm											      ' 공장명 
Dim strItemNm											      ' 품목명 
Dim MsgDisplayFlag
Dim iStatusFlag

Const C_SHEETMAXROWS_D  = 30                                          

    Call HideStatusWnd 

	lgPageNo       = UNICInt(Trim(Request("txtHlgPageNo")),0)    	
	lgSelectList   = Request("txtHlgSelectList")                                
	lgTailList     = Request("txtHlgTailList")                                 
	lgSelectListDT = Split(Request("txtHlgSelectListDT"), gColSep)             
	
    Call FixUNISQLData()									 
    Call QueryData()										 
    
'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
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

    If  iLoopCount < C_SHEETMAXROWS_D Then                             
        lgPageNo = ""                                                 
    End If
  	
	rs0.Close
    Set rs0 = Nothing 

End Sub

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Function SetConditionData()
    On Error Resume Next
    
	SetConditionData = False

    If Not(rs1.EOF Or rs1.BOF) Then
        strPtnBpNm =  rs1(1)
        Set rs1 = Nothing
    Else
		Set rs1 = Nothing

		If Len(Request("txtHPtnBpCd")) And MsgDisplayFlag = False Then
			Call DisplayMsgBox("970000", vbInformation, "주문처", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    MsgDisplayFlag = True	
		%>
		<Script language=vbs> Parent.frm1.txtPtnBpCd.focus </Script>
		<%	 
		End If
	End If   	
    
     
	If Not(rs2.EOF Or rs2.BOF) Then
        strSalesGrpNm =  rs2(1)
        Set rs2 = Nothing
    Else
		Set rs2 = Nothing
		If Len(Request("txtHSalesGrp")) And MsgDisplayFlag = False Then
			Call DisplayMsgBox("970000", vbInformation, "영업그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    MsgDisplayFlag = True	
		%>
		<Script language=vbs> Parent.frm1.txtSalesGrp.focus </Script>
		<%	 
		End If			
    End If   	
    
    If Not(rs3.EOF Or rs3.BOF) Then
        strPlantNm =  rs3(1)
        Set rs3 = Nothing
    Else
		Set rs3 = Nothing
		If Len(Request("txtHPlant")) And MsgDisplayFlag = False Then
			Call DisplayMsgBox("970000", vbInformation, "공장", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    MsgDisplayFlag = True	
		%>
		<Script language=vbs> Parent.frm1.txtPlant.focus </Script>
		<%	 
		End If				
    End If      

    If Not(rs4.EOF Or rs4.BOF) Then
        strItemNm =  rs4(1)
        Set rs4 = Nothing
    Else
		Set rs4 = Nothing
		If Len(Request("txtHItem")) And MsgDisplayFlag = False Then
			Call DisplayMsgBox("970000", vbInformation, "품목", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    MsgDisplayFlag = True	
		%>
		<Script language=vbs> Parent.frm1.txtItem.focus </Script>
		<%	 
		End If				
    End If      

	If MsgDisplayFlag = True Then Exit Function

	SetConditionData = True

End Function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

	On Error Resume Next

    Dim strVal
    Dim arrVal(3)
    Redim UNISqlId(4)                                                     '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(4,2)

    UNISqlId(0) = "S3135PA301"
    UNISqlId(1) = "s0000qa002"					'주문처명 
    UNISqlId(2) = "s0000qa005"					'영업그룹명 
    UNISqlId(3) = "s0000qa009"					'공장명  
    UNISqlId(4) = "s0000qa001"					'품목명  
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------

    UNIValue(0,0) = Trim(lgSelectList)                                      '☜: Select list
    
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------

	iStatusFlag = Request("txtHStatusFlag")

	strVal = " "
	
	If Len(Request("txtHSoNo")) Then
		strVal = "AND SH.SO_NO = " & FilterVar(Request("txtHSoNo"), "''", "S") & " "	
	Else
		strVal = ""
	End If

	If Len(Request("txtHPtnBpCd")) Then
		strVal = strVal & " AND SH.SOLD_TO_PARTY = " & FilterVar(Request("txtHPtnBpCd"), "''", "S") & " "			
	End If		
	arrVal(0) = FilterVar(Trim(Request("txtHPtnBpCd")), "", "S")
		   
	If Len(Request("txtHSalesGrp")) Then
		strVal = strVal & " AND SH.SALES_GRP = " & FilterVar(Request("txtHSalesGrp"), "''", "S") & " "				
	End If		
	arrVal(1) = FilterVar(Trim(Request("txtHSalesGrp")), "", "S")
    
 	If Len(Request("txtHPlant")) Then
		strVal = strVal & " AND SD.PLANT_CD = " & FilterVar(Request("txtHPlant"), "''", "S") & " "				
	End If	    
	arrVal(2) = FilterVar(Trim(Request("txtHPlant")), "", "S") 
	
 	If Len(Request("txtHItem")) Then
		strVal = strVal & " AND BI.ITEM_CD = " & FilterVar(Request("txtHItem"), "''", "S") & " "				
	End If	    
	arrVal(3) = FilterVar(Trim(Request("txtHItem")), "", "S") 

    If Len(Request("txtHFromDt")) Then
		strVal = strVal & " AND SH.SO_DT >= " & FilterVar(UNIConvDate(Request("txtHFromDt")), "''", "S") & ""			
	End If		
	
	If Len(Request("txtHToDt")) Then
		strVal = strVal & " AND SH.SO_DT <= " & FilterVar(UNIConvDate(Request("txtHToDt")), "''", "S") & ""		
	End If

	Select Case UCase(iStatusFlag)
		Case "SO" 
			strVal = strVal & " AND SH.REL_DN_FLAG = " & FilterVar("Y", "''", "S") & "  AND SD.GI_QTY > 0 "
		Case "ML", "MS"
			strVal = strVal & " AND SH.LC_FLAG = " & FilterVar("M", "''", "S") & "  AND SD.CFM_QTY > SD.LC_QTY "
		Case "MA"
			strVal = strVal & " AND SH.LC_FLAG = " & FilterVar("M", "''", "S") & "  AND SD.LC_QTY > 0 "
		Case "LL"
			strVal = strVal & " AND SH.LC_FLAG = " & FilterVar("L", "''", "S") & "  AND SD.CFM_QTY > SD.LC_QTY "
		Case "LD"	
			strVal = strVal & " AND SD.REQ_QTY > 0 "
		Case "LA"	
			strVal = strVal & " AND SH.LC_FLAG = " & FilterVar("L", "''", "S") & "  AND SD.LC_QTY > 0 "
		Case "DN"
			strVal = strVal & " AND SD.CFM_QTY > SD.REQ_QTY "
		Case "CS"	
			strVal = strVal & " AND EXISTS(SELECT * FROM S_SO_SCHD SC WHERE SC.SO_NO = SD.SO_NO AND SC.SO_SEQ = SD.SO_SEQ AND SC.CFM_QTY > SC.CC_QTY) "
		Case "CL"	
			strVal = strVal & " AND EXISTS(SELECT * FROM S_SO_SCHD SC WHERE SC.SO_NO = SD.SO_NO AND SC.SO_SEQ = SD.SO_SEQ AND SD.LC_QTY > SC.CC_QTY) "	
		Case "BS"
			strVal = strVal & " AND SD.CFM_QTY > SD.BILL_QTY AND SH.REL_DN_FLAG = " & FilterVar("N", "''", "S") & "  AND SH.LC_FLAG = " & FilterVar("N", "''", "S") & "  "
		Case "BL"	
			strVal = strVal & " AND SD.CFM_QTY > SD.BILL_QTY AND SH.REL_DN_FLAG = " & FilterVar("N", "''", "S") & "  AND SH.LC_FLAG = " & FilterVar("L", "''", "S") & "  "	
		Case "BD"				
			strVal = strVal & " AND SD.GI_QTY > SD.BILL_QTY "
		Case "ES"
			strVal = strVal & " AND SD.CFM_QTY > SD.BILL_QTY AND SH.REL_DN_FLAG = " & FilterVar("N", "''", "S") & "  AND SH.EXPORT_FLAG = " & FilterVar("Y", "''", "S") & "  "
		Case "EL"
			strVal = strVal & " AND SD.LC_QTY > SD.BILL_QTY AND SH.REL_DN_FLAG = " & FilterVar("N", "''", "S") & "  AND SH.LC_FLAG = " & FilterVar("M", "''", "S") & "  "
		Case "EC"
			strVal = strVal & " AND EXISTS(SELECT * FROM S_SO_SCHD SC WHERE SC.SO_NO = SD.SO_NO AND SC.SO_SEQ = SD.SO_SEQ AND SC.CC_QTY > SC.BILL_QTY AND GI_QTY > 0) "
	End Select

    UNIValue(0,1) = strVal   
    UNIValue(1,0) = arrVal(0)	
    UNIValue(2,0) = arrVal(1)
    UNIValue(3,0) = arrVal(2)	
    UNIValue(4,0) = arrVal(3)	
    
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
    
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4)

	Set lgADF   = Nothing
	
    iStr = Split(lgstrRetMsg,gColSep)

	If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If 
         
    If SetConditionData = False Then Exit Sub

    If  rs0.EOF And rs0.BOF And FalsechkFlg =  False Then
        Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
        rs0.Close
        Set rs0 = Nothing
        MsgDisplayFlag = True
    %>
		<Script language=vbs> Parent.frm1.txtPtnBpCd.focus </Script>
	<%	 
	Else    
        Call  MakeSpreadSheetData()
    End If  
End Sub

%>

<Script Language=vbscript>
    With parent
		.frm1.txtPtnBpNm.value			= "<%=ConvSPChars(strPtnBpNm)%>" 
		.frm1.txtSalesGrpNm.value		= "<%=ConvSPChars(strSalesGrpNm)%>" 
		.frm1.txtPlantNm.value			= "<%=ConvSPChars(strPlantNm)%>" 
        .frm1.txtItemNm.value			= "<%=ConvSPChars(strItemNm)%>"	

		.ggoSpread.Source    = .frm1.vspdData 
		.ggoSpread.SSShowDataByClip "<%=lgstrData%>"                            '☜: Display data 
		.lgPageNo			 =  "<%=lgPageNo%>"							  '☜: Next next data tag
		.DbQueryOk

	End with
</Script>	
