<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<%
'********************************************************************************************************
'*  1. Module Name          : Procurement																*
'*  2. Function Name        :																			*
'*  3. Program ID           : m3112rb5.asp															*
'*  4. Program Name         : 발주내역참조(입고등록)ADO																			*
'*  5. Program Desc         : Purchase Order Detail 참조 PopUp ASP										*
'*  7. Modified date(First) : 2000/03/22																*
'*  8. Modified date(Last)  : 2003/05/07																*
'*  9. Modifier (First)     :																			*
'* 10. Modifier (Last)      : Kim Jin Ha																*
'* 11. Comment              :																			*
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"									*
'*                            this mark(⊙) Means that "may  change"										*
'*                            this mark(☆) Means that "must change"										*
'* 13. History              : 1. 2000/04/08 : Coding Start												*
'********************************************************************************************************
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%

	On Error Resume Next
	Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3           '☜ : DBAgent Parameter 선언 
	Dim lgStrData                                               '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
	Dim iTotstrData
	Dim lgTailList
	Dim lgSelectList
	Dim lgSelectListDT
	Dim lgDataExist
	Dim lgPageNo
   
	Dim strPlantNm
	Dim strPurGrpNm
		   
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
    
    Call  FixUNISQLData()                                                '☜ : DB-Agent로 보낼 parameter 데이타 set
    call  QueryData()                                                    '☜ : DB-Agent를 통한 ADO query

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
    On Error Resume Next
    SetConditionData = false
    
    If Not(rs1.EOF Or rs1.BOF) Then
        strPurGrpNm = rs1("Pur_Grp_Nm")
   		Set rs1 = Nothing
    Else
		Set rs1 = Nothing
		If Len(Trim(Request("txtGroup"))) Then
			Call DisplayMsgBox("970000", vbInformation, "구매그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    Exit Function
		End If
	End If  
	

    If Not(rs2.EOF Or rs2.BOF) Then
        strPlantNm = rs2("plant_nm")
   		Set rs2 = Nothing
    Else
		Set rs2 = Nothing
		If Len(Trim(Request("txtPlant"))) Then
			Call DisplayMsgBox("970000", vbInformation, "공장", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    Exit Function
		End If
	End If  

	SetConditionData = true
	
End Function
'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()
    Dim strVal
    Redim UNISqlId(2)                                                     '☜: SQL ID 저장을 위한 영역확보 
    Redim UNIValue(2,2)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 


	IF  Trim(Request("txtRcptflg")) = "Y" THEN                                                                           '    parameter의 수에 따라 변경함 
	     UNISqlId(0) = "m3112ra502" 
	ELSE 
	     UNISqlId(0) = "m3112ra501" 
	END IF
	UNISqlId(1) = "S0000QA019"	'구매그룹 
	UNISqlId(2) = "M2111QA302"	'공장 

	'--- 2004-08-20 by Byun Jee Hyun for UNICODE	
    UNIValue(0,0) = Trim(lgSelectList)		                              '☜: Select 절에서 Summary    필드 
	
	If Len(Request("txtPoNo")) Then
		strVal = strVal & " AND A.PO_NO = " & FilterVar(Trim(Request("txtPoNo")), " " , "S") & " "
		
	End If

    If Len(Trim(Request("txtFrPoDt"))) Then
		strVal = strVal & " AND F.PO_DT >= " & FilterVar(UNIConvDate(Request("txtFrPoDt")), "''", "S") & ""
	End If		
	
	If Len(Trim(Request("txtToPoDt"))) Then
		strVal = strVal & " AND F.PO_DT <= " & FilterVar(UNIConvDate(Request("txtToPoDt")), "''", "S") & ""		
	End If

	If Len(Request("txtSupplier")) Then
		strVal = strVal & " AND F.BP_CD = " & FilterVar(Trim(Request("txtSupplier")), " " , "S") & " "
	End If

	If Len(Request("txtGroup")) Then
		strVal = strVal & " AND F.PUR_GRP = " & FilterVar(Trim(Request("txtGroup")), " " , "S") & " "
	End If

	If Trim(Request("txtRcptflg")) = "Y" Then
		strVal = strVal & " AND F.RCPT_TYPE = " & FilterVar(Trim(Request("txtRcptType")), " " , "S") & " "
	ELSE 
		strVal = strVal & " AND F.ISSUE_TYPE = " & FilterVar(Trim(Request("txtRcptType")), " " , "S") & " "
	End If	
	
	If Len(Request("txtPlant")) Then
		strVal = strVal & " AND A.PLANT_CD = " & FilterVar(Trim(Request("txtPlant")), " " , "S") & " "
	End If	
	'2003.07 TrackingNo 추가 
    If Len(Request("txtTrackingNo")) Then
		strVal = strVal & " AND A.TRACKING_NO = " & FilterVar(Trim(UCase(Request("txtTrackingNo"))), " " , "S") & "  "		
	End If	
	
    UNIValue(0,1) = strVal  
    UNIValue(1,0) = FilterVar(Trim(UCase(Request("txtGroup"))), "" , "S") 						'구매그룹   
    UNIValue(2,0) = " " & FilterVar(Trim(UCase(Request("txtPlant"))), "" , "S") & " "				'구매그룹   
   
    UNIValue(0,UBound(UNIValue,2)) = " " & UCase(Trim(lgTailList))
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
        
 	If SetConditionData = False Then Exit Sub

    If rs0.EOF And rs0.BOF Then
       Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)		'No Data Found!!
       rs0.Close
       Set rs0 = Nothing
       Exit Sub
    Else 
        Call  MakeSpreadSheetData()
    End If
    
End Sub

%>
<Script Language=vbscript>
With Parent	
	.frm1.txtGroupNm.value = "<%=ConvSPChars(strPurGrpNm)%>" 
	.frm1.txtPlantNm.value = "<%=ConvSPChars(strPlantNm)%>" 

    If "<%=lgDataExist%>" = "Yes" Then
       If "<%=lgPageNo%>" = "1" Then 
			.frm1.hdnPoNo.value		= "<%=ConvSPChars(Request("txtPoNO"))%>"   
			.frm1.hdnFrPoDt.value		= "<%=Request("txtFrPoDt")%>"
			.frm1.hdnToPoDt.value		= "<%=Request("txtToPoDt")%>"
			.frm1.hdnGroupCd.value	= "<%=ConvSPChars(Request("txtGroup"))%>"
			.frm1.hdnGroupNm.value	= "<%=ConvSPChars(strPurGrpNm)%>" 
			.frm1.hdnPlantCd.value	= "<%=ConvSPChars(Request("txtPlant"))%>"  
       End If
	   
       .ggoSpread.Source  = .frm1.vspdData
       .frm1.vspdData.Redraw = False
       .ggoSpread.SSShowData "<%=iTotstrData%>","F"          '☜ : Display data
       
		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,.GetKeyPos("A",13),.GetKeyPos("A",11),"C","I","X","X")
		Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,.GetKeyPos("A",13),.GetKeyPos("A",12),"A","I","X","X")
		.frm1.vspdData.Redraw = True
       
       .lgPageNo      =  "<%=lgPageNo%>"               '☜ : Next next data tag
       .DbQueryOk
       
    End If
End With       
</Script>	
