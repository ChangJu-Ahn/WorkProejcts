<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<%
'======================================================================================================
'*  1. Module Name          : 구매 
'*  2. Function Name        : 입고관리 
'*  3. Program ID           : m5112rb2
'*  4. Program Name         : 매입내역참조 
'*  5. Program Desc         : ADO Query
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2003-05-28
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Kim Jin Ha
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change"
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'*                            2000/12/09
'*                            2001/12/18	Date표준적용 
'=======================================================================================================
%>
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
On Error Resume Next


Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3               '☜ : DBAgent Parameter 선언 
Dim lgstrData                                                              '☜ : data for spreadsheet data
Dim iTotstrData
Dim lgTailList                                                             '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT

Dim strPurGrpNm
Dim strPlantNm
Dim lgDataExist
Dim lgPageNo

Dim iPrevEndRow
Dim iEndRow 

Dim lgCurrency  

    Call HideStatusWnd 
    Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "RB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "RB")

    lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
    lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
    lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
	lgDataExist      = "No"
	iPrevEndRow = 0
    iEndRow = 0
    
	lgCurrency   = Request("txtCurrency")	                               '☜ : Currency

    Call FixUNISQLData()
    Call QueryData()
    
'----------------------------------------------------------------------------------------------------------
' Query Data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
    Dim  ColCnt
    Dim  iLoopCount
    Dim  iRowStr
    Dim  PvArr
    
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
    Redim UNIValue(2,2)

    UNISqlId(0) = "M5112ra101"									'* : 데이터 조회를 위한 SQL문     
	UNISqlId(1) = "S0000QA019"	'구매그룹 
    UNISqlId(2) = "M2111QA302"	'공장 
	
    UNIValue(0,0) = lgSelectList                                          '☜: Select list
    
	strVal = " "
	If Len(Request("txtIvNo")) Then
		strVal = strVal & " AND A.IV_NO = " & FilterVar(UCase(Request("txtIvNo")), "''", "S") & " "
	End If
	
	If Len(Request("txtPoNo")) Then
		strVal = strVal & " AND A.PO_NO = " & FilterVar(UCase(Request("txtPoNo")), "''", "S") & " "
	End If
	
	If Len(Trim(Request("txtIvFrDt"))) Then
		strVal = strVal & " AND F.IV_DT >= " & FilterVar(UNIConvDate(Request("txtIvFrDt")), "''", "S") & ""		
	End If		
	
	If Len(Trim(Request("txtIvToDt"))) Then
		strVal = strVal & " AND F.IV_DT <= " & FilterVar(UNIConvDate(Request("txtIvToDt")), "''", "S") & ""		
	End If
	
	If Len(Trim(Request("hdnBpCd"))) Then
		strVal = strVal & " AND F.BP_CD = " & FilterVar(UCase(Request("hdnBpCd")), "''", "S") & ""		
	End If
	
	If Len(Trim(Request("txtGroup"))) Then
		strVal = strVal & " AND H.PUR_GRP =  " & FilterVar(UCase(Request("txtGroup")), "''", "S") & " "
	End If
	
	If Len(Request("txtPlant")) Then
		strVal = strVal & " AND A.PLANT_CD = " & FilterVar(Request("txtPlant"), "''", "S") & " "
	End If
	
	'---2003.07 TrackingNo 추가 
    If Len(Request("txtTrackingNo")) Then
		strVal = strVal & " AND A.TRACKING_NO = " & FilterVar(UCase(Request("txtTrackingNo")), "''", "S") & "  "		
	End If		

     If Request("gPlant") <> "" Then
        strVal = strVal & " AND A.PLANT_CD=" & FilterVar(Request("gPlant"),"''","S")
     End If
     If Request("gPurGrp") <> "" Then
        strVal = strVal & " AND F.PUR_GRP=" & FilterVar(Request("gPurGrp"),"''","S")
     End If
     If Request("gPurOrg") <> "" Then
        strVal = strVal & " AND F.PUR_ORG=" & FilterVar(Request("gPurOrg"),"''","S")
     End If
     If Request("gBizArea") <> "" Then
        strVal = strVal & " AND F.IV_BIZ_AREA=" & FilterVar(Request("gBizArea"),"''","S")
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
    Dim lgADF                                                                  '☜ : ActiveX Data Factory 지정 변수선언 
	Dim lgstrRetMsg                                                            '☜ : Record Set Return Message 변수선언 
    Dim iStr
	
    Set lgADF   = Server.CreateObject("prjPublic.cCtlTake")
    lgstrRetMsg = lgADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0,rs1,rs2,rs3) '* : Record Set 의 갯수 조정 
    
    iStr = Split(lgstrRetMsg,gColSep)


    If iStr(0) <> "0" Then
        Call ServerMesgBox(lgstrRetMsg , vbInformation, I_MKSCRIPT)
    End If    
        
	If SetConditionData = False Then Exit Sub

	If  rs0.EOF And rs0.BOF Then
	    Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
	    rs0.Close
	    Set rs0 = Nothing    
	Else    
	    Call  MakeSpreadSheetData()
	End If
End Sub

%>
<Script Language=vbscript>
	parent.frm1.txtGroupNm.value = "<%=ConvSPChars(strPurGrpNm)%>" 
	parent.frm1.txtPlantNm.value = "<%=ConvSPChars(strPlantNm)%>" 

    If "<%=lgDataExist%>" = "Yes" Then
		With parent
		    .ggoSpread.Source   = .frm1.vspdData 
		    .frm1.vspdData.Redraw = False
		    .ggoSpread.SSShowData "<%=iTotstrData%>","F"          '☜ : Display data
       
			Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,.GetKeyPos("A",13),.GetKeyPos("A",11),"C","I","X","X")
			Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iPrevEndRow+1%>",<%=iEndRow%>,.GetKeyPos("A",13),.GetKeyPos("A",12),"A","I","X","X")
			
		    .frm1.hdnIvNo.value		= "<%=ConvSPChars(Request("txtIvNo"))%>"
		    .frm1.hdnPoNo.value     = "<%=ConvSPChars(Request("txtPoNo"))%>"
		    .frm1.hdnIvFrDt.value   = "<%=Request("txtIvFrDt")%>"
		    .frm1.hdnIvToDt.value   = "<%=Request("txtIvToDt")%>"
		    .frm1.hdnGroupNm.value	= "<%=ConvSPChars(strPurGrpNm)%>" 
			.frm1.hdnGroupCd.value	= "<%=ConvSPChars(Request("txtGroup"))%>" 
			.frm1.hdnPlantCd.value	= "<%=ConvSPChars(Request("txtPlant"))%>"  
			
		    .DbQueryOk
		    .frm1.vspdData.Redraw = True
		End with
	End if	
</Script>	
<%
    Response.End													'☜: 비지니스 로직 처리를 종료함 
%>
