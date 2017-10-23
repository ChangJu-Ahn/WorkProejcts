<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit    %>
<!--
'**********************************************************************************************
'*  1. Module Name          : Prucurement
'*  2. Function Name        : 
'*  3. Program ID           : M3111PB1
'*  4. Program Name         : 발주번호 
'*  5. Program Desc         : 발주번호 
'*  6. Component List       : 
'*  7. Modified date(First) : 2000/01/09
'*  8. Modified date(Last)  : 2003/05/22
'*  9. Modifier (First)     : Min Hak-jun
'* 10. Modifier (Last)      : Kang Su Hwan
'* 11. Comment              :
'* 12. Common Coding Guide  : this mark(☜) means that "Do not change" 
'*                            this mark(⊙) Means that "may  change"
'*                            this mark(☆) Means that "must change"
'* 13. History              :
'**********************************************************************************************
-->
<!-- #Include file="../../inc/incSvrMain.asp" -->
<!-- #Include file="../../inc/incSvrDate.inc" -->
<!-- #Include file="../../inc/incSvrNumber.inc" -->
<!-- #Include file="../../inc/incSvrDBAgent.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->

<%
On Error Resume Next
                                                                         
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0 , rs1, rs2, rs3      '☜ : DBAgent Parameter 선언 
Dim lgStrData                                                 '☜ : Spread sheet에 보여줄 데이타를 위한 변수 
Dim lgTailList                                                '☜ : Orderby절에 사용될 field 리스트 
Dim lgSelectList
Dim lgSelectListDT
Dim lgDataExist
Dim lgPageNo
Dim SortNo													  ' Sort 종류 

Dim PotypeNm														'☜ : 발주형태명 저장 
Dim GroupNm										   				    '☜ : 구매그룹명 저장 
Dim SupplierNm														'☜ : 공급처명 저장 
Dim iFrPoint
   iFrPoint=0
   
    Call HideStatusWnd 
	Call LoadBasisGlobalInf()
	Call LoadInfTB19029B("I", "*", "NOCOOKIE", "PB")
	Call LoadBNumericFormatB("I", "*", "NOCOOKIE", "PB")

	lgPageNo       = UNICInt(Trim(Request("lgPageNo")),0)    '☜: "0"(First),"1"(Second),"2"(Third),"3"(...)
	lgSelectList   = Request("lgSelectList")                               '☜ : select 대상목록 
	lgTailList     = Request("lgTailList")                                 '☜ : Orderby value
	lgSelectListDT = Split(Request("lgSelectListDT"), gColSep)             '☜ : 각 필드의 데이타 타입 
	lgDataExist    = "No"
	    
    Call FixUNISQLData()									 '☜ : DB-Agent로 보낼 parameter 데이타 set
    Call QueryData()										 '☜ : DB-Agent를 통한 ADO query
    
'----------------------------------------------------------------------------------------------------------
' Make srpread sheet data
'----------------------------------------------------------------------------------------------------------
Sub MakeSpreadSheetData()
	Const C_SHEETMAXROWS_D  = 100
    Dim iLoopCount                                                                     
    Dim iRowStr
    Dim ColCnt
	Dim PvArr
    
    lgDataExist    = "Yes"
    lgstrData      = ""
  
    If CLng(lgPageNo) > 0 Then
       rs0.Move     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)                  'C_SHEETMAXROWS_D:Max Fetched Count at once , lgStrPrevKeyIndex : Previous PageNo
	   iFrPoint     = CLng(C_SHEETMAXROWS_D) * CLng(lgPageNo)	
    End If
    
    iLoopCount = -1
    
   ReDim PvArr(C_SHEETMAXROWS_D - 1)
   Do while Not (rs0.EOF Or rs0.BOF)
   
        iLoopCount =  iLoopCount + 1
        iRowStr = ""
        
		For ColCnt = 0 To UBound(lgSelectListDT) - 1 
            iRowStr = iRowStr & Chr(11) & FormatRsString(lgSelectListDT(ColCnt),rs0(ColCnt))
		Next
 
        If iLoopCount < C_SHEETMAXROWS_D Then
           lgstrData = lgstrData & iRowStr & Chr(11) & Chr(12)
        Else
           lgPageNo = lgPageNo + 1
           Exit Do
        End If
        PvArr(iLoopCount) = lgstrData
        lgstrData=""
        
        rs0.MoveNext
	Loop
    lgstrData = Join(PvArr,"")

    If iLoopCount < C_SHEETMAXROWS_D Then                                 '☜: Check if next data exists
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
    
	If Not(rs1.EOF Or rs1.BOF) Then
		PotypeNm = rs1("PO_TYPE_NM")
		Set rs1 = Nothing
	Else
		Set rs1 = Nothing
		If Len(Request("txtPotypeCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "발주형태", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    Exit Function
		End If
	End If   	
	
	If Not(rs2.EOF Or rs2.BOF) Then
		SupplierNm = rs2("BP_NM")
		Set rs2 = Nothing
	Else
		Set rs2 = Nothing
		If Len(Request("txtSupplierCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "공급처", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    Exit Function
		End If
	End If   	
	
	If Not(rs3.EOF Or rs3.BOF) Then
		GroupNm = rs3("PUR_GRP_NM")
		Set rs3 = Nothing
	Else
		Set rs3 = Nothing
		If Len(Request("txtGroupCd")) Then
			Call DisplayMsgBox("970000", vbInformation, "구매그룹", "", I_MKSCRIPT)	'⊙: you must release this line if you change msg into code	
		    Exit Function
		End If
	End If   	

    SetConditionData = True
End Function

'----------------------------------------------------------------------------------------------------------
' Set DB Agent arg
'----------------------------------------------------------------------------------------------------------
Sub FixUNISQLData()

    Dim strVal
	dim sTemp
	Redim UNISqlId(3)                                                     '☜: SQL ID 저장을 위한 영역확보 
'--------------- 개발자 coding part(실행로직,Start)----------------------------------------------------
    Redim UNIValue(3,1)                                                  '⊙: DB-Agent로 전송될 parameter를 위한 변수 
                                                                          '    parameter의 수에 따라 변경함 
	strVal = ""
    UNISqlId(0) = "M3111PA101"
    UNISqlId(1) = "s0000qa020"
    UNISqlId(2) = "s0000qa002"
    UNISqlId(3) = "s0000qa019"
    
    '--- 2004-08-19 by Byun Jee Hyun for UNICODE
    UNIValue(1,0) = FilterVar("zzzzz", " " , "S")
    UNIValue(2,0) = FilterVar("zzzzzzzzzz", " " , "S")
    UNIValue(3,0) = FilterVar("zzzz", " " , "S")
    
    sTemp = "1"
    
	If Len(Trim(Request("txtFrPoDt"))) Then
		If UNIConvDate(Request("txtFrPoDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtFrPoDt", 0, I_MKSCRIPT)
		    Exit Sub
		End If
	End If
	
	If Len(Trim(Request("txtToPoDt"))) Then
		If UNIConvDate(Request("txtToPoDt")) = "" Then
		    Call DisplayMsgBox("122116", vbInformation, "", "", I_MKSCRIPT)
		    Call LoadTab("parent.frm1.txtToPoDt", 0, I_MKSCRIPT)
		    Exit Sub
		End If
	End If

    '발주형태                    
    If Len(Trim(Request("txtPotypeCd"))) Then
		if sTemp = "1" then
			strVal = strVal & " WHERE A.PO_TYPE_CD =  " & FilterVar(Trim(UCase(Request("txtPotypeCd"))), " " , "S") & "  "	
			sTemp = "2"
		else
			strVal = strVal & " AND A.PO_TYPE_CD =  " & FilterVar(Trim(UCase(Request("txtPotypeCd"))), " " , "S") & "  "	
		end if	
		UNIValue(1,0) = FilterVar(Trim(UCase(Request("txtPotypeCd"))), " " , "S")
	End If
    '확정여부			
    If Trim(Request("txtRadio")) = "Y" then
		if sTemp = "1" then
			strVal = strVal & " WHERE A.RELEASE_FLG = " & FilterVar("Y", "''", "S") & "  "
			sTemp = "2"
		else
			strVal = strVal & " AND A.RELEASE_FLG = " & FilterVar("Y", "''", "S") & "  "
		end if	
	ElseIf Trim(Request("txtRadio")) = "N" then
		if sTemp = "1" then
			strVal = strVal & " WHERE A.RELEASE_FLG = " & FilterVar("N", "''", "S") & "  "
			sTemp = "2"
		else
			strVal = strVal & " AND A.RELEASE_FLG = " & FilterVar("N", "''", "S") & "  "
		end if		
	End If
	'공급처 
    If Len(Trim(Request("txtSupplierCd"))) Then
		if sTemp = "1" then
			strVal = strVal & " WHERE A.BP_CD =  " & FilterVar(Trim(UCase(Request("txtSupplierCd"))), " " , "S") & "  "	
			sTemp = "2"
		else
			strVal = strVal & " AND A.BP_CD =  " & FilterVar(Trim(UCase(Request("txtSupplierCd"))), " " , "S") & "  "	
		end if		    
		UNIValue(2,0) = FilterVar(Trim(UCase(Request("txtSupplierCd"))), " " , "S")
	End If
    '발주일 
    If Len(Trim(Request("txtFrPoDt"))) Then
		if sTemp = "1" then
			strVal = strVal & " WHERE A.PO_DT >=  " & FilterVar(UniConvDate(Request("txtFrPoDt")), "''", "S") & " "	
			sTemp = "2"
		else
			strVal = strVal & " AND A.PO_DT >=  " & FilterVar(UniConvDate(Request("txtFrPoDt")), "''", "S") & " "	
		end if		      
	End If
			
    If Len(Trim(Request("txtToPoDt"))) Then
		if sTemp = "1" then
			strVal = strVal & " WHERE A.PO_DT <=  " & FilterVar(UniConvDate(Request("txtToPoDt")), "''", "S") & " "	
			sTemp = "2"
		else
			strVal = strVal & " AND A.PO_DT <=  " & FilterVar(UniConvDate(Request("txtToPoDt")), "''", "S") & " "	
		end if		      
	End If
	'구매구룹 
	If Len(Trim(Request("txtGroupCd"))) Then
		if sTemp = "1" then
			strVal = strVal & " WHERE A.PUR_GRP =  " & FilterVar(Trim(UCase(Request("txtGroupCd"))), " " , "S") & "  "	
			sTemp = "2"
		else
			strVal = strVal & " AND A.PUR_GRP =  " & FilterVar(Trim(UCase(Request("txtGroupCd"))), " " , "S") & "  "	
		end if		  	
		UNIValue(3,0) = FilterVar(Trim(UCase(Request("txtGroupCd"))), " " , "S")
	End If
    '반품 
	If Trim(Request("hdnRetFlg")) = "Y" then
		if sTemp = "1" then
			strVal = strVal & " WHERE A.RET_FLG = " & FilterVar("Y", "''", "S") & "  "
			sTemp = "2"
		else
			strVal = strVal & " AND A.RET_FLG = " & FilterVar("Y", "''", "S") & "  "
		end if		  	
	ElseIf Trim(Request("hdnRetFlg")) = "N" then
		if sTemp = "1" then
			strVal = strVal & " WHERE A.RET_FLG = " & FilterVar("N", "''", "S") & "  "
			sTemp = "2"
		else
			strVal = strVal & " AND A.RET_FLG = " & FilterVar("N", "''", "S") & "  "
		end if		  	
	End If
	'2002/12/14 
	'STO flag
    If Trim(Request("hdnSTOFlg")) = "Y" then
		if sTemp = "1" then
			strVal = strVal & " WHERE A.STO_FLG <> " & FilterVar("N", "''", "S") & "  "
			sTemp = "2"
		else
			strVal = strVal & " AND A.STO_FLG <> " & FilterVar("N", "''", "S") & "  "
		end if	
	ElseIf Trim(Request("hdnSTOFlg")) = "N" then
		if sTemp = "1" then
			strVal = strVal & " WHERE A.STO_FLG <> " & FilterVar("Y", "''", "S") & "  "
			sTemp = "2"
		else
			strVal = strVal & " AND A.STO_FLG <> " & FilterVar("Y", "''", "S") & "  "
		end if		
	End If

     If Request("gPurGrp") <> "" Then
        if sTemp = "1" then
            strVal = strVal & " WHERE 1=1" 
            sTemp = "2"
        End If
        strVal = strVal & " AND a.PUR_GRP=" & FilterVar(Request("gPurGrp"),"''","S")
     End If
     If Request("gPurOrg") <> "" Then
        if sTemp = "1" then
            strVal = strVal & " WHERE 1=1" 
            sTemp = "2"
        End If
        strVal = strVal & " AND a.PUR_ORG=" & FilterVar(Request("gPurOrg"),"''","S")
     End If
     If Request("gBizArea") <> "" Then
        if sTemp = "1" then
            strVal = strVal & " WHERE 1=1" 
            sTemp = "2"
        End If
        strVal = strVal & " AND a.PUR_BIZ_AREA=" & FilterVar(Request("gBizArea"),"''","S")
     End If   
	
'--------------- 개발자 coding part(실행로직,End)------------------------------------------------------
	UNIValue(0,0) = Trim(lgSelectList)		                              '☜: Select 절에서 Summary    필드 
	UNIValue(0,1) = strVal & UCase(Trim(lgTailList)) 

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
    With parent
		.frm1.txtPotypeNm.value = "<%=ConvSPChars(PotypeNm)%>"
		.frm1.txtSupplierNm.value = "<%=ConvSPChars(SupplierNm)%>"
		.frm1.txtGroupNm.value = "<%=ConvSPChars(GroupNm)%>"
		
		If "<%=lgDataExist%>" = "Yes" Then
			'Set condition data to hidden area
			If "<%=lgPageNo%>" = "1" Then           ' "1" means that this query is first and next data exists
				.frm1.hdnPotype.value	= "<%=ConvSPChars(Request("txtPotypeCd"))%>"
				.frm1.hdnSupplier.value	= "<%=ConvSPChars(Request("txtSupplierCd"))%>"
				.frm1.hdnFrDt.value 	= "<%=ConvSPChars(Request("txtFrPoDt"))%>"
				.frm1.hdnToDt.value 	= "<%=ConvSPChars(Request("txtToPoDt"))%>"
				.frm1.hdnGroup.value	= "<%=ConvSPChars(Request("txtGroupCd"))%>"
				.frm1.hdtxtRadio.value	= "<%=ConvSPChars(Request("txtRadio"))%>"
				.frm1.hdnRetFlg.value	= "<%=ConvSPChars(Request("hdnRetFlg"))%>"					
			End If    
			'Show multi spreadsheet data from this line
			.ggoSpread.Source    = .frm1.vspdData 
			Parent.frm1.vspdData.Redraw = False
			.ggoSpread.SSShowData "<%=lgstrData%>", "F"                  '☜: Display data 
			
			Call .ReFormatSpreadCellByCellByCurrency(.frm1.vspdData,"<%=iFrPoint+1%>",.frm1.vspddata.maxrows,.GetKeyPos("A",8),.GetKeyPos("A",7),"A","I","X","X")
			
			.lgPageNo			 =  "<%=lgPageNo%>"				    '☜: Next next data tag
			.DbQueryOk
			Parent.frm1.vspdData.Redraw = True
		End If
	End with
</Script>	
