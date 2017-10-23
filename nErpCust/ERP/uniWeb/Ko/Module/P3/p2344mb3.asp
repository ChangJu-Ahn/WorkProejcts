<%@LANGUAGE = VBScript%>
<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgent.inc" -->
<!-- #Include file="../../inc/IncSvrDBAgentVariables.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<!--'======================================================================================================
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : p2344mb3.asp
'*  4. Program Name         :
'*  5. Program Desc         :
'*  6. Modified date(First) :
'*  7. Modified date(Last)  : 
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : Jung Yu Kyung
'* 10. Comment              :
'* 11. Common Coding Guide  : this mark(☜) means that "Do not change"
'=======================================================================================================-->
<% 

Call LoadBasisGlobalInf
Call LoadInfTB19029B("Q", "P", "NOCOOKIE", "MB") 

On Error Resume Next

Dim ADF1 
Dim strRetMsg
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0
Dim strQuryMod
Dim i

Const C_SHEETMAXROWS = 100

Dim lgStrPrevKey21	' 이전 값 
Dim lgStrPrevKey22	' 이전 값 

Call HideStatusWnd

strQryMode = Request("lgIntFlgMode")

On Error Resume Next

Dim strItemCd
Dim strTrackingNo
Dim strConvType1
Dim strConvType2
Dim strStartDt
Dim strEndDt
Dim txtPurOrg

	lgStrPrevKey21 = UCase(Trim(Request("lgStrPrevKey21")))
	lgStrPrevKey22 = UCase(Trim(Request("lgStrPrevKey22")))
	
	Redim UNISqlId(0)
	Redim UNIValue(0, 11)
	
	UNISqlId(0) = "P2344MB1B"	
	
	IF Request("txtItemCd") = "" Then
		strItemCd = "|"
	Else
		StrItemCd = FilterVar(Trim(Request("txtItemCd"))	, "''", "S")
	End IF

	IF Request("txtTrackingNo") = "" Then
		strTrackingNo = "|"
	Else
		StrTrackingNo = FilterVar(Trim(Request("txtTrackingNo"))	, "''", "S")
	End IF

	IF Request("rdoConvType") = "A" THEN
       strConvType1 = "|"
       strConvType2 = "|"
    ELSEIF Request("rdoConvType") = "NL" then
		strConvType1 = "" & FilterVar("NL", "''", "S") & ""
		strConvType2 = "|"		
    Else	
		strConvType1 = "|"
		strConvType2 = "" & FilterVar("NL", "''", "S") & ""
    END IF
    
	IF Request("txtStartDt") = "" THEN
	   strStartDt = "|"
	ELSE
	   strStartDt = FilterVar(UniConvDate(Request("txtStartDt"))	, "''", "S")
	END IF

    IF Request("txtEndDt") = "" THEN
    	strEndDt = "|"
    ELSE
    	strEndDt = FilterVar(UniConvDate(Request("txtEndDt"))	, "''", "S")
    END IF      
    
    IF Request("txtPurOrg") = "" THEN
    	txtPurOrg = "|"
    ELSE
    	txtPurOrg = FilterVar(Trim(Request("txtPurOrg"))	, "''", "S")
    END IF   
			
	UNIValue(0, 0) = "^"
	UNIValue(0, 1) = FilterVar(Trim(Request("txtPlantCd"))	, "''", "S")
	UNIValue(0, 2) = FilterVar(Trim(Request("lgStrPrevKey21"))	, "''", "S")
	UNIValue(0, 3) = strTrackingNo		
	UNIValue(0, 4) = strConvType1
	UNIValue(0, 5) = strConvType2	
	UNIValue(0, 6) = "a.item_cd > " & FilterVar(lgStrPrevKey21	, "''", "S") & " or (a.item_cd = " & FilterVar(lgStrPrevKey21	, "''", "S")	
	UNIValue(0, 7) = FilterVar(lgStrPrevKey22	, "''", "S")
	UNIValue(0, 8) = strStartDt
	UNIValue(0, 9) = strEndDt	
	UNIValue(0, 10) = txtPurOrg
	IF Request("txtItemGroupCd") = "" Then
		UNIValue(0, 11) = "|"
	Else
		UNIValue(0, 11) = "b.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & FilterVar(Trim(Request("txtItemGroupCd"))	, "''", "S") & " ))"
	End IF
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
    Set ADF1 = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF1.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)
      
	If rs0.EOF And rs0.BOF Then
		Call DisplayMsgBox("900014", vbInformation, "", "", I_MKSCRIPT)
		rs0.Close		
		Set rs0 = Nothing
		Set ADF = Nothing		
		Response.End
	End If
	
%>

<Script Language=vbscript>
Dim LngMaxRow2
Dim strData
Dim arrVal
ReDim arrVal(0)

With parent	
	LngMaxRow2 = .frm1.vspdData2.MaxRows
		
<%  
	If Not(rs0.EOF And rs0.BOF) Then
		For i=0 to rs0.RecordCount-1 
			IF i < C_SHEETMAXROWS Then
%>
				strData = ""
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_CD"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ITEM_NM"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("SPEC"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("tracking_no"))%>"
				strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany(rs0("PLAN_QTY"),ggQty.DecPoint, ggQty.RndPolicy, ggQty.RndUnit,0)%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("BASIC_UNIT"))%>"			'단위 
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("START_PLAN_DT"))%>"
				strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0("END_PLAN_DT"))%>"		

<%				IF Trim(rs0("PLAN_STATUS")) ="NL" Then%>
				   strData = strData & Chr(11) & "Plan"
<%				ELSEIF Trim(rs0("PLAN_STATUS")) = "OP" Then%>
					strData = strData & Chr(11) & "Open"
<%				ELSEIF Trim(rs0("PLAN_STATUS")) = "RL" Then%>
					strData = strData & Chr(11) & "Release"
<%				ELSEIF Trim(rs0("PLAN_STATUS")) = "ST" Then%>
					strData = strData & Chr(11) & "Start"
<%				ELSE%>
					strData = strData & Chr(11) & "Close"
<%				END IF%>

				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("ORDER_NO"))%>"  'MRP_RUN_NO '제조오더번호 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("PUR_ORG"))%>" 
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_cd"))%>"
				strData = strData & Chr(11) & "<%=ConvSPChars(rs0("item_group_nm"))%>"
				strData = strData & Chr(11) & LngMaxRow2 + "<%=i%>"
				strData = strData & Chr(11) & Chr(12)
				
				ReDim Preserve arrVal(<%=i%>)
				arrVal(<%=i%>) = strData
<%		
				rs0.MoveNext
			End If
		Next
%>
		.ggoSpread.Source = .frm1.vspdData2
		.ggoSpread.SSShowData Join(arrVal,"")
		
		.lgStrPrevKey21 = "<%=ConvSPChars(rs0("ITEM_CD"))%>"
		.lgStrPrevKey22 = "<%=ConvSPChars(rs0("PLAN_ORDER_NO"))%>"
		
<%
	End If
%>

End With

</Script>	
<%
rs0.Close
Set rs0 = Nothing

Set ADF = Nothing
%>
