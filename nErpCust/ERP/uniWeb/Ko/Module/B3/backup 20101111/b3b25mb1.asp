<%@LANGUAGE = VBScript%>
<%'**********************************************************************************************
'*  1. Module Name          : Production
'*  2. Function Name        : 
'*  3. Program ID           : b3b25mb1.asp
'*  4. Program Name         : Copy Item By Plant
'*  5. Program Desc         :
'*  6. Comproxy List        : 
'*  7. Modified date(First) : 
'*  8. Modified date(Last)  : 2003/02/10
'*  9. Modifier (First)     : 
'* 10. Modifier (Last)      : Park In Sik
'* 11. Comment              :
'**********************************************************************************************%>

<%Option Explicit%>
<!-- #Include file="../../inc/IncSvrMain.asp" -->
<!-- #Include file="../../inc/IncSvrNumber.inc" -->
<!-- #Include file="../../inc/IncSvrDate.inc" -->
<!-- #Include file="../../ComAsp/LoadInfTB19029.asp" -->
<%
Call LoadBasisGlobalInf
Call LoadInfTB19029B("I", "*", "NOCOOKIE", "MB")
													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 
Dim ADF														'ActiveX Data Factory ���� �������� 
Dim strRetMsg												'Record Set Return Message �������� 
Dim UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3 ,rs4 ,rs5				'DBAgent Parameter ���� 
Dim strQryMode
Dim BlankchkFlg

Const C_SHEETMAXROWS_D = 100

Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

'=======================================================================================================
'	�Ʒ� ����Ǿ� �ִ� �������� COOL:Gen �� Record Return Count �� ���ѿ� ���� ���̴�.
'	����, ADO�� ����� ��� �׿Ͱ��� �������� ���� ������ �Ʒ��� �������� ������� ������ ���� 
'	uniERP2000 ���� �ѹ��� ��ȸ�Ǵ� Record Count �� ���� 30���� �����ϰ� �ִ� ��ŭ �׿� ���� 
'	ǥ���� ���ÿ� �߰��� �����̹Ƿ� ���������� ���� �ʰ� �״�� ���д�.
'=======================================================================================================
Dim StrNextKey		' ���� �� 
Dim lgStrPrevKey	
Dim i

Call HideStatusWnd

Dim strYear, strMonth, strDay, StartDate

Call ExtractDateFrom(GetSvrDate, gServerDateFormat, gServerDateType, strYear, strMonth, strDay)
StartDate = UniConvYYYYMMDDToDate(gDateFormat, strYear, strMonth, strDay)

strMode = Request("txtMode")						'�� : ���� ���¸� ���� 
strQryMode = Request("lgIntFlgMode")

On Error Resume Next
Err.Clear

Dim strPlantCd
Dim strPlantNm
Dim strItemCd
Dim strItemNm
Dim strItemCd1
Dim strClassCd
Dim strClassNm
Dim strItemGroup
Dim strItemGroupNm


lgStrPrevKey = FilterVar(UCase(Request("lgStrPrevKey")), "''", "S")
	
'======================================================================================================
'	ǰ���̸� ó�����ִ� �κ� 
'======================================================================================================
	Redim UNISqlId(5)
	Redim UNIValue(5, 1)
		
	UNISqlId(0) = "122700sab"	'plant_nm
	UNISqlId(1) = "122600sac"	'item_nm
	UNISqlId(2) = "127400saa"	'item_group_nm
	UNISqlId(3) = "122600sac"	'item_nm
	UNISqlId(4) = "b3b25mb1a"	'class_nm
	UNISqlId(5) = "122600SAG"	'ProcType
		
	strItemCd = FilterVar(Request("txtItemCd"),"''","S")

	strItemGroup = FilterVar(Request("txtItemGroupCd"),"''","S")


	   strPlantCd = FilterVar(Request("txtPlantCd"),"''","S")

	   strItemCd1 = FilterVar(Request("txtItemCd1"),"''","S")

		strClassCd = FilterVar(Request("txtClassCd"),"''","S")

	'WHERE ���� 
		
	UNIValue(0, 0) = strPlantCd
	UNIValue(1, 0) = strItemCd
	UNIValue(2, 0) = strItemGroup
	UNIValue(3, 0) = strItemCd1
	UNIValue(4, 0) = strClassCd
	UNIValue(5, 0) = strPlantCd
	UNIValue(5, 1) = strItemCd1
	
	UNILock = DISCONNREAD :	UNIFlag = "1"
	
	BlankchkFlg = False
	    
    Set ADF = Server.CreateObject("prjPublic.cCtlTake")
    strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0, rs1, rs2, rs3, rs4, rs5)
		
	Call  SetConditionData()
		
	'-----------------------
    'Com action result check area(OS,internal)
    '-----------------------
    If Err.Number <> 0 Then
		Set rs0 = Nothing
		Set rs1 = Nothing
		Set rs2 = Nothing
		Set rs3 = Nothing
		Set rs4 = Nothing
		Set rs5 = Nothing
		Set ADF = Nothing												'��: ComProxy Unload
		Call ServerMesgBox(Err.description, vbCritical, I_MKSCRIPT)						'��:		
		Response.End														'��: �����Ͻ� ���� ó���� ������ 
	End If
		
	If rs3.EOF And rs3.BOF Then
		Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtItemNm1.value = """"" & vbCrLf
		Response.Write "parent.frm1.txtItemSpec1.value = """"" & vbCrLf
		Response.Write "parent.frm1.txtItemProcType1.value = """"" & vbCrLf
		Response.Write "</script>" & vbCr
	Else
		Response.Write "<Script Language = VBScript>" & vbCrLf
		Response.Write "parent.frm1.txtItemNm1.value = """ & ConvSPChars(rs3(0)) & """" & vbCrLf
		Response.Write "parent.frm1.txtItemSpec1.value = """ & ConvSPChars(rs3(1)) & """" & vbCrLf
		Response.Write "parent.frm1.txtItemProcType1.value = """ & ConvSPChars(rs5(1)) & """" & vbCrLf
		Response.Write "parent.frm1.htxtItemProcType1.value = """ & ConvSPChars(rs5(0)) & """" & vbCrLf
		Response.Write "</Script>" & vbCr
	End If

	rs3.Close
	Set rs3 = Nothing
	Set ADF = Nothing				'��: ActiveX Data Factory Object Nothing

	
'=======================================================================================================
'	����, ������ ������ �迭�̶�� �Ʒ��Ͱ��� Fix �� �迭�� Redim �� �ؼ� �Ѱ���� �Ѵ�.
'=======================================================================================================

Redim UNISqlId(0)
Redim UNIValue(0,4)
	
UNISqlId(0) = "B3B25MB1"
	
IF Request("txtPlantCd") = "" Then
   strPlantCd = "|"
ELSE
   strPlantCd = FilterVar(Request("txtPlantCd"),"''","S")
END IF
	
IF Request("txtItemCd") = "" Then
   strItemCd = "|"
ELSE
   strItemCd = FilterVar(Request("txtItemCd"),"''","S")
END IF
	
	
IF Request("txtItemGroupCd") = "" Then
   strItemGroup = "|"
ELSE
   strItemGroup = FilterVar(Request("txtItemGroupCd"),"''","S")
END IF
	
IF Request("txtClassCd") = "" Then
   strClassCd = "|"
ELSE
   strClassCd = FilterVar(Request("txtClassCd"),"''","S")
END IF
		
UNIValue(0, 0) = "^"
UNIValue(0, 1) = strPlantCd
	
Select Case strQryMode
	Case CStr(OPMD_CMODE)
		UNIValue(0, 2) = strItemCd
	Case CStr(OPMD_UMODE) 
		UNIValue(0, 2) = lgStrPrevKey
End Select

UNIValue(0, 3) = strClassCd
IF Request("txtItemGroupCd") = "" Then
	UNIValue(0,4) = "|"
Else
	UNIValue(0,4) = "a.item_group_cd in (select item_group_cd from ufn_P_ListItemGrp(" & strItemGroup & " ))"
End IF
	
UNILock = DISCONNREAD :	UNIFlag = "1"
	
Set ADF = Server.CreateObject("prjPublic.cCtlTake")

strRetMsg = ADF.QryRs(gDsnNo, UNISqlId, UNIValue, UNILock, UNIFlag, rs0)

If BlankchkFlg = False Then	
	If  rs0.EOF And rs0.BOF And BlankchkFlg =  False Then
	    Call DisplayMsgBox("900014", vbOKOnly, "", "", I_MKSCRIPT)
	    rs0.Close
	    Set rs0 = Nothing
        ' Modify Focus Events    
        %>
            <Script language=vbs>
            Parent.frm1.txtPlantCd.focus    
            </Script>
        <%
	Else    
	    %>

	    <Script Language=vbscript>
		Dim LngMaxRow
		Dim strData
		Dim TmpBuffer
		Dim iTotalStr

		With parent																	'��: ȭ�� ó�� ASP �� ��Ī�� 
			LngMaxRow = .frm1.vspdData.MaxRows										'Save previous Maxrow
		<%		
				If C_SHEETMAXROWS_D < rs0.RecordCount Then 
		%>
					ReDim TmpBuffer(<%=C_SHEETMAXROWS_D - 1%>)
		<%
				Else
		%>			
					ReDim TmpBuffer(<%=rs0.RecordCount - 1%>)
		<%
				End If

				For i=0 to rs0.RecordCount-1 
					
					If i < C_SHEETMAXROWS_D Then
		%>				
						strData = ""	
						strData = strData & Chr(11) & ""											'1:C_Select
						strData = strData & Chr(11) & "<%=ConvSPChars(rs0(0))%>"					'2:C_Item
						strData = strData & Chr(11) & "<%=ConvSPChars(rs0(1))%>"					'3:C_ItmNm
						strData = strData & Chr(11) & "<%=ConvSPChars(rs0(2))%>"					'4:C_ItmSpec
						strData = strData & Chr(11) & "S"											'5:C_PrcCtrlInd
						strData = strData & Chr(11) & ""											'6:C_PrcCtrlIndNm
						strData = strData & Chr(11) & "<%=UniConvNumberDBToCompany("",ggUnitCost.DecPoint,ggUnitCost.RndPolicy, ggUnitCost.RndUnit,0)%>"	'7:C_UnitPrice
						strData = strData & Chr(11) & "<%=startdate%>"								'8:C_ValidFromDt
						strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0(16))%>"			'9:C_ValidToDt
							
						strData = strData & Chr(11) & "<%=ConvSPChars(rs0(3))%>"					'10: Ŭ���� 
						strData = strData & Chr(11) & "<%=ConvSPChars(rs0(4))%>"					'11: Ŭ������ 
						strData = strData & Chr(11) & "<%=ConvSPChars(rs0(5))%>"					'12: ��簪1
						strData = strData & Chr(11) & "<%=ConvSPChars(rs0(6))%>"					'13: ��簪��1
						strData = strData & Chr(11) & "<%=ConvSPChars(rs0(7))%>"					'14: ��簪2
						strData = strData & Chr(11) & "<%=ConvSPChars(rs0(8))%>"					'15: ��簪��2
							
						strData = strData & Chr(11) & ""											'16: ǰ�����nm
						strData = strData & Chr(11) & "<%=ConvSPChars(rs0(9))%>"					'17: ǰ����� 
						strData = strData & Chr(11) & "<%=ConvSPChars(rs0(10))%>"					'18: ���ش��� 
						strData = strData & Chr(11) & "<%=ConvSPChars(rs0(11))%>"					'18: Phantom���� 
						strData = strData & Chr(11) & "<%=ConvSPChars(rs0(12))%>"					'19: ǰ��׷� 
						strData = strData & Chr(11) & "<%=ConvSPChars(rs0(13))%>"					'20:gourpNm
						strData = strData & Chr(11) & "<%=rs0(14)%>"								'21:C_DefaultFlg
						strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0(15))%>"			'22:C_ValidFromDt
						strData = strData & Chr(11) & "<%=UNIDateClientFormat(rs0(16))%>"			'23:C_ValidToDt
							
						strData = strData & Chr(11) & LngMaxRow + "<%=i%>"							'31:
						strData = strData & Chr(11) & Chr(12)
						
						TmpBuffer(<%=i%>) = strData 
				
		<%			
						rs0.MoveNext
					End If
				Next
		%>
				iTotalStr = Join(TmpBuffer,"")
				.ggoSpread.Source = .frm1.vspdData
				.ggoSpread.SSShowDataByClip iTotalStr
					
				.lgStrPrevKey = "<%=Trim(rs0(0))%>"
				
				.frm1.txtPlantNm.value			= "<%=ConvSPChars(strPlantNm)%>"
				.frm1.txtClassNm.value			= "<%=ConvSPChars(strClassNm)%>"
				.frm1.txtItemNm.value			= "<%=ConvSPChars(strItemNm)%>"
				.frm1.txtHighItemGroupNm.value  = "<%=ConvSPChars(strItemGroupNm)%>"
					
				.frm1.hItemCd.value			= "<%=ConvSPChars(Request("txtItemCd"))%>"
				.frm1.hPlantCd.value		= "<%=ConvSPChars(Request("txtPlantCd"))%>"
				.frm1.htxtClassCd.value		= "<%=Request("txtClassCd")%>"	
				.frm1.hItemGroupCd.value	= "<%=ConvSPChars(Request("txtItemGroupCd"))%>"
		<%			
				rs0.Close
				Set rs0 = Nothing
		%>
			.DbQueryOk(LngMaxRow + 1)
		End With	
		</Script>	
						

<%		    

	End If
End If	

Set ADF = Nothing												'��: ActiveX Data Factory Object Nothing
'++++++++++++++++++++++++++++++++++++++++++++++++++++++

'----------------------------------------------------------------------------------------------------------
' Name : SetConditionData
' Desc : set value in condition area
'----------------------------------------------------------------------------------------------------------
Sub SetConditionData()
    On Error Resume Next
	
    If Not(rs0.EOF Or rs0.BOF) Then
       strPlantNm =  rs0(0)
			%>
                <Script language=vbs>
				Parent.frm1.txtPlantNm.value			= "<%=ConvSPChars(strPlantNm)%>"
				</Script>
            <%		
    Else
   		If Len(Request("txtPlantCd")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "����", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		   BlankchkFlg = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtPlantCd.focus
                Parent.frm1.txtPlantNm.value = ""    
                </Script>
            <%		   	
		End If	
    End If   
    
    Set rs0 = Nothing 
	
	If Not(rs4.EOF Or rs4.BOF) Then
       strClassNm =  rs4(0)
			%>
                <Script language=vbs>
				Parent.frm1.txtClassNm.value			= "<%=ConvSPChars(strClassNm)%>"
				</Script>
            <%		   	
    Else
   		If Len(Request("txtClassCd")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "Ŭ����", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		   BlankchkFlg = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtClassCd.focus
                Parent.frm1.txtClassNm.value = ""        
                </Script>
            <%		   	
		End If	
    End If   
    
    Set rs4 = Nothing 

	If Not(rs1.EOF Or rs1.BOF) Then
       strItemNm =  rs1(0)
			%>
                <Script language=vbs>
				Parent.frm1.txtItemNm.value				= "<%=ConvSPChars(strItemNm)%>"
				</Script>
            <%		   	
    Else
   		If Len(Request("txtItemCd")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "ǰ��", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		   BlankchkFlg = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtItemCd.focus
                Parent.frm1.txtItemNm.value = ""
                </Script>
            <%		   	
		End If	
    End If   
    
    Set rs1 = Nothing 

	If Not(rs2.EOF Or rs2.BOF) Then
		strItemGroupNm = rs2(0)
			%>
                <Script language=vbs>
				Parent.frm1.txtHighItemGroupNm.value  = "<%=ConvSPChars(strItemGroupNm)%>"
				</Script>
            <%		   	
    Else
   		If Len(Request("txtItemGroupCd")) And BlankchkFlg = False Then
		   Call DisplayMsgBox("970000", vbInformation, "ǰ��׷�", "", I_MKSCRIPT)	'��: you must release this line if you change msg into code
		   BlankchkFlg = True
            ' Modify Focus Events    
            %>
                <Script language=vbs>
                Parent.frm1.txtHighItemGroupCd.focus
                Parent.frm1.txtHighItemGroupNm.value = ""
                </Script>
            <%		   	
		End If	
    End If   
    
    Set rs2 = Nothing
	
End Sub	
%>




