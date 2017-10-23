<%@ LANGUAGE=VBSCript %>
<%Option Explicit%>
<%
'======================================================================================================
'*  1. Module Name          : Finance
'*  2. Function Name        : F_Notes
'*  3. Program ID           : f5102ma1
'*  4. Program Name         : ������ǥ��ȣ��� 
'*  5. Program Desc         : ����/��ǥå�� ���/����/����/��ȸ 
'*  6. Modified date(First) : 1999/09/10
'*  7. Modified date(Last)  : 2002/08/19
'*  8. Modifier (First)     : 
'*  9. Modifier (Last)      : Shin Myoung_Ha
'* 10. Comment              : 
'* 11. Common Coding Guide  : this mark(��) means that "Do not change"
'* 12. History				: 1. (ǥ��) FilterVar()�Լ� ���� 
'*							  2. ó�� �ε�� Ȱ��ȭ ���� �ʾƾ� �� ��ư�� Ȱ��ȭ �� - 2002/08/19
'*                            3. �ؽ�ƮŰ���� ���� ��ȸ���¿����� ���������� ��ȸ������ �ؽ�ƮŰ���� �������� ��ȸ������ 
'*								 ������ ���� ������ �����͵� ���� ��ȸ��(����) - 2002/09/06

'=======================================================================================================

%>

<!-- #Include file="../../inc/IncSvrMain.asp"  -->
<!-- #Include file="../../inc/IncSvrDate.inc"  -->

<%													'�� : ���⼭ ���� ������ �����Ͻ� ������ ó���ϴ� ������ ���۵ȴ� 

On Error Resume Next								'��:
ERR.CLEAR 


Dim strMode											'��: ���� MyBiz.asp �� ������¸� ��Ÿ�� 

Dim StrNextKey		' ���� �� 
Dim lgStrPrevKey	' ���� �� 
Dim LngMaxRow		' ���� �׸����� �ִ�Row
Dim LngRow
Dim GroupCount          


Call LoadBasisGlobalInf()

Call HideStatusWnd

strMode = Request("txtMode")						'�� : ���� ���¸� ���� 

'Multi SpreadSheet
lgLngMaxRow = Request("txtMaxRows")					'��: Read Operation Mode (CRUD)

Select Case strMode
    Case CStr(UID_M0001)							'��: Query
         Call SubBizQueryMulti()
    Case CStr(UID_M0002)                            '��: Save,Update             
         Call SubBizSaveMulti()
End Select


'============================================================================================================
' Name : SubBizQuery
' Desc : Query Data from Db
'============================================================================================================
Sub SubBizQueryMulti()
	
	Dim PAFG510
	Dim importArray
	Dim exportData
	Dim exportData1	
	Dim lgMaxCount
    Dim I2_f_note_no			'Parameter(����, ��ǥ)
	Dim I1_b_bank
	Dim iStrData				'��ȸ����Ÿ ���庯�� 
    Dim iIntLoopCount
    Dim iLngRow
    Dim strBankCd
    Dim strBankNm
    Dim iStrPrevKey
    Dim iPlantCd
    
	On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
	
	Const C_SHEETMAXROWS_D = 100				 			'�� ȭ�鿡 �������� �ִ밹��*1.5
	
	'###########################################
	'IMPORT PARAMETER
	Const C_MaxFetchRc		= 0
	Const C_PrevtKey        = 1
	Const C_NOTE_NO			= 2
	Const C_NOTE_KIND		= 3
	Const C_STS				= 4
	Const C_ISSUE_DT		= 5
	'###########################################
		
	'###########################################
	'EXPORT PARAMETER
	Const C_ENOTENO				= 0
	Const C_ENOTEKIND			= 1
	Const C_ENOTEKINDNM			= 2
	Const C_ENOTESTS			= 3
	Const C_ENOTESTSNM			= 4
	Const C_ENOTEISSUEDT		= 5
	Const C_ENOTEINSRTDT		= 6	
	Const C_ENOTEINSRTUSERID	= 7
	Const C_ENOTEUPDTDT			= 8
	Const C_ENOTEUPDTUSERID		= 9
	Const C_EBANKCD				= 10
	Const C_EBANKNM				= 11
	'###########################################
        
    lgMaxCount  = C_SHEETMAXROWS_D						'��: Fetch count at a time for VspdData
    
    Redim I2_f_note_no(6)
    
    '###########################################
    lgStrPrevKey = Request("lgStrPrevKey")         '��: Next Key Value
	
 	'Key ���� �о�´� 
	iPlantCd= Request("txtNoteNo")
	    
    '##############################################
        
    '######################################################
    'Data manipulate  area(import view match)        
    '����Ʈ ����Ÿ ���� 
    
    I2_f_note_no(C_MaxFetchRc) = lgMaxCount
    If lgStrPrevKey <> "" Then
		'I2_f_note_no(C_PrevtKey) = FilterVar(Trim(lgStrPrevKey),"","S")
		I2_f_note_no(C_PrevtKey) = Trim(lgStrPrevKey)
	Else
		'I2_f_note_no(C_PrevtKey) = FilterVar(Trim(Request("txtNoteNo")),"","S")
		I2_f_note_no(C_NOTE_NO) = Request("txtNoteNo")			'����, ��ǥ����		
	End If	
    I2_f_note_no(C_NOTE_KIND) = Request("cboNoteKind")
    I2_f_note_no(C_STS) = Request("txtSts")
    I2_f_note_no(C_ISSUE_DT) = UNIConvDate(Request("txtIssueDt"))
            
    'I1_b_bank = FilterVar(Trim(Request("txtBankCd")),"","S")	'�������� 
    I1_b_bank = Trim(Request("txtBankCd"))						'�������� 
    '######################################################
    
    '########################################
    'DEBUG
    'Response.Write I2_f_note_no(C_PrevtKey)
    'Response.Write I2_f_note_no(C_ISSUE_DT)
    '########################################
    
    
    Set PAFG510 = Server.CreateObject("PAFG510.cFListNoteNoSvr")
	If CheckSYSTEMError(Err, True) = True Then		
       Exit Sub    
    End If   
    
    '###################################################################################################
    'PROTOTYPE
    'LIST_NOTE_NO_SVR(ByVal pvStrGlobalCollection As String, Optional ByVal I1_f_note_no As Variant, _
	'				Optional ByVal I2_b_bank As Variant, _
	'				Optional ByRef EG1_export_group As Variant, _
	'				Optional ByRef E1_f_note_no As Variant)
	
	Call PAFG510.LIST_NOTE_NO_SVR(gStrGlobalCollection,I2_f_note_no,I1_b_bank,exportData)
	
	If CheckSYSTEMError(Err, True) = True Then					
       Set PAFG510 = Nothing
       Response.Write " <Script Language=vbscript> " & vbCr
		Response.Write " parent.DbQueryOk            " & vbCr
		Response.Write "</Script>                   " & vbCr
       Exit Sub
    End If    
    '###################################################################################################
   
	iStrData = ""
    iIntLoopCount = 0	
	
	'###################################################################################################

	'Export Data ���� 
	strBankCd = ConvSPChars(Trim(exportData(0,C_EBANKCD)))
	strBankNm = ConvSPChars(Trim(exportData(0,C_EBANKNM)))
	
	For iLngRow = 0 To UBound(exportData, 1)
		iIntLoopCount = iIntLoopCount + 1
		If  iLngRow < lgMaxCount Then
			istrData = istrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, C_ENOTEKINDNM)))				
			istrData = istrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, C_ENOTEKIND)))				
			istrData = istrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, C_EBANKCD)))
			istrData = istrData & Chr(11) & ""
			istrData = istrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, C_EBANKNM)))
			istrData = istrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, C_ENOTENO)))			
			istrData = istrData & Chr(11) & UNIDateClientFormat(Trim(exportData(iLngRow, C_ENOTEISSUEDT))) 'UNIDateClientFormat
			istrData = istrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, C_ENOTESTS)))
			istrData = istrData & Chr(11) & ConvSPChars(Trim(exportData(iLngRow, C_ENOTESTSNM)))
			istrData = istrData & Chr(11) & Cint(LngMaxRow) + iLngRow + 1
			istrData = istrData & Chr(11) & Chr(12)
			'Response.Write ILNGROW & "<BR>"
			'Response.Write 	Trim(exportData(ILNGROW, C_ENOTENO)) & "<BR>"
		Else
			'Response.Write ILNGROW & "<BR>"
			'Response.Write "ISTRPREVKEY = " & Trim(exportData(ILNGROW, C_ENOTENO)) & "<BR>"
			iStrPrevKey = Trim(exportData(ILNGROW, C_ENOTENO))			
		End If		
	Next
	'#####################################################################################################
	
	'###################################
	'Debug
	'Response.Write "ISTRPREVKEY = " & Trim(exportData(ILNGROW, C_ENOTENO))
	'###################################
	
	If  iLngRow < lgMaxCount Then
		iStrPrevKey = ""		
	End If	


	'#####################################################################################################
	'��: ȭ�� ó�� ASP �� ��Ī�� 
	Response.Write " <Script Language=vbscript>			"	& vbCr
    Response.Write " With parent						"	& vbCr
	Response.Write " .frm1.txtBankCd.value =			"""	& strBankCd & """		" & vbCr
	Response.Write " .frm1.txtBankNm.value =			"""	& strBankNm & """		" & vbCr
	Response.Write " .frm1.vspdData.Redraw = False		"	& vbCr
	Response.Write " .ggoSpread.Source = .frm1.vspdData "	& vbCr
	Response.Write " .ggoSpread.SSShowData 				"""   & istrData & """"			& vbCr	
	Response.Write " .frm1.vspdData.Redraw = True       "	& vbCr	
	Response.Write " .lgStrPrevKey =                    """	& iStrPrevKey & """"	& vbCr
	Response.Write " .frm1.hBankCd.value   =			"""	& strBankCd & """" & vbCr
	Response.Write " .frm1.hNoteKind.value =			"""	& I2_f_note_no(C_NOTE_KIND) & """" & vbCr	
	Response.Write " .frm1.hIssueDt.value  =			"""	& I2_f_note_no(C_ISSUE_DT) & """" & vbCr	
	Response.Write " .DbQueryOk							"	& vbCr	
	Response.Write " End With							"	& vbCr	
	Response.Write " </Script>							"	& vbCr
'#######################################################################################################


End Sub

'============================================================================================================
' Name : SubBizSaveMulti
' Desc : Save Data 
'============================================================================================================
Sub SubBizSaveMulti()
	
	Dim PAFG510
	Dim arrTemp
	Dim iErrorPosition

	On Error Resume Next                                                             '��: Protect system from crashing
    Err.Clear                                                                        '��: Clear Error status
    
    Set PAFG510 = Server.CreateObject("PAFG510.cFMngNoteNoSvr")

    If CheckSYSTEMError(Err, True) = True Then					
       Exit Sub
    End If    

    Call PAFG510.MANAGE_NOTE_NO_SVR(gStrGlobalCollection, Trim(Request("txtSpread")), iErrorPosition)
		
    If CheckSYSTEMError2(Err, True,iErrorPosition & "��","","","","") = True Then
       Set PAFG510 = Nothing
       Exit Sub
    End If    
    
    Set PAFG510 = Nothing
		
	Response.Write " <Script Language=vbscript> " & vbCr
	Response.Write " parent.DbSaveOk            " & vbCr
    Response.Write "</Script>                   " & vbCr
    
end sub	

%>
