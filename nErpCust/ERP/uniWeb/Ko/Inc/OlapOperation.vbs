'************************************************************************************************
'*  1. Module Name          : BA
'*  2. Function Name        : 
'*  3. Program ID           : 0lapOperation.vbs
'*  4. Program Name         : 
'*  5. Program Desc         : 
'*  6. Comproxy List        : 
'                             +
'*  7. Modified date(First) : 2001/12/10
'*  8. Modified date(Last)  : 2001/12/10
'*  9. Modifier (First)     : LEE SEOK GON
'* 10. Modifier (Last)      : LEE SEOK GON
'* 11. Comment              : 
'* 12. Common Coding Guide  : 
'* 13. History              :
'*                           
'************************************************************************************************
<!-- #Include file="./Operation.vbs"  -->

'================================================================================================
' ConnectToCube()
'
' Purpose: Connects the PivotTable to an OLAP cube
' In:      ptable   reference to the PivoTable component
'          sConn    connection string for the data source
'          sCube    name of the cube to open
'================================================================================================
Sub ConnectToCube(ptable, sConn, sCube)
    Dim IntRetCD 
    On Error Resume Next
    Err.Clear
 
    ' Set the PivotTable's ConnectionString property///////////////////////
    ptable.ConnectionString = sConn
    
    ' Set the DataMember property to the cube name/////////////////////////
    ptable.DataMember = sCube

     If err.number <> 0 Then
		If err.number = -2147217843 Then 
			IntRetCD = DisplayMsgBox("990029", VB_YES_NO, "X", "X")
			Exit  Sub
		Else
			msgbox err.number & " : " & err.Description 
			Exit  Sub
		End If	
     End If    

End Sub 'ConnectToCube()

'================================================================================================
' QuickPivot()
'
' Purpose: Configures the PivotTable's view to show a report
' In:      ptable     reference to the PivotTable component
'          ptotal     total to show in the data area
'          fsRows     FieldSet to show on the rows axis
'          fsCols     FieldSet to show on the columns axis
'          fsFilter   FieldSet to show in the filter area
'================================================================================================
Sub QuickPivot(ptable, ptotal, fsRows, fsCols, fsFilter)
    ' Local variables
    Dim pview    ' Reference to the view
    Dim i
    
    ' Grab a reference to the view
    set pview = ptable.ActiveView
    
    ' Clear the view
    pview.AutoLayout()
    
    ' Row의 Non-Time Dimension 셋팅////////////////////////////////////////////
    For i = 0 to ubound(fsRows)
		pview.RowAxis.InsertFieldSet pview.FieldSets(fsRows(i))
    Next
    'pview.RowAxis.InsertFieldSet pview.FieldSets(fsRows(0))
    ' Column의 Time Dimension 셋팅/////////////////////////////////////////////
    pview.ColumnAxis.InsertFieldSet pview.FieldSets(fsCols)
    
    ' Filter Axis의 Dimension 셋팅/////////////////////////////////////////////
    For i = 0 to ubound(fsFilter)
		pview.FilterAxis.InsertFieldSet pview.FieldSets(fsFilter(i))
    Next
    
    'Data Area에 Total값 셋팅//////////////////////////////////////////////////
    For i = 0 to ubound(ptotal)
		pview.DataAxis.InsertTotal pview.Totals(ptotal(i))
    Next
    
End Sub 'QuickPivot()

'================================================================================================
' BindChartToPivot()
'
' Purpose: Display by changing the chart type 
'================================================================================================
Sub BindChartToPivot(cspace, ptable, chCategoiesvalue, chtotalvalue, chtypevalue, ptotal)
    ' Local variables
    Dim cht    ' Chart object that we'll create in the chart space
    Dim ax     ' Temp axis reference
    Dim fnt    ' Temp font reference
    Dim nChartType ' Variable to save ChartType constant////////////////////////
    Dim c
    
    ' Grab the Constants object so that we can use constant names in
    ' the script. Note: This is needed only in VBScript -- do not include
    ' this in VBA code.
    set c = cspace.Constants
        
    ' Clear out anything that is in the chart space
    cspace.Clear

    ' First tell the chart that its data is coming from the PivotTable 
    set cspace.DataSource = ptable
        
    ' Create a  chart in the chart space
    set cht = cspace.Charts.Add()
    cht.HasLegend = True
    
    '리스트박스에서 차트변환상수를 받아오는 부분///////////////////////////////
    nChartType = CLng(chtypevalue)
	cht.Type = nChartType 
	
    If chtotalvalue = "A" Then
		Call MultiTotals(cspace, c, cht, chCategoiesvalue)
	Else
		Call SingleTotals(cspace, c, Cht, ptable, chCategoiesvalue, chtotalvalue, ptotal)
	End IF		
    
    'ax.NumberFormat = ptable.ActiveView.DataAxis.Totals(0).NumberFormat
    
End Sub 'BindChartToPivot()

'================================================================================================
' MultiTotals()
'
' Purpose: Display Multitotals in the chart
'================================================================================================
Function MultiTotals(Cspace, c , Chtobj, chCategoiesvalue)
	
	Dim fSeriesInCols
	
	MultiTotals = False
	
	'Chart의 Series값과 Categories의 Value를 바꿈////////////////////////////// 
	
	fSeriesInCols = chCategoiesvalue 
	
	if fSeriesInCols then
	    Chtobj.SetData c.chDimSeriesNames, 0, 	c.chPivotColAggregates		'c.chPivotColumns
	    Chtobj.SetData c.chDimCategories,  0, c.chPivotRowAggregates		'c.chPivotRows 
	else
	    Chtobj.SetData c.chDimSeriesNames,  0, c.chPivotRowAggregates	'c.chPivotRows
	    Chtobj.SetData c.chDimCategories, 0, c.chPivotColAggregates	'c.chPivotColumns  		
	end if 'fSeriesInCols
		
	Chtobj.SetData c.chDimValues,  0, 0	
	
	 ' Finally, let's add title to the value
    cspace.HasChartSpaceTitle = True
    cspace.ChartSpaceTitle.Caption = strTitleBar   
    set fnt = cspace.ChartSpaceTitle.Font 
    fnt.Name = "Tahoma"
    fnt.Size = 12
    fnt.Bold = True
    
	MultiTotals = True
	
End Function     

'================================================================================================
' MultiTotals()
'
' Purpose: Display Single totals in the chart
'================================================================================================
Function SingleTotals(Cspace, c, Chtobj, ptable, chCategoiesvalue, chtotalvalue, ptotal)
	
	Dim i
	Dim chdimvalueNum
	Dim fSeriesInCols
	
	SingleTotals = False
	
	'Chart의 Series값과 Categories의 Value를 바꿈////////////////////////////// 
	
	fSeriesInCols = chCategoiesvalue
	
	if fSeriesInCols then
	    Chtobj.SetData c.chDimSeriesNames, 0, 	c.chPivotColumns
	    Chtobj.SetData c.chDimCategories,  0, c.chPivotRows 
	else
	    Chtobj.SetData c.chDimSeriesNames,  0, c.chPivotRows
	    Chtobj.SetData c.chDimCategories, 0, c.chPivotColumns  		
	end if 'fSeriesInCols

	for  i = 0  to ubound(ptotal)
		If ptotal(i) = chtotalvalue Then
			chdimvalueNum  = i 
			exit for
		End If			 
	next 

	Chtobj.SetData c.chDimValues,  0, chdimvalueNum
	
	 ' Finally, let's add title to the value
    cspace.HasChartSpaceTitle = True
    cspace.ChartSpaceTitle.Caption = ptable.ActiveView.DataAxis.Totals(chtotalvalue).Caption 
    set fnt = cspace.ChartSpaceTitle.Font 
    fnt.Name = "Tahoma"
    fnt.Size = 12
    fnt.Bold = True
    
	SingleTotals = True
	
End Function

'================================================================================================
' LoadSelectWithFieldsets()
'
' Purpose: Loads an HTML SELECT control with the fieldsets in a 
'          given PivotTable control
' In:      ptable   PivotTable control
'          sel      HTML SELECT control
'================================================================================================
Sub LoadSelectWithFieldsets(ptable, sel)
    ' Local Variables
    Dim fs     ' Temporary fieldset pointer
    Dim opt    ' New OPTION element for the HTML SELECT control
    
    ' Load the HTML SELECT control with all the available fieldsets
    for each fs in ptable.ActiveView.FieldSets
        set opt = document.createElement("OPTION")
        opt.Text = fs.Caption
        opt.Value = fs.Name
        sel.options.add opt
    next 'fs
    
    Set opt = Nothing
    
End Sub 'LoadSelectWithFieldsets()

'================================================================================================
' LoadSelectWithTotals()
'
' Purpose: Loads an HTML SELECT control with the Totals in a 
'          given PivotTable control
' In:      ptable   PivotTable control
'          sel      HTML SELECT control
'================================================================================================
Sub LoadSelectWithTotals(ptable,sel)
	Dim MeasureSel
	Dim pview 
	Dim objEl
	Dim i
	
	For each fs in ptable.ActiveView.Totals      
		Set objEl = Document.CreateElement("OPTION")	
		objEl.Text = fs.Caption
		objEl.Value = fs.Name
		sel.Add(objEl)
    next 'fs
    	
	Set opjEI = Nothing
		
End Sub 'LoadSelectWithTotals()

