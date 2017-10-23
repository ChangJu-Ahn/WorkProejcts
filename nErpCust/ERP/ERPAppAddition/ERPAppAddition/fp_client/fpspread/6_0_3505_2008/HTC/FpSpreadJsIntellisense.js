//
//	  Copyright (C) 2003-2010 GrapeCity Inc.	All rights reserved.
//

if (typeof(FarPoint)=="undefined")
  FarPoint = {}
FarPoint.Web = {}
FarPoint.Web.Spread = {} 

FarPoint.Web.Spread.Range = function() {}

FarPoint.Web.Spread.Range.prototype = {
  type: "Cell",
  row: -1,
  col: -1,
  rowCount: 0,
  colCount: 0
}

FarPoint.Web.Spread.TitleInfo = function() {}

FarPoint.Web.Spread.TitleInfo.prototype = {
  GetHeight: function() {
    /// <summary>Gets the height of the Title. </summary>
    ///	<returns type="Number" integer="true">
    ///		Integer: the height of the Title. 
    ///	</returns>
  },
  SetHeight: function(height) {
    /// <summary>Sets the height of the Title. </summary>
    ///	<param name="height" type="Number" integer="true">
    ///		Integer, the height of the Title. 
    ///	</param>
  },
  GetVisible: function() {
    /// <summary>Gets the visibility of the Title. </summary>
    ///	<returns type="Boolean">
    ///		Boolean: the visibility of the Title. 
    ///	</returns>
  },
  SetVisible: function(visibility) {
    /// <summary>Sets the visibility of the Title. </summary>
    ///	<param name="visibility" type="Boolean">
    ///		Boolean, the visibility of the Title.
    ///	</param>
  },
  GetValue: function() {
    /// <summary>Gets the value of the Title. </summary>
    ///	<returns type="String">
    ///		String: the value of the Title.
    ///	</returns>
  },
  SetValue: function(value) {
    /// <summary>Sets the value of the Title. </summary>
    ///	<param name="value" type="String">
    ///		String: the value of the Title.
    ///	</param>
  }
}

FarPoint.Web.Spread.Cell = function() {}

FarPoint.Web.Spread.Cell.prototype = {
  GetValue: function() {
    /// <summary>Gets the value of the Cell. </summary>
    ///	<returns type="String">
    ///		String: the value of the Cell.
    ///	</returns>
  },
  SetValue: function(value) {
    /// <summary>Sets the value of the Cell. </summary>
    ///	<param name="value" type="String">
    ///		String: the value of the Cell.
    ///	</param>
  },
  GetBackColor: function() {
    /// <summary>
    /// Gets the BackColor for the background color of the Cell. 
    /// </summary>
    ///	<returns type="String">
    ///		String: the BackColor of the Cell. 
    ///	</returns>
  },
  SetBackColor: function(color) {
    /// <summary>
    /// Sets the BackColor for the background color of the Cell. 
    ///</summary>
    ///	<param name="color" type="String">
    ///		String, the BackColor of the Cell.
    ///	</param>
  },
  GetForeColor: function() {
    /// <summary>
    /// Gets the ForeColor (TextColor) of the Cell.
    /// </summary>
    ///	<returns type="String">
    ///		String: the ForeColor of the Cell. 
    ///	</returns>
  },
  SetForeColor: function(color) {
    /// <summary>
    /// Sets the ForeColor (TextColor) of the Cell. 
    ///</summary>
    ///	<param name="color" type="String">
    ///		String, the ForeColor of the Cell.
    ///	</param>
  },
  GetTabStop: function() {
    /// <summary>Gets the TabStop setting of the Cell. </summary>
    ///	<returns type="Boolean">
    ///		Boolean: the tabstop setting of the Cell. 
    ///	</returns>
  },
  SetTabStop: function(IsStop) {
    /// <summary>Sets the TabStop setting of the Cell. </summary>
    ///	<param name="IsStop" type="Boolean">
    ///		Boolean, true if the user can give the focus to the cell using the TAB key; otherwise, false.
    ///	</param>
  },
  GetCellType: function() {
    /// <summary>Gets the cell type of the Cell. </summary>
    ///	<returns type="String">
    ///		Boolean: the cell type of the Cell. 
    ///	</returns>
  },
  //HorizontalAlign
  GetHAlign: function() {
    /// <summary>Gets the HorizontalAlign of the Cell. </summary>
    ///	<returns type="String">
    ///		String, the horizontal alignment of the Cell.
    ///	</returns>
  },
  SetHAlign: function(algin) {
    /// <summary>Sets the HorizontalAlign of the Celll. </summary>
    ///	<param name="algin" type="String">
    ///		String, the horizontal alignment of the Cell.
    ///	</param>
  },
  //VerticalAlign 
  GetVAlign: function() {
    /// <summary>Gets the VerticalAlign of the Cell. </summary>
    ///	<returns type="String">
    ///		Boolean: the VerticalAlign of the Cell. 
    ///	</returns>
  },
  SetVAlign: function(value) {
    /// <summary>Sets the VerticalAlign of the Cell. </summary>
    ///	<param name="value" type="String">
    ///		String, the VerticalAlign of the Cell.
    ///	</param>
  },
  GetLocked: function() {
    /// <summary>Gets the Locked of the Cell. </summary>
    ///	<returns type="Boolean">
    ///		Boolean: the Locked of the Cell. 
    ///	</returns>
  },
  GetFont_Name: function() {
    /// <summary>Gets the font name of the Cell. </summary>
    ///	<returns type="String">
    ///		String: the font name of the Cell. 
    ///	</returns>
  },
  SetFont_Name: function(name) {
    /// <summary>Sets the font name of the Cell. </summary>
    ///	<param name="name" type="String">
    ///		String, the font name of the Cell.
    ///	</param>
  },
  GetFont_Size: function() {
    /// <summary>Gets the font size of the Cell. </summary>
    ///	<returns type="String">
    ///		Boolean: the font size of the Cell. 
    ///	</returns>
  },
  SetFont_Size: function(size) {
    /// <summary>Sets the the font size of the Title. </summary>
    ///	<param name="size" type="String">
    ///		string, the font size of the Title.
    ///	</param>
  },
  GetFont_Bold: function() {
    /// <summary>
    /// Gets a value that indicates whether the font is bold for the Cell.
    /// </summary>
    ///	<returns type="Boolean">
    ///		Boolean: true if the font is bold; otherwise, false.
    ///	</returns>
  },
  SetFont_Bold: function(IsBold) {
    /// <summary>
    /// Sets a value that indicates whether the font is bold for the Cell.
    /// </summary>
    ///	<param name="IsBold" type="Boolean">
    ///		Boolean, true if the cell is bold; false otherwise.
    ///	</param>
  },
  GetFont_Italic: function() {
    /// <summary>
    /// Gets a value that indicates whether the font is italic for the Cell.
    /// </summary>
    ///	<returns type="Boolean">
    ///		Boolean: true if the font is italic; otherwise, false.
    ///	</returns>
  },
  SetFont_Italic: function(IsItalic) {
    /// <summary>
    /// Sets a value that indicates whether the font is italic for the Cell. 
    ///</summary>
    ///	<param name="IsItalic" type="Boolean">
    ///		Boolean: true if the font is italic; otherwise, false.
    ///	</param>
  },
  GetFont_Overline: function() {
    /// <summary>    
    /// Gets a value that indicates whether the font is overlined.
    /// </summary>
    ///	<returns type="Boolean">
    ///		Boolean: true if the font is italic; otherwise, false.
    ///	</returns>
  },
  SetFont_Overline: function(IsOverLine) {
    /// <summary>
    /// Sets a value that indicates whether the font is overlined.
    /// </summary>
    ///	<param name="IsOverLine" type="Boolean">
    /// Boolean: true if the fonit is overlined; otherwise, false.
    ///	</param>
  },
  GetFont_Strikeout: function() {
    /// <summary>
    /// Gets a value that indicates whether the font is strikethrough.
    /// </summary>
    ///	<returns type="Boolean">
    ///		Boolean: true if the font is struck throuth; otherwise, false.
    ///	</returns>
  },
  SetFont_Strikeout: function(IsStrikeout) {
    /// <summary>
    /// Sets a value that indicates whether the font is strikethrough.
    /// </summary>
    ///	<param name="IsStrikeout" type="Boolean">
    /// Boolean: true if the font is struck through; otherwise, false.
    ///	</param>
  },
  GetFont_Underline: function() {
    /// <summary>
    /// Gets a value that indicates whether the font is underlined.
    ///</summary>
    ///	<returns type="Boolean">
    ///		Boolean: true if the font is underlined; otherwise, false.
    ///	</returns>
  },
  SetFont_Underline: function(value) {
    /// <summary>
    /// Sets a value that indicates whether the font is underlined. </summary>
    ///	<param name="value" type="Boolean">
    ///		Boolean:  true if the font is underlined; otherwise, false.
    ///	</param>
  }
}

FarPoint.Web.Spread.Row = function() {}

FarPoint.Web.Spread.Row.prototype = {
  GetHeight: function() {
    /// <summary>
    /// Gets the height of the row.  
    /// </summary>
    ///	<returns type="Number" integer="true">
    ///		Integer: the height of the row. 
    ///	</returns>
  },
  SetHeight: function(height) {
    /// <summary>
    /// Sets the height of the row. 
    /// </summary>
    ///	<param name="height" type="Number" integer="true">
    ///		Integer, the height of the row. 
    ///	</param>
  }
}

FarPoint.Web.Spread.Column = function() {}

FarPoint.Web.Spread.Column.prototype = {
  GetWidth: function() {
    /// <summary>Gets the width of the column.  </summary>
    ///	<returns type="Number" integer="true">
    ///		Integer: the width of the column. 
    ///	</returns>
  },
  SetWidth: function(width) {
    /// <summary>Sets the width of the column.</summary>
    ///	<param name="width" type="Number" integer="true">
    ///		Integer, the width of the column. 
    ///	</param>
  }
}

FarPoint.Web.Spread.FpSpread = function() {}

FarPoint.Web.Spread.FpSpread.prototype = {  
  //start: <Methods>
  Add: function() {
    /// <summary>Inserts a new row at the end of the sheet. </summary>
  },
  AddKeyMap: function(keycode, ctrl, shift, alt, action) {
    /// <summary>
    /// This method allows you to map a keyboard key which will cause an action such as moving to the next or previous cell or row, or the first or last column. You can also scroll to a cell.
    /// </summary>
    ///	<param name="keycode" type="Number" integer="true">
    ///		Integer, key being pressed. 
    ///	</param>
    ///	<param name="ctrl" type="Boolean">
    ///		Boolean, Control key.
    ///	</param>
    ///	<param name="shift" type="Boolean">
    ///		Boolean, Shift key. 
    ///	</param>
    ///	<param name="alt" type="Boolean">
    ///		Boolean, Alt key.
    ///	</param>
    ///	<param name="action" type="String">
    ///		Constant, MoveToPrevCell, MoveToNextCell, MoveToNextRow, MoveToPrevRow, MoveToFirstColumn, MoveToLastColumn, ScrollTo.
    ///	</param>
    /// Remarks
    /// This method allows you to set a javascript code action based on a key the user presses. The action can move to the next or previous cell or row, first or last column, or scroll to a cell. 

    /// Example
    /// This is a sample that contains the method. On the client side, the script that contains the method would look like this:\r\n

    /// &lt;SCRIPT language=javascript&gt;
    ///   function setMap() { 
    ///       var ss = document.getElementById("FpSpread1"); 
    ///       if (ss != null){ 
    ///       ss.AddKeyMap(13,true,true,false,"this.MoveToLastColumn()");
    ///   }
    /// &lt;/SCRIPT&gt;
  },
  AddSelection: function(row, column, rowcount, columncount) {
    /// <summary>
    /// Adds a selection. This method is similar to selecting cells with the mouse.
    /// </summary>
    ///	<param name="row" type="Number" integer="true">
    ///		Integer, row index.
    ///	</param>
    ///	<param name="column" type="Number" integer="true">
    ///		Integer, column index.
    ///	</param>
    ///	<param name="rowcount" type="Number" integer="true">
    ///		Integer, number of rows.
    ///	</param>
    ///	<param name="columncount" type="Number" integer="true">
    ///		Integer, number of columns. 
    ///	</param>
  },
  CallBack: function(cmd) {
    /// <summary>
    /// Calls back to the ASPX page with the specified command.
    /// </summary>
    ///	<param name="cmd" type="String" >
    ///		String, the command to post back See the list of reserved commands below.
    ///	</param>
    /*
    command list:
    ActiveSpread: reserved for internal use for down-level browsers; makes a specified Spread component the active component 
    Add: adds a row and raises the event 
    Cancel: cancels the operation and raises the event 
    CellClick: clicks a cell using the left mouse button 
    ChildView: displays a specified child view when the HierarchicalView property is set to False and raises the event 
    ColumnHeaderClick:	same as click in column header and raises the event 
    Delete: deletes a row and raises the event 
    Edit:	same as Edit button and raises the event 
    ExpandView:	expands or collapses a specified row; raises the event; when row is expanded to display a child view, raises the ChildViewCreated event 
    Insert:	 adds a row at a particular place and raises the event 
    Next: 	moves to the next page on the sheet 
    Page:	 moves to a specified page on the sheet 
    Prev:	 moves to the previous page on the sheet 
    RowHeaderClick:	clicks a cell in the row header 
    Select:	 reserved for internal use for down-level browsers; selects a specified row 
    SelectView: moves to a specified sheet
    SortColumn: sorts a column 
    TabLeft: displays the previous sheet tabs to the left 
    TabRight:	 displays the next sheet tabs to the right 
    Update:	saves the changes 
    */
  },
  Cancel: function() {
    /// <summary>
    /// Cancels the changes since the most recent postback. 
    /// </summary>
  },
  Cells: function(r, c) {
    /// <summary>
    /// This method allows you to get or set various cell properties. 
    /// </summary>
    ///	<param name="r" type="Number" integer="true">
    ///		Integer, row index . 
    ///	</param>
    ///	<param name="c" type="Number" integer="true">
    ///		Integer, column index.
    ///	</param>
    /// <returns type="FarPoint.Web.Spread.Cell" />
  },
  Clear: function() {
    /// <summary>
    /// Clears the contents of the selected cells. 
    /// </summary>
  },
  ClearSelection: function() {
    /// <summary>
    /// Clears all the selections in the display.
    /// </summary>
  },
  Columns: function(c) {
    /// <summary>
    /// Gets a column dom object to allows you to get or set the width of the column.
    /// </summary>
    ///	<param name="c" type="Number" integer="true">
    ///		Integer, column index .
    ///	</param>
    /// <returns type="FarPoint.Web.Spread.Column" />
  },
  Copy: function() {
    /// <summary>
    /// Copies the contents of the selected cells to the Clipboard.
    /// </summary>
  },
  Delete: function() {
    /// <summary>
    /// Deletes the row with the active cell.
    /// </summary>
  },
  Edit: function(row) {
    /// <summary>
    /// Starts edit mode for one row when the edit template is enabled. 
    /// </summary>
    ///	<param name="row" type="Number" integer="true">
    ///		Integer, row index. 
    ///	</param>
  },
  EndEdit: function() {
    /// <summary>
    /// Ends edit mode for the cell being edited.
    /// </summary>
  },
  ExpandRow: function(row) {
    /// <summary>
    /// Expands or collapses the specified row.
    /// </summary>
    ///	<param name="row" type="Number" integer="true">
    ///		Integer, row index . 
    ///	</param>
  },
  GetActiveRow: function() {
    /// <summary>Gets the index of the row of the active cell. </summary>    
  },
  GetActiveCol: function() {
    /// <summary>Gets the index of the column of the active cell. </summary>
  },
  GetTopRowIndex: function(){
    /// <summary>Gets the current page scroll position top row index. </summary>
  },
  GetLeftColIndex: function() {
    /// <summary>Gets the current page scroll position left column index. </summary>
  },
  GetColByKey: function(key) {
    /// <summary>
    ///		Gets the client side column index for the specified server side column index.
    /// </summary>
    ///	<param name="key" type="Number" integer="true">
    ///		Integer, index of the column key.
    ///	</param>
    ///	<returns type="Integer">
    ///		the index of the column.
    ///	</returns>
  },
  GetColKeyFromCol : function(column){
    /// <summary>
    ///		Gets the server side column index for the specified client side column index.
    /// </summary>
    ///	<param name="column" type="Number" integer="true">
    ///		Integer, index of the column.
    ///	</param>
    ///	<returns type="Integer">
    ///		the index of the column key.
    ///	</returns>
  },
  GetRowByKey: function(key) {
    /// <summary>
    ///		Gets the client side row index for the specified server side row index.
    /// </summary>
    ///	<param name="key" type="Number" integer="true">
    ///		Integer, index of the row key.
    ///	</param>
    ///	<returns type="Integer">
    ///		the index of the row.
    ///	</returns>
  },
  GetRowKeyFromRow: function(row) {
    /// <summary>
    ///		Gets the server side row index for the specified client side row index.
    /// </summary>
    ///	<param name="row" type="Number" integer="true">
    ///		Integer, index of the row.
    ///	</param>
    ///	<returns type="Integer">
    ///		the index of the row key.
    ///	</returns>
  },  
  GetActiveChildSheetView: function() {
    /// <summary>
    /// Gets the sheet view object of the active child sheet in hierarchical mode. 
    /// </summary>
    ///	<returns type="FarPoint.Web.Spread.FpSpread">
    ///		 the sheet view object of the active child sheet in hierarchical mode.
    ///	</returns>
  },
  GetCellByRowCol: function(row, column) {
    /// <summary>
    /// Gets the table cell for the specified row and column. 
    /// </summary>
    ///	<param name="row" type="Number" integer="true">
    ///		Integer, row index of the specified cell. 
    ///	</param>
    ///	<param name="column" type="Number" integer="true">
    ///		Integer, column index of the specified cell .
    ///	</param>
    ///	<returns type="Element">
    ///		 TD element (HTML element).
    ///	</returns>
  },
  GetChildSpread: function(row, relation) {
    /// <summary>
    /// Gets the child FarPoint Spread object on the client of the specified relation at the specified row. 
    /// </summary>
    ///	<param name="row" type="Number" integer="true">
    ///		Integer, row index . 
    ///	</param>
    ///	<param name="relation" type="Number">
    ///		Integer, relation index .
    ///	</param>
    ///	<returns type="FarPoint.Web.Spread.FpSpread" />
  },
  GetChildSpreads: function() {
    /// <summary>
    /// Gets an array of the child FarPoint Spread objects of the displayed page on the client.
    /// </summary>
    ///	<returns type="Array">
    /// Array of FarPoint Spread objects (HTML elements) 
    /// </returns>
  },
  GetColCount: function() {
    /// <summary>
    /// Gets the number of columns in the displayed page. 
    /// </summary>
    ///	<returns type="Number" integer="true">
    /// Integer, number of columns. 
    /// </returns>
  },
  GetFormula: function(row, column) {
    /// <summary>
    /// Gets the formula in the cell of the specified row and column. 
    /// </summary>
    ///	<param name="row" type="Number" integer="true">
    ///		Integer, row index of cell. 
    ///	</param>
    ///	<param name="column" type="Number" integer="true">
    ///		Integer, column index of cell.
    ///	</param>
    ///	<returns type="String">
    /// String with the formula in the specified cell or an empty string if no formula is specified.
    /// </returns>
  },
  GetHiddenValue: function(row, columnName) {
    /// <summary>
    /// Gets the value of a programmatically hidden cell at the specified row and column.
    /// </summary>
    ///	<param name="row" type="Number" integer="true">
    ///		Integer, row index . 
    ///	</param>
    ///	<param name="columnName" type="String">
    ///		String, name of the column .
    ///	</param>
    ///	<returns type="String">
    /// String, with the value of the cell. 
    /// </returns>
  },
  GetPageCount: function() {
    /// <summary>
    /// Gets the number of pages. 
    /// </summary>
    ///	<returns type="Number" integer="true">
    /// Integer, a count of pages. 
    /// </returns>
  },
  IsValid: function() {
    /// <summary>
    /// Gets current Spread is valid.
    /// </summary>
    ///	<returns type="Boolean">
    ///		Boolean: true if Spread is valid; otherwise, false.
    ///	</returns>
  },
  GetParentRowIndex: function() {
    /// <summary>
    /// Gets the row index of the parent FarPoint Spread object of the displayed FarPoint Spread object.
    /// </summary>
    ///	<returns type="Number" integer="true">
    /// Integer, index of row of the parent FarPoint Spread object. 
    /// </returns>
  },
  GetCurrentPageIndex: function() {
    /// <summary>
    /// Gets the page index of the current displayed FarPoint Spread object.
    /// </summary>
    ///	<returns type="Number" integer="true">
    /// Integer, index of the page of current FarPoint Spread object.
    /// </returns>
  },
  GetParentSpread: function() {
    /// <summary>
    /// Gets the parent FarPoint Spread object of the displayed FarPoint Spread object.
    /// </summary>
    ///	<returns type="FarPoint.Web.Spread.FpSpread">
    /// FarPoint Spread object (HTML element). 
    /// </returns>
  },
  GetRowCount: function() {
    /// <summary>
    /// Gets the number of rows in the displayed page on the client. 
    /// </summary>
    ///	<returns type="Number" integer="true">
    /// Integer, number of rows. 
    /// </returns>
  },
  GetSelectedRange: function() {
    /// <summary>
    /// Gets the range of cells that are selected on the displayed page.
    /// </summary>
    ///	<returns type="FarPoint.Web.Spread.Range">
    /// The returned value is a JavaScript object with the following properties: type, row, column, rowCount, and columnCount. The row and column properties are the indexes of the starting cell in the range. The rowCount and columnCount are the number of rows and columns in the range. The type property, which specifies the type of range, can be one of the following: Cell (if range is a range of cells but not entire row or column), Row (if range is an entire row or multiple rows) or Column (if range is entire column or multiple columns).
    /// </returns>
  },
  GetSelectedRanges: function() {
    /// <summary>
    /// Gets the ranges of cells that are selected on the displayed page.
    /// </summary>
    ///	<returns type="Array">
    /// The return type is an array of selected ranges. The returned value is a JavaScript object with the following properties: type, row, column, rowCount, and columnCount. The row and column properties are the indexes of the starting cell in the range. The rowCount and columnCount are the number of rows and columns in the range. The type property, which specifies the type of range, can be one of the following: Cell (if range is a range of cells but not entire row or column), Row (if range is an entire row or multiple rows) or Column (if range is entire column or multiple columns).
    /// </returns>
  },
  GetSheetColIndex: function(column) {
    /// <summary>
    /// Gets the column index of the SheetView object on the server for the specified column of the FarPoint Spread object on the displayed page on the client.
    /// </summary>
    ///	<param name="column" type="Number" integer="true">
    ///		Integer, column index (on the client) . 
    ///	</param>
    ///	<returns type="Number" integer="true">
    /// Integer, index of the column in the sheet (SheetView object on the server)
    /// </returns>
  },
  GetSheetRowIndex: function(row) {
    /// <summary>
    /// Gets the row index of the SheetView object on the server for the specified row of the FarPoint Spread displayed on the client. 
    /// </summary>
    ///	<param name="row" type="Number" integer="true">
    ///		Integer, row index (on the client) . 
    ///	</param>
    ///	<returns type="Number" integer="true">
    /// Integer, index of the row in the sheet (SheetView object on the server). 
    /// </returns>
  },
  GetSpread: function(element) {
    /// <summary>
    /// Gets the FarPoint Spread object (an HTML element) on the client that contains the specific HTML element. If there is no FarPoint Spread object, the method returns Null. 
    /// </summary>
    ///	<param name="element" type="Element">
    ///		HTML element . 
    ///	</param>
    ///	<returns type="FarPoint.Web.Spread.FpSpread">
    /// FarPoint Spread object (HTML element) . 
    /// </returns>
  },
  GetTitleInfo: function() {
    /// <summary>
    /// Gets the TitleInfo JavaScript object of the Spread component, the subtile or other properties of the TitleInfo class.
    /// </summary>
    ///	<returns type="FarPoint.Web.Spread.TitleInfo" />
  },
  GetTotalRowCount: function() {
    /// <summary>
    /// Gets the total number of rows for the active sheet. If the AllowPage property is true, this method may return more rows than the GetRowCount method. 
    /// </summary>
    ///	<returns type="Number" integer="true">
    /// Integer, the total number of rows for the active sheet. 
    /// </returns>
  },
  GetValue: function(row, column) {
    /// <summary>
    /// Gets the value of a cell. The row and column indexes must be valid.
    /// </summary>
    ///	<param name="row" type="Number" integer="true">
    ///		Integer, row index . 
    ///	</param>
    ///	<param name="column" type="Number" integer="true">
    ///		Integer, column index.
    ///	</param>
    ///	<returns type="String">
    /// String, the value of the cell. 
    /// </returns>
  },
  HideMessage: function(row, column) {
    /// <summary>
    /// Hides the error message from the ShowMessage method. 
    /// </summary>
    ///	<param name="row" type="Number" integer="true">
    ///		Integer, row index . 
    ///	</param>
    ///	<param name="column" type="Number" integer="true">
    ///		Integer, column index . 
    ///	</param>
  },
  Insert: function() {
    /// <summary>
    /// Inserts a new row before the row with the active cell.
    /// It is similar to pressing the Insert button on the command bar. 
    /// This method causes a postback to occur. 
    /// When this method is called, the Insert event is rasied on the server. 
    /// </summary>
  },
  MoveToFirstColumn: function() {
    /// <summary>
    /// Moves the active cell to the first column. 
    /// </summary>
  },
  MoveToLastColumn: function() {
    /// <summary>
    /// Moves the active cell to the last column. 
    /// </summary>
  },
  MoveToNextCell: function() {
    /// <summary>
    /// Moves the active cell to the next cell.
    /// </summary>
  },
  MoveToNextRow: function(wrap) {
    /// <summary>
    /// Moves the active cell to the next row.
    /// </summary>
    ///	<param name="wrap" type="Boolean">
    ///		Boolean, true if wrap while move to next row; false otherwise.
    ///	</param>
  },
  MoveToPrevCell: function() {
    /// <summary>
    /// Moves the active cell to the previouse cell.
    /// </summary>
  },
  MoveToPrevRow: function(wrap) {
    /// <summary>
    /// Moves the active cell to the previouse row.
    /// </summary>
    ///	<param name="wrap" type="Boolean">
    ///		Boolean, true if wrap while move to previous row; false otherwise.
    ///	</param>
  },
  Next: function() {
    /// <summary>
    /// Changes the display to show the next page.
    /// It is similar to pressing the Next button on the command bar or the other page navigation aids; 
    /// it advances one page to display the next set of rows. 
    /// If there is not a next page, the FarPoint Spread stays on the current page. 
    /// This method causes a postback to occur. 
    /// When this method is called, the TopRowChanged event is raised on the server. 
    /// </summary>
  },
  Paste: function() {
    /// <summary>
    /// Pastes the Clipboard contents to the cells.
    /// </summary>
  },
  Prev: function() {
    /// <summary>
    /// Changes the display to show the previous page.
    /// It is similar to pressing the Previous button on the command bar or the other page nativation aids;
    /// it goes back one page to display the set of rows on the previous page.
    /// If there is not a previouse page, the FarPoint Spread stays on the current page.
    /// This method causes a postback to occur.
    /// When this method is called, the TopRowChanged event is fired on the server.
    /// </summary>
  },
  Print: function() {
    /// <summary>
    /// Prints the sheet.
    /// </summary>
  },
  RemoveKeyMap: function(keycode, ctrl, shift, alt, action) {
  /// <summary>
  /// Remove the KeyMap of spread component.
  /// </summary>
  ///	<param name="keycode" type="Number" integer="true">
  ///		Integer, key being pressed. 
  ///	</param>
  ///	<param name="ctrl" type="Boolean">
  ///		Boolean, Control key.
  ///	</param>
  ///	<param name="ctrl" type="Boolean">
  ///		Boolean, Shift key. 
  ///	</param>
  ///	<param name="alt" type="Boolean">
  ///		Boolean, Alt key.
  ///	</param>
  ///	<param name="action" type="String">
  ///		Constant, MoveToPrevCell, MoveToNextCell, MoveToNextRow, MoveToPrevRow, MoveToFirstColumn, MoveToLastColumn, ScrollTo.
  ///	</param>
  },
  Rows: function(row) {
    /// <summary>
    ///  Gets a row JavaScript Object to get or set the height of the row. 
    /// </summary>
    ///	<param name="row" type="Number" integer="true">
    ///		Integer, row index of row.
    ///	</param>
    /// <returns type="FarPoint.Web.Spread.Row" />
  },
  ScrollTo: function(row, column) {
    /// <summary>
    /// Moves the active cell to a specified cell.
    /// </summary>
    ///	<param name="row" type="Number" integer="true">
    ///		Integer, row index of row. 
    ///	</param>
    ///	<param name="column" type="Number" integer="true">
    ///		Integer, column index of column. 
    ///	</param>
  },
  SetActiveCell: function(row, column) {
    /// <summary>
    /// Sets the active cell to the cell at the specified row and column.
    /// Both row and column indexes must be valid.
    /// </summary>
    ///	<param name="row" type="Number" integer="true">
    ///		Integer, row index of the specified cell. 
    ///	</param>
    ///	<param name="column" type="Number" integer="true">
    ///		Integer, column index of the specified cell. 
    ///	</param>
  },
  setColWidth: function(column, width) {
    /// <summary>
    /// Sets a column to a specified width. 
    /// This method sets the specified column to the specified width. 
    /// The column index must be valid. Remember that the column index is zero-based, so the first column is 0. 
    /// This method can be used to hide a column by setting the width to zero.     
    /// </summary>
    ///	<param name="column" type="Number" integer="true">
    ///		Integer, column index . 
    ///	</param>    
    ///	<param name="width" type="Number" integer="true">
    ///		Integer, number of pixels in width of column . 
    ///	</param>

  },
  SetFormula: function(row, column, value) {
    /// <summary>
    /// Sets the formula in a cell at the specified row and column. 
    /// Both row and column indexes must be integer values and must be valid. 
    /// The formula is parsed and evaluated when the data is posted back to the server. 
    /// This method does not cause a postback to occur. 
    /// For a list of the operators and functions that appear in formulas, refer to the FarPoint Spread for .NET Formula Reference. 
    /// Be sure to start the formula with an equals sign (=). 
    /// </summary>
    ///	<param name="row" type="Number" integer="true">
    ///		Integer, row index of cell. 
    ///	</param>
    ///	<param name="column" type="Number" integer="true">
    ///		Integer, column index of cell . 
    ///	</param>
    ///	<param name="value" type="String">
    ///		String, the formula to put in the cell.
    ///	</param>
  },
  SetSelectedRange: function(row, column, rowCount, columnCount) {
    /// <summary>
    /// Sets the specified range of cells as selected.
    /// </summary>
    ///	<param name="row" type="Number" integer="true">
    ///		Integer, starting row index. 
    ///	</param>
    ///	<param name="column" type="Number" integer="true">
    ///		Integer, starting column index. 
    ///	</param>
    ///	<param name="rowCount" type="Number" integer="true">
    ///		Integer, number of rows in selection . 
    ///	</param>
    ///	<param name="columnCount" type="Number" integer="true">
    ///		Integer, number of columns in selection. 
    ///	</param>
  },
  SetValue: function(row, column, value, noEvent) {
    /// <summary>
    /// Sets the value of a cell at the specified row and column and triggers the onDataChanged event if specified. 
    /// If noEvent is true, the method does not trigger an event. 
    /// Otherwise, the method triggers a onDataChanged event. 
    /// </summary>
    ///	<param name="row" type="Number" integer="true">
    ///		Integer, row index . 
    ///	</param>
    ///	<param name="column" type="Number" integer="true">
    ///		Integer, column index .
    ///	</param>
    ///	<param name="value" type="String">
    ///		String, value . 
    ///	</param>
    ///	<param name="noEvent" type="Boolean">
    ///		Boolean, whether to trigger the onDataChanged event .
    ///	</param>
  },
  ShowMessage: function(message, row, column, time) {
    /// <summary>
    /// Shows an error message. 
    /// </summary>
    ///	<param name="message" type="String">
    ///		String, message text .
    ///	</param>
    ///	<param name="row" type="Number" integer="true">
    ///		Integer, row index . 
    ///	</param>
    ///	<param name="column" type="Number" integer="true">
    ///		Integer, column index . 
    ///	</param>
    ///	<param name="time" type="Number" integer="true">
    ///		Integer, time. 
    ///	</param>
  },
  SizeToFit: function() {
    /// <summary>
    /// Sets the column width to the size of the largest text in the column. 
    /// </summary>
  },
  SortColumn: function(column) {
    /// <summary>
    /// Sorts a specified column. 
    /// The AllowSort property must be set to True for the sheet. 
    /// </summary>
    ///	<param name="column" type="Number" integer="true">
    ///		Integer, column index. 
    ///	</param>
  },
  StartEdit: function(cell) {
    /// <summary>
    /// Puts a cell into edit mode to allow editing the cell. 
    /// </summary>
    ///	<param name="cell" type="Element">
    ///		Table element for the cell. 
    ///	</param>
  },
  Update: function() {
    /// <summary>
    /// Saves the changes to the FarPoint Spread object on the client back to the server.
    /// </summary>
  },
  UpdatePostbackData: function() {
    /// <summary>
    /// Saves the changes to the postback data.
    /// </summary>
  },
  addEventListener: function(event, handler, useCapture) {
    /// <summary>
    /// Add an event handler for the specified event.
    /// </summary>
  },

  removeEventListener: function(event, handler, useCapture) {
    /// <summary>
    /// Remove the specified event handler.
    /// </summary>
  },
  LoadData: function(asyncCallBack) {
    /// <summary>
    /// Gets next block of rows from server. 
    /// </summary>
    ///	<param name="asyncCallBack" type="Boolean">
    ///		Boolean, whether the callback is asynchronous. 
    ///	</param>
  },
  
  SuspendLayout: function() {
    /// <summary>
    /// Temporarily suspends the layout logic for the spread.
    /// </summary>
  },
  
  ResumeLayout: function(performLayout) {
    /// <summary>
    /// Resumes normal layout logic and optionally forces the component to apply layout logic to its child controls.
    /// </summary>
    ///	<param name="performLayout" type="Boolean">
    ///		Boolean, Whether to perform layout logic on the child controls.
    ///	</param>
  },
  // end: <Methods />

  ActiveCol: {
  },

  ActiveRow: {
  }
}

function FpSpread(spread) {
  ///	<summary>
  ///		Gets the FpSpread that support JavaScript Intellisence in desgin time.
  ///	</summary>
  ///	<param name="spread" type="String">
  ///		1: string - the id of the spread.
  ///		2: Dom Object - the dom object of the spread.
  ///	</param>
  ///	<returns type="FarPoint.Web.Spread.FpSpread" />
  //spread = document.getElementById("FpSpread1");
}


