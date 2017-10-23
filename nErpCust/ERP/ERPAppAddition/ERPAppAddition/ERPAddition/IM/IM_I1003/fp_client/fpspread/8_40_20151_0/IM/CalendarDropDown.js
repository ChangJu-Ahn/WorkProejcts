//
//	  Copyright (C) 2003-2015 GrapeCity Inc.	All rights reserved.
//

var FarPoint;(function(a){var b;(function(a){var b;(function(b){var a=function(){function a(a){var b=this;this.calendar=null;this.show=false;this.imageList=[];this.element=a;a.Toggle=function(a,c){b.Toggle(a,c)};a.Show=function(a){b.Show(a)};a.Hide=function(){b.Hide()};a.OnKeyDown=function(a,c){b.OnKeyDown(a,c)};a.Initialized=true}a.prototype.OnKeyDown=function(c,a){if(!this.calendar)return;this.ddController=c;if(a.target==this.ddController.editor.element||$$.j.contains(this.ddController.editor.element,a.target)){var b=GCIMCalendar.Key;if(a.keyCode==b.Escape||a.keyCode==b.Return){if(a.keyCode==b.Return)this.ddController.SetValueAndClose(this.calendar.GetSelectedDate());else this.ddController.editor.input.focus();this.Hide();$$.utils.cancelBubble(a,true)}else if(this.calendar.UIProcess.onKeyDownHandler(a)){GCIMCalendar.Utility.PreventDefault(a);GCIMCalendar.Utility.CancelBubble(a)}}};a.prototype.InitCalendar=function(c){var a=this;if(this.calendar==null){var b=GCIMCalendar;this.ddController=c;this.InitImageList();this.calendar=new b.GcCalendar(this.element,this.imageList);this.LoadCalendarSetting();this.OverrideRealSetFocus();this.SyncCulture();this.calendar.OnClickDate(function(){var e=a.calendar.GetSelectedDate();if(a.calendar.GetCalendarType()===b.CalendarType.YearMonth){var d=c.GetValue();if(d==null)d=new Date;e.setDate(d.getDate())}a.ddController.SetValueAndClose(e);a.Hide()});this.calendar.OnScrolled(function(){setTimeout(function(){var b=$$.j(a.ddController.editor.element);b.css("display")!="none"&&b.focus()},800)})}};a.prototype.SyncCulture=function(){var b=this.ddController.container.getAttribute("hostspreadid"),a=document.getElementById(b);if(a){var c=a.getAttribute("culture");this.calendar.SetSpreadCulture(c)}};a.prototype.OverrideRealSetFocus=function(){var a=GCIMCalendar;if(a.GcCalendar.prototype._oldRealSetFocus)return;a.GcCalendar.prototype._oldRealSetFocus=a.GcCalendar.prototype._realSetFocus;a.GcCalendar.prototype._realSetFocus=function(a){$$.j(this.Render.CalendarSectionDom.OutterContainerDiv).is(":visible")&&this._oldRealSetFocus(a);return this}};a.prototype.LoadCalendarSetting=function(){this.calendar.SuspendLayout();var f=this.element.getAttribute("IMCalendarType");f&&this.calendar.SetCalendarType(f);var j=this.element.getAttribute("IMBackColor");j&&this.calendar.SetBackColor(j);var k=this.element.getAttribute("IMForeColor");k&&this.calendar.SetForeColor(k);var b=this.element.getAttribute("IMHeaderBackColor");b&&this.calendar.SetHeaderBackColor(b);var c=this.element.getAttribute("IMHeaderForeColor");c&&this.calendar.SetHeaderForeColor(c);var e=this.element.getAttribute("IMTodayMarkColor");e&&this.calendar.SetTodayMarkColor(e);var n=this.element.getAttribute("IMShowToday");n&&this.calendar.SetShowToday(true);var g=this.element.getAttribute("IMHeaderFormat");g&&this.calendar.SetHeaderFormat(g);var d=this.element.getAttribute("IMYearMonthFormat");d&&this.calendar.SetYearMonthFormat(d);var h=this.element.getAttribute("IMFontFamily");h&&this.calendar.SetFontFamily(h);var m=this.element.getAttribute("IMFontSize");m&&this.calendar.SetFontSize(m);var l=this.element.getAttribute("IMFontBold");l&&this.calendar.SetFontWeight(l);var i=this.element.getAttribute("IMFontItalic");i&&this.calendar.SetFontStyle(i);var a=this.element.getAttribute("IMTextDecoration");a&&this.calendar.SetTextDecoration(a);this.calendar.ResumeLayout()};a.prototype.SyncImageList=function(){this.InitImageList();for(var a=0;a<this.imageList.length;a++){var b=this.element.getAttribute(this.imageList[a]);this.calendar.SetServerImagePath(this.imageList[a],b)}this.calendar.UpdateControl()};a.prototype.InitImageList=function(){this.imageList.push("Calendar_OutlookArrow_Left_Chrome_Normal_Svg");this.imageList.push("Calendar_OutlookArrow_Left_Ipad_Normal_Svg");this.imageList.push("Calendar_OutlookArrow_Left_Normal_Normal_Svg");this.imageList.push("Calendar_OutlookArrow_Left_Vista_Normal_Svg");this.imageList.push("Calendar_OutlookArrow_Left_Vista_Normal_Gif");this.imageList.push("Calendar_OutlookArrow_Left_Win8_Normal_Svg");this.imageList.push("Calendar_OutlookArrow_Left_Win8_Normal_Gif");this.imageList.push("Calendar_OutlookArrow_Right_Chrome_Normal_Svg");this.imageList.push("Calendar_OutlookArrow_Right_Ipad_Normal_Svg");this.imageList.push("Calendar_OutlookArrow_Right_Normal_Normal_Svg");this.imageList.push("Calendar_OutlookArrow_Right_Vista_Normal_Svg");this.imageList.push("Calendar_OutlookArrow_Right_Vista_Normal_Gif");this.imageList.push("Calendar_OutlookArrow_Right_Win8_Normal_Svg");this.imageList.push("Calendar_OutlookArrow_Right_Win8_Normal_Gif");this.imageList.push("Calendar_ZoomButton_ZoomIn_Chrome_Normal_Svg");this.imageList.push("Calendar_ZoomButton_ZoomIn_Ipad_Normal_Svg");this.imageList.push("Calendar_ZoomButton_ZoomIn_Normal_Normal_Svg");this.imageList.push("Calendar_ZoomButton_ZoomIn_Normal_White_Normal_Svg");this.imageList.push("Calendar_ZoomButton_ZoomIn_Normal_White_Normal_Gif");this.imageList.push("Calendar_ZoomButton_ZoomIn_Vista_Normal_Gif");this.imageList.push("Calendar_ZoomButton_ZoomIn_Vista_Normal_Svg");this.imageList.push("Calendar_ZoomButton_ZoomIn_Win8_Normal_Svg");this.imageList.push("Calendar_ZoomButton_ZoomIn_Win8_Normal_Gif");this.imageList.push("Calendar_ZoomButton_ZoomOut_Chrome_Normal_Svg");this.imageList.push("Calendar_ZoomButton_ZoomOut_Ipad_Normal_Svg");this.imageList.push("Calendar_ZoomButton_ZoomOut_Normal_Normal_Svg");this.imageList.push("Calendar_ZoomButton_ZoomOut_Normal_White_Normal_Svg");this.imageList.push("Calendar_ZoomButton_ZoomOut_Normal_White_Normal_Gif");this.imageList.push("Calendar_ZoomButton_ZoomOut_Vista_Normal_Gif");this.imageList.push("Calendar_ZoomButton_ZoomOut_Vista_Normal_Svg");this.imageList.push("Calendar_ZoomButton_ZoomOut_Win8_Normal_Svg");this.imageList.push("Calendar_ZoomButton_ZoomOut_Win8_Normal_Gif");this.imageList.push("Today_Gif");this.imageList.push("Today_Svg")};a.prototype.RefreshCalendar=function(b,d){this.calendar==null&&this.InitCalendar(b);this.ddController=b;if(d){var c=b.container;c!=this.element.parentNode&&c.appendChild(this.element);var a=b.GetValue();if(a==null)a=new Date;if(a instanceof Date){this.calendar.SetSelectedDate(a);this.calendar.SetFocusDate(a)}}};a.prototype.Toggle=function(a,b){if(!a.editor.IsDropDownControlShowing()){this.ddController=a;this.Show(b)}else this.Hide()};a.prototype.Show=function(a){if(this.ddController){if(a)GCIMCalendar.Utility.SetZoomStyle(this.element,"1.5");else GCIMCalendar.Utility.SetZoomStyle(this.element,"");this.RefreshCalendar(this.ddController,true);this.ddController.Show();this.ddController.editor.element.focus()}};a.prototype.Hide=function(){if(this.ddController){this.ddController.Hide();this.calendar&&this.calendar._dispose()}};a.Initialize=function(b){return b.Initialized?void 0:new a(b)};return a}();b.CalendarDropDown=a})(b=a.Spread||(a.Spread={}))})(b=a.Web||(a.Web={}))})(FarPoint||(FarPoint={}))