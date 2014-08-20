<% 
If Session("LoggedIn") <> True Then 
	Response.Redirect("dispatchlogin.asp") 
End If 
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01//EN" "http://www.w3.org/TR/html4/strict.dtd">
<html xmlns="http://www.w3.org/1999/xhtml"> 
	<head>
		<title>Technician Dispatch Board</title>
		<style type="text/css">
			html, body{
				 height: 100%;
				 margin: 0;
				 padding: 5px;
				 border: none;
				 font-size: 14px;
			} 
			
			div.dispatchDay{
				margin-bottom: 20px;
			}
			
			table.dispatch{
				width: 100%;
				margin-bottom: 10px;
				border-top: 1px solid #000000;
				font-size: 12px;
			}
			
			table.assembly{
				width: 100%;
				margin-bottom: 10px;
				font-size: 12px;
			}
			
			table.today{
				width: 400px;
			}
			
			td.partslist{
				color: #15428B;
			}
			
			h1.dispatchDay{
				margin-bottom: 0px;
				background-color: #C1D3EA;
				font-size: 15px;
				font-weight: bold;
			}
			
		</style>
	</head>
	<body>
<%


' ******* CONNECT TO CRM SQL DATABASE ********
'SERVER IP: 192.168.0.##
connString = "Provider=SQLNCLI10;Data Source=192.168.0.##\INSTANCENAME;Initial Catalog=DATABASENAME;User ID=SA;Password=********"
set Conn = Server.CreateObject("ADODB.Connection")
Conn.ConnectionString = connString
Conn.Open

' Retrieve the Rep Group view. If there isn't one selected, default to the Technicians.
repGroup = Request.Querystring("repgroup")
if(repGroup = "") then
	repGroup = "Technicians"
end if

' User can choose to display just the day or the entire week. Build links to access other views.
Response.Write ("Change Rep Group: <a href=""dispatchboard.asp?display=week&repgroup=Technicians"">Technicians</a> | <a href=""dispatchboard.asp?display=week&repgroup=Operations"">Operations</a><br/>")
Response.Write ("Display Type: <a href=""dispatchboard.asp?display=day&repgroup=" & repGroup & """>Today</a> | <a href=""dispatchboard.asp?display=week&repgroup=" & repGroup & """>This Week</a> | <a href=""dispatchboard.asp?display=nextweek&repgroup=" & repGroup & """>Next Week</a>")

' Retrieve display type, default to the weekly view if it is empty
displayType = Request.Querystring("display")
if(displayType = "") then
	displayType = "week"
end if

' If the display type is week or next week, we need to print out all tasks for all 7 days of the week.
if (displayType = "week" OR displayType = "nextweek") then
	' If the display type is this week, we'll choose today as the base date.
	if (displayType = "week") then
		baseDate = Date
	' If the display type is next week, we'll add a week to today's date.
	else
		weekFromToday = DateAdd("ww", 1, Date)
		selectedWeekday = Weekday(weekFromToday, vbMonday)
		baseDate = DateAdd("d", -(selectedWeekday - 1), weekFromToday)
	end if
%>
		<br/>
		<table width="100%">
			<tr>
				<td colspan="3" align="center">
					<h1>Dispatch Board for the week of 
<%
	' Reiterating to the technician which display method they are looking at
	Response.Write (baseDate)
%>
					</h1>
				</td>
			</tr>
<%

	' Our calendar starts on Monday. This is to convert the date into a weekday number rather than the full date with the Monday start in mind.
	selectedWeekday = Weekday(baseDate, vbMonday)

	Dim arrayDays(7)
	Dim arrayDayStrings(7)
	Dim query
	Dim i

	' Need to subtract days from the current weekday number to get to Monday.
	arrayDays(0) = DateAdd("d", -(selectedWeekday - 1), baseDate) ' Monday
	arrayDays(1) = DateAdd("d", 1, arrayDays(0)) ' Tuesday
	arrayDays(2) = DateAdd("d", 2, arrayDays(0)) ' Wednesday
	arrayDays(3) = DateAdd("d", 3, arrayDays(0)) ' Thursday
	arrayDays(4) = DateAdd("d", 4, arrayDays(0)) ' Friday
	arrayDays(5) = DateAdd("d", 5, arrayDays(0)) ' Saturday
	arrayDays(6) = DateAdd("d", 6, arrayDays(0)) ' Sunday

	' Strings for all of the corresponding day names
	arrayDayStrings(0) = "Monday"
	arrayDayStrings(1) = "Tuesday"
	arrayDayStrings(2) = "Wednesday"
	arrayDayStrings(3) = "Thursday"
	arrayDayStrings(4) = "Friday"
	arrayDayStrings(5) = "Saturday"
	arrayDayStrings(6) = "Sunday"


	set rs = Server.CreateObject("ADODB.recordset")

	''' BUILD THE WEEKLY TASK DISPLAY TABLE
	'''
	''' The weekly display table is a 3 x 3 grid with Saturday and Sunday crammed together.
	'''
	''' The current implementation of this is not ideal. It was a temporary solution and can be greatly improved and made to be more dynamic.
	'''

	For i = 0 to UBound(arrayDays) - 1
		' If the day is Monday or Thursday, we start a new row
		if (i = 0 OR i = 3) then
	%>
			<tr>
	<%
		end if

		' If the day isn't Sunday, start a new table cell (Sunday gets put into Saturday's cell)
		if (i <> 6) then
	%>	
				<td valign="top" width="33%">
					<div class="dispatchDay">
	<%	
		end if
		
		' Day and Date title for daily task lists
	%>
		<h1 class="dispatchDay">
	<%
		Response.Write (arrayDayStrings(i) & ", " & arrayDays(i))
	%>
						</h1>
						<br/>
	<%
		query = buildRecordSetQuery(arrayDays(i), repGroup)


		rs.Open query, Conn
		' Display table for date, passing the recordset and date to the subroutine.
		displayDispatchTask rs, arrayDays(i)
		rs.Close
		
		' If the day is Saturday, we're not ready to close the cell
		if (i <> 5) then
	%>
					</div>
				</td>
	<%
		end if

		' If the day is Wednesday or Sunday, we close the row
		if (i = 2 OR i = 6) then
	%>
			</tr>
	<%
		end if
		
		' If the day is Sunday, we close the whole table
		if (i = 6) then
	%>
			</table>
	<%
		end if
	Next
	''' END THE WEEKLY TASK DISPLAY TABLE

' If the display type is today, we only need to print out the tasks for today.
elseif(displayType = "day") then
	' The day version of the calendar switches at 9pm because tasks are not scheduled that late and it enables them to see what is on
	' the agenda for tomorrow before going to sleep.
	if ((Time > Cdate("21:00:00")) AND (Time < Cdate("23:59:59"))) then
		currentDay = DateAdd("d", 1, Date)
	Else
		currentDay = Date
	End if
	
	set rs = Server.CreateObject("ADODB.recordset")
	query = buildRecordSetQuery(currentDay, repGroup)
	rs.Open query, Conn
	
	''' BUILD TODAY's TASK DISPLAY TABLE
	'''
%>
	<table class="today">
		<tr>
			<td colspan="3" align="center"><h1>Dispatch Board for <% Response.Write currentDay %>, Current Time: <% Response.Write Time %><br/>(day switches at 9pm)<br/></h1>
<%
	displayDispatchTask rs, currentDay
	rs.Close
%>
			</td>
		</tr>
	</table>
<%

' If the display type is parts, we need to print out a list of the service order parts
elseif(displayType = "parts") then
	soNumber = Request.Querystring("so")
	
	if(isNumeric(soNumber)) then
%>

	<br/><br/>Parts for SO <% Response.Write soNumber %>
	<table class="dispatch">
	
<%		
		set rs = Server.CreateObject("ADODB.recordset")
		soPartsQuery = "SELECT tblSOPartsUsed.SONumber, tblSOPartsUsed.ItemID, tblSOPartsUsed.ItemDescription, tblSOPartsUsed.Type, tblSOPartsUsed.Quantity, tblSOPartsUsed.SOPartsUsedKeyID FROM tblSOPartsUsed WHERE tblSOPartsUsed.SONumber = " & soNumber
		rs.Open soPartsQuery, Conn
		
		if NOT rs.EOF then
			rs.MoveFirst
		else
			' If there are no records, there are no parts
			Response.Write "<tr><td>No Parts</td></tr>"
		End If

		Do While Not rs.EOF
			' We do not want to put Labor or "Special" parts on the parts list.
			if(rs("Type") <> "S" AND rs("Type") <> "L") then
				Response.Write "<tr><td>" & rs("Quantity") & "</td><td>" & rs("ItemDescription")
				if(rs("Type") = "A") then
					Response.Write " (Assembly)"
				end if
				Response.Write "</td></tr>"
			end if
			
			' If the part is an assembly, we need to gather all of its components.
			if(rs("Type") = "A") then
				assemblyQuery = "SELECT tblSOPartsUsedAssemblyDetail.ItemDescription, tblSOPartsUsedAssemblyDetail.Quantity FROM tblSOPartsUsedAssemblyDetail WHERE tblSOPartsUsedAssemblyDetail.FKSOPartsUsed = " & rs("SOPartsUsedKeyID")
				set rs2 = Server.CreateObject("ADODB.recordset")
				rs2.Open assemblyQuery, Conn
				
				if NOT rs2.EOF then
					rs2.MoveFirst
					Response.Write("<tr><td></td><td><table class=""assembly"">")
					
					Do While Not rs2.EOF
						Response.Write "<tr><td>" & rs2("Quantity") & "</td><td>" & rs2("ItemDescription") & "</td></tr>"
						rs2.MoveNext
					Loop
					
					Response.Write("</table></tr></td>")
				end if
				
			end if
			rs.MoveNext
		Loop
	else
		Response.Write "<br/><br/>INVALID SO NUMBER"
	end if
end if	


''' buildRecordSetQuery
'''
''' Subroutine builds a task query string for a given day and repgroup
'''
''' @param String day
''' @param String repGroup
''' @return String query - string containing task query for a given day and repgroup
Function buildRecordSetQuery(day, repGroup)

	queryStart = "SELECT tblReps.FirstName as techFirstName, tblReps.LastName as techLastName, tblTasks.StartTime as StartTime, tblTasks.EndTime, tblTasks.StartDate, tblTasks.EndDate, tblTasks.TaskCompletedIndicator, tblTasks.Location, tblAccounts.AccountName, tblTasks.Subject, tblAccounts.Address1, tblAccounts.Address2, tblAccounts.City, tblAccounts.State, tblAccounts.PostalCode, tblAccounts.PrimaryPhoneNumber, tblTasks.TaskKeyID, tblTasks.SONumber, tblTasks.TaskComment FROM tblTasks, tblAssignedRepGroups, tblReps, tblAccounts WHERE tblTasks.AccountNumber = tblAccounts.AccountNumber AND StartDate <= '"
	queryMid = "' AND EndDate >= '"
	queryEnd = "' AND ScheduledForRepNumber = tblReps.RepNumber AND ScheduledForRepNumber = tblAssignedRepGroups.RepNumber AND RepGroup = '" & repGroup & "' ORDER BY StartTime"
	
	buildRecordSetQuery = queryStart & day & queryMid & day & queryEnd
End Function
' End buildRecordSetQuery Function



''' displayDispatchTask
'''
''' Subroutine outputs a technician's task in a table for a given recordset and date.
'''
''' @param Recordset rs
''' @param String date
''' @return None
Sub displayDispatchTask(rs, date)

	if NOT rs.EOF then
		rs.MoveFirst
	End If

	' Begin Recordset Loop
	Do While Not rs.EOF

		' The dispatch board is task driven. If a technician completes a task, it should no longer be shown on the board.
		if (not rs("TaskCompletedIndicator")) then
%>
		<table width="100%" class="dispatch">
			<tr>
				<td width="40%">
					<strong>
<%		
			Response.Write (rs("techFirstName") & " " & rs("techLastName"))
%>
					</strong>
				</td>
				<td align="right">
<%
			' Some tasks span multiple days. This will display arrows to show that the task carries on to the day before or after.
			' These conditionals are not an if / elseif or a switch case because the technician will need to know when to be there on
			' that particular day even if the task spans the days before and/or after.
			if (rs("StartDate") < date) then
				Response.Write "<< "
			end if
			if (rs("StartDate") = date) then
				Response.Write (rs("StartTime"))
			end if
			if (rs("StartDate") < date AND NOT rs("EndDate") > date) then
				if(not isnull(rs("EndTime"))) then
					Response.Write (" until " & rs("EndTime"))
				end if
			end if
			if (rs("StartDate") < date AND rs("EndDate") > date) then
				Response.Write " All Day "
			end if
			if (rs("EndDate") > date) then
				Response.Write " >>"
			end if
%>
				</td>
			</tr>
			<tr>
				<td colspan="2"><strong>
<%		
			' Not all tasks have locations assigned. This is a manually entered field by the dispatcher.
			if(not isnull(rs("Location"))) then
				Response.Write rs("Location")
			else
				Response.Write "Verify Service Location"
			end if
%>			
				</strong></td>
			</tr>
			<tr>
				<td class="partslist" colspan="2">
					<img src="dispatchimages/screwdriver.jpg">
<%
			' Create a link to the service order parts list so that road techs can access it for verification.
			Response.Write ("<a href=""dispatchboard.asp?display=parts&so=" & rs("SONumber") & """><strong>" & rs("AccountName") & "</a></strong> (click for parts list)")

%>
				</td>
			</tr>
			<tr>
				<td colspan="2">
<%	
			' Display the description of the task. Sometimes these descriptions can have characters in them such as < and > which are not
			' HTML friendly. These need to be HTML encoded to display properly.
			if(not isnull(rs("Subject"))) then
				Response.Write (Server.HTMLEncode(rs("Subject")))
			end if
%>
				</td>
			</tr>
			<tr>
				<td colspan="2">
<%
			Response.Write (rs("Address1"))
%>
				<br/>
<%
			' Clients do not have to have the Address 2 field.
			if(not isnull(rs("Address2"))) then
					Response.Write (rs("Address2") & "<br/>")
			end if
			
			Response.Write(rs("City") & ", " & rs("State") & " " & rs("PostalCode"))

%>
				</td>
			</tr>
<%
			
			' A second recordset to pull contact data for a particular task ID.
			set rs2 = Server.CreateObject("ADODB.recordset")
			rs2.Open "SELECT tblContacts.ContactName FROM tblContacts, tblTasks WHERE tblContacts.ContactNumber = tblTasks.ContactNumber AND tblTasks.TaskKeyID = " & rs("TaskKeyID"), Conn
%>
			<tr>
				<td colspan="2">
<%
			' Not all tasks have an assigned client contact.
			if(not rs2.EOF AND not isnull(rs2("ContactName"))) then
				Response.Write (rs2("ContactName") & " ")
			end if
			Response.Write (rs("PrimaryPhoneNumber"))
%>
				</td>
			</tr>
<%			
			' Google Map Link Construction
			googleMapAddress = ""
			
			if(not isnull(rs("Address1"))) then
				googleMapAddress = rs("Address1")
			end if
			
			if(not isnull(rs("Address2"))) then
				googleMapAddress = googleMapAddress & " " & rs("Address2")
			end if
			
			if(not isnull(rs("City"))) then
				googleMapAddress = googleMapAddress & ", " & rs("City")
			end if
			
			if(not isnull(rs("State"))) then
				googleMapAddress = googleMapAddress & ", " & rs("State")
			end if
			
			if(not isnull(rs("PostalCode"))) then
				googleMapAddress = googleMapAddress & " " & rs("PostalCode")
			end if
			
			if(googleMapAddress <> "") then
				Response.Write "<tr><td colspan=""2"" align=""right""><a href=""http://maps.google.com/maps?q=" & Server.URLEncode(googleMapAddress) & """>Google Maps</a></td></tr>"
			end if
			' End Google Map Link Construction
			
			' Not all tasks have a comment.
			if(not isnull(rs("TaskComment"))) then
				Response.Write ("<tr><td colspan=""2""><strong>Task Comments:</strong><br/>")
				Response.Write (rs("TaskComment") & " ")
				Response.Write ("</td></tr>")
			end if
			
			' A third recordset to pull internal service order comments. This recordset could be joined with the second recordset for efficiency.
			set rs3 = Server.CreateObject("ADODB.recordset")
			rs3.Open "SELECT tblServiceOrders.InternalSOComments FROM tblServiceOrders, tblTasks WHERE tblTasks.SONumber = tblServiceOrders.SONumber AND tblTasks.TaskKeyID = " & rs("TaskKeyID"), Conn
			
			if(not rs3.EOF AND not isnull(rs3("InternalSOComments"))) then			
				Response.Write ("<tr><td colspan=""2""><b>SO Internal Comments:</b><br/>")
				Response.Write (rs3("InternalSOComments") & " ")
				Response.Write ("</td></tr>")
			end if

			
%>
		</table>
<%
		End If

		rs.MoveNext
	Loop
	' End Recordset Loop

	
End Sub
' End displayDispatchTask Subroutine

Set rs = Nothing
Set rs2 = Nothing
Set rs3 = Nothing
%>
</body>
</html>