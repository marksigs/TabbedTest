<%@  Language=JScript %>
<html>
<%
/*
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Workfile:      scTableListScroll.htm
Copyright:     Copyright © 1999 Marlborough Stirling

Description:   File containing list table functionality
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
History:

Prog	Date		Description

~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
MARS Specific History :

Prog	Date		AQR			Description
GHun	22/07/2005	MAR14		Apply ING Style Sheet and GUI Images
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
*/
%>
<head>
	<meta name="vs_targetSchema" content="http://schemas.marlborough-stirling.com/intellisense/ie5_omiga4"/>
	<link href="stylesheet.css" rel="stylesheet" type="text/css"/>
</head>

<body>
<div id="divNumbers0" class="msgMenu3" style="LEFT: 0px; POSITION: absolute; TOP: 1px; HEIGHT: 22px; WIDTH: 154px; background-color:FFEFE5">
	<div id="divNumbers1" class="msgMenu3" style="LEFT: 1px; POSITION: absolute; TOP: 2px; HEIGHT: 18px; WIDTH: 151px; background-color:FFEFE5; color:616161">
		<div id="divNumbers2" style="LEFT: 4px; POSITION: absolute; TOP: 2px; VISIBILITY: hidden">
			<span style="LEFT: 0px; POSITION: absolute; TOP: 0px">records</span>
			<span id="spStartingRecNo" style="LEFT: 40px; POSITION: absolute; TOP: 0px">999</span> 
			<span style="LEFT: 64px; POSITION: absolute; TOP: 0px">to</span>
			<span id="spMaxOnScreen" style="LEFT: 80px; POSITION: absolute; TOP: 0px">999</span>
			<span style="LEFT: 104px; POSITION: absolute; TOP: 0px">of</span>
			<span id="spNoOfRecords" style="LEFT: 120px; POSITION: absolute; TOP: 0px">9999</span>
		</div>
	</div>
</div>

<div id="idScrollButtons" style="HEIGHT: 24px; LEFT: 156px; POSITION: absolute; TOP: 0px; WIDTH: 148px; background-color:FFEFE5">
	<form id="frmScrollButtons" style="VISIBILITY: hidden">
		<button id="btnStart" class="msgButton" style="WIDTH: 24px" type="button">|&lt;</button> 
		<button id="btnPageUp" class="msgButton" style="WIDTH: 24px" type="button">&lt;&lt;</button> 
		<button id="btnOneUp" class="msgButton" style="WIDTH: 24px" type="button">&lt;</button> 
		<button id="btnOneDown" class="msgButton" style="WIDTH: 24px" type="button">&gt;</button> 
		<button id="btnPageDown" class="msgButton" style="WIDTH: 24px" type="button">&gt;&gt;</button> 
		<button id="btnEnd" class="msgButton" style="WIDTH: 24px" type="button">&gt;|</button> 
	</form>
</div> 
</body>

<script language="javascript" type="text/javascript">

public_description = new TableScrollListMgr;
	
<% /* ---- Table data ---- */ %>
var m_refTable = null;
var m_refKeyInput = null;
var m_nKeyCol = null;
var m_indexRowSelected = null;
var m_aidActualRowsSelected = new Array(); <% /* Array of Visible Rows selected	*/ %>
var m_bIsDisabled = null;
var m_bIsTableMultiSelect = null; 
var m_CtrlKey = null;
var m_ShiftKey = null;

<% /* ---- Scroll data ---- */ %>
var m_fnShowList = null;
var m_nStart = null; <% /* first visible record */ %>
var m_nEnd = null; <% /* last visible record */ %>
var m_nTotalRecords = null;
<% /*the number of rows in the table field, excluding any header row */ %>
var m_nTableLength = null;
var m_aFGColor = new Array();

function TableScrollListMgr()
{
<%	// Table data
%>	this.initialiseTable = initialiseTable;
	this.clear = clear;
	this.getRowSelected = getRowSelected;
	this.getRowSelectedIndex = getRowSelectedIndex;
	this.getArrayofRowsSelected = getArrayofRowsSelected;
	this.setRowSelected = setRowSelected;
	this.setAllRowsSelected = setAllRowsSelected;
	this.setMultiRowSelected = setMultiRowSelected;
	this.DisableTable = DisableTable;
	this.EnableTable = EnableTable;
	this.EnableMultiSelectTable = EnableMultiSelectTable;
	this.DisableMultiSelectTable = DisableMultiSelectTable;
	this.redisplaySelection = redisplaySelection;
	this.setMultiRowUnselected = setMultiRowUnselected;
	this.getFirstVisibleRecord = getFirstVisibleRecord;
	this.getLastVisibleRecord = getLastVisibleRecord;
	this.getTotalRecords = getTotalRecords;
	this.setRowSelectedIndex = setRowSelectedIndex;

<%	// Scroll data
%>	this.getOffset = getOffset;
	this.RowDeleted = RowDeleted;
	this.oneDown = oneDown;
	this.oneUp = oneUp;
	this.pageDown = pageDown;
	this.pageUp = pageUp;
	this.toTop = toTop;
	this.toEnd = toEnd;
	this.setRecordSelected = setRecordSelected;
	

	frmScrollButtons.btnOneDown.onclick = oneDown;
	frmScrollButtons.btnOneUp.onclick = oneUp;
	frmScrollButtons.btnPageDown.onclick = pageDown;
	frmScrollButtons.btnPageUp.onclick = pageUp;
	frmScrollButtons.btnStart.onclick = toTop;
	frmScrollButtons.btnEnd.onclick = toEnd;
}	

function initialiseTable(refTable, nKeyCol, refKeyInput, fnShowList, nTableLength, nTotalRecords, aFGColor)
{
	m_refTable = refTable;
	m_refKeyInput = refKeyInput;
	m_nKeyCol = nKeyCol;
	m_refTable.onmouseover = table_mouseover;
	m_refTable.onmouseout = table_mouseout;
	m_refTable.onclick = table_onclick;
	m_refTable.onkeydown = table_onkeydown;
	m_refTable.onkeyup = table_onkeyup;
	m_bIsDisabled = false;
	m_bIsTableMultiSelect = false; <% /* MCS 22/10/99 Multiselect table default to single selection */ %>
	m_CtrlKey = false;
	m_ShiftKey = false;

	m_fnShowList = fnShowList;
<%	// set the number of rows in the table control, excluding any header row
	//               (was previously hard-coded as 10)
%>	m_nTableLength = nTableLength;
	m_nTotalRecords = nTotalRecords;
	m_aFGColor = aFGColor
	
	divNumbers0.style.backgroundColor = "FFEFE5";
	divNumbers2.style.visibility = "hidden";
	frmScrollButtons.style.visibility = "hidden";

	if(m_nTotalRecords > 0)
	{
		m_nStart = 1;
		if(m_nTotalRecords > m_nTableLength)
		{
			m_nEnd = m_nTableLength;
			frmScrollButtons.style.visibility = "visible";
			divNumbers0.style.backgroundColor = "616161";
			divNumbers2.style.visibility = "visible";
			showPosition();
		}
		else m_nEnd = m_nTotalRecords;
	}

	doButtonStates();
}

function DisableTable()
{
	m_bIsDisabled = true;
}

function EnableTable()
{
	m_bIsDisabled = false;
}

function DisableMultiSelectTable()
{
	if(!m_bIsDisabled) m_bIsTableMultiSelect = false;
}

function EnableMultiSelectTable()
{
	if(!m_bIsDisabled) m_bIsTableMultiSelect = true;
}

function clear()
{
	if( m_refTable != null )
	{	
		for(var i = 0; i < m_refTable.rows.length; i++)
		{
			if(m_refTable.rows(i).id != "rowTitles")
			{
				m_refTable.rows(i).style.background = "#FFFFFF";
				for(var j = 0; j < m_refTable.rows(i).cells.length; j++)
				{
					m_refTable.rows(i).cells(j).innerText = " ";
					m_refTable.rows(i).cells(j).title = "";
					//m_refTable.rows(i).cells(j).style.color = "616161";
					 <% /* Add 1 to take into account the titles */ %>
					m_refTable.rows(i).cells(j).style.color = m_aFGColor[m_nStart + i - 1];
				}
			}
		}
		m_indexRowSelected = null;
		
		<% /* Clearing a multi-select table resets the selected row array. */ %>
		m_aidActualRowsSelected = new Array();
	}
}

function getRowSelectedIndex()
{
	return m_indexRowSelected;
}

function getFirstVisibleRecord()
{
	return m_nStart;
}	
function getLastVisibleRecord()
{
	return m_nEnd;
}
function getTotalRecords()
{
	return m_nTotalRecords;
}	
function setRowSelectedIndex(iIndex)
{
	m_indexRowSelected = iIndex;
}

function getArrayofRowsSelected()
{
	return m_aidActualRowsSelected;
}

function getRowSelected()
{
	if(m_indexRowSelected != null)
		for(var i = 0; i < m_refTable.rows.length; i++)
			if(m_refTable.rows(i).rowIndex == ( m_indexRowSelected - getOffset())) return i;

	return -1;
}

function setRowSelected(nRow)
{
	for(var i = 0; i < m_refTable.rows.length; i++)
	{
		if(m_refTable.rows(i).id != "rowTitles")
		{
			if(i == nRow)
			{
				m_indexRowSelected = m_aidActualRowsSelected[0]= m_refTable.rows(i).rowIndex + getOffset();
				m_refTable.rows(i).style.background = "FFEFE5";
				
				for(var j = 0; j < m_refTable.rows(i).children.length; j++)
					m_refTable.rows(i).children(j).style.color = "616161";
			}
			else
			{
				m_refTable.rows(i).style.background = "white";
				for(var j = 0; j < m_refTable.rows(i).children.length; j++)
					//m_refTable.rows(i).children(j).style.color = "616161";
					<% /* Add 1 to take into account the titles */ %>
					m_refTable.rows(i).children(j).style.color = m_aFGColor[m_nStart + i - 1];
			}
		}
	}

	if(nRow == -1)
	{
		m_indexRowSelected = null;
		m_aidActualRowsSelected.length=0;
	}
}

function setMultiRowSelected(nRow)
{
	if(m_bIsTableMultiSelect)
	{
		var rowsSelectedSize = m_aidActualRowsSelected.length;
		var bRowSelected = false;
		for(var i = 0; i < m_refTable.rows.length && bRowSelected == false; i++)
		{
			if(m_refTable.rows(i).id != "rowTitles")
			{
				if(i == nRow)
				{
					// avoid saving the same row twice. This could happen when re-highlighting after scrolling the listbox
					var bAlreadySaved = false;
					for(var j = 0; j < rowsSelectedSize && bAlreadySaved == false; j++)
					{
						if(m_aidActualRowsSelected[j] - getOffset() == i) bAlreadySaved = true;
					}
					if(!bAlreadySaved)m_indexRowSelected = m_aidActualRowsSelected[rowsSelectedSize]= m_refTable.rows(i).rowIndex + getOffset();
					else m_indexRowSelected = m_refTable.rows(i).rowIndex + getOffset();
					m_refTable.rows(i).style.background = "FFEFE5";
					bRowSelected = true;
					for(var j = 0; j < m_refTable.rows(i).children.length; j++)
						m_refTable.rows(i).children(j).style.color = "616161";
				}
			}
		}
	}
}

function setAllRowsSelected()
{
	if(m_bIsTableMultiSelect)
	{
		for(var i = 1; i <= m_nTotalRecords; i++)
		{
			if(i >= m_nStart && i <= m_nEnd )
			{
				if(m_refTable.rows(i - getOffset()).id != "rowTitles" )
				{
					m_indexRowSelected = m_nStart;

					m_refTable.rows(i - getOffset()).style.background = "FFEFE5";
					for(var j = 0; j < m_refTable.rows(i - getOffset()).children.length; j++)
						m_refTable.rows(i - getOffset()).children(j).style.color = "616161";
				}
			}
			m_aidActualRowsSelected[i-1]=i;
		}
	}
}

function table_mouseover()
{
	var thisEvent = this.document.parentWindow.event.srcElement;

	if(!m_bIsDisabled && thisEvent.parentElement.id != "rowTitles" && thisEvent.tagName == "TD" && 
		thisEvent.parentElement.rowIndex != m_indexRowSelected - getOffset())
	{
		var bModify = false;

		if(m_bIsTableMultiSelect)
		{
			var iArraySize = m_aidActualRowsSelected.length;

<%			// nothing selected
			//just do normal processing as we don't have a current selection
%>			if( iArraySize == 0 ) bModify = true;
			else
			{
				bModify = true;
<%				//need to loop round all of the possible rows that are selected
%>				for(var loop = 0; loop < iArraySize && bModify == true;loop++)
				{
					idRowSelected = m_aidActualRowsSelected[loop];
					if(idRowSelected >= m_nStart && idRowSelected <= m_nEnd)
						if(thisEvent.parentElement.rowIndex == idRowSelected - getOffset()) bModify = false;
				}
			}
		}
		else if(thisEvent.parentElement.rowIndex != ( m_indexRowSelected - getOffset() )) bModify = true;

		if (bModify && thisEvent.parentElement.cells(m_nKeyCol).innerText.length > 0 && 
			thisEvent.parentElement.cells(m_nKeyCol).innerText != " ")
		{
			thisEvent.parentElement.style.background = "D4DDE9";
			thisEvent.style.cursor = "hand";
		}
	}
}

function table_onkeydown()
{
	if(!m_bIsDisabled && m_bIsTableMultiSelect)
	{
		var thisEvent = this.document.parentWindow.event;
		m_ShiftKey = thisEvent.shiftKey;
		m_CtrlKey = thisEvent.ctrlKey;
	}
}

function table_onkeyup()
{
	if(!m_bIsDisabled && m_bIsTableMultiSelect)
	{
		var thisEvent = this.document.parentWindow.event;
		m_ShiftKey = thisEvent.shiftKey;
		m_CtrlKey = thisEvent.ctrlKey;
	}
}

function table_mouseout()
{
	var thisEvent = this.document.parentWindow.event.srcElement;

	if(!m_bIsDisabled && thisEvent.tagName == "TD" && thisEvent.parentElement.id != "rowTitles" && 
		thisEvent.parentElement.rowIndex != m_indexRowSelected - getOffset())
	{
		var bModify = false;

		if (m_bIsTableMultiSelect)
		{
			var iArraySize = m_aidActualRowsSelected.length;

<%			//just do normal processing as we don't have a current selection
%>			if( iArraySize == 0 ) bModify = true;
			else
			{
				bModify	= true;
<%				//need to loop round all of the possible rows that are selected
%>				for(var loop = 0; loop < iArraySize && bModify	== true;loop++)
				{
					idRowSelected = m_aidActualRowsSelected[loop];
					if(idRowSelected >= m_nStart && idRowSelected <= m_nEnd)
						if(thisEvent.parentElement.rowIndex == idRowSelected - getOffset()) bModify = false;
				}
			}
		}
		else bModify = true;

		if(bModify)
		{
			thisEvent.parentElement.style.background = "FFFFFF";
			thisEvent.parentElement.style.color = "616161";
					//m_refTable.rows(i).children(j).style.color = m_aFGColor[i];
			thisEvent.style.cursor = "";
		}
	}
}

function table_onclick()
{
	var thisEvent = this.document.parentWindow.event.srcElement;

	if(!m_bIsDisabled && thisEvent.tagName == "TD" && thisEvent.parentElement.id != "rowTitles")
	{
		var bInCurrentSelection = false;
		var idRowSelected;
		var iArraySize = m_aidActualRowsSelected.length;
		var iValue;

<%		// loop round all the elements in the array
%>		for(var loop = 0; loop < iArraySize;loop++)
		{
			idRowSelected = m_aidActualRowsSelected[loop];

			if(idRowSelected >= m_nStart && idRowSelected <= m_nEnd)
				if(thisEvent.parentElement.rowIndex == idRowSelected - getOffset())
				{
					bInCurrentSelection = true;
					iValue=loop;
				}
		}

		var bDeSelect = false;
		var bDeSelectMulti = false;
		var bDeSelectShift = false;
		var bSelect = false;
		var bShiftSelect = false;
		var bDeSelectAll = false;
		var bSelectMulti = false;

		if(m_bIsTableMultiSelect)
		{
			if(m_CtrlKey && !m_ShiftKey)
			{
<%				//if we have more than one element de-select current selection
				//else add this current selection to the array
%>				if(bInCurrentSelection && iArraySize > 1) bDeSelectMulti = true;
				else if(!bInCurrentSelection) bSelectMulti = true;

			}
			else if(m_ShiftKey && !m_CtrlKey)
			{
<%				//select all elements upto and including this element
				//else select from current selection to this selection
%>				if(bInCurrentSelection && iArraySize > 1) bDeSelectShift = true;
				else bShiftSelect = true;
			}
			else if(!m_ShiftKey && !m_CtrlKey)
			{
				if(iArraySize > 1 )
				{
					bDeSelectAll = true;
					bSelect = true;
				}
				else if (!bInCurrentSelection)
				{
					bDeSelect = true;
					bSelect = true;
				}
			}
		}
		else if( (!bInCurrentSelection && iArraySize > 0) && !bSelectMulti)
		{
<%			//single selection 
			//de-select previous
			//select this one
%>			bDeSelect = true;
			bSelect = true;
		}
		else if(m_indexRowSelected == null) 
		{
			bSelect = true;
			iArraySize = 1;
		}

		if(bDeSelect || bDeSelectAll || bDeSelectMulti)
		{
			if (thisEvent.parentElement.cells(m_nKeyCol).innerText.length > 0 && thisEvent.parentElement.cells(m_nKeyCol).innerText != " ")
			{
				var bLoop = true;

<%				//loop round all the elements in the array
%>				for(var loop = 0; loop < iArraySize && bLoop;loop++)
				{
					idRowSelected = m_aidActualRowsSelected[loop];

					if(idRowSelected >= m_nStart && idRowSelected <= m_nEnd)
					{
						if(((bDeSelect || bDeSelectAll) && thisEvent.parentElement.rowIndex != idRowSelected - getOffset())
						|| (bDeSelectMulti && thisEvent.parentElement.rowIndex == idRowSelected - getOffset()))
						{
							var bContinue = true;
							var bModify = true;
							for(var i = 0; i < m_refTable.rows.length && bContinue; i++)
							{
								if(m_refTable.rows(i).id != "rowTitles")
								{
									if(bDeSelectMulti)
									{
										if(m_refTable.rows(i).rowIndex == thisEvent.parentElement.rowIndex)
										{
											bContinue = false;
											bLoop = false;
											bModify = true;
										}
										else bModify = false; <% /* dont do modify */ %>
									}

									if(bModify)
									{
										m_refTable.rows(i).style.background = "FFFFFF";

										for(var j = 0; j < m_refTable.rows(i).children.length; j++)
											<% /* Add 1 to take into account the titles */ %>
											m_refTable.rows(i).children(j).style.color = m_aFGColor[m_nStart + i - 1];
											//m_refTable.rows(i).children(j).style.color = "616161";

										m_refKeyInput.value = m_refTable.rows(i).cells(m_nKeyCol).innerText;
									}

									if(!bDeSelectAll) bLoop = false;
								}
							}
						}
					}
				}
			}

			if(bDeSelectMulti)
			{
				var aidOldCurrentRowSelected = new Array();
<%				//copy the currentarray to the oldarray
%>				aidOldCurrentRowSelected = m_aidActualRowsSelected;

				for(var loop2 = iValue; loop2 < iArraySize;loop2++)
				{
					m_indexRowSelected = m_aidActualRowsSelected[loop2] = aidOldCurrentRowSelected[loop2+1];
					bInCurrentSelection	= true;
				}

				m_aidActualRowsSelected.length = iArraySize-1;
			}

			if(bDeSelectAll)
			{
				m_aidActualRowsSelected.length=1;
				iArraySize=1;
			}
		}

		if(bSelect || bSelectMulti || bShiftSelect)
		{
			if(bShiftSelect)
			{
				//NO shift processing yet
			}

			if (thisEvent.parentElement.cells(m_nKeyCol).innerText.length > 0 && thisEvent.parentElement.cells(m_nKeyCol).innerText != " ")
			{
			
				thisEvent.parentElement.style.background = "FFEFE5";

				for(var i = 0; i < thisEvent.parentElement.children.length; i++)
					thisEvent.parentElement.children(i).style.color = "616161";

				if(bSelect)
				{
					if(iArraySize == 0)m_indexRowSelected = m_aidActualRowsSelected[iArraySize] = thisEvent.parentElement.rowIndex + getOffset();
					else m_indexRowSelected = m_aidActualRowsSelected[iArraySize-1] = thisEvent.parentElement.rowIndex + getOffset();
				}
				if(bSelectMulti)
					m_indexRowSelected = m_aidActualRowsSelected[iArraySize] = thisEvent.parentElement.rowIndex + getOffset();
			}
		}
	}
}

function redisplaySelection()
{
	m_fnShowList(m_nStart - 1);
	doButtonStates();
	showPosition();

	var iArraySize = m_aidActualRowsSelected.length;

<%	//total No. of  visible rows
%>	for(var i = 0; i <= m_nTableLength ; i++)
	{
		var bSetWhite = true;

		if(m_refTable.rows(i).id != "rowTitles" )
		{
			var bExit=false;

			for(var loop = 0; loop < iArraySize && !bExit;loop++)
			{
				idAbsoluteRowSelected = m_aidActualRowsSelected[loop];

				idNewRowSelected = idAbsoluteRowSelected - getOffset();

				if(idNewRowSelected == i )
				{
					bExit = true;
					bSetWhite = false;
				}
				else if(idNewRowSelected > i ) <% /* gone passed the current index */ %>
					bExit	=	true;
			}

			if(i >= (m_nStart - getOffset() ) && i <= (m_nEnd - getOffset()))
			{
				if(!bSetWhite)
				{
					m_refTable.rows(i).style.background = "FFEFE5";
					for(var j = 0; j < m_refTable.rows(i).children.length; j++)
						m_refTable.rows(i).children(j).style.color = "616161";
				}
				else
				{
					m_refTable.rows(i).style.background = "FFFFFF";
					for(var j = 0; j < m_refTable.rows(i).children.length; j++)
						<% /* Add 1 to take into account the titles */ %>
						m_refTable.rows(i).children(j).style.color = m_aFGColor[m_nStart + i - 1];
						//m_refTable.rows(i).children(j).style.color = "616161";
				}
			}
		}
	}
}
<%
//---------------------------------------END TABLE FUNCTIONALITY HERE------------------------------------------

//---------------------------------------BEGIN SCROLL FUNCTIONALITY HERE---------------------------------------
%>
function showPosition()
{
<%	// put position numbers in boxes
%>	spStartingRecNo.innerHTML = m_nStart;
	spMaxOnScreen.innerHTML = m_nEnd;
	spNoOfRecords.innerHTML = m_nTotalRecords;
}

function toEnd()
{
	if(m_nTotalRecords > m_nTableLength && (m_nStart < m_nTotalRecords - (m_nTableLength - 1)))
	{
		m_nStart = m_nTotalRecords - (m_nTableLength - 1);
		m_nEnd = m_nTotalRecords;
		redisplaySelection();
	}
}

function toTop()
{
	if(m_nTotalRecords > m_nTableLength && m_nStart > 1)
	{
		m_nStart = 1;
		m_nEnd = m_nTableLength;
		redisplaySelection();
	}
}

function oneDown()
{
	if(m_nTotalRecords > m_nTableLength && (m_nStart + m_nTableLength <= m_nTotalRecords))
	{
		m_nStart++;
		m_nEnd++;
		redisplaySelection();
	}
}

function oneUp()
{
	if(m_nTotalRecords > m_nTableLength && m_nStart > 0)
	{
		m_nStart--;
		m_nEnd--;
		redisplaySelection();
	}
}

function pageUp()
{
	if(m_nTotalRecords > m_nTableLength && m_nStart > 0)
	{
		m_nStart -= m_nTableLength;
		if(m_nStart < 1) m_nStart = 1;
		m_nEnd = m_nStart + (m_nTableLength - 1);
		redisplaySelection();
	}
}

function pageDown()
{
	if(m_nTotalRecords > m_nTableLength && (m_nStart < m_nTotalRecords - (m_nTableLength - 1)))
	{
		m_nStart += m_nTableLength;
		if(m_nStart + m_nTableLength > m_nTotalRecords) m_nStart = m_nTotalRecords - (m_nTableLength - 1);
		m_nEnd = m_nStart + (m_nTableLength - 1);
		redisplaySelection();
	}
}

function setRecordSelected(nRec)
{
	if (nRec < 1 || nRec > m_nTotalRecords) alert("Invalid selection");
	if (nRec > m_nTableLength) m_nStart = (scMath.Truncate(nRec / m_nTableLength) * m_nTableLength) + 1;

	if (nRec >  m_nStart + m_nTableLength) m_nEnd = m_nStart + (m_nTableLength - 1);
	else m_nEnd = nRec;

	redisplaySelection();
	setRowSelected(nRec - (m_nStart - 1));
}

function doButtonStates()
{
	if(m_nStart == 1)
	{
		frmScrollButtons.btnStart.disabled = true;
		frmScrollButtons.btnPageUp.disabled = true;
		frmScrollButtons.btnOneUp.disabled = true;
	}
	else
	{
		frmScrollButtons.btnStart.disabled = false;
		frmScrollButtons.btnPageUp.disabled = false;
		frmScrollButtons.btnOneUp.disabled = false;
	}

	if(m_nStart < (m_nTotalRecords - (m_nTableLength - 1)))
	{
		frmScrollButtons.btnOneDown.disabled = false;
		frmScrollButtons.btnPageDown.disabled = false;
		frmScrollButtons.btnEnd.disabled = false;
	}
	else
	{
		frmScrollButtons.btnOneDown.disabled = true;
		frmScrollButtons.btnPageDown.disabled = true;
		frmScrollButtons.btnEnd.disabled = true;
	}

<%	// If the total rows fall below the table length, hide the scroll controls
%>	if(m_nTotalRecords <= m_nTableLength)
	{
		divNumbers0.style.backgroundColor = "FFEFE5";
		divNumbers2.style.visibility = "hidden";
		frmScrollButtons.style.visibility = "hidden";
	}
}

function getOffset()
{
	return (m_nStart -1);
}

function RowDeleted()
{
<%	// Decrement the total number of rows
	// If the end number is then greater than this set it to the total
%>	m_nTotalRecords--;

	if(m_nEnd > m_nTotalRecords) m_nEnd = m_nTotalRecords;

<%	// If the number of rows displayed is less than the length of the table
	// and we are not displaying record #1, move the display up one row
%>	if(m_nStart > 1 && m_nEnd - m_nStart < (m_nTableLength - 1)) m_nStart--;
		
<%	// Redisplay
%>	clear();
	m_fnShowList(m_nStart-1);
	redisplaySelection();
		
}

function SetRowUnselected(nRow)
{
	for(var i = 0; i < m_refTable.rows.length; i++)
	{
		if(m_refTable.rows(i).id != "rowTitles")
		{
			if(i == nRow)
			{
				m_indexRowSelected = m_aidActualRowsSelected[0]= m_refTable.rows(i).rowIndex + getOffset();
				m_refTable.rows(i).style.background = "616161";
				
				for(var j = 0; j < m_refTable.rows(i).children.length; j++)
				m_refTable.rows(i).children(j).style.color = "616161";
			}
			else
			{
				m_refTable.rows(i).style.background = "FFFFFF";
				for(var j = 0; j < m_refTable.rows(i).children.length; j++)
					 <% /* Add 1 to take into account the titles */ %>
					m_refTable.rows(i).children(j).style.color = m_aFGColor[m_nStart + i - 1];
					//m_refTable.rows(i).children(j).style.color = "616161";
			}
		}
	}

	if(nRow == -1)
	{
		m_indexRowSelected = null;
		m_aidActualRowsSelected.length=0;
	}
}

function setMultiRowUnselected(nRow)
{
	if(m_bIsTableMultiSelect)
	{	
		var rowsSelectedSize = m_aidActualRowsSelected.length;
		var bRowFound = false;
					
		for (var i=0; i<m_refTable.rows.length && bRowFound == false; i++)
			{
			if(m_refTable.rows(i).id != "rowTitles") 
			{	 
				if (i = nRow)
				{	
					//Find it in array list 
					for (var j = 0; j < rowsSelectedSize && bRowFound == false; j++)
					{
						if(m_aidActualRowsSelected[j]- getOffset() == i)
						{
							bRowFound = true;
						}		  
					}
					// Take it out of the array list
					if(bRowFound == true)
					{
						var aidOldCurrentRowsSelected = new Array();
						aidOldCurrentRowsSelected = m_aidActualRowsSelected;
						for(var k = j; k < rowsSelectedSize; k++)
						{
							m_indexRowSelected = m_aidActualRowsSelected[k] = aidOldCurrentRowsSelected[k+1];
						}	
						m_aidActualRowsSelected.length= rowsSelectedSize - 1;
						//Change colour of row to unselected
						m_refTable.rows(i).style.background = "FFFFFF";
						//Loop through the children of the row
						for(var l = 0; l < m_refTable.rows(l).children.length;l++)
						{
							<% /* Add 1 to take into account the titles */ %>
							m_refTable.rows(i).children(j).style.color = m_aFGColor[m_nStart + i - 1];
							//m_refTable.rows(i).children(l).style.color="616161";
						} //for	l		
 					}// if
				}//if i	
			}//if	
		}//for
	}//if				
}//function

</script>
</body>
</html>
