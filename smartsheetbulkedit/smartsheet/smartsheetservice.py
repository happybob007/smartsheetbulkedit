import logging
import sys

from smartsheetclient import Column
from smartsheetclient import RowPositionProperties
from smartsheetclient import SmartsheetClient
from smartsheetclient import SheetInfo

from smartsheetbulkedit import SmartsheetBulkEditError

class SmartsheetService(object):
	__logger = logging.getLogger(__name__)

	def __init__(self, token):
		self.__smartsheetClient = SmartsheetClient(token, logger=logging.getLogger(SmartsheetClient.__name__))
		self.__smartsheetClient.connect()

	def updateCell(self, sheet, rowNumber, columnIndex=None, columnTitle=None, value=None):
		if columnIndex is not None and columnTitle is not None:
			raise SmartsheetBulkEditError('one but not both "columnIndex" and "columnTitle" must be specified')
		elif columnTitle is not None:
			columnIndex = sheet.getColumnsInfo().getColumnByTitle(columnTitle).index
		elif columnIndex is None:
			raise SmartsheetBulkEditError('either "columnIndex" or "columnTitle" must be specified')
		row = sheet[rowNumber]
		row[columnIndex] = value
		row.getCellByIndex(columnIndex).save(propagate=False)

	def updateCellInAllSheets(self, rowNumber, workspace=None, columnIndex=None, columnTitle=None, value=None):
		for sheetInfo in self.getSheetInfos(workspace):
			sheet = self.__getSheetIfInWorkspace(sheetInfo, workspace)
			if sheet is not None:
				self.updateCell(sheet, rowNumber, columnIndex=columnIndex, columnTitle=columnTitle, value=value)

	def updateCellInSheetList(self, rowNumber, columnIndex=None, columnTitle=None, value=None, sheetInfoList=None):
		'''
		Update the specified Cell in the list of sheets.
		The Cell is identified by rowNumber and columnIndex or columnTitle.

		Two lists are returned, the first is a list of the SheetInfo objects
		for the Sheets that were successfully updated.
		The second returned list is a list of 3-tuples:
		    (SheetInfo, Exception, stacktrace) for the sheets that were NOT
		updated successfully.
		'''
		good, bad = [], []
		if columnIndex is not None and columnTitle is not None:
			raise SmartsheetBulkEditError('one but not both "columnIndex" and "columnTitle" must be specified')
		elif columnTitle is not None:
			columnIndex = sheet.getColumnsInfo().getColumnByTitle(columnTitle).index
		elif columnIndex is None:
			raise SmartsheetBulkEditError('either "columnIndex" or "columnTitle" must be specified')
		for sheetInfo in sheetInfoList:
			try:
				sheet = sheetInfo.loadSheet()
				self.updateCell(sheet, rowNumber=rowNumber, columnIndex=columnIndex, columnTitle=columnTitle, value=value)
				good.append(sheet)
			except Exception as e:
				bad.append((sheetInfo, e, sys.exc_info()[2]))
		return good, bad

	def addColumn(self, sheet, title, index=None, type=None, options=None, symbol=None, isPrimary=None, systemColumnType=None, autoNumberFormat=None, width=None):
		params = {}
		if sheet is not None:
			params["sheet"] = sheet
		if index is not None:
			params["index"] = index
		if type is not None:
			params["type"] = type
		if options is not None:
			params["options"] = options
		if symbol is not None:
			params["symbol"] = symbol
		if isPrimary is not None:
			params["primary"] = isPrimary
		if systemColumnType is not None:
			params["systemColumnType"] = systemColumnType
		if autoNumberFormat is not None:
			params["autoNumberFormat"] = autoNumberFormat
		column = Column(title, **params)
		sheet.insertColumn(column, column.index)

	def addColumnInAllSheets(self, title, workspace=None, index=None, type=None, options=None, symbol=None, isPrimary=None, systemColumnType=None, autoNumberFormat=None, width=None):
		for sheetInfo in self.getSheetInfos(workspace):
			sheet = self.__getSheetIfInWorkspace(sheetInfo, workspace)
			if sheet is not None:
				self.addColumn(
					sheet, 
					title, 
					index=index, 
					type=type, 
					options=options, 
					symbol=symbol, 
					isPrimary=isPrimary, 
					systemColumnType=systemColumnType, 
					autoNumberFormat=autoNumberFormat, 
					width=width)

	def addColumnInSheetList(self, title, workspace=None, index=None, type=None, options=None, symbol=None, isPrimary=None, systemColumnType=None, autoNumberFormat=None, width=None, sheetInfoList=None):
		good, bad = [], []
		for sheetInfo in sheetInfoList:
			try:
				sheet = sheetInfo.loadSheet()
				self.addColumn(
					sheet,
					title,
					index=index,
					type=type,
					options=options,
					symbol=symbol,
					isPrimary=isPrimary,
					systemColumnType=systemColumnType,
					autoNumberFormat=autoNumberFormat,
					width=width)
				good.append(sheetInfo)
			except Exception as e:
				bad.append((sheetInfo, e, sys.exc_info()[2]))
		return good, bad


	def updateColumn(self, sheet, oldTitle, newTitle=None, index=None, type=None, options=None, symbol=None, systemColumnType=None, autoNumberFormat=None, width=None, format=None):
		column = sheet.getColumnsInfo().getColumnByTitle(oldTitle)
		if newTitle is not None:
			column.title = newTitle
		if index is not None:
			column.index = index
		if type is not None:
			column.type = type
		if options is not None:
			column.options = options
		if symbol is not None:
			column.symbol = symbol
		if systemColumnType is not None:
			column.systemColumnType = systemColumnType
		if autoNumberFormat is not None:
			column.autoNumberFormat = autoNumberFormat
		if width is not None:
			column.width = width
		if format is not None:
			column.format = format
		column.update()

	def updateColumnInAllSheets(self, oldTitle, workspace=None, newTitle=None, index=None, type=None, options=None, symbol=None, systemColumnType=None, autoNumberFormat=None, width=None, format=None):
		for sheetInfo in self.getSheetInfos(workspace):
			sheet = self.__getSheetIfInWorkspace(sheetInfo, workspace)
			if sheet is not None:
				self.updateColumn(
					sheet, 
					oldTitle, 
					newTitle=newTitle, 
					index=index, 
					type=type, 
					options=options, 
					symbol=symbol, 
					systemColumnType=systemColumnType, 
					autoNumberFormat=autoNumberFormat, 
					width=width, 
					format=format)

	def updateColumnInSheetList(self, oldTitle, workspace=None, newTitle=None, index=None, type=None, options=None, symbol=None, systemColumnType=None, autoNumberFormat=None, width=None, format=None, sheetInfoList=None):
		good, bad = [], []
		for sheetInfo in sheetInfoList:
			try:
				sheet = sheetInfo.loadSheet()
				self.updateColumn(
					sheet,
					oldTitle,
					newTitle=newTitle,
					index=index,
					type=type,
					options=options,
					symbol=symbol,
					systemColumnType=systemColumnType,
					autoNumberFormat=autoNumberFormat,
					width=width,
					format=format)
				good.append(sheetInfo)
			except Exception as e:
				bad.append((sheetInfo, e, sys.exc_info()[2]))
		return good, bad

	def addRow(self, sheet, rowDictionary, rowNumber=None):
		row = sheet.makeRow(**rowDictionary)
		if rowNumber is None:
			# add as last row
			sheet.addRow(row)
		elif rowNumber in (0, 1):
			# add as first row
			sheet.addRow(row, position=RowPositionProperties.Top)
		else:
			# new row is inserted below sibling, so the sibling above will be:
			# if rowNumber < 0, the row currently at the desired row number
			# if rowNumber > 1, the row 1 above the desired row number
			siblingAboveRowId = sheet.getRowByRowNumber(rowNumber if rowNumber < 0 else rowNumber - 1).id
			sheet.addRow(row, siblingId=siblingAboveRowId)

	def addRowInAllSheets(self, rowDictionary, workspace=None, rowNumber=None):
		for sheetInfo in self.getSheetInfos(workspace):
			sheet = self.__getSheetIfInWorkspace(sheetInfo, workspace)
			if sheet is not None:
				self.addRow(sheet, rowDictionary, rowNumber)

	def addRowInSheetList(self, rowDictionary, rowNumber=None, sheetInfoList=None):
		good, bad = [], []
		for sheetInfo in sheetInfoList:
			try:
				sheet = sheetInfo.loadSheet()
				self.addRow(sheet, rowDictionary, rowNumber=rowNumber)
				good.append(sheetInfo)
			except Exception as e:
				bad.append((sheetInfo, e, sys.exc_info()[2]))
		return good, bad
		
	def expandAllRows(self, sheet, isExpanded=True):
		# operate only on rows referenced to be parent rows
		parentRowNumbers = frozenset([row.parentRowNumber for row in sheet.rows if row.parentRowNumber])
		for parentRowNumber in parentRowNumbers:
			row = sheet[parentRowNumber]
			if row.expanded != isExpanded:
				row.expanded = isExpanded
				row.save()

	def expandAllRowsInAllSheets(self, workspace=None, isExpanded=True):
		for sheetInfo in self.getSheetInfos(workspace):
			sheet = self.__getSheetIfInWorkspace(sheetInfo, workspace)
			if sheet is not None:
				self.expandAllRows(sheet, isExpanded)

	def expandAllRowsInSheetList(self, isExpanded=True, sheetInfoList=None):
		good, bad = [], []
		for sheetInfo in sheetInfoList:
			try:
				sheet = sheetInfo.loadSheet()
				self.expandAllRows(sheet, isExpanded=isExpanded)
				good.append(sheetInfo)
			except Exception as e:
				bad.append((sheetInfo, e, sys.exc_info()[2]))
		return good, bad

	def getSheetInfos(self, workspace=None):
		# Smartsheet Python SDK cannot filter by workspace
		return self.__smartsheetClient.fetchSheetList()

	def getSheetInfosInWorkspace(self, workspaceID=''):
		'''
		Get a list of the SheetInfo objects in the specified workspace.
		'''
		# Uses the Smartsheet Python SDK client directly, since this
		# functionality isn't yet implemented in the high-level Python SDK.
		sheetInfoList = []
		try:
			workspacePath = '/workspace/%s' % workspaceID
			workspace = self.__smartsheetClient.GET(workspacePath)
			for sheet_fields in workspace['sheets']:
				sheetInfoList.append(SheetInfo(sheet_fields, self.__smartsheetClient))
			return sheetInfoList
		except Exception as e:
			self.__logger.exception("Error getting list of sheets from workspace ID: %s: %r", workspaceID, e)
			raise
			raise (SmartsheetBulkEditError, ("Error getting list of sheets from workspace ID: %s: %r" % (workspaceID, e)), sys.exc_info()[2])

	def getWorkspacesByName(self, workspaceName):
		'''
		Get a list of workspaces that have the given name.
		The returned workspaces are Python dicts, not objects from the
		Smartsheet client library -- it does not yet support workspaces.
		'''
		# Uses the Smartsheet Python SDK client directly, since this
		# functionality isn't yet implemented by the high-level Python SDK.
		workspaces = []
		try:
			for workspace in self.__smartsheetClient.GET('/workspaces'):
				if workspace['name'] == workspaceName:
					workspaces.append(workspace)
			return workspaces
		except Exception as e:
			raise (SmartsheetBulkEditError, ("Error getting workspace list: %r" % e), sys.exc_info()[2])
		raise SmartsheetBulkEditError("Unable to find workspace named: '%s'" % workspaceName)

	def __getSheetIfInWorkspace(self, sheetInfo, workspace):
		""" Returns a Sheet if it belongs to the specified workspace
		or if workspace == None.  Returns None if the sheet does not 
		belong to the workspace.

		:param sheetInfo: the SheetInfo for the desired sheet
		:param workspace: the desired workspace name, or None to 
		disable workspace checking and always return the associated Sheet.
		"""
		sheet = sheetInfo.loadSheet()
		if (sheet):
			sheetWorkspace = sheet.workspace["name"]
			isSheetInWorkspace = not workspace or sheetWorkspace == workspace
			if (not isSheetInWorkspace):
				self.__logger.debug('sheet %s workspace "%s" != "%s"' % (sheetInfo, sheetWorkspace, workspace))
		return sheet

