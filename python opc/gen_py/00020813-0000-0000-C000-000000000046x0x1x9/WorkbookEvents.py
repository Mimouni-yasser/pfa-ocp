# -*- coding: mbcs -*-
# Created by makepy.py version 0.5.01
# By python version 3.8.3 (tags/v3.8.3:6f8c832, May 13 2020, 22:37:02) [MSC v.1924 64 bit (AMD64)]
# From type library '{00020813-0000-0000-C000-000000000046}'
# On Sun Aug 27 18:59:30 2023
'Microsoft Excel 16.0 Object Library'
makepy_version = '0.5.01'
python_version = 0x30803f0

import win32com.client.CLSIDToClass, pythoncom, pywintypes
import win32com.client.util
from pywintypes import IID
from win32com.client import Dispatch

# The following 3 lines may need tweaking for the particular server
# Candidates are pythoncom.Missing, .Empty and .ArgNotFound
defaultNamedOptArg=pythoncom.Empty
defaultNamedNotOptArg=pythoncom.Empty
defaultUnnamedArg=pythoncom.Empty

CLSID = IID('{00020813-0000-0000-C000-000000000046}')
MajorVersion = 1
MinorVersion = 9
LibraryFlags = 8
LCID = 0x0

class WorkbookEvents:
	CLSID = CLSID_Sink = IID('{00024412-0000-0000-C000-000000000046}')
	coclass_clsid = IID('{00020819-0000-0000-C000-000000000046}')
	_public_methods_ = [] # For COM Server support
	_dispid_to_func_ = {
		1610612736 : "OnQueryInterface",
		1610612737 : "OnAddRef",
		1610612738 : "OnRelease",
		1610678272 : "OnGetTypeInfoCount",
		1610678273 : "OnGetTypeInfo",
		1610678274 : "OnGetIDsOfNames",
		1610678275 : "OnInvoke",
		     1923 : "OnOpen",
		      304 : "OnActivate",
		     1530 : "OnDeactivate",
		     1546 : "OnBeforeClose",
		     1547 : "OnBeforeSave",
		     1549 : "OnBeforePrint",
		     1550 : "OnNewSheet",
		     1552 : "OnAddinInstall",
		     1553 : "OnAddinUninstall",
		     1554 : "OnWindowResize",
		     1556 : "OnWindowActivate",
		     1557 : "OnWindowDeactivate",
		     1558 : "OnSheetSelectionChange",
		     1559 : "OnSheetBeforeDoubleClick",
		     1560 : "OnSheetBeforeRightClick",
		     1561 : "OnSheetActivate",
		     1562 : "OnSheetDeactivate",
		     1563 : "OnSheetCalculate",
		     1564 : "OnSheetChange",
		     1854 : "OnSheetFollowHyperlink",
		     2157 : "OnSheetPivotTableUpdate",
		     2158 : "OnPivotTableCloseConnection",
		     2159 : "OnPivotTableOpenConnection",
		     2266 : "OnSync",
		     2283 : "OnBeforeXmlImport",
		     2285 : "OnAfterXmlImport",
		     2287 : "OnBeforeXmlExport",
		     2288 : "OnAfterXmlExport",
		     2610 : "OnRowsetComplete",
		     2895 : "OnSheetPivotTableAfterValueChange",
		     2896 : "OnSheetPivotTableBeforeAllocateChanges",
		     2897 : "OnSheetPivotTableBeforeCommitChanges",
		     2898 : "OnSheetPivotTableBeforeDiscardChanges",
		     2899 : "OnSheetPivotTableChangeSync",
		     2900 : "OnAfterSave",
		     2901 : "OnNewChart",
		     3075 : "OnSheetLensGalleryRenderComplete",
		     3076 : "OnSheetTableUpdate",
		     3077 : "OnModelChange",
		     3079 : "OnSheetBeforeDelete",
		}

	def __init__(self, oobj = None):
		if oobj is None:
			self._olecp = None
		else:
			import win32com.server.util
			from win32com.server.policy import EventHandlerPolicy
			cpc=oobj._oleobj_.QueryInterface(pythoncom.IID_IConnectionPointContainer)
			cp=cpc.FindConnectionPoint(self.CLSID_Sink)
			cookie=cp.Advise(win32com.server.util.wrap(self, usePolicy=EventHandlerPolicy))
			self._olecp,self._olecp_cookie = cp,cookie
	def __del__(self):
		try:
			self.close()
		except pythoncom.com_error:
			pass
	def close(self):
		if self._olecp is not None:
			cp,cookie,self._olecp,self._olecp_cookie = self._olecp,self._olecp_cookie,None,None
			cp.Unadvise(cookie)
	def _query_interface_(self, iid):
		import win32com.server.util
		if iid==self.CLSID_Sink: return win32com.server.util.wrap(self)

	# Event Handlers
	# If you create handlers, they should have the following prototypes:
#	def OnQueryInterface(self, riid=defaultNamedNotOptArg, ppvObj=pythoncom.Missing):
#	def OnAddRef(self):
#	def OnRelease(self):
#	def OnGetTypeInfoCount(self, pctinfo=pythoncom.Missing):
#	def OnGetTypeInfo(self, itinfo=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, pptinfo=pythoncom.Missing):
#	def OnGetIDsOfNames(self, riid=defaultNamedNotOptArg, rgszNames=defaultNamedNotOptArg, cNames=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg
#			, rgdispid=pythoncom.Missing):
#	def OnInvoke(self, dispidMember=defaultNamedNotOptArg, riid=defaultNamedNotOptArg, lcid=defaultNamedNotOptArg, wFlags=defaultNamedNotOptArg
#			, pdispparams=defaultNamedNotOptArg, pvarResult=pythoncom.Missing, pexcepinfo=pythoncom.Missing, puArgErr=pythoncom.Missing):
#	def OnOpen(self):
#	def OnActivate(self):
#	def OnDeactivate(self):
#	def OnBeforeClose(self, Cancel=defaultNamedNotOptArg):
#	def OnBeforeSave(self, SaveAsUI=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnBeforePrint(self, Cancel=defaultNamedNotOptArg):
#	def OnNewSheet(self, Sh=defaultNamedNotOptArg):
#	def OnAddinInstall(self):
#	def OnAddinUninstall(self):
#	def OnWindowResize(self, Wn=defaultNamedNotOptArg):
#	def OnWindowActivate(self, Wn=defaultNamedNotOptArg):
#	def OnWindowDeactivate(self, Wn=defaultNamedNotOptArg):
#	def OnSheetSelectionChange(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnSheetBeforeDoubleClick(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnSheetBeforeRightClick(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnSheetActivate(self, Sh=defaultNamedNotOptArg):
#	def OnSheetDeactivate(self, Sh=defaultNamedNotOptArg):
#	def OnSheetCalculate(self, Sh=defaultNamedNotOptArg):
#	def OnSheetChange(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnSheetFollowHyperlink(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnSheetPivotTableUpdate(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnPivotTableCloseConnection(self, Target=defaultNamedNotOptArg):
#	def OnPivotTableOpenConnection(self, Target=defaultNamedNotOptArg):
#	def OnSync(self, SyncEventType=defaultNamedNotOptArg):
#	def OnBeforeXmlImport(self, Map=defaultNamedNotOptArg, Url=defaultNamedNotOptArg, IsRefresh=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnAfterXmlImport(self, Map=defaultNamedNotOptArg, IsRefresh=defaultNamedNotOptArg, Result=defaultNamedNotOptArg):
#	def OnBeforeXmlExport(self, Map=defaultNamedNotOptArg, Url=defaultNamedNotOptArg, Cancel=defaultNamedNotOptArg):
#	def OnAfterXmlExport(self, Map=defaultNamedNotOptArg, Url=defaultNamedNotOptArg, Result=defaultNamedNotOptArg):
#	def OnRowsetComplete(self, Description=defaultNamedNotOptArg, Sheet=defaultNamedNotOptArg, Success=defaultNamedNotOptArg):
#	def OnSheetPivotTableAfterValueChange(self, Sh=defaultNamedNotOptArg, TargetPivotTable=defaultNamedNotOptArg, TargetRange=defaultNamedNotOptArg):
#	def OnSheetPivotTableBeforeAllocateChanges(self, Sh=defaultNamedNotOptArg, TargetPivotTable=defaultNamedNotOptArg, ValueChangeStart=defaultNamedNotOptArg, ValueChangeEnd=defaultNamedNotOptArg
#			, Cancel=defaultNamedNotOptArg):
#	def OnSheetPivotTableBeforeCommitChanges(self, Sh=defaultNamedNotOptArg, TargetPivotTable=defaultNamedNotOptArg, ValueChangeStart=defaultNamedNotOptArg, ValueChangeEnd=defaultNamedNotOptArg
#			, Cancel=defaultNamedNotOptArg):
#	def OnSheetPivotTableBeforeDiscardChanges(self, Sh=defaultNamedNotOptArg, TargetPivotTable=defaultNamedNotOptArg, ValueChangeStart=defaultNamedNotOptArg, ValueChangeEnd=defaultNamedNotOptArg):
#	def OnSheetPivotTableChangeSync(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnAfterSave(self, Success=defaultNamedNotOptArg):
#	def OnNewChart(self, Ch=defaultNamedNotOptArg):
#	def OnSheetLensGalleryRenderComplete(self, Sh=defaultNamedNotOptArg):
#	def OnSheetTableUpdate(self, Sh=defaultNamedNotOptArg, Target=defaultNamedNotOptArg):
#	def OnModelChange(self, Changes=defaultNamedNotOptArg):
#	def OnSheetBeforeDelete(self, Sh=defaultNamedNotOptArg):


win32com.client.CLSIDToClass.RegisterCLSID( "{00024412-0000-0000-C000-000000000046}", WorkbookEvents )