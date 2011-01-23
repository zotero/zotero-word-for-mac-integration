#!/usr/bin/env python

"""
    ***** BEGIN LICENSE BLOCK *****
	
	Copyright (c) 2009  Zotero
	                    Center for History and New Media
						George Mason University, Fairfax, Virginia, USA
						http://zotero.org
	
	This program is free software: you can redistribute it and/or modify
	it under the terms of the GNU General Public License as published by
	the Free Software Foundation, either version 3 of the License, or
	(at your option) any later version.
	
	This program is distributed in the hope that it will be useful,
	but WITHOUT ANY WARRANTY; without even the implied warranty of
	MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
	GNU General Public License for more details.
	
	You should have received a copy of the GNU General Public License
	along with this program.  If not, see <http://www.gnu.org/licenses/>.
    
    ***** END LICENSE BLOCK *****
"""

from xpcom import components, ServerException, nsError
import sys, os, subprocess

# Make sure exec_prefix is corect
sys.exec_prefix = sys.prefix
# Hack to fix pylib loading errors
pylib = os.path.realpath(os.path.dirname(__file__)+"/../pylib")
if not pylib in sys.path:
	# If no pylib not in path, add it
	sys.path.insert(0, pylib)
elif sys.path[-1].endswith('zoteroMacWordIntegration@zotero.org/pylib'):
	# Move appscript to start of path
	sys.path.insert(0, sys.path.pop())

from MacWord import Document, Document_2004, Field, Bookmark
import appscript

class Application_2008:
	_com_interfaces_ = [components.interfaces.zoteroIntegrationApplication]
	_reg_clsid_ = "{ea584d70-2797-4cd1-8015-1a5f5fb85af7}"
	_reg_contractid_ = "@zotero.org/Zotero/integration/application?agent=MacWord2008;1"
	_reg_desc_ = "MacWord 2008 Application"
	primaryFieldType = "Field"
	secondaryFieldType = "Bookmark"
	
	def getDocument(self, wordPath):
		return Document(appscript.app(wordPath))
	
	def getActiveDocument(self):
		return Document(appscript.app(id='com.Microsoft.Word'))

class Application_2004:
	_com_interfaces_ = [components.interfaces.zoteroIntegrationApplication]
	_reg_clsid_ = "{b063dd87-5615-45c5-ac3d-4b0583034616}"
	_reg_contractid_ = "@zotero.org/Zotero/integration/application?agent=MacWord2004;1"
	_reg_desc_ = "MacWord 2004 Application"
	primaryFieldType = "Field"
	secondaryFieldType = "Bookmark"
	
	def getDocument(self, wordPath):
		return Document(appscript.app(wordPath))
	
	def getActiveDocument(self):
		return Document_2004(appscript.app('Microsoft Word'))