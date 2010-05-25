#!/usr/bin/env python


import sys, os
import shutil
import wx
import time
import tarfile

from threading import *

from elaphe import barcode

from pyPdf import PdfFileReader, PdfFileWriter
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfgen import canvas

import sqlite3

# Global functions
ID_START = wx.NewId()
ID_STOP = wx.NewId()

EVT_RESULT_ID = wx.NewId()

def EVT_RESULT(win, func):
	"""Define Result Event."""
	win.Connect(-1, -1, EVT_RESULT_ID, func)

def returnListOfFiles(dir):
	return os.listdir(dir)

def _mkdir(newdir):
	"""works the way a good mkdir should :)
	- already exists, silently complete
	- regular file in the way, raise an exception
	- parent directory(ies) does not exist, make them as well
	"""
	if os.path.isdir(newdir):
		pass
	elif os.path.isfile(newdir):
		raise OSError("a file with the same name as the desired " \
					"dir, '%s', already exists." % newdir)
	else:
		head, tail = os.path.split(newdir)
		if head and not os.path.isdir(head):
			_mkdir(head)
		if tail:
			os.mkdir(newdir)

def bkFiles(source, destination):
	targetBackup = destination + time.strftime('%Y%m%d%H%M%S') + '.tar.gz'
	tar = tarfile.open(targetBackup, "w:gz")
	tar.add(source)
	tar.close()

#---------------------------------------------------------------------------------------

class logger:
	def __init__(self):
		wxSP = wx.StandardPaths.Get()
		self.path = path = wxSP.GetDocumentsDir() + '/BarCodeLabeling/BarCodeLabeling.db'
		if not os.path.isfile(path):
			conn = sqlite3.connect(self.path)
			c = conn.cursor()
			c.execute("""
				CREATE TABLE FICHAS
				(
					nm_paciente text, dt_ficha text
				)
			""")
			conn.commit()
			c.close()
	def logFicha(self, name):
		conn = sqlite3.connect(self.path)
		c = conn.cursor()
		c.execute("""
			INSERT INTO FICHAS(rowid, nm_paciente, dt_ficha)
			VALUES (NULL, '%s', '%s')
		"""%(name,  time.strftime('%Y-%m-%d %H:%M:%S'),))
		conn.commit()
		id = c.lastrowid
		c.close()
		return id
class barcode_generator:
	"""
		Class for creating bar code
	"""
	def __init__(self, id):
		self.id = id
	def get(self):
		ps = barcode('ean8', '%08d'%self.id, options=dict(version=9, eclevel='M'), margin=10)
		return ps

class ResultEvent(wx.PyEvent):
	"""Simple event to carry arbitrary result data."""
	def __init__(self, data):
		"""Init Result Event."""
		wx.PyEvent.__init__(self)
		self.SetEventType(EVT_RESULT_ID)
		self.data = data


class WorkerThread(Thread):
	def __init__(self, notify_window):
		Thread.__init__(self)
		self._notify_window = notify_window
		self._want_abort = 0
		self.start()
	def run(self):
		"""Run Worker Thread."""
		lg = logger()
		dir = self._notify_window.pdfDir.GetValue()
		wxSP = wx.StandardPaths.Get()
		bkDir = wxSP.GetDocumentsDir() + '/BarCodeLabeling/.backup/'
		if not os.path.isdir(bkDir):
			_mkdir(bkDir)
		listOfFiles = returnListOfFiles(dir)
		if True: #In future this can be an user's option
			bkFiles(dir, bkDir)

		for pdf_file in listOfFiles:
			if self._want_abort:
				wx.PostEvent(self._notify_window, ResultEvent(None))
				return
			#Create barcode
			name = pdf_file.rsplit('.', 1)
			id = lg.logFicha(name[0])
			bc = barcode_generator(id)
			ps = bc.get()
			ps.save('.bc.ps')
			#Create barcode label page
			c = canvas.Canvas('.temp.pdf', pagesize=letter)
			width, height = letter
			c.drawImage('.bc.ps', inch/2, height - 1.5* inch)
			c.save()
			input  = PdfFileReader(file(dir +'/'+ pdf_file, 'rb'))
			output = PdfFileWriter()
			bc_pdf = PdfFileReader(file('.temp.pdf', 'rb'))
			for npage in range(input.getNumPages()):
				page = input.getPage(npage)
				page.mergePage(bc_pdf.getPage(0))
				output.addPage(page)
			outputStream = file(dir +'/'+ name[0] + '_1.pdf', "wb")
			output.write(outputStream)
			os.remove('.temp.pdf')
			os.remove('.bc.ps')
			shutil.move(dir +'/'+ name[0] + '_1.pdf', dir+'/'+pdf_file)
			del bc
		wx.PostEvent(self._notify_window, ResultEvent(10))

	def abort(self):
		"""abort worker thread."""
		self._want_abort = 1

class MyFileDropTarget(wx.FileDropTarget):
	def __init__(self, window):
		wx.FileDropTarget.__init__(self)
		self.window = window

	def OnDropFiles(self, x, y, filenames):
		self.window.SetInsertionPointEnd()
		for file in filenames:
			self.window.WriteText(file)

class JoinerPanel(wx.Panel):
	def __init__(self, parent):
		"""Constructor"""
		wx.Panel.__init__(self, parent=parent)
		lblSize = (70,-1)
		pdfLblDir = wx.StaticText(self, label=u"Arquivo:", size=lblSize)
		self.pdfDir = wx.TextCtrl(self)
		dt = MyFileDropTarget(self.pdfDir)
		self.pdfDir.SetDropTarget(dt)
		pdfDirBtn = wx.Button(self, label="Browse", name="pdfDirBtn")
		pdfDirBtn.Bind(wx.EVT_BUTTON, self.onBrowse)

		widgets = [(pdfLblDir, self.pdfDir, pdfDirBtn)]
		joinBtn = wx.Button(self, label="Etiquetar Fichas")
		joinBtn.Bind(wx.EVT_BUTTON, self.onJoinPdfs)

		# Set up event handler for any worker thread results
		EVT_RESULT(self,self.OnResult)
		self.mainSizer = wx.BoxSizer(wx.VERTICAL)
		for widget in widgets:
			self.buildRows(widget)
		self.mainSizer.Add(joinBtn, 0, wx.ALL|wx.CENTER, 5)
		self.SetSizer(self.mainSizer)
		self.worker = None

	def buildRows(self, widgets):
		""""""
		sizer = wx.BoxSizer(wx.HORIZONTAL)
		for widget in widgets:
			if isinstance(widget, wx.StaticText):
				sizer.Add(widget, 0, wx.ALL|wx.CENTER, 5)
			elif isinstance(widget, wx.TextCtrl):
				sizer.Add(widget, 1, wx.ALL|wx.EXPAND, 5)
			else:
				sizer.Add(widget, 0, wx.ALL, 5)
		self.mainSizer.Add(sizer, 0, wx.EXPAND)
 
	def onBrowse(self, event):
		"""
		Browse for PDFs
		"""
		widget = event.GetEventObject()
		name = widget.GetName()
 
		dlg = wx.DirDialog(
			self, message="Escolha um arquivo",
			style=wx.OPEN | wx.CHANGE_DIR
			)
		if dlg.ShowModal() == wx.ID_OK:
			path = dlg.GetPath()
			if name == "pdfDirBtn":
				self.pdfDir.SetValue(path)
			self.currentPath = os.path.dirname(path)
		dlg.Destroy()
 
	def onJoinPdfs(self, event):
		"""
		Join the two PDFs together and save the result to the desktop
		"""
		pdfDir = self.pdfDir.GetValue()
		if not os.path.exists(pdfDir):
			msg = "A escolha do arquivo %s foi errada " % pdfDir
			dlg = wx.MessageDialog(None, msg, 'Erro', wx.OK|wx.ICON_EXCLAMATION)
			dlg.ShowModal()
			dlg.Destroy()
			return
		if not self.worker:
			self.worker = WorkerThread(self)

	def OnStop(self, event):
		"""Stop Computation."""
		# Flag the worker thread to stop if running
		if self.worker:
			self.worker.abort()

	def OnResult(self, event):
		"""Show Result status."""
		if event.data is None:
			# Thread aborted (using our convention of None return)
			msg = 'Fim da operacao'
			dlg = wx.MessageDialog(None, msg, 'ERRO', wx.OK|wx.ICON_INFORMATION)
			dlg.ShowModal()
			dlg.Destroy()
		else:
			msg = 'Fim da operacao'
			dlg = wx.MessageDialog(None, msg, 'Fichas Cadastradas', wx.OK|wx.ICON_INFORMATION)
			dlg.ShowModal()
			dlg.Destroy()
		self.worker = None

class JoinerFrame(wx.Frame):
	def __init__(self):
		wx.Frame.__init__(self, None, wx.ID_ANY,
						 "Barcode Labeling", size=(550, 200))
		panel = JoinerPanel(self)

#----------------------------------------------------------------------
# Run the program
if __name__ == "__main__":
	app = wx.App(False)
	frame = JoinerFrame()
	frame.Show()
	app.MainLoop()

