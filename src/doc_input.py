#!/usr/bin/env python
# vim:fileencoding=utf-8
from __future__ import (unicode_literals, division, absolute_import, print_function)

__license__ = 'GPL v3'
__copyright__ = '2013, Kovid Goyal <kovid at kovidgoyal.net>'

from PyQt5 import (QtCore, QtWidgets)
from PyQt5.Qt import QFileDialog
from calibre.gui2.convert import Widget
from calibre.gui2.preferences.conversion import InputOptions as BaseInputOptions
from calibre.gui2 import SanitizeLibraryPath
import os


class PluginWidget(Widget):

    TITLE = _('DOC Input')
    COMMIT_NAME = 'doc_input'
    ICON = I('mimetypes/docx.png')
    HELP = _('Options specific to the input format.')

    def __init__(self, parent, get_option, get_help, db=None, book_id=None):
        self.db = db                # db is set for conversion, but not default preferences
        self.book_id = book_id      # book_id is set for individual conversion, but not bulk

        Widget.__init__(self, parent, ['docx_no_cover', 'wordconv_exe_path'])
        self.initialize_options(get_option, get_help, db, book_id)


    def setupUi(self, Form):
        Form.setObjectName("Form")
        Form.setWindowTitle("Form")
        Form.resize(518, 353)
        self.verticalLayout = QtWidgets.QVBoxLayout(Form)
        self.verticalLayout.setSizeConstraint(QtWidgets.QLayout.SetDefaultConstraint)
        self.verticalLayout.setObjectName("verticalLayout")
        self.label = QtWidgets.QLabel(Form)
        self.label.setWordWrap(True)
        self.label.setObjectName("label")
        self.verticalLayout.addWidget(self.label)
        self.hboxlayout = QtWidgets.QHBoxLayout()
        self.hboxlayout.setObjectName("hboxlayout")
        self.opt_wordconv_exe_path = QtWidgets.QLineEdit(Form)
        self.opt_wordconv_exe_path.setObjectName("opt_wordconv_exe_path")
        self.hboxlayout.addWidget(self.opt_wordconv_exe_path)
        self.fileChoose = QtWidgets.QPushButton(Form)
        self.fileChoose.setObjectName("fileChoose")
        self.fileChoose.clicked.connect(self.fileSearch)
        self.hboxlayout.addWidget(self.fileChoose)
        self.verticalLayout.addLayout(self.hboxlayout)
        self.opt_docx_no_cover = QtWidgets.QCheckBox(Form)
        self.opt_docx_no_cover.setObjectName("opt_docx_no_cover")
        self.verticalLayout.addWidget(self.opt_docx_no_cover)
        spacerItem = QtWidgets.QSpacerItem(20, 40, QtWidgets.QSizePolicy.Minimum, QtWidgets.QSizePolicy.Expanding)
        self.verticalLayout.addItem(spacerItem)

        self.label.setText("<html><head/><body><p>LibreOffice conversion program soffice.exe. Usually is placed in &quot;C:\\Program Files\\LibreOffice\\program\\soffice.exe&quot;. You can dowload it from <a href=\"https://www.libreoffice.org\"><span style=\" text-decoration: underline; color:#0000ff;\">www.libreoffice.org.</span></a></p></body></html>")
        self.fileChoose.setText("...")
        self.opt_docx_no_cover.setText("Do not try to autodetect a &cover from images in the document")
        QtCore.QMetaObject.connectSlotsByName(Form)


    def fileSearch(self):
        filterString = 'soffice.exe (soffice.exe)'

        with SanitizeLibraryPath():
            parent = self.parent()
            if parent is None:
                raise ValueError('parent is None')
            selectedFile = QFileDialog.getOpenFileName(parent=self, caption=_('Find soffice.exe'), directory=self.opt_wordconv_exe_path.text(), filter=filterString)
            if selectedFile:
                selectedFile = selectedFile[0] if isinstance(selectedFile, tuple) else selectedFile
                if selectedFile and os.path.exists(selectedFile):
                    self.opt_wordconv_exe_path.setText(selectedFile)
