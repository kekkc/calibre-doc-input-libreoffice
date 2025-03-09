#!/usr/bin/env python
# vim:fileencoding=utf-8
from calibre.ptempfile import PersistentTemporaryDirectory
from pathlib import Path
import subprocess # Import subprocess module
import os # We will use the exists() function from this module to know if the file was created.

__license__ = 'GPL v3'
__copyright__ = '2025, kekkc / originally by David Ignjic <ignjic at gmail.com>'


from calibre.customize.conversion import InputFormatPlugin, OptionRecommendation

from calibre.customize.builtins import plugins
     
class DOCInput(InputFormatPlugin):

    name = 'DOC Input'
    author = 'kekkc / originally by David Ignjic'
    description = _('DOC-File (.doc) conversion to docx & HTML via LibreOffice (enabling full text search)')
    supported_platforms = ['windows']
    file_types = {'doc'}
    minimum_calibre_version = (7, 0, 0)
    version = (2, 0, 1)

    options = {
        OptionRecommendation(name='wordconv_exe_path', recommended_value='C:\Program Files\LibreOffice\program\soffice.exe', 
            help=_('Path to LibreOffice. Usually it is in "C:\Program Files\LibreOffice\program\soffice.exe". If you don\'t have it you can download it from https://www.libreoffice.org ')),
        OptionRecommendation(name='docx_no_cover', recommended_value=False,
            help=_('Normally, if a large image is present at the start of the document that looks like a cover, '
                   'it will be removed from the document and used as the cover for created ebook. This option '
                   'turns off that behavior.')),

    }

    recommendations = set([('page_breaks_before', '/', OptionRecommendation.MED)])

    def gui_configuration_widget(self, parent, get_option_by_name,
        get_option_help, db, book_id=None):
        from calibre_plugins.doc_input.doc_input import PluginWidget
        return PluginWidget(parent, get_option_by_name, get_option_help, db, book_id)

    def convert(self, stream, options, file_ext, log, accelerators):
        from calibre.ebooks.docx.to_html import Convert
        doc_temp_directory = PersistentTemporaryDirectory('doc_input')
        log.debug('Convert doc ' + stream.name + ' to docx via ' + options.wordconv_exe_path)
        
        doc_file = Path(os.path.join(doc_temp_directory,os.path.basename(stream.name)))
        docx_file = str(doc_file.with_suffix('.docx'))
        log.debug('Temp directory '+ doc_temp_directory + ' temp output file'+docx_file)
        stream.close()
        if  not os.path.exists(options.wordconv_exe_path):
            raise ValueError('Not found ' + options.wordconv_exe_path)

        def convert_doc_to_odt(LibreOffice,file_path, output_dir):
            subprocess.run(
                    [
                        LibreOffice,
                        "--headless",
                        "--convert-to",
                        "docx",
                        file_path,
                        "--outdir",
                        output_dir,
                    ],
                    cwd=os.getcwd(),
                    shell=True,
                    check=True,
                )

        convert_doc_to_odt(options.wordconv_exe_path,stream.name, doc_temp_directory)

        if  not os.path.exists(docx_file):
            raise ValueError('Not converted ' + docx_file)
        
        return Convert(docx_file, detect_cover=not options.docx_no_cover, log=log)()