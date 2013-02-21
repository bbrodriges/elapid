# -*- coding: utf-8 -*-

# STANDARD MODULES
import os
import zipfile
import subprocess

# THIRD-PARTY MODULES
import comtypes.client
import pyPdf

class Elapid:

    """ Elapid is a universal file to images converter. """

    def __init__(self, file, folder):

        """
            @type file str
            @type folder str
            @return void

            Instanciates Elapid.
        """

        # default conversion successfulness
        self.success = False

        # checking passed params
        if os.path.isfile(file) and os.path.isdir(folder):

            # list of convertable types of files
            convertables = ['ppt', 'pptx', 'pdf', 'zip', 'ps']
            # list of acceptable types from zip archive
            self.acceptables = ['jpg', 'jpeg', 'png', 'gif']

            # checking file type
            type = self._get_type(file)
            if type in convertables:
                # deciding what conversion method to use
                if type in ['pptx', 'ppt']:
                    self.success = self.ppt(file, folder)
                elif type in ['pdf', 'ps']:
                    self.success = self.pdf_ps(file, folder)
                elif type == 'zip':
                    self.success = self.zip(file, folder)

    def ppt(self, file, folder):

        """
            @type file str
            @type folder str
            @return bool

            Converts .ppt/.pptx file.
        """

        # creating powerpoint object
        powerpoint = comtypes.client.CreateObject("Powerpoint.Application")
        if powerpoint:
            # opening powerpoint window (not working without this line)
            powerpoint.Visible = True
            # opening presentation file with powerpoint
            powerpoint.Presentations.Open(file)
            # exporting presentation to desired format
            powerpoint.ActivePresentation.Export(folder, 'jpg')
            # closing presentation
            powerpoint.Presentations[1].Close()
            # closing powerpoint and checking exit status
            exit_code = powerpoint.Quit()
            if exit_code == 0:
                return True
        return False

    def pdf_ps(self, file, folder):

        """
            @type file str
            @type folder str
            @return bool

            Converts .pdf or PostScript file.
        """

        # adding quotations to path to avoid spaces
        file, folder = r'"%s"' % file, r'"%s"' % folder
        # preparing command
        gs_path = r'"C:\Program Files\gs\gs9.06\bin\gswin64c.exe"'
        command = r'%s -dNOPAUSE -dBATCH -dSAFER -sDEVICE=jpeg -dJPEG=80 -r96 -dTextAlphaBits=4 -dGraphicsAlphaBits=4 -sOutputFile=%s %s'
        command = command % (gs_path, folder+r"\Slide%d.jpg", file)
        # running command
        child = subprocess.Popen(command, shell=True, stdout=subprocess.PIPE)
        streamdata = child.communicate()[0]
        # checking return code
        if child.returncode == 0:
            return True
        return False

    def zip(self, file, folder):

        """
            @type file str
            @type folder str
            @return bool

            Extracts images from zip file.
        """

        zip_file = zipfile.ZipFile(file, 'r')
        for inner_file in zip_file.namelist():
            # getting extension of file in zip archive
            extension = os.path.splitext(inner_file)[1][1:]
            if extension in self.acceptables:
                zip_file.extract(inner_file, folder)

        return True

    def _get_type(self, file):

        """
            @type file str
            @return str or None

            Confirms type of file.
        """

        # default result
        result = None
        # getting file extension
        extension = os.path.splitext(file)[1][1:]
        # deep check of file
        if extension == 'pptx':
            opened = zipfile.ZipFile(file, 'r')
            # every .pptx contains 'ppt/presentation.xml' file
            result = extension if 'ppt/presentation.xml' in opened.namelist() else None
        elif extension == 'ppt':
            opened = open(file, 'r')
            first_line = '%r' % opened.readline()
            # '\xd0\xcf\x11\xe0\xa1\xb1' is a first line of all PPTs
            result = extension if first_line == r"'\xd0\xcf\x11\xe0\xa1\xb1'" else None
            opened.close()
        elif extension == 'zip':
            result = extension if zipfile.is_zipfile(file) else None
        elif extension == 'pdf':
            # if pyPdf can open it - this is PDF
            result = extension if pyPdf.PdfFileReader(open(file, 'rb')) else None
        elif extension == 'ps':
            result = extension

        return result

input = r'C:\Users\bender\PycharmProjects\comtypes_test\slides.zip'
output = r'C:\Users\bender\PycharmProjects\comtypes_test\out'
converter = Elapid(input, output)
