#!d:\pdf\env\scripts\python.exe

"""
Converts PDF text content (though not images containing text) to plain text, html, xml or "tags".
"""

from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfparser import PDFParser
from pdfminer.pdfinterp import PDFResourceManager, PDFPageInterpreter
from pdfminer.pdfdevice import PDFDevice, TagExtractor
from pdfminer.pdfpage import PDFPage
from pdfminer.converter import XMLConverter, HTMLConverter, TextConverter
from pdfminer.cmapdb import CMapDB
from pdfminer.image import ImageWriter

import io
import argparse
import logging
import six
import sys
import pdfminer.settings
pdfminer.settings.STRICT = False
#import pdfminer.high_level
import pdfminer.layout
from pdfminer.image import ImageWriter

def extract_text_to_fp(inf, outfp,
                    _py2_no_more_posargs=None,  # Bloody Python2 needs a shim
                    output_type='text', codec='utf-8', laparams = None,
                    maxpages=0, page_numbers=None, password="", scale=1.0, rotation=0,
                    layoutmode='normal', output_dir=None, strip_control=False,
                    debug=False, disable_caching=False, **other):
    """
    Parses text from inf-file and writes to outfp file-like object.
    Takes loads of optional arguments but the defaults are somewhat sane.
    Beware laparams: Including an empty LAParams is not the same as passing None!
    Returns nothing, acting as it does on two streams. Use StringIO to get strings.
    
    output_type: May be 'text', 'xml', 'html', 'tag'. Only 'text' works properly.
    codec: Text decoding codec
    laparams: An LAParams object from pdfminer.layout.
        Default is None but may not layout correctly.
    maxpages: How many pages to stop parsing after
    page_numbers: zero-indexed page numbers to operate on.
    password: For encrypted PDFs, the password to decrypt.
    scale: Scale factor
    rotation: Rotation factor
    layoutmode: Default is 'normal', see pdfminer.converter.HTMLConverter
    output_dir: If given, creates an ImageWriter for extracted images.
    strip_control: Does what it says on the tin
    debug: Output more logging data
    disable_caching: Does what it says on the tin
    """
    if six.PY2 and sys.stdin.encoding:
        password = password.decode(sys.stdin.encoding)

    imagewriter = None
    if output_dir:
        imagewriter = ImageWriter(output_dir)
    
    rsrcmgr = PDFResourceManager(caching=not disable_caching)

    if output_type == 'text':
        device = TextConverter(rsrcmgr, outfp, codec=codec, laparams=laparams,
                               imagewriter=imagewriter)

    if six.PY3 and outfp == sys.stdout:
        outfp = sys.stdout.buffer

    if output_type == 'xml':
        device = XMLConverter(rsrcmgr, outfp, codec=codec, laparams=laparams,
                              imagewriter=imagewriter,
                              stripcontrol=strip_control)
    elif output_type == 'html':
        device = HTMLConverter(rsrcmgr, outfp, codec=codec, scale=scale,
                               layoutmode=layoutmode, laparams=laparams,
                               imagewriter=imagewriter)
    elif output_type == 'tag':
        device = TagExtractor(rsrcmgr, outfp, codec=codec)

    interpreter = PDFPageInterpreter(rsrcmgr, device)

    #增加一个空列表用于保存每页提取的文本内容
    t_list= []
    seek_pos = 0 #缓冲区首指针
    for page in PDFPage.get_pages(inf,
                                  page_numbers,
                                  maxpages=maxpages,
                                  password=password,
                                  caching=not disable_caching,
                                  check_extractable=True):
        page.rotate = (page.rotate + rotation) % 360
        interpreter.process_page(page)
        #将指针定位到未读取位置
        outfp.seek(seek_pos)
        #将提取的文本内容保存到列表中 
        t_list.append(outfp.read())
        #将read()后的缓冲区底部指针保存，以备下次从此位置读取文本
        seek_pos = outfp.tell() 
    device.close()
    #增加了一个返回列表
    return t_list

def extract_text(files=[], outfile='-',
            _py2_no_more_posargs=None,  # Bloody Python2 needs a shim
            no_laparams=False, all_texts=None, detect_vertical=None, # LAParams
            word_margin=None, char_margin=None, line_margin=None, boxes_flow=None, # LAParams
            output_type='text', codec='utf-8', strip_control=False,
            maxpages=0, page_numbers=None, password="", scale=1.0, rotation=0,
            layoutmode='normal', output_dir=None, debug=False,
            disable_caching=False, **other):
    if _py2_no_more_posargs is not None:
        raise ValueError("Too many positional arguments passed.")
    if not files:
        raise ValueError("Must provide files to work upon!")

    # If any LAParams group arguments were passed, create an LAParams object and
    # populate with given args. Otherwise, set it to None.
    if not no_laparams:
        laparams = pdfminer.layout.LAParams()
        for param in ("all_texts", "detect_vertical", "word_margin", "char_margin", "line_margin", "boxes_flow"):
            paramv = locals().get(param, None)
            if paramv is not None:
                setattr(laparams, param, paramv)
    else:
        laparams = None

    imagewriter = None
    if output_dir:
        imagewriter = ImageWriter(output_dir)

    if output_type == "text" and outfile != "-":
        for override, alttype in (  (".htm", "html"),
                                    (".html", "html"),
                                    (".xml", "xml"),
                                    (".tag", "tag") ):
            if outfile.endswith(override):
                output_type = alttype

    if outfile == "-":
        #outfp = sys.stdout
        #将原定向于屏幕的文本流，重定向为StringIO字符串流
        outfp = io.StringIO()
        if outfp.encoding is not None:
            codec = 'utf-8'
    else:
        outfp = open(outfile, "wb")

    for fname in files:
        with open(fname, "rb") as fp:
            #pdfminer.high_level.extract_text_to_fp(fp, **locals())
            #经过合并，extract_text_to_fp是在同一文件的函数
            #因此去掉pdfminer.high_level，直接引用
            #并增加了获得返回文本的列表
            t_list = extract_text_to_fp(fp, **locals())
    #return outfp
    #增加了一个返回列表
    return outfp,t_list