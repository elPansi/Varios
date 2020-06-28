# -*- coding: utf-8 -*-

from distutils.core import setup 
import py2exe 
 
setup(name="appEng", 
 version="1.0", 
 description="Cambia la configuracion del idioma", 
 author="Fernando", 
 author_email="fernandohc@gmail.com", 
 url="", 
 license="open", 
 scripts=["appLang.py"], 
 windows=["appLang.py"], 
 options={"py2exe": {"bundle_files": 1}}, 
 zipfile=None,
)