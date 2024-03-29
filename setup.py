#!/usr/bin/python
# -*- coding: utf-8 -*-
# Python 3
  
"""
Icone sous Windows: il faut:
=> un xxx.ico pour integration dans le exe, avec "icon=xxx.ico"
=> un xxx.png pour integration avec PyQt4 + demander la recopie avec includefiles.
"""
  
import sys, os
from cx_Freeze import setup, Executable
  
#############################################################################
# preparation des options
  
# chemins de recherche des modules
# ajouter d'autres chemins (absolus) si necessaire: sys.path + ["chemin1", "chemin2"]
path = sys.path
# options d'inclusion/exclusion des modules
includes = []  # nommer les modules non trouves par cx_freeze
excludes = []
packages = []  # nommer les packages utilises
  
# copier les fichiers non-Python et/ou repertoires et leur contenu:
includefiles = ["E:\Lab\Denomb\icon.ico"]
  
if sys.platform == "win32":
    pass
    # includefiles += [...] : ajouter les recopies specifiques à Windows
elif sys.platform == "linux2":
    pass
    # includefiles += [...] : ajouter les recopies specifiques à Linux
else:
    pass
    # includefiles += [...] : cas du Mac OSX non traite ici
  
# pour que les bibliotheques binaires de /usr/lib soient recopiees aussi sous Linux
binpathincludes = []
if sys.platform == "linux2":
    binpathincludes += ["/usr/lib"]
  
# niveau d'optimisation pour la compilation en bytecodes
optimize = 0
  
# si True, n'affiche que les warning et les erreurs pendant le traitement cx_freeze
silent = False
  
# construction du dictionnaire des options
options = {"path": path,
           "includes": includes,
           "excludes": excludes,
           "packages": packages,
           "include_files": includefiles,
           "bin_path_includes": binpathincludes,
           #"create_shared_zip": False,  # <= ne pas generer de fichier zip
           #"include_in_shared_zip": False,  # <= ne pas generer de fichier zip
           #"compressed": False,  # <= ne pas generer de fichier zip
           "optimize": optimize,
           "silent": False
           }
  

  
#############################################################################
# preparation des cibles
base = None
if sys.platform == "win32":
    base = "Win32GUI"  # pour application graphique sous Windows
    # base = "Console" # pour application en console sous Windows
  
icone = None
if sys.platform == "win32":
    icone = "icon.ico"
  
cible_1 = Executable(
    script="denombrement.py",
    base=base,
    #compress=False,  # <= ne pas generer de fichier zip
    #copyDependentFiles=True,
    #appendScriptToExe=True,
    #appendScriptToLibrary=False,  # <= ne pas generer de fichier zip
    icon= icone
    #icon_resources: [(1, "test.ico")]
    )
  
#cible_2 = Executable(
 #   script="monprogramme2.pyw",
  #  base=base,
   # compress=False,
    #copyDependentFiles=True,
    #appendScriptToExe=True,
    #appendScriptToLibrary=False,
    #icon=icone
    #)
  
#############################################################################
# creation du setup
setup(
    name = "Dénombrement_1.2.1",
    version = "1.2.1",
    description = "Calcul le nombre de ufc/g d'un ou groupe de microbe se trouvant dans un produit",
    author = "Ibrahima Gaye",
    options = {"build_exe":options},
    executables = [cible_1]
    )