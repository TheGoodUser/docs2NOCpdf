import dataExtractor as dxe
from os import mkdir, listdir, system, path

try:
    import pandas as pd # type: ignore
    from docx import Document # type: ignore 
    import docxedit # type: ignore
except:
    system(f"pip install -r requirements.txt")



# creating a output folder
if not path.exists("outputs"):
    mkdir("outputs")

dxe.dataRetrieverFromExcel()