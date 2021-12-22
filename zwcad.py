from pyzwcad import ZwCAD, APoint
import win32com.client
import os
import time

acad = win32com.client.Dispatch("ZWCAD.Application")
root = 'C:\\Users\\patrick.vieira\\Documents\\PY\\Zwcad'
#acad = ZwCAD()
#print(acad.doc.Name)


for file in os.listdir(root):
    if file.endswith('.dwg'):
        print(file)
        print(os.path.join(root, file))
        acad.Documents.Open(os.path.join(root,file))
        time.sleep(2)
        for entity in acad.ActiveDocument.PaperSpace:
            name = entity.EntityName
            if name == 'AcDbBlockReference':
                HasAttributes = entity.HasAttributes
                if HasAttributes:
                    for attrib in entity.GetAttributes():
                        print("  {}: {}".format(attrib.TagString, attrib.TextString))
                        #print(attrib.TextString)
        acad.Documents.Close()