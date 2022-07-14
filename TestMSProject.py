import win32com.client

doc = 'C:\\Work\\pp.mpp'


try:
  mpp = win32com.client.Dispatch("MSProject.Application")
  mpp.Visible = 1
  
  try:
    mpp.FileOpen(doc)
    proj = mpp.ActiveProject
    print (proj.BuiltinDocumentProperties(1))
    print (proj.BuiltinDocumentProperties(2))
    print (proj.BuiltinDocumentProperties(3))
    print (proj.BuiltinDocumentProperties(11))
    print (proj.BuiltinDocumentProperties(12))

    print (proj.Tasks.Item(3).Name)
    print (proj.Tasks.Item(3).Start)

    print (proj.Tasks.count)

  except Exception as e:
    print ("Error" + e)

  mpp.FileSave()
  mpp.Quit()
except Exception as e:
  print ("Error opening file" + e)
