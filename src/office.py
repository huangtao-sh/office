#!/usr/bin/env python3
from win32com.client import Dispatch

class Application:
    def __init__(self):
        self.application=Dispatch(self.Object)

    def __del__(self):
        self.application=None

class Workbook(Application):
    Object="Excel.Application"
    def __init__(self,file_name=None,template=''):
        super().__init__()
        books=self.application.Workbooks
        if file_name:
            self.workbook=books.Open(file_name)
        else:
            self.workbook=books.Add(Template=template)
        
    @property
    def sheets(self):
        return self.workbook.Worksheets

    @property
    def sheet_name(self):
        return self.active_sheet.Name

    @property
    def active_sheet(self):
        return self.application.ActiveSheet
    
    def iter_sheets(self):
        for i in range(1,self.sheets.Count+1):
            self.activate_sheet(i)
            yield self.sheet_name

    def iter_rows(self):
        for row in self.active_sheet.UsedRange.Rows:
            yield row

    def iter_columns(self):
        for column in self.active_sheet.UsedRange.Columns:
            yield column        
    
    def activate_sheet(self,index):
        self.sheets(index).Activate()
        
    def cells(self,row,column):
        return self.application.Cells(row,column)
          
    def range(self,ref=None,row=1,column=1):
        if ref:
            return self.application.Range(ref)
        else:
            return self.application.Cells(row,column)

    def text(self,ref=None,row=1,column=1):
        return self.range(ref,row,column).Text

    def value(self,ref=None,row=1,column=1):
        return self.range(ref,row,column).Value

    def formula(self,ref=None,row=1,column=1):
        return self.range(ref,row,column).Formula

    def set_value(self,ref=None,row=1,column=1,value=None):
        self.range(ref,row,column).Value=value

    def set_formula(self,ref=None,row=1,column=1,formula=None):
        self.range(ref,row,column).Formula=formula

    def save_as(self,file_name):
        self.workbook.SaveAs(file_name)

    def save(self):
        self.workbook.Save()

    def close(self):
        self.workbook.Close()
        
class Document(Application):
    Object='Word.Application'
    def __init__(self,file_name=None,template=''):
        super().__init__()
        docs=self.application.Documents
        if file_name:
            self.document=docs.Open(file_name)
        else:
            self.document=docs.Add(Template=template)
        self.selection=self.application.Selection

    def pstyle(self,style):
        self.selection.ParagraphFormat.Style=style

    def style(self,style):
        self.selection.Style=style

    @property
    def font(self):
        return self.selection.Font

    def list_gallery(self,gallery=1,index=1):
        self.selection.Range.ListFormat.\
        ApplyListTemplateWithLevel(
            ListTemplate=self.application.\
            ListGalleries(gallery).ListTemplates(index),
            ContinuePreviousList=0,
            ApplyTo=0,
            DefaultListBehavior=2
            )

    def list_indent(self):
        self.selection.Range.ListFormat.ListIndent()

    def list_outdent(self):
        self.selection.Range.ListFormat.ListOutdent()
        
    def typetext(self,text):
        self.selection.TypeText(text)

    def newp(self):
        self.selection.TypeParagraph()

    def save_as(self,file_name):
        self.document.SaveAs(file_name)

    def save(self):
        self.document.Save()

    def close(self):
        self.document.Close(0)
        self.document=None

if __name__=='__main__':
    xls=Workbook('d:/全行通讯录.xls')
    try:
        for name in xls.iter_sheets():
            if name=='全行机构地址':
                for row in xls.iter_rows():
                    print(row.cells(1,1).Value)
                    print(row.cells(1,6).Text)
            
    finally:
        xls.close()
