# -*- coding: utf-8 -*_
import docx
import os
import wx


class windows(wx.Frame):
    def __init__(self):
        super(windows, self).__init__(None, size=(360, 180),title='only master can use')
        self.panel = wx.Panel(parent=self)
        self.sizer = wx.GridBagSizer(5, 5)
        self.textinpath = wx.TextCtrl(self.panel)
        self.sizer.Add(self.textinpath, pos=(1, 1), span=(1, 15), flag=wx.EXPAND | wx.RIGHT, border=5)
        self.buttonin = wx.Button(self.panel, label='选择输入文件夹')
        self.sizer.Add(self.buttonin, pos=(1, 16), flag=wx.ALIGN_CENTER, border=5)

        self.textoutpath = wx.TextCtrl(self.panel)
        self.sizer.Add(self.textoutpath, pos=(2, 1), span=(1, 15), flag=wx.EXPAND | wx.RIGHT, border=5)
        self.buttonout = wx.Button(self.panel, label='选择输出文件夹')
        self.sizer.Add(self.buttonout, pos=(2, 16), flag=wx.ALIGN_CENTER, border=5)

        self.textsearch = wx.TextCtrl(self.panel)
        self.sizer.Add(self.textsearch, pos=(3, 1), span=(1, 15), flag=wx.EXPAND | wx.RIGHT, border=5)
        self.buttonsearch = wx.Button(self.panel, label='开始搜索关键词')
        self.sizer.Add(self.buttonsearch, pos=(3, 16), flag=wx.ALIGN_CENTER, border=5)

        self.Bind(event=wx.EVT_BUTTON, source=self.buttonin, handler=self.inputdir)
        self.Bind(event=wx.EVT_BUTTON, source=self.buttonout, handler=self.outputdir)
        self.Bind(event=wx.EVT_BUTTON, source=self.buttonsearch, handler=self.search)
        self.SetMinSize(size=(360, 180))
        self.SetMaxSize(size=(360, 180))
        self.panel.SetSizerAndFit(self.sizer)
        self.Show()

    def inputdir(self, event):
        dia = wx.DirDialog(self)
        if dia.ShowModal() == wx.ID_OK:
            self.inpath = dia.GetPath()  # 返回的若含有中文，就返回unicode类型
            self.textinpath.write(self.inpath)

    def outputdir(self, event):
        dia = wx.DirDialog(self)
        if dia.ShowModal() == wx.ID_OK and os.path.isdir(dia.GetPath()):
            self.outpath = dia.GetPath()
            self.textoutpath.write(self.outpath)
        else:
            wx.MessageBox(message='输入的不是文件夹！', style=wx.ICON_ERROR)

    def search(self, event):
        pass
        self.search = self.textsearch.GetValue()
        if self.search == '':
            wx.MessageBox(message='输入为空，请输入关键词', style=wx.ICON_ERROR)
        else:
            for (root, dirs, files) in os.walk(self.inpath):
                for name in files:
                    index = name.rfind('.')
                    filename = root + '\\' + name
                    document = docx.Document(filename)
                    txt = ''
                    for para in document.paragraphs:
                        txt = txt + para.text
                    if self.search in txt:
                        newpath = self.outpath + '\\' + name[:index - 1] + '.docx'
                        document.save(newpath)
                        print name+u'含有关键字：'+self.search
                    # else:
                    #     print name+u'不含有关键字：'+self.search
        print u'完成搜索'
        wx.MessageBox(message=u'完成搜索', style=wx.ICON_INFORMATION)



app = wx.App()
windows()
app.MainLoop()
