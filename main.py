#!/usr/bin/env python


import wx
import fitz
import os
import time


def ts():
    return time.strftime('%Y%m%d%H%M%S', time.localtime(time.time()))


class Task(object):
    def __init__(self, file, start, end):
        self.file = file
        self.start = start
        self.end = end


class RagdollApp(wx.App):

    def OnInit(self):
        frame = MainWindow(None, title="Ragdoll")
        self.SetTopWindow(frame)
        frame.Show()
        return True


def do_merge(source: Task, target: Task):
    with fitz.open(source.file) as s:
        with fitz.open(target.file) as t:
            with fitz.open() as n:
                if target.start > 0:
                    n.insert_pdf(t, to_page=target.start)
                if source.end == 0:
                    n.insert_pdf(s, from_page=source.start - 1)
                else:
                    n.insert_pdf(s, from_page=source.start - 1, to_page=source.end)
                n.insert_pdf(t, from_page=target.start)
                n.save(os.path.splitext(target.file)[0] + '_' + ts() + '.pdf')
    os.startfile(os.path.split(target.file)[0])


class MainWindow(wx.Frame):

    def __init__(self, *args, **kwargs):
        super(MainWindow, self).__init__(*args, **kwargs)
        self.SetIcon(wx.Icon('/Users/zhangyebai/code/py/Ragdoll/res/ragdoll.ico'))
        self.SetSize((800, 600))
        self.Center()
        panel = wx.Panel(self)
        v_box = wx.BoxSizer(wx.VERTICAL)
        v_box.Add(
            wx.StaticText(
                panel,
                wx.ID_ANY,
                u'Please Select Source Pdf File Below',
                wx.DefaultPosition,
                wx.DefaultSize,
                0
            ),
            0, wx.EXPAND | wx.ALL, 10
        )
        self.source_file_picker = wx.FilePickerCtrl(
            panel,
            message='Plz Select Source Pdf File',
            wildcard='PDF Files (*.pdf)|*.pdf',
            style=wx.FLP_DEFAULT_STYLE | wx.FLP_USE_TEXTCTRL
        )
        v_box.Add(self.source_file_picker, 0, wx.EXPAND | wx.ALL, 10)
        v_box.Add(
            wx.StaticText(
                panel,
                wx.ID_ANY,
                u'Please Input Number Below For Source Pdf Page Start Index, Default Value 1 For First Page If Not',
                wx.DefaultPosition,
                wx.DefaultSize,
                0),
            0, wx.EXPAND | wx.ALL, 10
        )
        self.source_start_index_input = wx.TextCtrl(panel, value='1', style=wx.TE_PROCESS_ENTER)
        v_box.Add(self.source_start_index_input, 0, wx.EXPAND | wx.ALL, 10)
        v_box.Add(
            wx.StaticText(
                panel,
                wx.ID_ANY,
                u'Please Input Number Below For Source Pdf Page End Index, No Limit If Not Or Less Than Start Index',
                wx.DefaultPosition,
                wx.DefaultSize,
                0),
            0, wx.EXPAND | wx.ALL, 10
        )
        self.source_end_index_input = wx.TextCtrl(panel, value='0', style=wx.TE_PROCESS_ENTER)
        # self.source_end_index_input.SetHint(
        #     'Source Pdf Page End Index, Default Value 0 For No Limit If Not Input Or Less Than Start Index'
        # )
        v_box.Add(self.source_end_index_input, 0, wx.EXPAND | wx.ALL, 10)
        v_box.Add(
            wx.StaticText(
                panel,
                wx.ID_ANY,
                u'Please Select Target Pdf File Below',
                wx.DefaultPosition,
                wx.DefaultSize,
                0
            ),
            0, wx.EXPAND | wx.ALL, 10
        )
        self.target_file_picker = wx.FilePickerCtrl(
            panel,
            message='Plz Select Target Pdf File',
            wildcard='PDF Files (*.pdf)|*.pdf',
            style=wx.FLP_DEFAULT_STYLE | wx.FLP_USE_TEXTCTRL
        )
        v_box.Add(self.target_file_picker, 0, wx.EXPAND | wx.ALL, 10)
        v_box.Add(
            wx.StaticText(
                panel,
                wx.ID_ANY,
                u'Please Input Number Below For Target Pdf Page Start Index, Default Value 0 For Append Last If Not',
                wx.DefaultPosition,
                wx.DefaultSize,
                0),
            0, wx.EXPAND | wx.ALL, 10
        )
        self.target_start_index_input = wx.TextCtrl(panel, value='0', style=wx.TE_PROCESS_ENTER)
        v_box.Add(self.target_start_index_input, 0, wx.EXPAND | wx.ALL, 10)

        # v_box.Add(
        #     wx.StaticText(
        #         panel,
        #         wx.ID_ANY,
        #         u'Please Input Number Below For Source Pdf Page End Index, No Limit If Not Or Less Than Start Index',
        #         wx.DefaultPosition,
        #         wx.DefaultSize,
        #         0),
        #     0, wx.EXPAND | wx.ALL, 10
        # )
        # self.target_end_index_input = wx.TextCtrl(panel, value='0', style=wx.TE_PROCESS_ENTER)
        # v_box.Add(self.target_end_index_input, 0, wx.EXPAND | wx.ALL, 10)

        self.merge_button = wx.Button(panel, label='Merge', size=(70, 30))
        self.merge_button.Bind(wx.EVT_BUTTON, self.on_merge)
        v_box.Add(self.merge_button, 0, wx.ALIGN_CENTER | wx.ALL, 10)

        panel.SetSizer(v_box)
        self.Bind(wx.EVT_TEXT_ENTER, self.on_merge) #, self.page_input)

    def on_merge(self, event):
        self.merge_button.Disable()
        try:
            source_file = self.source_file_picker.GetPath()
            source_start = int(self.source_start_index_input.GetValue())
            if source_start < 0:
                source_start = 0
            source_end = int(self.source_end_index_input.GetValue())
            if source_end < 0 or source_end < source_start:
                source_end = 0
            source_task = Task(file=source_file, start=source_start, end=source_end)
            target_file = self.target_file_picker.GetPath()
            target_start = int(self.target_start_index_input.GetValue())
            if target_start > 0:
                target_start = target_start - 1
            else:
                target_start = 0
            target_task = Task(file=target_file, start=target_start, end=0)
            do_merge(source_task, target_task)
        except Exception as ex:
            print(str(ex))
        self.merge_button.Enable(True)

    def on_test(self, event):
        with fitz.open('/Users/zhangyebai/Desktop/日历&日程参数配置信息.pdf') as source:
            with fitz.open('/Users/zhangyebai/Desktop/简账户对接文档.pdf') as target:
                with fitz.open() as tt:
                    tt.insert_pdf(target, to_page=3)
                    tt.insert_pdf(source)
                    tt.insert_pdf(target, from_page=4)
                    tt.save('/Users/zhangyebai/Desktop/my_test.pdf')


if __name__ == '__main__':
    app = RagdollApp()
    app.MainLoop()
