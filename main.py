import os
import pandas as pd
from zipfile import BadZipFile
import wx
import threading

# def activate_window():
#     script = '''tell application "System Events"
#                     set proc to the first process whose unix id is {}
#                     if visible of proc is false then
#                         set visible of proc to true
#                     end if
#                     if frontmost of proc is false then
#                         set frontmost of proc to true
#                     end if
#                 end tell'''.format(os.getpid())
#     os.system(f"osascript -e '{script}'")


class SearchThread(threading.Thread):
    def __init__(self, folder_paths, search_values, progress_bar, result_text, search_button):
        threading.Thread.__init__(self)
        self.folder_paths = folder_paths
        self.search_values = search_values
        self.progress_bar = progress_bar
        self.result_text = result_text
        self.search_button = search_button
        self.stop_event = threading.Event()

    def run(self):
        max_value = 0
        for folder_path in self.folder_paths:
            for dirpath, dirnames, filenames in os.walk(folder_path):
                max_value += len(filenames)
        wx.CallAfter(self.progress_bar.SetRange, max_value)
        messages = []
        count = 0

        for folder_path in self.folder_paths:
            for dirpath, dirnames, filenames in os.walk(folder_path):
                for filename in filenames:
                    if self.stop_event.is_set():
                        return
                    if filename.endswith('.xlsx') and not filename.startswith('~$') and not filename.startswith('.'):
                        file_path = os.path.join(dirpath, filename)
                        try:
                            xls = pd.read_excel(
                                file_path, engine='openpyxl', sheet_name=None)
                            for sheet_name, df in xls.items():
                                for search_value in self.search_values:
                                    mask = df.applymap(
                                        lambda x: search_value in str(x))
                                    if mask.any().any():
                                        result = df[mask.any(axis=1)]
                                        messages.append(
                                            f'{search_value} found in {filename}, sheet: {sheet_name}, folder: {folder_path}')
                        except BadZipFile:
                            messages.append(
                                f'{filename} is not a valid .xlsx file')
                    count += 1
                    wx.CallAfter(self.progress_bar.SetValue, count)
                    wx.CallAfter(self.search_button.SetLabel,
                                 f"Stop ({int(count / max_value * 100)}%)")
                    wx.Yield()

        wx.CallAfter(self.result_text.SetLabel, "\n".join(messages))
        wx.CallAfter(self.search_button.SetLabel, "Start")

    def stop(self):
        self.stop_event.set()


class MyFileDropTarget(wx.FileDropTarget):
    def __init__(self, window):
        wx.FileDropTarget.__init__(self)
        self.window = window

    def OnDropFiles(self, x, y, filenames):
        folder_paths = [
            filename for filename in filenames if os.path.isdir(filename)]
        self.window.folder_paths.extend(folder_paths)
        self.window.update_label()
        return True


class MyFrame(wx.Frame):
    def __init__(self, *args, **kwargs):
        super(MyFrame, self).__init__(*args, **kwargs)
        self.SetMinSize((600, 400))
        self.folder_paths = []
        self.search_thread = None
        self.panel = wx.Panel(self)
        self.sizer = wx.BoxSizer(wx.VERTICAL)
        self.label = wx.StaticText(self.panel, label="Drag and drop folders here", size=(
            250, 75), style=wx.SIMPLE_BORDER)
        self.sizer.Add(self.label, 0, wx.ALL | wx.ALIGN_CENTER_HORIZONTAL, 5)
        drop_target = MyFileDropTarget(self)
        self.label.SetDropTarget(drop_target)
        self.search_ctrl = wx.SearchCtrl(self.panel, style=wx.TE_PROCESS_ENTER)
        self.sizer.Add(self.search_ctrl, 0, wx.ALL | wx.EXPAND, 5)
        self.search_button = wx.Button(self.panel, label="Start")
        self.search_button.Bind(wx.EVT_BUTTON, self.on_search_button)
        self.sizer.Add(self.search_button, 0, wx.ALL | wx.CENTER, 5)
        self.progress_bar = wx.Gauge(self.panel)
        self.sizer.Add(self.progress_bar, 0, wx.ALL | wx.EXPAND, 5)
        self.result_text = wx.StaticText(self.panel)
        self.sizer.Add(self.result_text, 0, wx.ALL | wx.CENTER | wx.EXPAND, 5)
        self.panel.SetSizer(self.sizer)
        self.Layout()

    def Show(self):
        super(MyFrame, self).Show()
        self.search_ctrl.SetFocus()

    def update_label(self):
        label_text = "Selected folders:\n" + "\n".join(self.folder_paths)
        self.label.SetLabel(label_text)

    def on_search_button(self, event):
        if not self.search_thread:
            search_values = self.search_ctrl.GetValue().split()
            if search_values:
                event.GetEventObject().SetLabel("Stop")

                # Create a new SearchThread object and start it
                self.search_thread = SearchThread(
                    self.folder_paths,
                    search_values,
                    self.progress_bar,
                    self.result_text,
                    event.GetEventObject()
                )
                self.search_thread.start()
        else:
            event.GetEventObject().SetLabel("Start")
            self.search_thread.stop()
            self.search_thread = None


app = wx.App()
frame = MyFrame(None, title="Excel Multi Search", style=wx.DEFAULT_FRAME_STYLE | wx.STAY_ON_TOP)
frame.Show()
app.MainLoop()

