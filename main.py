# import os
# import pandas as pd
# from zipfile import BadZipFile
# from tqdm import tqdm

# folder_paths = ['/Volumes/Hams Hall Workspace/ Retouch Log/Week 28'] # Change this to a list of the paths of the folders containing the spreadsheets
# search_values = ['62259944'] # Change this to a list of the values you're searching for

# messages = []  # Store the messages in a list while the script is running, and then 
#                # print all the messages at once after the progress bar has completed. 

# for folder_path in folder_paths:
#     for filename in tqdm(os.listdir(folder_path), desc='Processing files'):
#         if filename.endswith('.xlsx') and not filename.startswith('~$') and not filename.startswith('.'):
#             file_path = os.path.join(folder_path, filename)
#             try:
#                 xls = pd.read_excel(file_path, engine='openpyxl', sheet_name=None)
#                 for sheet_name, df in xls.items():
#                     for search_value in search_values:
#                         mask = df.applymap(lambda x: search_value in str(x))
#                         if mask.any().any():
#                             result = df[mask.any(axis=1)]
#                             messages.append(f'{search_value} found in {filename}, sheet: {sheet_name}')
#             except BadZipFile:
#                 messages.append(f'{filename} is not a valid .xlsx file')

# print('\n'.join(messages))


import os
import pandas as pd
from zipfile import BadZipFile
import wx
import threading

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
                            xls = pd.read_excel(file_path, engine='openpyxl', sheet_name=None)
                            for sheet_name, df in xls.items():
                                for search_value in self.search_values:
                                    mask = df.applymap(lambda x: search_value in str(x))
                                    if mask.any().any():
                                        result = df[mask.any(axis=1)]
                                        messages.append(f'{search_value} found in {filename}, sheet: {sheet_name}, folder: {folder_path}')
                        except BadZipFile:
                            messages.append(f'{filename} is not a valid .xlsx file')
                    count += 1
                    wx.CallAfter(self.progress_bar.SetValue, count)
                    wx.CallAfter(self.search_button.SetLabel, f"Stop ({int(count / max_value * 100)}%)")
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
        folder_paths = [filename for filename in filenames if os.path.isdir(filename)]
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
        self.label = wx.StaticText(self.panel, label="Drag and drop folders here", size=(250, 75), style=wx.SIMPLE_BORDER)

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

    def update_label(self):
      label_text = "Selected folders:\n" + "\n".join(self.folder_paths)
      self.label.SetLabel(label_text)

    def on_search_button(self,event):
      if not self.search_thread:
          search_values=self.search_ctrl.GetValue().split()
          if search_values:
              event.GetEventObject().SetLabel("Stop")
            #   event.GetEventObject().Bind(wx.EVT_ENTER_WINDOW,self.on_enter_window)
            #   event.GetEventObject().Bind(wx.EVT_LEAVE_WINDOW,self.on_leave_window)

              # Create a new SearchThread object and start it
              self.search_thread=SearchThread(
                  self.folder_paths,
                  search_values,
                  self.progress_bar,
                  self.result_text,
                  event.GetEventObject()
              )
              self.search_thread.start()
      else:
          event.GetEventObject().SetLabel("Start")
          event.GetEventObject().Unbind(wx.EVT_ENTER_WINDOW)
          event.GetEventObject().Unbind(wx.EVT_LEAVE_WINDOW)
          self.search_thread.stop()
          self.search_thread=None

    # def on_enter_window(self,event):
    #   event.GetEventObject().SetBackgroundColour(wx.Colour(255,0,0))
    #   event.GetEventObject().Refresh()

    # def on_leave_window(self,event):
    #   event.GetEventObject().SetBackgroundColour(wx.NullColour)
    #   event.GetEventObject().Refresh()

app=wx.App()
frame=MyFrame(None,title="Folder Search")
frame.Show()
app.MainLoop()












# 17606263, 21406412, 35486536

# import os
# import pandas as pd
# from zipfile import BadZipFile
# import wx
# import wx.lib.agw.multidirdialog as MDD
# import multiprocessing # import multiprocessing module
# import logging


# # Create a logger object
# logger = logging.getLogger('excel_search')
# logger.setLevel(logging.INFO)

# # Create a file handler and a stream handler
# file_handler = logging.FileHandler('excel_search.log')
# stream_handler = logging.StreamHandler()

# # Create a formatter and add it to the handlers
# formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
# file_handler.setFormatter(formatter)
# stream_handler.setFormatter(formatter)

# # Add the handlers to the logger
# logger.addHandler(file_handler)
# logger.addHandler(stream_handler)


# def search_files(folder_path, search_values): 
#     results = []
#     for dirpath, dirnames, filenames in os.walk(folder_path):
#         for filename in filenames:
#             if filename.endswith('.xlsx') and not filename.startswith('~$') and not filename.startswith('.'):
#                 file_path = os.path.join(dirpath, filename)

#                 logger.info(f'Scanning {file_path}')
#                 try:
#                     xls = pd.read_excel(file_path, sheet_name=None)
#                     for sheet_name, df in xls.items():
#                         for search_value in search_values:
#                             try:
#                                 if search_value in df.values or int(search_value) in df.values:
#                                     results.append(f'{search_value} found in {filename}, sheet: {sheet_name}')
#                             except ValueError:
#                                 logger.warning(f'Could not convert search value {search_value} to integer')
#                 except BadZipFile:
#                     results.append(f'{filename} is not a valid .xlsx file')
#     return results



# def select_folder():
#     app = wx.App(None)
#     dialog = MDD.MultiDirDialog(None, 'Select Items')
#     if dialog.ShowModal() == wx.ID_OK:
#         items = dialog.GetPaths()
#         return items
#     return []


# class MyFrame(wx.Frame):
#     def __init__(self):
#         super().__init__(parent=None, title='Excel Search')
#         panel = wx.Panel(self)
#         self.select_button = wx.Button(panel, label='Select Folder')
#         self.select_button.Bind(wx.EVT_BUTTON, self.on_select)
#         self.selected_paths_label = wx.StaticText(panel, label='Selected Folder:')
#         self.selected_paths_value = wx.StaticText(panel, label='')
#         self.search_values_label = wx.StaticText(panel, label='Search Values:')
#         self.search_values_entry = wx.TextCtrl(panel)
#         self.search_button = wx.Button(panel, label='Search')
#         self.search_button.Bind(wx.EVT_BUTTON, self.on_search)
#         self.results_label = wx.StaticText(panel, label='Results:')
#         self.results_value = wx.TextCtrl(panel, style=wx.TE_MULTILINE | wx.TE_READONLY)
#         # Add a status bar to show the messages
#         self.status_bar = self.CreateStatusBar()
#         sizer = wx.BoxSizer(wx.VERTICAL)
#         sizer.Add(self.select_button, 0, wx.ALL | wx.CENTER, 5)
#         sizer.Add(self.selected_paths_label, 0, wx.ALL, 5)
#         sizer.Add(self.selected_paths_value, 0, wx.ALL | wx.EXPAND, 5)
#         sizer.Add(self.search_values_label, 0, wx.ALL, 5)
#         sizer.Add(self.search_values_entry, 0, wx.ALL | wx.EXPAND, 5)
#         sizer.Add(self.search_button, 0, wx.ALL | wx.CENTER, 5)
#         sizer.Add(self.results_label, 0, wx.ALL, 5)
#         sizer.Add(self.results_value, 1, wx.ALL | wx.EXPAND, 5)
#         panel.SetSizer(sizer)
#         self.folder_paths = []
#         self.Show()

#     def on_select(self,event):
#       self.folder_paths=select_folder()
#       if self.folder_paths:
#           self.selected_paths_value.SetLabel(', '.join(self.folder_paths))
#           self.Layout()

#     def on_search(self,event):
#       search_values=self.search_values_entry.GetValue().replace(',', ' ').split()
#       pool = multiprocessing.Pool() # create a pool object
#       results=pool.starmap(search_files ,zip(self.folder_paths ,[search_values]*len(self.folder_paths ))) # use starmap to apply search_files to each folder_path and search_values pair
#       pool.close() # close the pool
#       pool.join() # wait for the processes to finish
#       result_text='\n'.join(sum(results ,[])) # flatten and join the results
#       self.results_value.SetValue(result_text )
#       # Update the UI with the messages from the done queue
#       while not self.doneQueue.empty():
#           message = self.doneQueue.get()
#           wx.CallAfter(self.update_ui, message)

#     def update_ui(self, message):
#         # Set the status bar text with the message
#         self.status_bar.SetStatusText(message)

# if __name__=='__main__':
#     app=wx.App()
#     frame=MyFrame()
#     app.MainLoop()
