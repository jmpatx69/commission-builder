import wx
import pyperclip

class FolderSelector(wx.Frame):
    def __init__(self, *args, **kw):
        super(FolderSelector, self).__init__(*args, **kw)
        self.InitUI()

    def InitUI(self):
        panel = wx.Panel(self)
        vbox = wx.BoxSizer(wx.VERTICAL)

        self.folder_path = wx.TextCtrl(panel)
        vbox.Add(self.folder_path, flag=wx.EXPAND | wx.LEFT | wx.RIGHT | wx.TOP, border=10)

        browse_btn = wx.Button(panel, label='Browse')
        browse_btn.Bind(wx.EVT_BUTTON, self.OnBrowse)
        vbox.Add(browse_btn, flag=wx.ALIGN_CENTER | wx.TOP | wx.BOTTOM, border=10)

        panel.SetSizer(vbox)

        self.SetTitle('Select Folder')
        self.Centre()

    def OnBrowse(self, event):
        dialog = wx.DirDialog(self, "Choose a directory:", style=wx.DD_DEFAULT_STYLE)
        if dialog.ShowModal() == wx.ID_OK:
            folder_path = dialog.GetPath()
            self.folder_path.SetValue(folder_path)
            pyperclip.copy(folder_path)
            wx.MessageBox(f'Folder path copied: {folder_path}', 'Info', wx.OK | wx.ICON_INFORMATION)
        dialog.Destroy()

def main():
    app = wx.App(False)
    frame = FolderSelector(None)
    frame.Show(True)
    app.MainLoop()

if __name__ == '__main__':
    main()