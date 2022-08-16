from window import *
# Основная функция, с которой происходит запуск программы
def main():
    app = wx.App()
    frame = MyFrame(None, title='Pallet Counter')
    frame.Center()
    frame.Show()
    app.MainLoop()


if __name__ == "__main__":
    main()