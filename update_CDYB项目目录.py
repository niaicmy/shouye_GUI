import PySimpleGUI


def show_windows():
    # sg.theme('DarkAmber')  # Keep things interesting for your users
    PySimpleGUI.theme('DarkBrown5')  # Keep things interesting for your users

    layout = [
        # 病案数据文件路径选择
        [PySimpleGUI.Input(default_text=r'医保通用项目编码', readonly=True, size=(69, 1), key=r'tybm')],
        [PySimpleGUI.Input(default_text=r'请选择病案首页数据文件...', readonly=True, size=(69, 1), key=r'file_path1')],


        [PySimpleGUI.Text(r'请输入数据行号：'), PySimpleGUI.Input(key=r'-num-', size=(21, 1)),
         PySimpleGUI.Checkbox(r'自动提取模式', key=r'-auto-', default=False)],

        [PySimpleGUI.Submit(button_text=r'插入数据', auto_size_button=True),
         PySimpleGUI.Submit(button_text=r'更新数据', auto_size_button=True),
         PySimpleGUI.Exit(button_text=r'退出程序', auto_size_button=True)],
    ]

    window = PySimpleGUI.Window(r'病案首页数据提取程序', layout)
    # 消息循环
    while True:  # The Event Loop
        event, values = window.read()
        if event == r'提取数据':
            # print(event, values)
            try:
                PySimpleGUI.Popup('', r'插入数据成功！')
            except Exception as e:
                print('提取数据出现错误：' + str(e))

        elif event == r'更新数据':
            break
        elif event == PySimpleGUI.WIN_CLOSED or event == r'退出程序':
            break

    window.close()


# 按间距中的绿色按钮以运行脚本。
if __name__ == '__main__':
    # print_hi('PyCharm')
    show_windows()
