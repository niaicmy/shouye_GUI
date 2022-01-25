# 这是一个示例 Python 脚本。

# 按 Shift+F10 执行或将其替换为您的代码。
# 按 双击 Shift 在所有地方搜索类、文件、工具窗口、操作和设置。
import PySimpleGUI
from DataDecode import DataDecode


# import os
# def print_hi(name):
#     # 在下面的代码行中使用断点来调试脚本。
#     print(f'Hi, {name}')  # 按 Ctrl+F8 切换断点。


def show_windows():
    # sg.theme('DarkAmber')  # Keep things interesting for your users
    PySimpleGUI.theme('DarkBrown5')  # Keep things interesting for your users

    layout = [
        # 第一行 病案数据文件路径选择
        [PySimpleGUI.Input(default_text=r'请选择病案首页数据文件...', readonly=True, size=(69, 1), key=r'file_path1'),
         PySimpleGUI.FileBrowse(button_text=r'打开病案首页数据文件...', key=r'data_path',
                                file_types=((r'Excel文件', r'*.xls;*.xlsx;*.csv;'), (r'所有文件', r'*.*')))],
        # 数据库配置文件路径
        # [sg.Input(default_text=r'', readonly=True, size=(69, 1), key=r'file_path2'),
        #  sg.FileBrowse(button_text=r'打开Database配置文件...', key=r'database_path',
        #                file_types=((r'Json文件', r'*.json;'), (r'所有文件', r'*.*')))],

        # 第二行 参数设置
        [PySimpleGUI.Text(r'请输入数据行号：'), PySimpleGUI.Input(key=r'-num-', size=(21, 1)),
         PySimpleGUI.Button(button_text=r'更新员工数据', auto_size_button=True),
         PySimpleGUI.Checkbox(r'自动提取模式', key=r'-auto-', default=False)],

        # 第三行
        [PySimpleGUI.Submit(button_text=r'提取数据', auto_size_button=True),
         PySimpleGUI.Exit(button_text=r'退出程序', auto_size_button=True)],
    ]

    window = PySimpleGUI.Window(r'病案首页数据提取程序', layout)

    while True:  # The Event Loop
        event, values = window.read()
        if event == r'提取数据':
            # print(event, values)
            if values['data_path'] != '':
                app_class = DataDecode(values['data_path'])
                try:
                    if values['-auto-']:
                        # app_class.update_staff_records()        # 更新员工数据
                        app_class.make_excel(values['-auto-'])
                    else:
                        # app_class.update_staff_records()  # 更新员工数据
                        app_class.make_excel(values['-auto-'], int(values['-num-']))

                    PySimpleGUI.Popup('', r'数据提取成功！')
                except Exception as e:
                    print('提取数据出现错误：' + str(e))

            elif values['file_path1'] == '请选择病案首页数据文件...' and values['data_path'] == '':
                PySimpleGUI.Popup('', r'请选择有效的文件路径！')

        elif event == r'更新员工数据':
            app_class = DataDecode()
            app_class.update_staff_records()  # 更新员工数据
            PySimpleGUI.Popup('', r'员工信息更新完成！')

        elif event == PySimpleGUI.WIN_CLOSED or event == r'退出程序':
            break

    window.close()


# 按间距中的绿色按钮以运行脚本。
if __name__ == '__main__':
    # print_hi('PyCharm')
    show_windows()

# 访问 https://www.jetbrains.com/help/pycharm/ 获取 PyCharm 帮助
