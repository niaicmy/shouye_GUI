from DataDecode import DataDecode
import pandas as pd


# 读取员工表目录
def my_staff_records():
    staff_records_name = r'staff_records.xlsx'
    # 读取 excel 数据
    df_source = pd.read_excel(staff_records_name)

    # 从源数据中读取需要的数据
    data_no = df_source.loc[:, '编号'].values  # 读所有行的title以及data列的值，这里需要嵌套列表
    data_name = df_source.loc[:, ['姓名', '身份证号', '联系电话']].values  # 读所有行的title以及data列的值，这里需要嵌套列表
    # staff_records_dict = dict(zip(data_no, data_name))
    # self.staff_dict = staff_records_dict

    return dict(zip(data_no, data_name))
    # print(staff_records_dict)


if __name__ == '__main__':
    # print_hi('PyCharm')
    data_path = '12月病案首页.xlsx'
    app_class = DataDecode(data_path)
    app_class.update_staff_records()
    staff_dict = my_staff_records()

    print(staff_dict)
