from scse_myschedular.MyScse_login import MyScse_login

if __name__ == '__main__':
    username = input("学号：")
    password = input("密码：")
    MyScse_login = MyScse_login(username,password)
    if not MyScse_login.to_login():
        print('获取失败！！！')
    MyScse_login.student_schedular()