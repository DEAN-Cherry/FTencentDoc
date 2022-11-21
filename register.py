import main


def register():
    main.load_config()

    name = input("请输入姓名: ")
    if main.registered_user.get(name) is not None:
        choose = input("已有该用户数据，是否修改数据(y/n): ")
        if choose != 'y':
            print("结束注册。")
            return

    wai = main.whereami
    re_pos = int(input(f"请输入相对于 {wai+3}行 的位置（你所在的行 - {wai+3}）: "))

    dict_temp = {name: re_pos}
    main.registered_user.update(dict_temp)

    main.save_config()

    print("注册完成。")


if __name__ == '__main__':
    register()
