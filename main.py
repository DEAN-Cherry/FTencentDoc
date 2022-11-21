# 这是一个示例 Python 脚本。

# 按 Shift+F10 执行或将其替换为您的代码。
# 按 双击 Shift 在所有地方搜索类、文件、工具窗口、操作和设置。

# import win32clipboard as w
# import win32con
import base64
import datetime
import json
import logging
import os
import time
from datetime import date

from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from tqdm import tqdm, trange
from webdriver_manager.chrome import ChromeDriverManager

path = os.path.dirname(__file__)
cookies = []
whereami = 206
doc_url = 'https://docs.qq.com/sheet/DZEx4eFBrYVNSb3di?tab=BB08J2&u=d4e0a9ea21c04264806e482dadeb398a'
reference_col = 0
reference_row = 0

column = 1
row = 2
information_col = 3
information_row = 3

registered_user = {
    "聂博洋": 0,
    "唐海成": 1,
    "李文锋": 2,
    "高泽森": -2,
    "卢国浩": 13
}


def log_init():
    log = logging.getLogger()
    log.setLevel(logging.INFO)
    rq = time.strftime('%Y-%m-%d', time.localtime(time.time()))
    # log_path = os.getcwd() + '/logs/'
    log_path = os.path.dirname(__file__) + '/logs/'
    log_name = log_path + rq + '.log'
    logfile = log_name
    fh = logging.FileHandler(logfile)
    fh.setLevel(logging.DEBUG)  # 输出到file的log等级的开关

    ch = logging.StreamHandler()
    ch.setLevel(logging.WARNING)

    # 第三步，定义handler的输出格式
    formatter = logging.Formatter(
        "%(asctime)s\t%(levelname)s --- %(filename)s[line:%(lineno)d] - %(funcName)s : %(message)s")
    fh.setFormatter(formatter)
    ch.setFormatter(formatter)
    # 第四步，将logger添加到handler里面
    log.addHandler(fh)
    log.addHandler(ch)
    return log


def update_config():
    global config
    config = {
        "whereami": whereami,
        "reference_pos": (reference_row, reference_col),
        "doc_url": doc_url,
        "user": registered_user,
        "cookies": cookies
    }
    print("config已更新")


def save_config():
    global config
    update_config()
    tf = open(f"{path}/localConfig.json", 'w')
    json.dump(config, tf)
    tf.close()
    print("config 已保存")


def load_config():
    global whereami, reference_row, reference_col, doc_url, registered_user, cookies
    tf = open(f"{path}/localConfig.json", "r")
    new_dict = json.load(tf)
    k = [i for i in new_dict]

    whereami = new_dict[k[0]]
    reference_row, reference_col = new_dict[k[1]]
    doc_url = new_dict[k[2]]
    registered_user = new_dict[k[3]]
    cookies = new_dict[k[4]]

    return new_dict, k


# def get_text():
#     time.sleep(1)
#     w.OpenClipboard()
#     is_available = True
#
#     try:
#         w.GetClipboardData(win32con.CF_TEXT)
#     except Exception as e:
#         print(e)
#         is_available = False
#
#     if is_available:
#         text = w.GetClipboardData(win32con.CF_TEXT)
#     else:
#         text = '该项目为空'.encode('gbk', 'xmlcharrefrepalce')
#
#     w.CloseClipboard()
#     return text.decode('gbk')
#
#
# def set_text(a):
#     w.OpenClipboard()
#     w.EmptyClipboard()
#     w.SetClipboardData(win32con.CF_TEXT, a)
#     w.CloseClipboard()
#
#
# def copy_cell():
#     ActionChains(driver).key_down(Keys.CONTROL).key_down('c').key_up(Keys.CONTROL).key_up('c').perform()
#     return get_text()


def get_cell():
    cell_input = driver.find_element(By.ID, 'alloy-simple-text-editor').text
    return cell_input


def get_xy():
    global row, column
    return row, column


def date_func(_x):
    _month = _x.split("月")[0]
    _day = _x.split("月")[1].split("日")[0]
    selected_date = date(date.today().year, int(_month), int(_day))
    logger.info(f"初始化位移后选中的日期是：{_month}月{_day}日")
    gap = date.today().toordinal() - selected_date.toordinal()
    logger.info(f"相差的单元格数为：{gap}\t正在移动中……")
    return gap


def cell_fill():
    if plain_text_fill():
        print("打卡成功")
        return True
    elif data_validation_fill():
        print("打卡成功")
        return True
    else:
        print("打卡失败")
        return False


def data_validation_fill():
    is_finished = False
    if field_validation('该项目为空'):
        ActionChains(driver).key_down(Keys.SPACE).perform()
        time.sleep(0.8)
        ActionChains(driver).key_down(Keys.DOWN).key_down(Keys.DOWN).perform()
        time.sleep(0.8)
        ActionChains(driver).key_down(Keys.ENTER).perform()
        get_focus()
        if field_validation("已做核酸"):
            is_finished = True
    elif field_validation("已做核酸"):
        print("今天已经签到过了")
        is_finished = True
    else:
        print(f"期望的内容`已做核酸`, 但是检测到`{get_cell()}`, 请检测您需要填入的信息。")
        is_finished = False

    return is_finished


def plain_text_fill(text="已做核酸"):
    is_finished = False

    if field_validation('该项目为空'):
        driver.find_element(By.ID, 'alloy-simple-text-editor').click()
        time.sleep(0.1)
        ActionChains(driver).double_click().perform()
        time.sleep(0.1)
        driver.find_element(By.ID, 'alloy-simple-text-editor').send_keys(text)
        time.sleep(0.5)
        ActionChains(driver).click().perform()
        time.sleep(0.1)
        if field_validation(text):
            is_finished = True
    elif field_validation(text):
        print("今天已经签过到了")
        is_finished = True
    else:
        print(f"期望的内容`{text}`, 但是检测到`{get_cell()}`, 请检测您需要填入的信息。")
        is_finished = False

    return is_finished


def column_redirection():
    driver.switch_to.parent_frame()
    normalization()
    move_down()

    if field_validation('年级'):
        logger.info("表格初始化定位成功")
    else:
        logger.error("表格初始化失败")
        exit()

    with tqdm(total=100, colour='blue') as pbar:
        pbar.set_description("列定位")
        for _ in range(0, 5):
            move_right()
            pbar.update(5)
            x, y = get_xy()
            pbar.set_postfix(text=get_cell(), row=x, column=y)

        cost = date_func(get_cell())
        time.sleep(0.5)
        for _ in range(0, cost):
            move_right()
            x, y = get_xy()
            pbar.set_postfix(text=get_cell(), row=x, column=y)
            pbar.update(75 / cost)



    time.sleep(0.1)
    logger.info(f"移动完成的位置是：{get_cell()}")


def row_redirection(p):
    with trange(p, colour='blue') as pbar:
        for _ in pbar:
            pbar.set_description("行定位")
            move_down()
            x, y = get_xy()
            pbar.set_postfix(text=get_cell(), row=x, column=y)

    print("Success\n")


def position_redirection():
    column_redirection()
    time.sleep(0.5)
    print("Success")
    row_redirection(whereami)
    print('位置初始化完成')


def position_validation(*p):
    global reference_row, reference_col, information_row, information_col
    bool_i, bool_j = True, True
    val_row, val_col = get_xy()

    with trange(val_col - information_col, colour='red') as pbar:
        pbar.set_description("行验证")
        for _ in pbar:
            move_left()
            x, y = get_xy()
            pbar.set_postfix(text=get_cell(), row=x, column=y)

    if field_validation("聂博洋"):
        logger.info("初始化行定位正确")
        print("行定位正确")

    else:
        logger.error("行定位失败")
        print('行定位失败')
        bool_i = False

    with trange(val_col - information_col, colour='green') as pbar:
        pbar.set_description("行复位")
        for _ in pbar:
            move_right()

    with trange(val_row - information_row, colour='red') as pbar:
        pbar.set_description("列验证")
        for _ in pbar:
            move_up()
            x, y = get_xy()
            pbar.set_postfix(text=get_cell(), row=x, column=y)

    if field_validation(today_date):
        logger.info("初始化行定位正确")
        print('列定位正确')
    elif field_validation(today_validation):
        logger.info("初始化行定位正确")
        print('列定位正确')
    else:
        logger.error("列定位失败")
        print('列定位失败')
        bool_j = False

    with trange(val_row - information_row, colour='green') as pbar:
        pbar.set_description("列复位")
        for _ in pbar:
            move_down()

    if bool_i:
        reference_row = val_row

        if bool_j:
            reference_col = val_col
            return True
        else:
            return False
    else:
        return False


def field_validation(text):
    copy_text = str(get_cell()).strip()

    if copy_text == '':
        copy_text = "该项目为空"

    if copy_text == text.strip():
        logger.info(f"`{text}`\t验证成功！")
        return True
    else:
        logger.error(f"验证失败，期望的值`{text.strip()}`，然而检验到的值为`{copy_text}`")
        return False


def cookies_validation():
    global cookies
    is_available = False
    login_method(2, driver)
    time.sleep(1)
    try:
        driver.find_element(By.XPATH, '//*[@id="canvasContainer"]/div[1]/div[2]').click()
        is_available = True
    except Exception as e:
        logger.info("cookies已过期")
        logger.info(repr(e))
        print("您的Cookies已过期")

    return is_available


def get_cookies():
    global cookies
    driver.refresh()  # 先刷新界面
    time.sleep(1)
    try:
        cookies = driver.get_cookies()
        logger.info(cookies)  # 获得cookie并打印
        return cookies

    except Exception as e:
        logger.warning("cookies获取失败")
        logger.warning(repr(e))
        print("cookies获取失败，请检查是否登陆了QQ")


def login_method(t, _driver):
    if t == 1:
        _driver.find_element(By.ID, 'blankpage-button-pc').click()
        # driver.find_element_by_class_name('unlogin-container').click()#点击登入按钮
        time.sleep(1)
        _driver.find_element(By.ID, 'qq-tabs-title').click()
        _driver.switch_to.frame(_driver.find_element(By.ID, 'login_frame'))
        try:
            _driver.find_element(By.CLASS_NAME, 'img_out_focus').click()
            time.sleep(0.5)
            _driver.switch_to.parent_frame()

        except Exception as e:
            # TODO 把密码隐藏
            logger.error("QQ点击登陆失败")
            logger.error(repr(e))
            print('未找到登录的QQ')
            print('尝试使用账号密码登陆,可能会需要手机验证码')
            _driver.find_element(By.ID, 'switcher_plogin').click()
            _driver.find_element(By.XPATH, '//*[@id="u"]').send_keys("772839031")
            time.sleep(1)
            _driver.find_element(By.XPATH, '//*[@id="p"]').send_keys(
                base64.b64decode('这里是QQ密码').decode("utf-8"))
            time.sleep(0.5)
            _driver.find_element(By.ID, 'login_button').click()
            # 这个不好用，因为一定会要求你使用手机验证码，废案。
    if t == 2:
        _driver.delete_all_cookies()
        for cookie in cookies:
            cookie_dict = {
                'domain': '.docs.qq.com',  # 这里是固定的每个网站都不同
                'name': cookie.get('name'),
                'value': cookie.get('value'),
                "expires": cookie.get('value'),
                'path': '/',
                'httpOnly': False,
                'HostOnly': False,
                'Secure': False}
            _driver.add_cookie(cookie_dict)
        _driver.refresh()  # 带着cookie重新加载
        time.sleep(1)


def normalization():
    global row, column
    row = 2
    column = 1
    ActionChains(driver).key_down(Keys.CONTROL).key_down(Keys.HOME).key_up(Keys.CONTROL).key_up(Keys.HOME).perform()


def move_down():
    global row
    row += 1
    ActionChains(driver).key_down(Keys.DOWN).perform()


def move_up():
    global row
    row -= 1
    ActionChains(driver).key_down(Keys.UP).perform()


def move_left():
    global column
    column -= 1
    ActionChains(driver).key_down(Keys.LEFT).perform()


def move_right():
    global column
    column += 1
    ActionChains(driver).key_down(Keys.RIGHT).perform()


def get_focus():
    driver.find_element(By.ID, 'alloy-simple-text-editor').click()


def user_service(user):
    user = dict(sorted(user.items(), key=lambda x: x[1]))
    step = 0
    step_list = [user[i] for i in user]
    user_list = [i for i in user]
    for i in range(len(step_list)):

        if step_list[i] == 0:
            if cell_fill():
                print(f"\x1b[1;37;44m用户`{user_list[i]}`签到已完成\x1b[0m")

        elif step_list[i] > 0:
            step = abs(step_list[i])
            for j in range(step):
                move_down()
            step_list = [*map(lambda x: x - step, step_list)]
            if cell_fill():
                print(f"\x1b[1;37;44m用户`{user_list[i]}`签到已完成\x1b[0m")

        elif step_list[i] < 0:
            step = abs(step_list[i])
            for j in range(step):
                move_up()
            step_list = [*map(lambda x: x + step, step_list)]
            if cell_fill():
                print(f"\x1b[1;37;44m用户`{user_list[i]}`签到已完成\x1b[0m")


# 按间距中的绿色按钮以运行脚本。
if __name__ == '__main__':

    logger = log_init()

    config, keys = load_config()

    today_date = str(datetime.datetime.now().year) + "/" + str(datetime.datetime.now().month) + "/" + str(
        datetime.datetime.now().day)
    today_validation = f'{date.today().month}月{date.today().day}日'

    chrome_options = webdriver.ChromeOptions()
    chrome_options.add_argument('--ignore-certificate-errors')
    chrome_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=chrome_options)
    # driver = webdriver.Edge()
    driver.get(doc_url)  # 将健康表的地址copy过来就行。

    if cookies_validation():
        print("成功登陆")
    else:
        login_method(1, driver)
        get_cookies()
        time.sleep(1)
        update_config()

    time.sleep(1)

    position_redirection()

    print("开始验证初始化位置正确性")
    if position_validation():
        print("验证成功")
        update_config()

    user_service(registered_user)
    time.sleep(1)
    get_cookies()
    time.sleep(5)
    save_config()
    os.system('pause')
    driver.quit()
    exit()
