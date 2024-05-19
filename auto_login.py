import socket
import uuid
import requests
import os
import random
import hashlib
import time
import win32com.client
import win32api, win32gui
from loguru import logger
import sys
import wmi

LAST_USER = ""
CHANGE_USER = False

def check_path_exist(path):
    if not os.path.exists(path):
        try:
            os.makedirs(path)
            logger.debug(f"Created log directory: {path}")
        except OSError as e:
            logger.debug(f"Failed to create log directory: {path}, Error: {e}")
            raise

def init_logger():
    current_path = os.path.dirname(os.path.abspath(__file__))
    current_user = get_current_user()
    log_path = os.path.join(current_path, current_user)
    check_path_exist(log_path)
    log_file_location = os.path.join(log_path, 'msg_queue.log')
    logger.remove(handler_id=None)
    logger.add(sink=log_file_location,level="DEBUG",compression="zip",enqueue=True,rotation="2 MB")
    logger.add(sink=sys.stderr,level="INFO")

def parse_string(string):
    result = {
        "CN":"",
        "OU":"",
        "DC":""
    }
    pairs = string.split(',')
    for pair in pairs:
        key, value = pair.split('=')
        if result[key] == "":
            result[key] = value
        elif key == "OU":
            result[key] = value + "/"+ result[key]
        elif key == "DC":
            result[key] = result[key]+ "." + value 
        else:
            result[key] = value
    return result


def get_ad_information():
    ad = win32com.client.Dispatch("ADSystemInfo")
    #domain = ad.DomainDNSName
    username = ad.UserName
    print(username)
    result = parse_string(username)
    return result



def get_random_and_md5():
    random_number = str(random.randint(10000, 99999))
    string = "chailease" + random_number
    hash_value = hashlib.md5(string.encode()).hexdigest()
    return random_number,hash_value

'''
这种方式判断切换使用者是失败的...
因为程序是跟着使用者的，所以切来切去current user永远是自己...
'''
def get_current_user():
    if os.name == 'nt':
        return os.environ['USERNAME']
    
def get_substring_after_backslash(string):
    #切换使用者中，会回传None
    if string == None:
        return ""
    index = string.rfind('\\')  
    if index != -1:
        substring = string[index + 1:]  
        return substring
    else:
        return ""  #没有找到反斜杠


def check_user_change():
    global LAST_USER
    c = wmi.WMI()
    query = "SELECT * FROM Win32_ComputerSystem"
    result = c.query(query)
    for item in result:
        username = get_substring_after_backslash(item.UserName)
        if LAST_USER == "":
            LAST_USER = username
        if username != LAST_USER:
            return True
    return False

def get_mac_address():
    mac = ':'.join(['{:02x}'.format((uuid.getnode() >> ele) & 0xff)
                    for ele in range(0, 8 * 6, 8)][::-1])
    return mac

def get_ip_address():
    s = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    s.connect(("8.8.8.8", 80))
    ip_address = s.getsockname()[0]
    s.close()
    return ip_address

def get_sso_login_data(current_user):
    result = get_ad_information()
    random_number, hash_value = get_random_and_md5()
    data = {
        "ip": get_ip_address(),
        "name": current_user,
        "show_name": result["CN"],
        "group": "/" + result["DC"] + "/" + result["OU"],
        "mac": get_mac_address(),
        "random": random_number,
        "md5": hash_value
    }
    return data

login_url = "http://acip:9999/v1/online-users"
logout_url = "http://acip:9999/v1/online-users?_method=DELETE"


def run_program():
    global LAST_USER,CHANGE_USER

    #第一次启动脚本 先做SSO认证
    _ = check_user_change()
    if LAST_USER != "local account":
        data = get_sso_login_data(LAST_USER)
        response = requests.post(login_url, json=data)
        logger.info("login...")

    i = 0
    while True:
        logger.info(LAST_USER)
        if check_user_change():
            CHANGE_USER = True
        elif check_user_change() is False and CHANGE_USER == True:
            CHANGE_USER = False
            if LAST_USER == "chailease":
                random_number, hash_value = get_random_and_md5()
                data = {
                    "ip" : get_ip_address(),
                    "random": random_number,
                    "md5": hash_value
                }
                response = requests.post(logout_url, json=data)
                logger.info("logout...")
            else:
                data = get_sso_login_data(LAST_USER)
                response = requests.post(login_url, json=data)
                logger.info("login...")
        if i == 60:
            i = 0
            #使用者 如果切换了 60s输出一次log 路径是在local temp底下
            logger.info("user change...")
               

        i = i + 1
        time.sleep(1)
        

if __name__ == '__main__':
    init_logger()
    #让程序在背景执行
    ct = win32api.GetConsoleTitle()
    hd = win32gui.FindWindow(0,ct)
    win32gui.ShowWindow(hd,0)
    run_program()
