# 自动定时版本 - 默认在每天中午12:00抢后天的场地
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import time
from datetime import datetime, timedelta
from venue_booking_info import VenueBooking
import subprocess
import os
import yaml  # 添加yaml导入

def setup_driver(headless=False):
    """优化的浏览器设置"""
    edge_options = Options()
    edge_options.add_argument('--disable-gpu')
    edge_options.add_argument('--no-sandbox')
    edge_options.add_argument('--disable-dev-shm-usage')
    edge_options.add_argument('--disable-extensions')
    edge_options.add_argument('--disable-logging')
    edge_options.add_argument('--log-level=3')
    edge_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    
    # 直接在浏览器中禁用代理
    edge_options.add_argument('--no-proxy-server')
    
    edge_options.add_experimental_option(
        'prefs', {
            'profile.managed_default_content_settings.images': 2,
            'profile.managed_default_content_settings.javascript': 1
        }
    )
    
    service = Service(EdgeChromiumDriverManager().install())
    driver = webdriver.Edge(service=service, options=edge_options)
    driver.implicitly_wait(3)
    return driver

def manage_proxy(disable=True):
    """管理系统代理设置
    
    Args:
        disable (bool): 如果为True，则禁用代理；如果为False，则恢复代理设置
    
    Returns:
        dict: 如果disable=True，返回原始代理设置；否则返回None
    """
    try:
        # 保存当前环境变量中的代理设置
        original_settings = {
            "http_proxy": os.environ.get("http_proxy", ""),
            "https_proxy": os.environ.get("https_proxy", ""),
            "HTTP_PROXY": os.environ.get("HTTP_PROXY", ""),
            "HTTPS_PROXY": os.environ.get("HTTPS_PROXY", ""),
            "no_proxy": os.environ.get("no_proxy", "")
        }
        
        if disable:
            # 临时禁用代理
            for proxy_var in ["http_proxy", "https_proxy", "HTTP_PROXY", "HTTPS_PROXY"]:
                if proxy_var in os.environ:
                    os.environ[proxy_var] = ""
            
            # 设置no_proxy为通配符，确保所有连接都不使用代理
            os.environ["no_proxy"] = "*"
            
            print("系统代理已临时禁用（通过环境变量）")
            
            # 尝试同时使用gsettings禁用系统代理（针对GNOME桌面环境）
            try:
                subprocess.call(["gsettings", "set", "org.gnome.system.proxy", "mode", "'none'"])
                print("系统代理已临时禁用（通过gsettings）")
            except:
                pass
                
            return original_settings
        else:
            # 恢复原始代理设置
            if isinstance(disable, dict):
                for key, value in disable.items():
                    if value:  # 只恢复非空值
                        os.environ[key] = value
                    elif key in os.environ:
                        del os.environ[key]
                
                # 尝试同时使用gsettings恢复系统代理（针对GNOME桌面环境）
                try:
                    mode = subprocess.check_output(
                        ["gsettings", "get", "org.gnome.system.proxy", "mode"]
                    ).decode().strip().replace("'", "")
                    
                    if mode == "none":
                        subprocess.call(["gsettings", "set", "org.gnome.system.proxy", "mode", "'manual'"])
                except:
                    pass
                
                print("系统代理设置已恢复")
            
            return None
    except Exception as e:
        print(f"管理代理设置时发生错误: {str(e)}")
        return None

def wait_until_time(target_hour=12, target_minute=0, target_second=0):
    """等待直到指定时间"""
    while True:
        current_time = datetime.now()
        if (current_time.hour == target_hour and 
            current_time.minute == target_minute and 
            current_time.second >= target_second):
            break
        time.sleep(0.1)

def prepare_booking(driver, booking, form_data, student_id, password):
    """提前准备表单"""
    try:
        print("正在打开登录页面...")
        driver.get("https://ids.shanghaitech.edu.cn/authserver/login?service=https%3A%2F%2Foa.shanghaitech.edu.cn%2Fworkflow%2Frequest%2FAddRequest.jsp%3Fworkflowid%3D14862")
        
        print("正在登录...")
        if not booking.login(student_id, password):
            raise Exception("登录失败")
        
        time.sleep(0.5)
        
        print("正在切换到表单...")
        if not booking.switch_to_form_frame():
            raise Exception("切换到表单frame失败")
        
        print("正在填写表单...")
        if not booking.fill_form(form_data):
            raise Exception("填写表单失败")
        
        # 定义各种场地列表
        venue_lists = {
            "羽毛球场": [
                "63_13",  # 羽毛球1号场地
                "63_14",  # 羽毛球2号场地
                "63_15",  # 羽毛球3号场地
                "63_16",  # 羽毛球4号场地
                "63_17",  # 羽毛球5号场地
                "63_18"   # 羽毛球6号场地
            ],
            "网球场": [
                "63_25",  # 网球场1号场地
                "63_26",  # 网球场2号场地
                "63_27"   # 网球场3号场地
            ],
            "乒乓球": [
                "63_19",  # 乒乓球1号场地
                "63_20",  # 乒乓球2号场地
                "63_21",  # 乒乓球3号场地
                "63_22",  # 乒乓球4号场地
                "63_23",  # 乒乓球5号场地
                "63_24"   # 乒乓球6号场地
            ]
        }
        
        # 根据用户选择的场地类型选择合适的场地列表
        venue_type = form_data.get("venue_type", "羽毛球场")
        available_venues = venue_lists.get(venue_type, venue_lists["羽毛球场"])
        
        print(f"场地类型: {venue_type}")
        print(f"可用场地列表: {available_venues}")
        
        # 从当前选择的场地开始
        current_venue_index = available_venues.index(form_data["venue_value"]) if form_data["venue_value"] in available_venues else 0
        
        return True, available_venues, current_venue_index
    except Exception as e:
        print(f"准备过程发生错误: {str(e)}")
        return False, None, None

def hold_page(driver, seconds=300):
    """保持页面打开一段时间，并显示倒计时"""
    print(f"\n页面将保持打开 {seconds} 秒")
    print("您可以随时按 Ctrl+C 关闭程序...")
    
    try:
        for remaining in range(seconds, 0, -1):
            minutes = remaining // 60
            secs = remaining % 60
            print(f"\r页面将在 {minutes:02d}:{secs:02d} 后关闭...", end='', flush=True)
            time.sleep(1)
    except KeyboardInterrupt:
        print("\n用户手动终止等待")
        return

def main():
    # 设置目标时间
    TARGET_HOUR = 12
    TARGET_MINUTE = 0
    PREPARE_SECONDS = 10  # 提前10秒开始准备
    PAGE_HOLD_TIME = 3600  # 提交后保持页面打开5分钟
    
    while True:
        current_time = datetime.now()
        next_run_time = current_time.replace(
            hour=TARGET_HOUR, 
            minute=TARGET_MINUTE, 
            second=0, 
            microsecond=0
        )
        
        # 如果当前时间超过了今天的目标时间，则设置为明天
        if current_time >= next_run_time:
            next_run_time += timedelta(days=1)
        
        # 计算准备时间
        prepare_time = next_run_time - timedelta(seconds=PREPARE_SECONDS)
        
        print(f"\n当前时间: {current_time.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"下次执行时间: {next_run_time.strftime('%Y-%m-%d %H:%M:%S')}")
        print(f"准备开始时间: {prepare_time.strftime('%Y-%m-%d %H:%M:%S')}")
        
        # 等待直到准备时间
        while datetime.now() < prepare_time:
            time.sleep(0.1)
        
        try:
            # 计算预订日期（两天后）
            booking_date = (datetime.now() + timedelta(days=2)).strftime('%Y-%m-%d')
            
            # 从params.yaml文件中读取表单数据
            try:
                with open('params.yaml', 'r', encoding='utf-8') as f:
                    params = yaml.safe_load(f)
                    form_data = {
                        "venue_type": params.get("venue_type", "羽毛球场"),
                        "participants": params.get("participants", "4"),
                        'time_value': params.get("time_value", "4"),
                        'venue_value': params.get("venue_value", "63_17"),
                        "people_category": params.get("people_category", "学生"),
                        "third_party": params.get("third_party", "无"),
                        "phone": params.get("phone", "19858214897"),
                        "take_charge_person": params.get("take_charge_person", "张嘉杰"),
                        # 使用计算出的后天日期，而不是从配置文件读取
                        "usage_date": booking_date
                    }
                    
                    # 读取登录信息
                    student_id = params.get("student_id", "2023233216")
                    password = params.get("password", "")
                    
                print(f"已从params.yaml加载配置: {form_data}")
                print(f"已从params.yaml加载登录信息: 学号={student_id}")
            except Exception as e:
                print(f"读取params.yaml失败: {e}，将使用默认配置")
                # 使用默认表单数据
                form_data = {
                    "venue_type": "羽毛球场", 
                    "participants": "2",
                    'time_value': "5",
                    'venue_value': "63_17",
                    "people_category": "学生",
                    "third_party": "无",
                    "phone": "19858214897",
                    "take_charge_person": "张嘉杰",
                    "usage_date": booking_date
                }
                # 默认登录信息
                student_id = "2023233216"
                password = ""
            
            # 时间段映射
            time_slots = {
                "1": "11:00-12:00",
                "2": "12:00-13:00",
                "3": "13:00-14:00",
                "4": "14:00-15:00",
                "5": "15:00-16:00",
                "6": "16:00-17:00",
                "7": "17:00-18:00",
                "8": "18:00-19:00",
                "9": "19:00-20:00",
                "10": "20:00-21:00",
                "11": "21:00-22:00"
            }
            
            # 场地名称映射
            venue_names = {
                "63_13": "羽毛球1号场地",
                "63_14": "羽毛球2号场地",
                "63_15": "羽毛球3号场地",
                "63_16": "羽毛球4号场地",
                "63_17": "羽毛球5号场地",
                "63_18": "羽毛球6号场地",
                "63_19": "乒乓球1号场地",
                "63_20": "乒乓球2号场地",
                "63_21": "乒乓球3号场地",
                "63_22": "乒乓球4号场地",
                "63_23": "乒乓球5号场地",
                "63_24": "乒乓球6号场地",
                "63_25": "网球场1号场地",
                "63_26": "网球场2号场地",
                "63_27": "网球场3号场地"
            }
            
            print(f"\n开始准备预订流程 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
            driver = setup_driver()
            booking = VenueBooking(driver)
            
            # 提前准备表单
            prepared, available_venues, current_venue_index = prepare_booking(driver, booking, form_data, student_id, password)
            
            if prepared:
                print(f"表单准备完成，等待提交时间 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                
                # 等待直到整点
                wait_until_time(TARGET_HOUR, TARGET_MINUTE, 0)
                
                # 立即提交
                print(f"开始提交表单 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                
                # 尝试提交当前场地
                current_venue = available_venues[current_venue_index]
                venue_name = venue_names.get(current_venue, current_venue)
                print(f"尝试预订场地: {venue_name} - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                submit_result = booking.submit_form()
                
                # 如果当前场地失败，尝试其他场地
                if not submit_result["success"]:
                    print(f"场地 {venue_name} 预订失败: {submit_result['reason']} - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                    
                    # 使用新方法尝试其他场地
                    success = booking.try_alternative_venues(form_data, available_venues, current_venue_index)
                    
                    # 显示最终结果
                    if success:
                        # 获取当前预订的场地代号
                        current_venue = booking.current_venue
                        venue_name = venue_names.get(current_venue, current_venue)
                        time_slot = time_slots.get(form_data["time_value"], f"时间段{form_data['time_value']}")
                        
                        print(f"\n恭喜! 场地预订成功! - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                        print(f"预订详情:")
                        print(f"  - 日期: {form_data['usage_date']}")
                        print(f"  - 时间段: {time_slot}")
                        print(f"  - 场地: {venue_name}")
                        print(f"  - 负责人: {form_data['take_charge_person']}")
                    else:
                        print(f"\n所有{form_data['venue_type']}场地都已被占用，无法预订 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                else:
                    time_slot = time_slots.get(form_data["time_value"], f"时间段{form_data['time_value']}")
                    # 更新当前预订的场地
                    current_venue = booking.current_venue if booking.current_venue else current_venue
                    venue_name = venue_names.get(current_venue, current_venue)
                    
                    print(f"场地 {venue_name} 预订成功! - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                    print(f"\n恭喜! 场地预订成功! - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                    print(f"预订详情:")
                    print(f"  - 日期: {form_data['usage_date']}")
                    print(f"  - 时间段: {time_slot}")
                    print(f"  - 场地: {venue_name}")
                    print(f"  - 负责人: {form_data['take_charge_person']}")
                
                # 显示最终结果
                if success or submit_result["success"]:
                    # 保持页面打开一段时间
                    hold_page(driver, PAGE_HOLD_TIME)
                else:
                    print(f"\n所有{form_data['venue_type']}场地都已被占用，无法预订 - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
                    # 提交失败也保持页面打开一段时间，方便查看错误信息
                    hold_page(driver, PAGE_HOLD_TIME)
            
            driver.quit()
            
            # 等待一段时间再开始下一次循环
            time.sleep(60)
            
        except Exception as e:
            print(f"发生错误: {str(e)}")
            # 发生错误时也保持页面打开一段时间
            try:
                hold_page(driver, PAGE_HOLD_TIME)
                driver.quit()
            except:
                pass
            
            # 出错后等待一段时间再重试
            time.sleep(60)
        
        except KeyboardInterrupt:
            print("\n程序已终止")
            try:
                driver.quit()
            except:
                pass
            break

if __name__ == "__main__":
    main()