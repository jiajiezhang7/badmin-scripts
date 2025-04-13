# 手动运行的版本
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import time
from venue_booking_info import VenueBooking
import subprocess
import os
import yaml

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

def setup_driver():
    """设置浏览器选项"""
    edge_options = Options()
    edge_options.add_argument('--log-level=3')
    edge_options.add_experimental_option('excludeSwitches', ['enable-logging'])
    
    # 直接在浏览器中禁用代理
    edge_options.add_argument('--no-proxy-server')
    edge_options.add_argument('--disable-extensions')
    edge_options.add_argument('--disable-gpu')
    edge_options.add_argument('--no-sandbox')
    edge_options.add_argument('--disable-dev-shm-usage')
    
    # 创建浏览器实例
    service = Service(EdgeChromiumDriverManager().install())
    driver = webdriver.Edge(service=service, options=edge_options)
    driver.implicitly_wait(10)
    return driver

def main():
    # 创建浏览器实例和代理设置变量
    driver = None
    original_proxy_settings = None
    
    try:
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
                    "usage_date": params.get("usage_date", "2025-3-05")
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
                "venue_type": "羽毛球场",  # 可选：羽毛球场、网球场、乒乓球
                "participants": "4",
                'time_value': "5",
                'venue_value': "63_18", # 羽毛球1-6号：63_13~63_18
                "people_category": "学生",
                "third_party": "无",
                "phone": "19858214897",
                "take_charge_person": "张嘉杰",
                "usage_date": "2025-3-09"
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
        
        # 创建浏览器实例
        driver = setup_driver()
        
        # 创建预订实例
        booking = VenueBooking(driver)
        
        # 打开登录页面
        print("正在打开登录页面...")
        driver.get("https://ids.shanghaitech.edu.cn/authserver/login?service=https%3A%2F%2Foa.shanghaitech.edu.cn%2Fworkflow%2Frequest%2FAddRequest.jsp%3Fworkflowid%3D14862")
        
        # 执行登录
        if not booking.login(student_id, password):
            raise Exception("登录失败")
        
        # 等待页面加载
        time.sleep(3)
        
        # 切换到表单frame
        if not booking.switch_to_form_frame():
            raise Exception("切换到表单frame失败")
        
        # 填写表单
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
                "63_25",  # 网球1号场地
                "63_26",  # 网球2号场地
                "63_27"   # 网球3号场地
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
        
        # 根据用户选择的场地类型选择合适的场地列表
        venue_type = form_data.get("venue_type", "羽毛球场")
        available_venues = venue_lists.get(venue_type, venue_lists["羽毛球场"])
        
        print(f"场地类型: {venue_type}")
        print(f"可用场地列表: {available_venues}")
        
        # 修改表单数据处理逻辑
        form_data_list = []
        time_values = form_data.pop("time_value")  # 从原始form_data中移除time_value
        if isinstance(time_values, str):
            time_values = [time_values]  # 兼容单个时间段的情况
            
        # 为每个时间段创建独立的form_data
        for time_value in time_values:
            current_form_data = form_data.copy()
            current_form_data["time_value"] = time_value
            form_data_list.append(current_form_data)
            
        # 按时间段优先级排序（可选）
        form_data_list.sort(key=lambda x: x["time_value"])
        
        # 存储所有成功预订的结果
        successful_bookings = []
        
        # 对每个时间段进行预订尝试
        for current_form_data in form_data_list:
            time_value = current_form_data["time_value"]
            print(f"\n尝试预订时间段 {time_slots[time_value]}")
            
            # 重置场地索引
            current_venue_index = available_venues.index(current_form_data["venue_value"]) if current_form_data["venue_value"] in available_venues else 0
            
            # 尝试提交当前场地
            print(f"尝试预订场地: {available_venues[current_venue_index]}")
            submit_result = booking.submit_form()
            
            if not submit_result["success"]:
                print(f"场地 {available_venues[current_venue_index]} 预订失败: {submit_result['reason']}")
                
                # 对当前时间段尝试其他场地
                success = booking.try_alternative_venues(current_form_data, available_venues, current_venue_index)
                
                if success:
                    successful_bookings.append({
                        "time_slot": time_slots[time_value],
                        "venue": booking.current_venue,
                        "date": current_form_data["usage_date"]
                    })
            else:
                successful_bookings.append({
                    "time_slot": time_slots[time_value],
                    "venue": available_venues[current_venue_index],
                    "date": current_form_data["usage_date"]
                })
            
            # 如果还有下一个时间段要预订，需要重新加载表单页面
            if current_form_data != form_data_list[-1]:
                if not booking.reload_form_page():
                    print("重新加载表单页面失败，无法继续预订其他时间段")
                    break
                
                # 重新填写基本表单信息
                if not booking.fill_form(current_form_data):
                    print("重新填写表单失败，无法继续预订其他时间段")
                    break
        
        # 显示所有成功预订的结果
        if successful_bookings:
            print("\n=== 预订成功的场地 ===")
            for booking_info in successful_bookings:
                venue_name = venue_names.get(booking_info["venue"], booking_info["venue"])
                print(f"日期: {booking_info['date']}")
                print(f"时间段: {booking_info['time_slot']}")
                print(f"场地: {venue_name}")
                print("------------------------")
        else:
            print("\n没有成功预订任何场地")
        
        # 保持页面打开一段时间
        time.sleep(600)
    except Exception as e:
        print(f"发生错误: {str(e)}")
    
    # 注意：这里不关闭浏览器，以便查看结果或调试

if __name__ == "__main__":
    main()
