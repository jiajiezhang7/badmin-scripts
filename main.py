# 手动运行的版本
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from webdriver_manager.microsoft import EdgeChromiumDriverManager
import time
from venue_booking_info import VenueBooking
import subprocess
import os

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
        # 设置表单数据
        # "1": "11:00-12:00",
        # "2": "12:00-13:00",
        # "3": "13:00-14:00",
        # "4": "14:00-15:00",
        # "5": "15:00-16:00",
        # "6": "16:00-17:00",
        # "7": "17:00-18:00",
        # "8": "18:00-19:00",
        # "9": "19:00-20:00",
        # "10": "20:00-21:00",
        # "11": "21:00-22:00"


        ### 羽毛球场
        # "63_13": "羽毛球1号场地",
        # "63_14": "羽毛球2号场地",
        # "63_15": "羽毛球3号场地",
        # "63_16": "羽毛球4号场地",
        # "63_17": "羽毛球5号场地",
        # "63_18": "羽毛球6号场地",
        
        ### 乒乓球场
        # | 代号 | 场地 |
        # |------|------|
        # | 63_19 | 乒乓球1号场地 |
        # | 63_20 | 乒乓球2号场地 |
        # | 63_21 | 乒乓球3号场地 |
        # | 63_22 | 乒乓球4号场地 |
        # | 63_23 | 乒乓球5号场地 |
        # | 63_24 | 乒乓球6号场地 |

        # ### 网球场
        # | 代号 | 场地 |
        # |------|------|
        # | 63_25 | 网球场1号场地 |
        # | 63_26 | 网球场2号场地 |
        # | 63_27 | 网球场3号场地 |
        form_data = {
            "venue_type": "羽毛球场",  # 可选：羽毛球场、网球场、乒乓球
            "participants": "4",
            'time_value': "4",
            'venue_value': "63_17", # 羽毛球1-6号：63_13~63_18
            "people_category": "学生",
            "third_party": "无",
            # TODO Your Phone
            "phone": "19858214897",
            # TODO Your Chinese Name
            "take_charge_person": "张嘉杰",
            # TODO Your Usage Date
            "usage_date": "2025-3-05"
        }
        
        # 创建浏览器实例
        driver = setup_driver()
        
        # 创建预订实例
        booking = VenueBooking(driver)
        
        # 打开登录页面
        print("正在打开登录页面...")
        driver.get("https://ids.shanghaitech.edu.cn/authserver/login?service=https%3A%2F%2Foa.shanghaitech.edu.cn%2Fworkflow%2Frequest%2FAddRequest.jsp%3Fworkflowid%3D14862")
        
        # 执行登录
        # TODO Your Student ID & Password
        if not booking.login("2023233216", "ZHANGjiajie123"):
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
                "63_19",  # 网球1号场地
                "63_20",  # 网球2号场地
                "63_21",  # 网球3号场地
                "63_22"   # 网球4号场地
            ],
            "乒乓球": [
                "63_23",  # 乒乓球1号场地
                "63_24",  # 乒乓球2号场地
                "63_25",  # 乒乓球3号场地
                "63_26"   # 乒乓球4号场地
            ]
        }
        
        # 根据用户选择的场地类型选择合适的场地列表
        venue_type = form_data.get("venue_type", "羽毛球场")
        available_venues = venue_lists.get(venue_type, venue_lists["羽毛球场"])
        
        print(f"场地类型: {venue_type}")
        print(f"可用场地列表: {available_venues}")
        
        # 从当前选择的场地开始
        current_venue_index = available_venues.index(form_data["venue_value"]) if form_data["venue_value"] in available_venues else 0
        
        # 尝试提交当前场地
        print(f"尝试预订场地: {available_venues[current_venue_index]}")
        submit_result = booking.submit_form()
        
        # 如果当前场地失败，尝试其他场地
        if not submit_result["success"]:
            print(f"场地 {available_venues[current_venue_index]} 预订失败: {submit_result['reason']}")
            
            # 使用新方法尝试其他场地
            success = booking.try_alternative_venues(form_data, available_venues, current_venue_index)
            
            # 显示最终结果
            if success:
                print("\n恭喜! 场地预订成功!")
            else:
                print(f"\n所有{venue_type}场地都已被占用，无法预订")
        else:
            print(f"场地 {available_venues[current_venue_index]} 预订成功!")
            print("\n恭喜! 场地预订成功!")
        
        # 等待几秒查看结果
        time.sleep(600)
    except Exception as e:
        print(f"发生错误: {str(e)}")
    
    # 注意：这里不关闭浏览器，以便查看结果或调试

if __name__ == "__main__":
    main()