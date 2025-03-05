from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait, Select
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException, UnexpectedAlertPresentException, NoAlertPresentException
import time

class VenueBooking:
    # 时间段映射表
    TIME_MAPPING = {
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
    
    # 场地号映射表
    VENUE_MAPPING = {
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
        "63_27": "网球场3号场地",
    }

    def __init__(self, driver):
        self.driver = driver
        self.form_data = None  # 存储表单数据以供后续使用
        self.current_venue = None  # 存储当前预订的场地

    def wait_for_element(self, by, value, timeout=10, action="visible"):
        try:
            if action == "visible":
                element = WebDriverWait(self.driver, timeout).until(
                    EC.visibility_of_element_located((by, value))
                )
            elif action == "clickable":
                element = WebDriverWait(self.driver, timeout).until(
                    EC.element_to_be_clickable((by, value))
                )
            elif action == "presence":
                element = WebDriverWait(self.driver, timeout).until(
                    EC.presence_of_element_located((by, value))
                )
            print(f"成功找到元素: {value}")
            return element
        except TimeoutException:
            print(f"超时: 未能找到元素 {value}")
            return None
        except Exception as e:
            print(f"查找元素 {value} 时发生错误: {str(e)}")
            return None

    def login(self, username, password):
        """登录系统"""
        try:
            print("正在执行登录...")
            username_element = self.wait_for_element(By.ID, "username")
            if username_element:
                username_element.send_keys(username)
            
            password_element = self.wait_for_element(By.ID, "password")
            if password_element:
                password_element.send_keys(password)
            
            login_button = self.wait_for_element(By.ID, "login_submit", action="clickable")
            if login_button:
                login_button.click()
                print("已点击登录按钮")
                return True
        except Exception as e:
            print(f"登录过程发生错误: {str(e)}")
            return False

    def switch_to_form_frame(self):
        """切换到表单所在的iframe"""
        try:
            print("正在切换到bodyiframe...")
            time.sleep(2)  # 等待iframe加载
            self.wait_for_element(By.ID, "bodyiframe", action="presence")
            self.driver.switch_to.frame("bodyiframe")
            return True
        except Exception as e:
            print(f"切换到iframe时发生错误: {str(e)}")
            return False

    def fill_form(self, form_data):
        """填写表单"""
        try:
            self.form_data = form_data  # 保存表单数据
            print("开始填写表单...")
            
            # 场馆类型 (下拉框)
            venue_type = self.wait_for_element(By.NAME, "field32340", action="presence")
            if venue_type:
                Select(venue_type).select_by_visible_text(form_data["venue_type"])
                print("已选择场馆类型")

            # 使用日期处理
            usage_date = self.wait_for_element(By.ID, "field31901", action="presence")
            verify_date = self.wait_for_element(By.NAME, "ecology_verify_field31901", action="presence")
            if usage_date and verify_date:
                date_value = form_data["usage_date"]
                self.driver.execute_script("""
                    arguments[0].value = arguments[2];
                    arguments[1].value = arguments[2];
                """, usage_date, verify_date, date_value)
                print("已设置使用日期")

            # 使用时间处理
            usage_time = self.wait_for_element(By.ID, "field31902", action="presence")
            verify_time_1 = self.wait_for_element(By.NAME, "field31902__", action="presence")
            verify_time_2 = self.wait_for_element(By.NAME, "ecology_verify_field31902", action="presence")
            verify_time_extra = self.wait_for_element(By.NAME, "ecology_verify_field31902__", action="presence")

            if usage_time and verify_time_1 and verify_time_2 and verify_time_extra:
                time_value = form_data["time_value"]
                self.driver.execute_script("""
                    arguments[0].value = arguments[4];
                    arguments[1].value = arguments[4];
                    arguments[2].value = arguments[4];
                    arguments[3].value = arguments[4];
                """, usage_time, verify_time_1, verify_time_2, verify_time_extra, time_value)
                print("已设置使用时间")
            
            # 使用场地处理
            venue_field = self.wait_for_element(By.ID, "field31883", action="presence")
            verify_venue_1 = self.wait_for_element(By.ID, "field31883__", action="presence")
            verify_venue_2 = self.wait_for_element(By.NAME, "ecology_verify_field31883", action="presence")
            verify_venue_extra = self.wait_for_element(By.NAME, "ecology_verify_field31883__", action="presence")
            if venue_field and verify_venue_1 and verify_venue_2 and verify_venue_extra:
                venue_value = form_data["venue_value"]
                self.driver.execute_script("""
                    arguments[0].value = arguments[4];
                    arguments[1].value = arguments[4];
                    arguments[2].value = arguments[4];
                    arguments[3].value = arguments[4];
                """, venue_field, verify_venue_1, verify_venue_2, verify_venue_extra, venue_value)
                print("已设置使用场地")

            # 参加人数
            participants = self.wait_for_element(By.NAME, "field31884", action="presence")
            if participants:
                participants.clear()
                participants.send_keys(form_data["participants"])
                print("已填写参加人数")
            
            # 现场责任人
            take_charge_person = self.wait_for_element(By.NAME, "field31888", action="presence")
            if take_charge_person:
                take_charge_person.clear()
                take_charge_person.send_keys(form_data["take_charge_person"])
                print("已填写现场责任人")

            # 人员类别 (下拉框)
            people_category = self.wait_for_element(By.NAME, "field31885", action="presence")
            if people_category:
                Select(people_category).select_by_visible_text(form_data["people_category"])
                print("已选择人员类别")
            
            # 有无第三方服务 (下拉框)
            third_party = self.wait_for_element(By.NAME, "field31892", action="presence")
            if third_party:
                Select(third_party).select_by_visible_text(form_data["third_party"])
                print("已选择第三方服务信息")

            # 联系电话 
            phone = self.wait_for_element(By.NAME, "field31889", action="presence")
            if phone:
                phone.clear()
                phone.send_keys(form_data["phone"])
                print("已填写联系电话")
            
            return True
            
        except UnexpectedAlertPresentException as e:
            alert = self.driver.switch_to.alert
            print(f"警告文本: {alert.text}")
            alert.accept()  # 关闭弹窗
            print("已关闭警告窗口")
            return False
        except Exception as e:
            print(f"填写表单时发生错误: {str(e)}")
            return False

    def reload_form_page(self):
        """重新加载表单页面"""
        try:
            print("正在重新加载表单页面...")
            # 切换回默认内容
            self.driver.switch_to.default_content()
            
            # 重新加载页面
            self.driver.get("https://oa.shanghaitech.edu.cn/workflow/request/AddRequest.jsp?workflowid=14862")
            time.sleep(2)
            
            # 切换到表单frame
            return self.switch_to_form_frame()
        except Exception as e:
            print(f"重新加载表单页面时发生错误: {str(e)}")
            return False

    def fill_venue(self, venue_value):
        """仅重新填写场地信息"""
        try:
            # 确保我们在正确的frame中
            try:
                # 先切换回默认内容
                self.driver.switch_to.default_content()
                print("已切换到默认内容")
                
                # 尝试切换到bodyiframe
                bodyiframe = self.wait_for_element(By.ID, "bodyiframe")
                if bodyiframe:
                    self.driver.switch_to.frame(bodyiframe)
                    print("已切换到bodyiframe")
            except Exception as e:
                print(f"切换到bodyiframe时发生错误: {str(e)}")
                return False
            
            # 选择场地
            venue_field = self.wait_for_element(By.ID, "field31883")
            venue_field_verify = self.wait_for_element(By.ID, "ecology_verify_field31883")
            
            if venue_field and venue_field_verify:
                # 清除原有选择
                venue_field.clear()
                venue_field_verify.clear()
                
                # 选择新场地
                venue_field.send_keys(venue_value.split("_")[1])
                venue_field_verify.send_keys(venue_value)
                
                # 确认选择
                verify_field = self.wait_for_element(By.ID, "ecology_verify_field31883__")
                if verify_field:
                    verify_field.click()
                    time.sleep(0.5)
                
                print(f"已重新设置使用场地: {venue_value}")
                return True
            else:
                print("未找到场地选择字段")
                return False
        except Exception as e:
            print(f"重新填写场地时发生错误: {str(e)}")
            return False

    def try_alternative_venues(self, form_data, available_venues, current_venue_index):
        """尝试预订其他场地"""
        tried_venues = [available_venues[current_venue_index]]
        success = False
        
        # 尝试其他场地
        for i in range(len(available_venues)):
            # 跳过已尝试的场地
            if available_venues[i] in tried_venues:
                continue
            
            print(f"\n尝试预订其他场地: {available_venues[i]}")
            
            # 重新加载表单页面
            if not self.reload_form_page():
                print("重新加载表单页面失败，无法尝试其他场地")
                continue
            
            # 重新填写表单，但使用新的场地
            new_form_data = form_data.copy()
            new_form_data["venue_value"] = available_venues[i]
            
            if not self.fill_form(new_form_data):
                print(f"重新填写表单失败，无法尝试场地 {available_venues[i]}")
                continue
            
            # 提交表单
            submit_result = self.submit_form()
            tried_venues.append(available_venues[i])
            
            if submit_result["success"]:
                print(f"场地 {available_venues[i]} 预订成功!")
                self.current_venue = available_venues[i]  # 更新当前预订的场地
                success = True
                break
            else:
                print(f"场地 {available_venues[i]} 预订失败: {submit_result['reason']}")
        
        return success

    def submit_form(self):
        """提交表单"""
        try:
            print("正在提交表单...")
            submit_button = self.wait_for_element(
                By.CLASS_NAME, "e8_btn_top_first", 
                action="clickable"
            )
            if submit_button:
                submit_button.click()
                print("已点击提交按钮")
                
                # 等待页面加载（可能包含警告信息）
                time.sleep(2)
                
                # 尝试处理成功提交的弹窗
                try:
                    # 切换到弹窗
                    alert = self.driver.switch_to.alert
                    alert_text = alert.text
                    print(f"检测到弹窗: {alert_text}")
                    
                    # 如果弹窗包含"Process of Success"，则表示提交成功
                    if "Process of Success" in alert_text:
                        print("检测到成功提交的弹窗")
                        # 接受弹窗
                        alert.accept()
                        # 如果表单数据存在，更新当前预订的场地
                        if self.form_data and "venue_value" in self.form_data:
                            self.current_venue = self.form_data["venue_value"]
                        return {"success": True, "reason": ""}
                    
                    # 接受弹窗
                    alert.accept()
                except Exception as e:
                    # 如果没有弹窗，继续检查其他成功标志
                    print(f"没有检测到弹窗或处理弹窗时出错: {str(e)}")
                
                # 检查当前URL，如果包含WorkflowDirection，则表示提交成功
                current_url = self.driver.current_url
                if "WorkflowDirection" in current_url:
                    print(f"检测到成功提交的URL: {current_url}")
                    # 如果表单数据存在，更新当前预订的场地
                    if self.form_data and "venue_value" in self.form_data:
                        self.current_venue = self.form_data["venue_value"]
                    return {"success": True, "reason": ""}
                
                # 确保我们在正确的frame中
                try:
                    # 先切换回默认内容
                    self.driver.switch_to.default_content()
                    print("已切换到默认内容")
                    
                    # 再次检查URL，因为可能在切换frame后URL已经变化
                    current_url = self.driver.current_url
                    if "WorkflowDirection" in current_url:
                        print(f"检测到成功提交的URL: {current_url}")
                        # 如果表单数据存在，更新当前预订的场地
                        if self.form_data and "venue_value" in self.form_data:
                            self.current_venue = self.form_data["venue_value"]
                        return {"success": True, "reason": ""}
                    
                    # 尝试切换到bodyiframe
                    bodyiframe = self.wait_for_element(By.ID, "bodyiframe")
                    if bodyiframe:
                        self.driver.switch_to.frame(bodyiframe)
                        print("已切换到bodyiframe")
                    
                    # 检查页面源代码中是否包含特定文本
                    page_source = self.driver.page_source
                    warning_texts = ["已被占用", "无法借用", "无法使用", "该场馆"]
                    
                    for warning_text in warning_texts:
                        if warning_text in page_source:
                            print(f"在bodyiframe中发现警告文本: '{warning_text}'")
                            
                            # 尝试找到并打印包含警告文本的具体元素
                            try:
                                warning_elements = self.driver.find_elements(By.XPATH, 
                                    f"//*[contains(text(), '{warning_text}')]")
                                if warning_elements:
                                    for element in warning_elements:
                                        print(f"警告元素文本: {element.text}")
                            except:
                                pass
                            
                            # 返回失败状态和原因
                            return {"success": False, "reason": "场地已被占用"}
                    
                    # 查找特定的错误图标和消息
                    try:
                        error_icons = self.driver.find_elements(By.XPATH, 
                            "//img[contains(@src, 'error') or contains(@src, 'warning') or contains(@src, 'failed')]")
                        if error_icons:
                            print("在bodyiframe中发现错误图标")
                            return {"success": False, "reason": "发现错误图标"}
                    except:
                        pass
                    
                    # 查找特定的错误消息区域
                    try:
                        error_divs = self.driver.find_elements(By.XPATH, 
                            "//div[contains(@class, 'error') or contains(@class, 'warning') or contains(@class, 'failed')]")
                        if error_divs:
                            for div in error_divs:
                                error_text = div.text
                                if error_text:
                                    print(f"错误区域文本: {error_text}")
                                    return {"success": False, "reason": "发现错误区域"}
                    except:
                        pass
                    
                    # 查找特定的红色文本
                    try:
                        red_texts = self.driver.find_elements(By.XPATH, 
                            "//*[contains(@style, 'color: red') or contains(@style, 'color:#f00') or contains(@style, 'color: #ff0000')]")
                        if red_texts:
                            for text in red_texts:
                                red_content = text.text
                                if red_content:
                                    print(f"红色文本: {red_content}")
                                    return {"success": False, "reason": "发现红色警告文本"}
                    except:
                        pass
                    
                    # 打印整个页面文本，帮助调试
                    print("页面文本:")
                    print(self.driver.find_element(By.TAG_NAME, "body").text)
                    
                except Exception as e:
                    # 检查错误消息中是否包含"Process of Success"
                    if "Process of Success" in str(e):
                        print("从错误消息中检测到成功提交的信号")
                        return {"success": True, "reason": ""}
                    
                    print(f"检查警告时发生错误: {str(e)}")
                    return {"success": False, "reason": f"检查警告时发生错误: {str(e)}"}
                
                # 如果没有发现任何警告或错误信息，则认为提交成功
                print("未发现警告信息，提交可能成功")
                return {"success": True, "reason": ""}
                
        except Exception as e:
            # 检查错误消息中是否包含"Process of Success"
            if "Process of Success" in str(e):
                print("从错误消息中检测到成功提交的信号")
                return {"success": True, "reason": ""}
            
            print(f"提交表单时发生错误: {str(e)}")
            return {"success": False, "reason": f"提交表单时发生错误: {str(e)}"}

    def get_booking_summary(self):
        """获取预订信息摘要"""
        if not self.form_data:
            return "无预订信息"
            
        time_str = self.TIME_MAPPING.get(self.form_data["time_value"], "未知时间段")
        venue_str = self.VENUE_MAPPING.get(self.form_data["venue_value"], "未知场地")
        
        return f"""
预订成功！详细信息：
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
使用日期：{self.form_data["usage_date"]}
时间段：{time_str}
场地：{venue_str}
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━"""