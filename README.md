# 羽毛球场地预订脚本使用说明 

这个项目包含两个主要的Python脚本用于上海科技大学场地预订系统的自动化预订：
- `auto.py`: 自动定时版本，默认在每天中午12:00自动抢后天的场地
- `main.py`: 手动运行版本，可以随时运行进行场地预订

## 时间段代号对照表

预订时间使用数字代号（1-11）表示不同的时间段：

| 代号 | 时间段 |
|------|--------|
| 1 | 11:00-12:00 |
| 2 | 12:00-13:00 |
| 3 | 13:00-14:00 |
| 4 | 14:00-15:00 |
| 5 | 15:00-16:00 |
| 6 | 16:00-17:00 |
| 7 | 17:00-18:00 |
| 8 | 18:00-19:00 |
| 9 | 19:00-20:00 |
| 10 | 20:00-21:00 |
| 11 | 21:00-22:00 |

## 场地代号对照表

场地使用特定代号（63_XX）表示不同的场地：

### 羽毛球场
| 代号 | 场地 |
|------|------|
| 63_13 | 羽毛球1号场地 |
| 63_14 | 羽毛球2号场地 |
| 63_15 | 羽毛球3号场地 |
| 63_16 | 羽毛球4号场地 |
| 63_17 | 羽毛球5号场地 |
| 63_18 | 羽毛球6号场地 |

### 乒乓球场
| 代号 | 场地 |
|------|------|
| 63_19 | 乒乓球1号场地 |
| 63_20 | 乒乓球2号场地 |
| 63_21 | 乒乓球3号场地 |
| 63_22 | 乒乓球4号场地 |
| 63_23 | 乒乓球5号场地 |
| 63_24 | 乒乓球6号场地 |

### 网球场
| 代号 | 场地 |
|------|------|
| 63_25 | 网球场1号场地 |
| 63_26 | 网球场2号场地 |
| 63_27 | 网球场3号场地 |

## 功能特点

1. **自动填写表单**：根据预设的表单数据自动填写场馆预订表单
2. **定时预订**：支持设定特定时间自动执行预订，确保在开放时间第一时间提交
3. **代理支持**：可配置代理服务器，避免IP限制
4. **警告检测**：能够检测网页中的警告信息，判断场地是否已被占用
5. **自动尝试其他场地**：当当前选择的场地已被占用时，会自动尝试预订其他可用场地，提高预订成功率
6. **多重成功检测**：通过多种方式检测表单是否提交成功
   - 检测"Process of Success"弹窗
   - 检查URL中是否包含"WorkflowDirection"
   - 分析错误消息中的成功信号
7. **表单重新加载**：在尝试其他场地前会重新加载表单页面，确保元素ID正确
8. **详细日志**：提供详细的操作日志，方便排查问题

## 使用方法

### 手动版本

1. 编辑 `params.yaml` 文件，设置您的学号、密码和预订信息
2. 运行脚本：`python main.py`
3. 脚本将自动填写表单并提交
4. 如果当前选择的场地已被占用，脚本会自动尝试预订其他可用场地
5. 脚本会通过多种方式检测预订是否成功，并给出详细的预订信息（包括日期、时间段、场地名称和负责人）

### 自动版本

1. 编辑 `params.yaml` 文件，设置您的学号、密码和预订信息
2. 设置预订时间（默认为每天中午12:00）
3. 运行脚本：`python auto.py`
4. 脚本将在指定时间自动执行预订
5. 如果当前选择的场地已被占用，脚本会自动尝试预订其他可用场地
6. 脚本会通过多种方式检测预订是否成功，并给出详细的预订信息（包括日期、时间段、场地名称和负责人）

## 环境配置

### 安装Python包

在使用脚本前，请确保已安装以下Python包：

```bash
pip install selenium==4.10.0
pip install webdriver-manager==3.8.6
```

或者使用requirements.txt文件安装：

```bash
pip install -r requirements.txt
```

### 浏览器要求

脚本默认使用Microsoft Edge浏览器，请确保已安装最新版本的Edge浏览器。如需使用其他浏览器，请修改代码中的浏览器驱动部分。

### 代理设置（可选）

如果需要使用代理服务器，请在`auto.py`文件中设置：

```python
USE_PROXY = True
PROXY_SERVER = "http://your-proxy-server:port"
```

## 配置文件说明

脚本使用`params.yaml`文件存储所有配置参数，包括登录信息和预订信息。以下是配置文件的主要参数：

```yaml
# 登录信息
student_id: "2023233216"  # 您的学号
password: "YourPassword"  # 您的密码

# 预订信息
venue_type: "羽毛球场"  # 场馆类型：羽毛球场、乒乓球、网球场
participants: "4"  # 参加人数
time_value: ["6", "7"]  # 时间段代号列表，例如["6", "7"]表示16:00-17:00和17:00-18:00
venue_value: "63_17"  # 场地代号，例如"63_17"表示羽毛球5号场地
people_category: "学生"  # 人员类别
third_party: "无"  # 有无第三方服务
phone: "12345678901"  # 联系电话
take_charge_person: "张三"  # 负责人姓名
usage_date: "2025-3-06"  # 使用日期，格式为YYYY-M-DD
```

### 时间段对照表

| 代号 | 时间段 |
|------|------|
| 1 | 11:00-12:00 |
| 2 | 12:00-13:00 |
| 3 | 13:00-14:00 |
| 4 | 14:00-15:00 |
| 5 | 15:00-16:00 |
| 6 | 16:00-17:00 |
| 7 | 17:00-18:00 |
| 8 | 18:00-19:00 |
| 9 | 19:00-20:00 |
| 10 | 20:00-21:00 |
| 11 | 21:00-22:00 |

## 注意事项

- 预订时请遵守场地预订系统的相关规定
- 建议提前几分钟运行脚本，以确保在目标时间能够准确预订
- 如果遇到网络问题或系统更新，可能需要调整脚本中的元素选择器
- 脚本会自动处理"Process of Success"弹窗，无需手动干预
- 在尝试其他场地前会重新加载表单页面，以确保元素ID正确
