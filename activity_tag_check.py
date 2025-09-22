import requests
import json
import pandas as pd
import os
import math
from datetime import date

# --- 1. 配置区域 ---
# 企业微信机器人 Webhook URL
WEBHOOK_URL = 'https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=5e02e1c5-9b8b-4af5-acd6-04c0f307794c' #MsgBot部门群
#WEBHOOK_URL = 'https://qyapi.weixin.qq.com/cgi-bin/webhook/send?key=6dd77fa6-0998-4a46-9c70-c8fab91c72b2' #Bot

# API 请求参数
API_URL = 'https://sales.mingya.com.cn/exercise/back/manage/list?code='
HEADERS = {
    'Accept': 'application/json, text/javascript, */*; q=0.01',
    'Content-Type': 'application/json',
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/139.0.0.0 Safari/537.36',
}
COOKIES = {
    'auth_token': 'MY_*0IfWYS-N4cbyPL3PwUUH7s_4ItMmB8koAV8p8N-rQ5aZr3ooRqEiR-kSraDPggCUNq7YdiCGUenK6ejEDlTNiLLg96F-Rcs9jBPUkkwzJsr8!',
}

# 数据筛选的起始ID (只统计 exerciseId 大于此值的活动)
CUTOFF_ID = 5659

# 分公司到大区的映射关系
ERP_DICT = {
    "北京分公司": "北区", "江苏分公司": "东区", "上海分公司": "东区",
    "陕西分公司": "西区", "深圳分公司": "南区", "重庆分公司": "西区",
    "河北分公司": "北区", "广西分公司": "南区", "内蒙古分公司": "北区",
    "四川分公司": "西区", "广东分公司": "南区", "山东分公司": "东区",
    "青岛分公司": "东区", "吉林分公司": "北区", "天津分公司": "北区",
    "浙江分公司": "东区", "湖北分公司": "西区", "河南分公司": "东区",
    "辽宁分公司": "北区", "福建分公司": "南区", "湖南分公司": "南区",
    "厦门分公司": "南区", "安徽分公司": "东区", "宁波分公司": "东区",
    "宁夏分公司": "西区", "新疆分公司": "西区", "海南分公司": "南区",
    "江西分公司": "南区", "山西分公司": "北区", "云南分公司": "西区",
    "贵州分公司": "西区", "黑龙江分公司": "北区", "大连分公司": "北区"
}

# --- 2. 功能函数定义 ---

def fetch_all_activities():
    """分页获取所有活动数据"""
    print("开始获取活动数据...")
    all_activities = []
    payload = {"page": 1, "limit": 90}
    
    try:
        response = requests.post(API_URL, headers=HEADERS, cookies=COOKIES, json=payload)
        response.raise_for_status()
        first_page_data = response.json()
        
        count = first_page_data.get('count', 0)
        if count == 0:
            print("未找到任何活动数据。")
            return []
            
        all_activities.extend(first_page_data.get('data', []))
        
        limit = payload['limit']
        total_pages = math.ceil(count / limit)
        print(f"总计 {count} 条数据，共 {total_pages} 页。")

        for page_num in range(2, total_pages + 1):
            print(f"正在获取第 {page_num}/{total_pages} 页...")
            payload['page'] = page_num
            response = requests.post(API_URL, headers=HEADERS, cookies=COOKIES, json=payload)
            response.raise_for_status()
            page_data = response.json().get('data', [])
            all_activities.extend(page_data)
            
    except requests.exceptions.RequestException as e:
        print(f"请求API失败: {e}")
        return None
    except json.JSONDecodeError:
        print("解析返回的JSON数据失败。")
        return None

    print(f"数据获取完成，共获得 {len(all_activities)} 条活动记录。")
    return all_activities

def process_and_summarize(activities, erp_mapping, cutoff_id):
    """根据ID和tag筛选并按大区汇总"""
    print(f"正在筛选和汇总数据 (只统计 exerciseId > {cutoff_id} 且tag为空的活动，排除测试活动)...")
    region_summary = {}

    for activity in activities:
        exercise_id = activity.get('exerciseId', 0)
        exercise_name = activity.get('exerciseName', '')
        
        # 全局过滤：排除测试活动
        if "测试" in exercise_name:
            continue

        # 核心筛选逻辑：tag为null或空字符串且exerciseId大于指定值
        tag = activity.get('tag')
        if (tag is None or tag == "") and exercise_id > cutoff_id:
            com_name = activity.get('workComName')
            
            if com_name == "总部":
                # 总部特殊处理：每个操作员作为独立区域
                operator_name = activity.get('operatorName', '未知操作员')
                region_key = f"总部-{operator_name}"
            else:
                # 普通分公司处理
                region_key = erp_mapping.get(com_name)
            
            if region_key:
                if region_key not in region_summary:
                    region_summary[region_key] = []
                region_summary[region_key].append(str(exercise_id))
    
    print("数据汇总完成。")
    return region_summary

def send_wechat_notification(summary):
    """发送企业微信通知"""
    if not summary:
        print(f"在 exerciseId > {CUTOFF_ID} 的活动中没有找到任何tag为空的项，无需发送通知。")
        return

    print("正在构造并发送企业微信通知...")
    today_str = date.today().strftime('%Y-%m-%d')
    total_count = sum(len(ids) for ids in summary.values())

    content_parts = [
        f"**【活动标签设置提醒】({today_str})**",
        f"> 无标签活动总计 <font color='warning'>{total_count}</font> 个，详情如下："
    ]

    # 按指定顺序输出区域：先四大区域，再总部操作员
    region_order = ["北区", "东区", "南区", "西区"]
    
    # 先输出四大区域
    for region in region_order:
        if region in summary:
            ids = summary[region]
            ids_str = '、'.join(ids)
            line = f"- **{region}** 总计：<font color='comment'>{len(ids)}</font> 个 ({ids_str})"
            content_parts.append(line)
    
    # 再输出总部操作员（按字母排序）
    headquarters_regions = [region for region in summary.keys() if region.startswith("总部-")]
    headquarters_regions.sort()
    
    for region in headquarters_regions:
        ids = summary[region]
        ids_str = '、'.join(ids)
        line = f"- **{region}** 总计：<font color='comment'>{len(ids)}</font> 个 ({ids_str})"
        content_parts.append(line)

    final_content = "\n\n".join(content_parts)
    
    webhook_payload = {
        "msgtype": "markdown",
        "markdown": {
            "content": final_content
        }
    }
    
    try:
        response = requests.post(WEBHOOK_URL, json=webhook_payload)
        response.raise_for_status()
        result = response.json()
        if result.get('errcode') == 0:
            print("企业微信通知发送成功！")
        else:
            print(f"企业微信通知发送失败: {result.get('errmsg')}")
    except requests.exceptions.RequestException as e:
        print(f"发送通知请求失败: {e}")

def save_to_excel(activities):
    """将全量数据保存到Excel文件"""
    if not activities:
        print("没有数据可以保存到Excel。")
        return
        
    print("正在将数据保存到Excel文件...")
    output_dir = 'Data'
    os.makedirs(output_dir, exist_ok=True)
    
    filename = os.path.join(output_dir, f"{date.today().strftime('%Y-%m-%d')}_activities_data.xlsx")
    
    try:
        df = pd.DataFrame(activities)
        df.to_excel(filename, index=False, engine='openpyxl')
        print(f"数据已成功保存到: {filename}")
    except Exception as e:
        print(f"保存Excel文件时出错: {e}")

def main():
    """主执行函数"""
    #script_dir = os.path.dirname(os.path.abspath(__file__))
    #os.chdir(script_dir)

    activities = fetch_all_activities()
    
    if activities is not None:
        summary = process_and_summarize(activities, ERP_DICT, CUTOFF_ID)
        send_wechat_notification(summary)
        #save_to_excel(activities)

if __name__ == '__main__':
    main()