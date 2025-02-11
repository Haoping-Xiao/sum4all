import requests
import json
import time
from datetime import datetime
from bridge import bridge
from common import const
from common.log import logger
from openai import OpenAI
from config import conf
import os

curdir = os.path.dirname(__file__)
config_path = os.path.join(curdir, "config.json")
if os.path.exists(config_path):
    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)
        NOTION_API_URL = config.get("keys", {}).get("NOTION_API_URL")
        NOTION_TOKEN = config.get("keys", {}).get("NOTION_TOKEN")
        assert NOTION_TOKEN, "miss NOTION_TOKEN"
        assert NOTION_API_URL, "miss NOTION_API_URL"

def get_tile_tag(note, content):
    system_msg_content = """你是专业的笔记整理专家，帮助用户总结笔记标题与标签，以便于检索。给定笔记内容<note>和参考内容<content>, 以JSON结构返回便于检索的标题<title> 和标签列表<tags>
    返回格式如下:
    {
        "title": "...",
        "tags": ["tag1", "tag2", ...]
    }
    """
    system_msg = {"role": "system", "content": system_msg_content}
    user_msg = {"role": "user", "content": f"笔记内容\n<note>{note}</note>\n\n参考内容<content>{content}</content>"}
    client = OpenAI(
        # 若没有配置环境变量，请用百炼API Key将下行替换为：api_key="sk-xxx",
        api_key= conf().get("dashscope_api_key"),
        base_url=conf().get("open_ai_api_base"),  # 填写DashScope服务的base_url
    )
    completion = client.chat.completions.create(
        model=conf().get("model"),
        messages=[system_msg, user_msg],
        response_format={"type": "json_object"}
    )
    json_string = completion.choices[0].message.content
    logger.debug(f"get json result {json_string}")
    result = json.loads(json_string)
    return result.get("title"), result.get("tags")
    


def create_notion_page(link, content, note):
    # get title and tag_list from LLM
    title, tag_list = get_tile_tag(note, content)
    tag_dict_list = [{"name": tag} for tag in tag_list]

    data = {
        "parent": { "database_id": '18f99608e4d180c09091d33e7e691e43' },
        "properties": {
            "Title": {
                "title": [
                    {
                        "text": {
                            "content": title
                        }
                    }
                ]
            },
            "Tag": {
                "multi_select": tag_dict_list
            },
            "Created Date":{
                "date":{
                    "start": datetime.now().isoformat()
                }
            },
            "URL": {
                "url": link
            }
        },
        "children":[
        {
            "object": "block",
            "type": "paragraph",
            "paragraph": {
            "rich_text": [
                {
                "type": "text",
                "text": {
                    "content": note,
                    "link": { "url": link }
                }
                }
            ]
            },
            }
        ]

    }
    
    # 请求头部
    headers = {
        "Authorization": f"Bearer {NOTION_TOKEN}",
        "Content-Type": "application/json",
        "Notion-Version": "2022-06-28"  # 使用 Notion API 的版本
    }

    try:
        logger.debug(f"notion requests body {data}")
        response = requests.post(NOTION_API_URL, headers=headers, data=json.dumps(data))
        response.raise_for_status()
        url = response.json()['url']
        return url
    except Exception as e:
        logger.error(e)
        return ''

        


