import requests
import json
import re
import traceback
import plugins
from bridge import bridge
from bridge.reply import Reply, ReplyType
from bridge.context import ContextType
from channel.chat_message import ChatMessage
from plugins import *
from common import const
from common.log import logger
from common.expired_dict import ExpiredDict
import os
from docx import Document
import markdown
import fitz
from openpyxl import load_workbook
import csv
from bs4 import BeautifulSoup
from pptx import Presentation
from PIL import Image
import base64
import html
from .video_to_text import url_to_text, upload_audio_to_oss, remove_audio_from_oss
from .notion_helper import create_notion_page


EXTENSION_TO_TYPE = {
    'pdf': 'pdf',
    'doc': 'docx', 'docx': 'docx',
    'md': 'md',
    'txt': 'txt',
    'xls': 'excel', 'xlsx': 'excel',
    'csv': 'csv',
    'html': 'html', 'htm': 'html',
    'ppt': 'ppt', 'pptx': 'ppt'
}

@plugins.register(
    name="sum4all",
    desire_priority=2,
    desc="A plugin for summarizing all things",
    version="0.7.11",
    author="fatwang2",
)

class sum4all(Plugin):
    def __init__(self):
        super().__init__()
        try:
            curdir = os.path.dirname(__file__)
            config_path = os.path.join(curdir, "config.json")
            if os.path.exists(config_path):
                with open(config_path, "r", encoding="utf-8") as f:
                    self.config = json.load(f)
            else:
                # 使用父类的方法来加载配置
                self.config = super().load_config()

                if not self.config:
                    raise Exception("config.json not found")
            # 设置事件处理函数
            self.handlers[Event.ON_HANDLE_CONTEXT] = self.on_handle_context
            self.params_cache = ExpiredDict(300)

            
            # 从配置中提取所需的设置
            self.keys = self.config.get("keys", {})
            self.url_sum = self.config.get("url_sum", {})
            self.search_sum = self.config.get("search_sum", {})
            self.file_sum = self.config.get("file_sum", {})
            self.image_sum = self.config.get("image_sum", {})
            self.note = self.config.get("note", {})

            self.sum4all_key = self.keys.get("sum4all_key", "")
            self.search1api_key = self.keys.get("search1api_key", "")
            self.gemini_key = self.keys.get("gemini_key", "")
            self.bibigpt_key = self.keys.get("bibigpt_key", "")
            self.outputLanguage = self.keys.get("outputLanguage", "zh-CN")
            self.opensum_key = self.keys.get("opensum_key", "")
            self.open_ai_api_key = self.keys.get("open_ai_api_key", "")
            self.model = self.keys.get("model", "gpt-3.5-turbo")
            self.open_ai_api_base = self.keys.get("open_ai_api_base", "https://api.openai.com/v1")
            self.azure_deployment_id = self.keys.get("azure_deployment_id", "")
            self.xunfei_app_id = self.keys.get("xunfei_app_id", "")
            self.xunfei_api_key = self.keys.get("xunfei_api_key", "")
            self.xunfei_api_secret = self.keys.get("xunfei_api_secret", "")
            self.perplexity_key = self.keys.get("perplexity_key", "")
            self.flomo_key = self.keys.get("flomo_key", "")

            # 提取sum服务的配置
            self.url_sum_enabled = self.url_sum.get("enabled", False)
            self.url_sum_api_key = self.url_sum.get("api_key", "")
            self.url_sum_service = self.url_sum.get("service", "")
            self.url_sum_group = self.url_sum.get("group", True)
            self.url_sum_qa_enabled = self.url_sum.get("qa_enabled", True)
            self.url_sum_qa_prefix = self.url_sum.get("qa_prefix", "问")
            self.url_sum_prompt = self.url_sum.get("prompt", "")
            self.url_sum_video_prompt = self.url_sum.get("video_prompt", "")

            self.search_sum_enabled = self.search_sum.get("enabled", False)
            self.search_sum_service = self.search_sum.get("service", "")
            self.search_service = self.search_sum.get("search_service", "duckduckgo")
            self.search_sum_group = self.search_sum.get("group", True)
            self.search_sum_search_prefix = self.search_sum.get("search_prefix", "搜")
            self.search_sum_prompt = self.search_sum.get("prompt", "")

            self.file_sum_enabled = self.file_sum.get("enabled", False)
            self.file_sum_service = self.file_sum.get("service", "")
            self.max_file_size = self.file_sum.get("max_file_size", 15000)
            self.file_sum_group = self.file_sum.get("group", True)
            self.file_sum_qa_prefix = self.file_sum.get("qa_prefix", "问")
            self.file_sum_prompt = self.file_sum.get("prompt", "")

            self.image_sum_enabled = self.image_sum.get("enabled", False)
            self.image_sum_service = self.image_sum.get("service", "")
            self.image_sum_group = self.image_sum.get("group", True)
            self.image_sum_qa_prefix = self.image_sum.get("qa_prefix", "问")
            self.image_sum_prompt = self.image_sum.get("prompt", "")

            self.note_user_id = self.note.get("user_id", "")
            self.note_enabled = self.note.get("enabled", False)
            self.note_service = self.note.get("service", "")
            self.note_prefix = self.note.get("prefix", "记")

            # 初始化成功日志
            logger.info("[sum4all] inited.")
        except Exception as e:
            # 初始化失败日志
            logger.warn(f"sum4all init failed: {e}")
    def on_handle_context(self, e_context: EventContext):
        context = e_context["context"]
        logger.info(f"on_handle_context: {context}")
        if context.type not in [ContextType.TEXT, ContextType.SHARING,ContextType.FILE,ContextType.IMAGE, ContextType.VIDEO]:
            return
        msg: ChatMessage = e_context["context"]["msg"]
        user_id = msg.from_user_id
        content = context.content
        isgroup = e_context["context"].get("isgroup", False)

        url_match = re.match('https?://(?:[-\w.]|(?:%[\da-fA-F]{2}))+', content)
        unsupported_urls = re.search(r'.*finder\.video\.qq\.com.*|.*support\.weixin\.qq\.com/update.*|.*support\.weixin\.qq\.com/security.*|.*mp\.weixin\.qq\.com/mp/waerrpage.*', content)

            # 检查输入是否以"搜索前缀词" 开头
        if content.startswith(self.search_sum_search_prefix) and self.search_sum_enabled:
            # 如果消息来自一个群聊，并且你不希望在群聊中启用搜索功能，直接返回
            if isgroup and not self.search_sum_group:
                return
            # Call new function to handle search operation
            self.call_service(content, e_context, "search")
            return
        
        if user_id in self.params_cache and ('last_file_content' in self.params_cache[user_id] or 'last_image_base64' in self.params_cache[user_id] or 'last_url' in self.params_cache[user_id]):
            # 如果存在最近一次处理的文件路径，触发文件理解函数
            if 'last_file_content' in self.params_cache[user_id] and content.startswith(self.file_sum_qa_prefix):
                logger.info('Content starts with the file_sum_qa_prefix.')
                # 去除关键词和紧随其后的空格
                new_content = content[len(self.file_sum_qa_prefix):]
                self.params_cache[user_id]['prompt'] = new_content
                logger.info('params_cache for user has been successfully updated.')            
                self.handle_file(self.params_cache[user_id]['last_file_content'], e_context)
            # 如果存在最近一次处理的图片路径，触发图片理解函数
            elif 'last_image_base64' in self.params_cache[user_id] and content.startswith(self.image_sum_qa_prefix):
                logger.info('Content starts with the image_sum_qa_prefix.')
                # 去除关键词和紧随其后的空格
                new_content = content[len(self.image_sum_qa_prefix):]
                self.params_cache[user_id]['prompt'] = new_content
                logger.info('params_cache for user has been successfully updated.')            
                self.handle_image(self.params_cache[user_id]['last_image_base64'], e_context)

            # 如果存在最近一次处理的URL，触发URL理解函数
            elif 'last_url' in self.params_cache[user_id] and content.startswith(self.url_sum_qa_prefix):
                logger.info('Content starts with the url_sum_qa_prefix.')
                # 去除关键词和紧随其后的空格
                new_content = content[len(self.url_sum_qa_prefix):]
                self.params_cache[user_id]['prompt'] = new_content
                logger.info('params_cache for user has been successfully updated.')            
                self.call_service(self.params_cache[user_id]['last_url'], e_context ,"sum")
            elif 'last_url' in self.params_cache[user_id] and content.startswith(self.note_prefix) and self.note_enabled and self.note_user_id == user_id:
                logger.info('Content starts with the note_prefix.')
                new_content = content[len(self.note_prefix):]
                self.params_cache[user_id]['note'] = new_content
                logger.info('params_cache for user has been successfully updated.')  
                self.call_service(self.params_cache[user_id]['last_url'], e_context, "note")
        if context.type == ContextType.FILE:
            if isgroup and not self.file_sum_group:
                # 群聊中忽略处理文件
                logger.info("群聊消息，文件处理功能已禁用")
                return
            logger.info("on_handle_context: 处理上下文开始")
            context.get("msg").prepare()
            file_path = context.content
            logger.info(f"on_handle_context: 获取到文件路径 {file_path}")
            
            # 检查是否应该进行文件总结
            if self.file_sum_enabled:
                # 更新params_cache中的last_file_content
                self.params_cache[user_id] = {}
                file_content = self.extract_content(file_path)
                if file_content is None:
                    logger.info("文件内容无法提取，跳过处理")
                else:
                    self.params_cache[user_id]['last_file_content'] = file_content
                    logger.info('Updated last_file_content in params_cache for user.')
                    self.handle_file(file_content, e_context)
            else:
                logger.info("文件总结功能已禁用，不对文件内容进行处理")
            # 删除文件
            os.remove(file_path)
            logger.info(f"文件 {file_path} 已删除")
        elif context.type == ContextType.IMAGE:
            if isgroup and not self.image_sum_group:
                # 群聊中忽略处理图片
                logger.info("群聊消息，图片处理功能已禁用")
                return
            logger.info("on_handle_context: 开始处理图片")
            context.get("msg").prepare()
            image_path = context.content
            logger.info(f"on_handle_context: 获取到图片路径 {image_path}")
            
            
            # 检查是否应该进行图片总结
            if self.image_sum_enabled:
                # 将图片路径转换为Base64编码的字符串
                url = upload_audio_to_oss(image_path)
                # base64_image = self.encode_image_to_base64(image_path)
                # 更新params_cache中的last_image_path
                self.params_cache[user_id] = {}
                self.params_cache[user_id]['last_image_base64'] = url
                logger.info('Updated last_image_base64 in params_cache for user.')
                self.handle_image(url, e_context)

            else:
                logger.info("图片总结功能已禁用，不对图片内容进行处理")
            # 删除文件
            os.remove(image_path)
            logger.info(f"文件 {image_path} 已删除")
        elif context.type == ContextType.VIDEO:
            try:
                # 处理视频号
                self.params_cache[user_id] = {}
                self.params_cache[user_id]['last_url'] = content
                self.handle_video(content, e_context)
            except Exception as e:
                logger.error(e)
                traceback.print_exc()
        elif context.type == ContextType.SHARING and self.url_sum_enabled:  #匹配卡片分享
            content = html.unescape(content)
            if unsupported_urls:  #匹配不支持总结的卡片
                if isgroup:  ##群聊中忽略
                    return
                else:  ##私聊回复不支持
                    logger.info("[sum4all] Unsupported URL : %s", content)
                    reply = Reply(type=ReplyType.TEXT, content="不支持总结小程序和视频号")
                    e_context["reply"] = reply
                    e_context.action = EventAction.BREAK_PASS
            else:  #匹配支持总结的卡片
                if isgroup:  #处理群聊总结
                    if self.url_sum_group:  #group_sharing = True进行总结，False则忽略。
                        logger.info("[sum4all] Summary URL : %s", content)
                        # 更新params_cache中的last_url
                        self.params_cache[user_id] = {}
                        self.params_cache[user_id]['last_url'] = content
                        logger.info('Updated last_url in params_cache for user.')
                        self.call_service(content, e_context, "sum")
                        return
                    else:
                        return
                else:  #处理私聊总结
                    logger.info("[sum4all] Summary URL : %s", content)
                    # 更新params_cache中的last_url
                    self.params_cache[user_id] = {}
                    self.params_cache[user_id]['last_url'] = content
                    logger.info('Updated last_url in params_cache for user.')
                    self.call_service(content, e_context, "sum")
                    return
        
        elif url_match and self.url_sum_enabled: #匹配URL链接
            if unsupported_urls:  #匹配不支持总结的网址
                logger.info("[sum4all] Unsupported URL : %s", content)
                reply = Reply(type=ReplyType.TEXT, content="不支持总结小程序和视频号")
                e_context["reply"] = reply
                e_context.action = EventAction.BREAK_PASS
            else:
                logger.info("[sum4all] Summary URL : %s", content)
                # 更新params_cache中的last_url
                self.params_cache[user_id] = {}
                self.params_cache[user_id]['last_url'] = content
                logger.info('Updated last_url in params_cache for user.')
                self.call_service(content, e_context, "sum")
                return
        elif context.type == ContextType.TEXT:
            bot = bridge.Bridge().find_chat_bot(const.QWEN_DASHSCOPE)
            bot.sessions.session_args["system_prompt"] = conf().get("character_desc", "")
            reply = bot.reply(content, context)
            e_context["reply"] = reply
            e_context.action = EventAction.BREAK_PASS
            return
    
    def call_service(self, content, e_context, service_type):
        if service_type == "search":
            if self.search_sum_service == "openai" or self.search_sum_service == "sum4all" or self.search_sum_service == "gemini" or self.search_sum_service == "azure":
                self.handle_search(content, e_context)
            elif self.search_sum_service == "perplexity":
                self.handle_perplexity(content, e_context)
        elif service_type == "sum":
            if self.url_sum_service == "bibigpt":
                self.handle_bibigpt(content, e_context)
            elif self.url_sum_service == "openai" or self.url_sum_service == "sum4all" or self.url_sum_service == "gemini" or self.search_sum_service == "azure":
                self.handle_url(content, e_context)
            elif self.url_sum_service == "opensum":
                self.handle_opensum(content, e_context)
        elif service_type == "note":
            if self.note_service == "flomo":
                self.handle_note(content, e_context)
            elif self.note_service == "notion":
                self.handle_notion(content, e_context)

    def handle_notion(self, link, e_context):
        user_id = e_context["context"]["msg"].from_user_id
        content = self.params_cache[user_id].get('content', '')
        note = self.params_cache[user_id].get('note', '')
        reply = Reply()
        reply.type = ReplyType.TEXT
        try:
            notion_url = create_notion_page(link, content, note)
            reply.content = f"已保存到Notion: {notion_url}"        
        except Exception as e:
            logger.error(f"create notion page failed {e}")
            reply.content = f"发送失败，错误信息：{e}"
        e_context["reply"] = reply
        e_context.action = EventAction.BREAK_PASS
        
    def handle_note(self,link,e_context):
        msg: ChatMessage = e_context["context"]["msg"]
        user_id = msg.from_user_id
        title = self.params_cache[user_id].get('title', '')
        content = self.params_cache[user_id].get('content', '')
        note = self.params_cache[user_id].get('note', '')
        # 将这些内容按照一定的格式整合到一起
        note_content = f"#sum4all\n{title}\n📒笔记：{note}\n{content}\n{link}"
        payload = {"content": note_content}
        # 将这个字典转换为JSON格式
        payload_json = json.dumps(payload)
        # 创建一个POST请求
        url = self.flomo_key
        headers = {'Content-Type': 'application/json'}
        # 发送这个POST请求
        response = requests.post(url, headers=headers, data=payload_json)
        reply = Reply()
        reply.type = ReplyType.TEXT
        if response.status_code == 200 and response.json()['code'] == 0:
            reply.content = f"已发送到{self.note_service}"        
        else:
            reply.content = "发送失败，错误码：" + str(response.status_code)
        e_context["reply"] = reply
        e_context.action = EventAction.BREAK_PASS   
    def short_url(self, long_url):
        url = "https://short.fatwang2.com"
        payload = {
            "url": long_url
        }        
        headers = {'Content-Type': "application/json"}
        response = requests.request("POST", url, json=payload, headers=headers)
        if response.status_code == 200:
            res_data = response.json()
            # 直接从返回的 JSON 中获取短链接
            short_url = res_data.get('shorturl', None)  
            
            if short_url:
                return short_url
        return None
    def handle_video(self, content, e_context):
        video_text = url_to_text(content)
        context = e_context["context"]
        # video_text = "普通的人没有什么资源，怎么找到这些好的项目呢？你多大？三十多？你现在有钱吗？比普通上班族好一些。那我想问你一个问题啊，你能不能告诉我你是怎么错过自媒体，错过电商，错过错过比特币的？你把这个问题悟透了，你就知道为什么你不行了。我觉得你们其实也可以悟一悟，都说我们这个时代没有机会。其实我们20年的经验过来，我跟你讲我们这个机会不是多是太多了。从房产到实体，20年前实体的大机会，到后面电商崛起，比特币啊，包括现在的短视频自媒体，在座的各位谁采用了？事实上你有没有发现，还有这些行业在开始的时候都是不投钱的，真的就是拿个手机就能挣钱。那你能告诉我你是怎么过的吗？你能把这个要想明白了，你就知道下一步怎么迎接了。那个时候意识不到这些东西，你为什么意识不到这些东西？理解不够？我告诉你不是啊，我给你个正确答案啊。因为你看事情永远要从三个角度来看。那这个时候刚做起来以后，那些人都开始了劳斯莱斯，我们的社会就在评价什么就傻逼了。因为第一呢我们一定要明白的道理，一件事情一旦有人做成了，他就是真的。美国有句谚语啊，如果一个生物走路像压，叫声像鸭子，它就是一只鸭子，你要信吗？明白一个点啊，一个事儿如果别人做成了，你不理解，通常是自己傻。成功即代表一切。然后第二呢我们要去分析他投多少钱，挣多少钱，他到底是不是盈利的。你不要想第三个东西，我给你一盘屎你吃不吃，那肯定不吃啊。给你一个亿是不是？你不能用赚钱以外的逻辑来看待这个行业。因为绝大部分行业在开始的时候真的很low。我之前做电商很早就能挣钱。07年08年的时候我在这个地方到10年左右，我拉我身边的人干很多都不干。你觉得不务正业。电商也是一样，比特币也是一样。这个东西你要记住了，挣钱就挣钱，你不能去因为他看起来不好觉的钱你知道吗？很多人是因为看钱以外的逻辑去看这件事，所以导致的错之商机。很多钱真的是脏钱搂钱。那么你觉得做网红抛头露面他就一定是好事情？他也不是，到现在也一直有人说，哎，那么大的老板，你看抖音不觉得丢人吗？一个大他妈吐死的干什么不丢人？你卖货就高级了，这个袜子就高级了？不会的。第三要看一个东西叫什么？叫操作难度，这个能不能干？比特币早些年入场的时候，你下场研究了吗？你买个比特币你会花多长时间？早几年去做抖音，拿个手机随便拍，随便报，这三个维度交织思考问题，你才不会错过商机。给你磨一下现实厂家，你周围一个傻子突然间看奥迪了，你就问问他兄弟你干嘛的？他跟你讲要给我收废品的，你不能说我这么姐，你得去了解你那个点儿在哪呀，你投了多少钱呀，你都怎么收的，你这样才有机会上车，明白我意思吧？绝大部分人他错过商机的本质什么？因为他的思维系统太过于复杂。人啊真正应该搭建的是什么？思考问题的方式，这个会很高级。森哥如果你现在是一个普通人的话，你会怎么去改变自己的这种现状啊？不好意思，现在不普通了，我的整个生涯也真的是运气占了绝大部分，真的没有你们想象的特别复杂。包括你说所谓的大老板，我认识的太多了，要么投胎好，要么婚姻好。因为现在我就碰到了一个行业，我就觉得这个行业不需要我什么能力，但是我确实是比以前赚的多了。意识到严哥你以前说的那个项目，真的是决定你的一些收入。所以说我现在就最近在思考，看不到这些好的项目。哎呀，那些事其实就一句话，挣钱了以后去报商学院，有钱报长江，没钱报我的看看世界，看看其他朋友在干什么。说信息圈终结了还能怎么看？真的以为老板到长江是为了去学什么东西了，都是交朋友，看看同等到的是干嘛的，一群没有钱的总会说一句话，有钱人被淘汰，我就想不明白，白一个职场混工资的人，你你怀怀疑老板干嘛的？是谁智商不行了？信息的价值往往是比工作要值钱100倍的，你见过哪个老板挣钱是闷在家里的，对不对？各种商学院去报一报，你报一圈，大家都是交8万的，都是交50万、60万的，他圈子质量是一样的。我有的时候我也看那些骂我的人什么样，我说你罗斯就就是这种小老板，要么就职场大学生，没办法，无解。你跟他讲信息值钱，交个朋友他们会下面。哎，有钱人都是傻子，他们花8万块钱交朋友被割了，这不是事实吗？为什么有这种想法？因为没搞到钱了。在你们的认知里面，一年搞个100万、200万就很好了。在有些圈子里面这个收入是不好意思说的。深耕都面临。"
        video_title = e_context["context"]["msg"].title

        msg: ChatMessage = context["msg"]
        # 刷新context type to text
        context["type"] = ContextType.TEXT
        user_id = msg.from_user_id
        user_params = self.params_cache.get(user_id, {})
        system_prompt = user_params.get('prompt', self.url_sum_video_prompt)
        logger.info(f"system_prompt: {system_prompt}")
        user_prompt = f"<title>小视频标题</title>\n\n{video_title}\n\n<transcript>小视频文案</transcript>\n\n{video_text}"
        logger.info(f"user_prompt: {user_prompt}")
        bot = bridge.Bridge().find_chat_bot(const.QWEN_DASHSCOPE)
        bot.sessions.session_args["system_prompt"] = system_prompt
        reply = bot.reply(user_prompt, context)

        self.params_cache[user_id]['content'] = reply.content
        self.params_cache[user_id]['title'] = video_title

        e_context["reply"] = reply
        e_context.action = EventAction.BREAK_PASS
        return
    
    def handle_url(self, content, e_context):
        logger.info('Handling Sum4All request...')
        # 根据sum_service的值选择API密钥和基础URL
        if self.url_sum_service == "openai":
            api_key = self.open_ai_api_key
            api_base = self.open_ai_api_base
            model = self.model
        elif self.url_sum_service == "sum4all":
            api_key = self.sum4all_key
            api_base = "https://pro.sum4all.site/v1"
            model = "sum4all"
        elif self.url_sum_service == "gemini":
            api_key = self.gemini_key
            model = "gemini"
            api_base = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent"
        else:
            logger.error(f"未知的sum_service配置: {self.url_sum_service}")
            return
        
        msg: ChatMessage = e_context["context"]["msg"]
        user_id = msg.from_user_id
        user_params = self.params_cache.get(user_id, {})
        isgroup = e_context["context"].get("isgroup", False)
        prompt = user_params.get('prompt', self.url_sum_prompt)
        headers = {
            'Content-Type': 'application/json',
            'Authorization': f'Bearer {api_key}'
        }
        payload = json.dumps({
            "link": content,
            "prompt": prompt,
            "model": model,
            "base": api_base
        })
        additional_content = ""  # 在 try 块之前初始化 additional_content

        try:
            logger.info('Sending request to LLM...')
            api_url = "https://ai.sum4all.site"
            response = requests.post(api_url, headers=headers, data=payload)
            response.raise_for_status()
            response_data = response.json()  # 解析响应的 JSON 数据
            logger.info('Received response from LLM. %s', response_data)
            if response_data.get("success"):
                content = response_data["content"].replace("\\n", "\n")  # 替换 \\n 为 \n
                self.params_cache[user_id]['content'] = content

                # 新增加的部分，用于解析 meta 数据
                meta = response_data.get("meta", {})  # 如果没有 meta 数据，则默认为空字典
                title = meta.get("og:title", "")  # 获取 og:title，如果没有则默认为空字符串
                self.params_cache[user_id]['title'] = title
                # 只有当 title 非空时，才加入到回复中
                if title:
                    additional_content += f"{title}\n\n"
                reply_content = additional_content + content  # 将内容加入回复
                
            else:
                reply_content = "Content not found or error in response"

        except requests.exceptions.RequestException as e:
            # 处理可能出现的错误
            logger.error(f"Error calling new combined api: {e}")
            reply_content = f"An error occurred"

        reply = Reply()
        reply.type = ReplyType.TEXT
        if not self.url_sum_qa_enabled:
            reply.content = remove_markdown(reply_content)
        elif isgroup or not self.note_enabled:
            reply.content = f"{remove_markdown(reply_content)}\n\n💬5min内输入{self.url_sum_qa_prefix}+问题，可继续追问"
        elif self.note_enabled:
            reply.content = f"{remove_markdown(reply_content)}\n\n💬5min内输入{self.url_sum_qa_prefix}+问题，可继续追问。\n\n📒输入{self.note_prefix}+笔记，可发送当前总结&笔记到{self.note_service}"
        e_context["reply"] = reply
        e_context.action = EventAction.BREAK_PASS
    def handle_bibigpt(self, content, e_context):    
        headers = {
            'Content-Type': 'application/json'
        }
        payload_params = {
            "url": content,
            "includeDetail": False,
            "promptConfig": {
                "outputLanguage": self.outputLanguage
            }
        }

        payload = json.dumps(payload_params)           
        try:
            api_url = f"https://bibigpt.co/api/open/{self.bibigpt_key}"
            response = requests.request("POST",api_url, headers=headers, data=payload)
            response.raise_for_status()
            data = json.loads(response.text)
            summary_original = data.get('summary', 'Summary not available')
            html_url = data.get('htmlUrl', 'HTML URL not available')
            # 获取短链接
            short_url = self.short_url(html_url) 
            
            # 如果获取短链接失败，使用 html_url
            if short_url is None:
                short_url = html_url if html_url != 'HTML URL not available' else 'URL not available'
            
            # 移除 "##摘要"、"## 亮点" 和 "-"
            summary = summary_original.split("详细版（支持对话追问）")[0].replace("## 摘要\n", "📌总结：").replace("## 亮点\n", "").replace("- ", "")
        except requests.exceptions.RequestException as e:
            reply = f"An error occurred"

        reply = Reply()
        reply.type = ReplyType.TEXT
        reply.content = f"{summary}\n\n详细链接：{short_url}"

        e_context["reply"] = reply
        e_context.action = EventAction.BREAK_PASS
    def handle_opensum(self, content, e_context):
        headers = {
            'Content-Type': 'application/json',
            'Authorization': f'Bearer {self.opensum_key}'
        }
        payload = json.dumps({"link": content})
        try:
            api_url = "https://read.thinkwx.com/api/v1/article/summary"
            response = requests.request("POST",api_url, headers=headers, data=payload)
            response.raise_for_status()
            data = json.loads(response.text)
            summary_data = data.get('data', {})  # 获取data字段                
            summary_original = summary_data.get('summary', 'Summary not available')
            # 使用正则表达式提取URL
            url_pattern = r'https:\/\/[^\s]+'
            match = re.search(url_pattern, summary_original)
            html_url = match.group(0) if match else 'HTML URL not available'            
            # 获取短链接
            short_url = self.short_url(html_url) if match else html_url
            # 用于移除摘要中的URL及其后的所有内容
            url_pattern_remove = r'https:\/\/[^\s]+[\s\S]*'
            summary = re.sub(url_pattern_remove, '', summary_original).strip()        

        except requests.exceptions.RequestException as e:
            summary = f"An error occurred"
            short_url = 'URL not available'
        
        reply = Reply()
        reply.type = ReplyType.TEXT
        reply.content = f"{summary}\n\n详细链接：{short_url}"

        e_context["reply"] = reply
        e_context.action = EventAction.BREAK_PASS    
    def handle_search(self, content, e_context):
        # 根据sum_service的值选择API密钥和基础URL
        if self.search_sum_service == "openai":
            api_key = self.open_ai_api_key
            api_base = self.open_ai_api_base
            model = self.model
        elif self.search_sum_service == "sum4all":
            api_key = self.sum4all_key
            api_base = "https://pro.sum4all.site/v1"
            model = "sum4all"
        elif self.search_sum_service == "gemini":
            api_key = self.gemini_key
            model = "gemini"
            api_base = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent"

        else:
            logger.error(f"未知的search_service配置: {self.search_sum_service}")
            return
        headers = {
            'Content-Type': 'application/json',
            'Authorization': f'Bearer {api_key}'
        }
        content = content[len(self.search_sum_search_prefix):]
        payload = json.dumps({
            "ur": content,
            "prompt": self.search_sum_prompt,
            "model": model,
            "base": api_base,
            "search1api_key": self.search1api_key,
            "search_service": self.search_service  
        })
        try:
            api_url = "https://ai.sum4all.site"
            response = requests.post(api_url, headers=headers, data=payload)
            response.raise_for_status()
            response_data = response.json()  # 解析响应的 JSON 数据
            if response_data.get("success"):
                content = response_data["content"].replace("\\n", "\n")  # 替换 \\n 为 \n
                reply_content = content  # 将内容加入回复

                # 解析 meta 数据
                meta = response_data.get("meta", {})  # 如果没有 meta 数据，则默认为空字典
                title = meta.get("og:title", "")  # 获取 og:title，如果没有则默认为空字符串
                og_url = meta.get("og:url", "")  # 获取 og:url，如果没有则默认为空字符串
                # 打印 title 和 og_url 以调试
                print("Title:", title)
                print("Original URL:", og_url)                
                # 只有当 title 和 url 非空时，才加入到回复中
                if title:
                    reply_content += f"\n\n参考文章：{title}"
                if og_url:
                    short_url = self.short_url(og_url)  # 获取短链接
                    reply_content += f"\n\n参考链接：{short_url}"                

            else:
                content = "Content not found or error in response"

        except requests.exceptions.RequestException as e:
            # 处理可能出现的错误
            logger.error(f"Error calling new combined api: {e}")
            reply_content = f"An error occurred"

        reply = Reply()
        reply.type = ReplyType.TEXT
        reply.content = f"{remove_markdown(reply_content)}"            
        e_context["reply"] = reply
        e_context.action = EventAction.BREAK_PASS
    def handle_perplexity(self, content, e_context):

        headers = {
            'Content-Type': 'application/json',
            'Authorization': f'Bearer {self.perplexity_key}'
        }
        data = {
            "model": "sonar-small-online",
            "messages": [
                {"role": "system", "content": self.search_sum_prompt},
                {"role": "user", "content": content}
        ]
        }
        try:
            api_url = "https://api.perplexity.ai/chat/completions"
            response = requests.post(api_url, headers=headers, json=data)
            response.raise_for_status()
            # 处理响应数据
            response_data = response.json()
            # 这里可以根据你的需要处理响应数据
            # 解析 JSON 并获取 content
            if "choices" in response_data and len(response_data["choices"]) > 0:
                first_choice = response_data["choices"][0]
                if "message" in first_choice and "content" in first_choice["message"]:
                    content = first_choice["message"]["content"]
                else:
                    print("Content not found in the response")
            else:
                print("No choices available in the response")
        except requests.exceptions.RequestException as e:
            # 处理可能出现的错误
            logger.error(f"Error calling perplexity: {e}")
        reply = Reply()
        reply.type = ReplyType.TEXT
        reply.content = f"{remove_markdown(content)}"            
        e_context["reply"] = reply
        e_context.action = EventAction.BREAK_PASS
    def get_help_text(self, verbose=False, **kwargs):
        help_text = "Help you summarize all things\n"
        if not verbose:
            return help_text
        help_text += "1.Share me the link and I will summarize it for you\n"
        help_text += f"2.{self.search_sum_search_prefix}+query,I will search online for you\n"
        return help_text
    def handle_file(self, content, e_context):
        logger.info("handle_file: 向LLM发送内容总结请求")
        # 根据sum_service的值选择API密钥和基础URL
        if self.file_sum_service == "openai":
            api_key = self.open_ai_api_key
            api_base = self.open_ai_api_base
            model = self.model
        elif self.file_sum_service == "azure":
            api_key = self.open_ai_api_key
            api_base = f"{self.open_ai_api_base}/openai/deployments/{self.azure_deployment_id}/chat/completions?api-version=2024-02-15-preview"
            model = self.model
        elif self.file_sum_service == "sum4all":
            api_key = self.sum4all_key
            api_base = "https://pro.sum4all.site/v1"
            model = "sum4all"
        elif self.file_sum_service == "gemini":
            api_key = self.gemini_key
            model = "gemini"
            api_base = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent"
        else:
            logger.error(f"未知的sum_service配置: {self.file_sum_service}")
            return
        msg: ChatMessage = e_context["context"]["msg"]
        user_id = msg.from_user_id
        user_params = self.params_cache.get(user_id, {})
        prompt = user_params.get('prompt', self.file_sum_prompt)
        if model == "gemini":
            headers = {
                'Content-Type': 'application/json',
                'x-goog-api-key': api_key
            }
            data = {
            "contents": [
                {"role": "user", "parts": [{"text": prompt}]},
                {"role": "model", "parts": [{"text": "okay"}]},
                {"role": "user", "parts": [{"text": content}]}
            ],
            "generationConfig": {
                "maxOutputTokens": 800
            }
            }
            api_url = api_base
        elif self.file_sum_service == "azure":
            headers = {
                "Content-Type": "application/json",
                "api-key": api_key
            }
            data = {
                "model": model,
                "messages": [
                    {"role": "system", "content": prompt},
                    {"role": "user", "content": content}
                ]
            }
            api_url = api_base
        else:
            headers = {
                'Content-Type': 'application/json',
                'Authorization': f'Bearer {api_key}'
            }
            data = {
                "model": model,
                "messages": [
                    {"role": "system", "content": prompt},
                    {"role": "user", "content": content}
                ]
            }
            api_url = f"{api_base}/chat/completions"
        try:
            response = requests.post(api_url, headers=headers, data=json.dumps(data))
            response.raise_for_status()
            response_data = response.json()
            
            # 解析 JSON 并获取 content
            if model == "gemini":
                if "candidates" in response_data and len(response_data["candidates"]) > 0:
                    first_candidate = response_data["candidates"][0]
                    if "content" in first_candidate:
                        if "parts" in first_candidate["content"] and len(first_candidate["content"]["parts"]) > 0:
                            response_content = first_candidate["content"]["parts"][0]["text"].strip()  # 获取响应内容
                            logger.info(f"Gemini API response content: {response_content}")  # 记录响应内容
                            reply_content = response_content.replace("\\n", "\n")  # 替换 \\n 为 \n
                        else:
                            logger.error("Parts not found in the Gemini API response content")
                            reply_content = "Parts not found in the Gemini API response content"
                    else:
                        logger.error("Content not found in the Gemini API response candidate")
                        reply_content = "Content not found in the Gemini API response candidate"
                else:
                    logger.error("No candidates available in the Gemini API response")
                    reply_content = "No candidates available in the Gemini API response"        
            else:
                if "choices" in response_data and len(response_data["choices"]) > 0:
                    first_choice = response_data["choices"][0]
                    if "message" in first_choice and "content" in first_choice["message"]:
                        response_content = first_choice["message"]["content"].strip()  # 获取响应内容
                        logger.info(f"LLM API response content")  # 记录响应内容
                        reply_content = response_content.replace("\\n", "\n")  # 替换 \\n 为 \n
                    else:
                        logger.error("Content not found in the response")
                        reply_content = "Content not found in the LLM API response"
                else:
                    logger.error("No choices available in the response")
                    reply_content = "No choices available in the LLM API response"

        except requests.exceptions.RequestException as e:
            logger.error(f"Error calling LLM API: {e}")
            reply_content = f"An error occurred while calling LLM API"

        reply = Reply()
        reply.type = ReplyType.TEXT
        reply.content = f"{remove_markdown(reply_content)}\n\n💬5min内输入{self.file_sum_qa_prefix}+问题，可继续追问" 
        e_context["reply"] = reply
        e_context.action = EventAction.BREAK_PASS
    def read_pdf(self, file_path):
        logger.info(f"开始读取PDF文件：{file_path}")
        doc = fitz.open(file_path)
        content = ' '.join([page.get_text() for page in doc])
        logger.info(f"PDF文件读取完成：{file_path}")

        return content
    def read_word(self, file_path):
        doc = Document(file_path)
        return ' '.join([p.text for p in doc.paragraphs])
    def read_markdown(self, file_path):
        with open(file_path, 'r', encoding='utf-8') as file:
            md_content = file.read()
            return markdown.markdown(md_content)
    def read_excel(self, file_path):
        workbook = load_workbook(file_path)
        content = ''
        for sheet in workbook:
            for row in sheet.iter_rows():
                content += ' '.join([str(cell.value) for cell in row])
                content += '\n'
        return content
    def read_txt(self, file_path):
        logger.debug(f"开始读取TXT文件: {file_path}")
        try:
            with open(file_path, 'r', encoding='utf-8') as file:
                content = file.read()
            logger.debug(f"TXT文件读取完成: {file_path}")
            logger.debug("TXT文件内容的前50个字符：")
            logger.debug(content[:50])  # 打印文件内容的前50个字符
            return content
        except Exception as e:
            logger.error(f"读取TXT文件时出错: {file_path}，错误信息: {str(e)}")
            return ""
    def read_csv(self, file_path):
        content = ''
        with open(file_path, 'r', encoding='utf-8') as csvfile:
            reader = csv.reader(csvfile)
            for row in reader:
                content += ' '.join(row) + '\n'
        return content
    def read_html(self, file_path):
        with open(file_path, 'r', encoding='utf-8') as file:
            soup = BeautifulSoup(file, 'html.parser')
            return soup.get_text()
    def read_ppt(self, file_path):
        presentation = Presentation(file_path)
        content = ''
        for slide in presentation.slides:
            for shape in slide.shapes:
                if hasattr(shape, "text"):
                    content += shape.text + '\n'
        return content
    def extract_content(self, file_path):
        logger.info(f"extract_content: 提取文件内容，文件路径: {file_path}")
        file_size = os.path.getsize(file_path) // 1000  # 将文件大小转换为KB
        if file_size > int(self.max_file_size):
            logger.warning(f"文件大小超过限制({self.max_file_size}KB),不进行处理。文件大小: {file_size}KB")
            return None
        file_extension = os.path.splitext(file_path)[1][1:].lower()
        logger.info(f"extract_content: 文件类型为 {file_extension}")

        file_type = EXTENSION_TO_TYPE.get(file_extension)

        if not file_type:
            logger.error(f"不支持的文件扩展名: {file_extension}")
            return None

        read_func = {
            'pdf': self.read_pdf,
            'docx': self.read_word,
            'md': self.read_markdown,
            'txt': self.read_txt,
            'excel': self.read_excel,
            'csv': self.read_csv,
            'html': self.read_html,
            'ppt': self.read_ppt
        }.get(file_type)

        if not read_func:
            logger.error(f"不支持的文件类型: {file_type}")
            return None
        logger.info("extract_content: 文件内容提取完成")
        return read_func(file_path)
    def encode_image_to_base64(self, image_path):
        # 打开图片
        img = Image.open(image_path)
        # 只有当图片的宽度大于1024像素时，才调整图片大小
        if img.width > 1024:
            img = img.resize((1024, int(img.height*1024/img.width)))
            # 将调整大小后的图片保存回原文件
            img.save(image_path)

        # 打开调整大小后的图片，读取并进行base64编码
        with open(image_path, "rb") as image_file:
            encoded = base64.b64encode(image_file.read()).decode('utf-8')
        return encoded
    def handle_image(self, base64_image, e_context):
        logger.info("handle_image: 解析图像处理API的响应")
        msg: ChatMessage = e_context["context"]["msg"]
        user_id = msg.from_user_id
        user_params = self.params_cache.get(user_id, {})
        prompt = user_params.get('prompt', self.image_sum_prompt)

        if self.image_sum_service == "openai":
            api_key = self.open_ai_api_key
            api_base = f"{self.open_ai_api_base}/chat/completions"
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}"
            }
            model = "qwen-vl-plus"
        elif self.image_sum_service == "azure":
            api_key = self.open_ai_api_key
            api_base = f"{self.open_ai_api_base}/openai/deployments/{self.azure_deployment_id}/chat/completions?api-version=2024-02-15-preview"
            headers = {
                "Content-Type": "application/json",
                "api-key": api_key
            }
            model = "gpt-4o-mini"
        elif self.image_sum_service == "xunfei":
            api_key = self.xunfei_api_key
            api_base = "https://spark.sum4all.site/v1/chat/completions"
            model = "spark-chat-vision"
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}"
            }
        elif self.image_sum_service == "sum4all":
            api_key = self.sum4all_key
            api_base = "https://pro.sum4all.site/v1/chat/completions"
            model = "sum4all-vision"
            headers = {
                "Content-Type": "application/json",
                "Authorization": f"Bearer {api_key}"
            }
        elif self.image_sum_service == "gemini":
            api_key = self.gemini_key
            api_base = "https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent"
            payload = {
                "contents": [
                    {
                        "parts": [
                            {"text": prompt},
                            {
                                "inline_data": {
                                    "mime_type":"image/png",
                                    "data": base64_image
                                }
                            }
                        ]
                    }
                ],
                "generationConfig": {
                    "maxOutputTokens": 800
                }
            }
            headers = {
                "Content-Type": "application/json",
                "x-goog-api-key": api_key
            }

        else:
            logger.error(f"未知的image_sum_service配置: {self.image_sum_service}")
            return

        if self.image_sum_service != "gemini":
            payload = {
                "model": model,
                "messages": [
                    {
                        "role": "user",
                        "content": [
                            {
                                "type": "text",
                                "text": prompt
                            },
                            {
                                "type": "image_url",
                                "image_url": {
                                    "url": base64_image
                                }
                            }
                        ]
                    }
                ],
                "max_tokens": 2000
            }
        try:
            response = requests.post(api_base, headers=headers, json=payload)
            response.raise_for_status()
            response_json = response.json()

            if self.image_sum_service == "gemini":
                reply_content = response_json.get('candidates', [{}])[0].get('content', {}).get('parts', [{}])[0].get('text', 'No text found in the response')
            else:
                if "choices" in response_json and len(response_json["choices"]) > 0:
                    first_choice = response_json["choices"][0]
                    if "message" in first_choice and "content" in first_choice["message"]:
                        response_content = first_choice["message"]["content"].strip()
                        logger.info("LLM API response content")
                        reply_content = response_content
                    else:
                        logger.error("Content not found in the response")
                        reply_content = "Content not found in the LLM API response"
                else:
                    logger.error("No choices available in the response")
                    reply_content = "No choices available in the LLM API response"
        except Exception as e:
            logger.error(f"Error processing LLM API response: {e}")
            reply_content = f"An error occurred while processing LLM API response"

        reply = Reply()
        reply.type = ReplyType.TEXT
        reply.content = f"{remove_markdown(reply_content)}\n\n💬5min内输入{self.image_sum_qa_prefix}+问题，可继续追问"
        remove_audio_from_oss(e_context["context"].content)
        e_context["reply"] = reply
        e_context.action = EventAction.BREAK_PASS
    
def remove_markdown(text):
    # 替换Markdown的粗体标记
    text = text.replace("**", "")
    # 替换Markdown的标题标记
    text = text.replace("### ", "").replace("## ", "").replace("# ", "")
    return text