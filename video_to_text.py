import dashscope
import json
import requests
from tqdm import tqdm
import os
import oss2
import time
from moviepy import VideoFileClip
from urllib import request
from oss2.credentials import EnvironmentVariableCredentialsProvider


curdir = os.path.dirname(__file__)
config_path = os.path.join(curdir, "config.json")
if os.path.exists(config_path):
    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)
        OSS_ACCESS_KEY_SECRET = config.get("keys", {}).get("OSS_ACCESS_KEY_SECRET")
        OSS_ACCESS_KEY_ID = config.get("keys", {}).get("OSS_ACCESS_KEY_ID")
        open_ai_api_key = config.get("keys", {}).get("open_ai_api_key")
        assert OSS_ACCESS_KEY_ID, "miss OSS_ACCESS_KEY_ID"
        assert OSS_ACCESS_KEY_SECRET, "miss OSS_ACCESS_KEY_SECRET"
        assert open_ai_api_key, "miss open_ai_api_key"
        os.environ['OSS_ACCESS_KEY_ID'] = OSS_ACCESS_KEY_ID
        os.environ['OSS_ACCESS_KEY_SECRET'] = OSS_ACCESS_KEY_SECRET


def download_large_file(url, output_path, chunk_size=50*1024*1024):  # 每次下载50MB
    try:
        # 获取文件总大小
        response = requests.head(url)
        file_size = int(response.headers.get('content-length', 0))
        
        # 如果文件已存在，获取已下载的大小
        if os.path.exists(output_path):
            first_byte = os.path.getsize(output_path)
        else:
            first_byte = 0
        
        # 显示总进度条
        progress = tqdm(total=file_size, initial=first_byte,
                       unit='iB', unit_scale=True, desc='下载进度')
        
        # 分片下载
        while first_byte < file_size:
            last_byte = min(first_byte + chunk_size - 1, file_size - 1)
            
            # 设置请求头，指定下载范围
            headers = {'Range': f'bytes={first_byte}-{last_byte}'}
            
            # 下载当前分片
            response = requests.get(url, headers=headers, stream=True, timeout=300)
            
            # 追加模式写入文件
            mode = 'ab' if first_byte else 'wb'
            with open(output_path, mode) as f:
                for chunk in response.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
                        progress.update(len(chunk))
            
            # 更新下载位置
            first_byte = last_byte + 1
            
        progress.close()
        
    except Exception as e:
        print(f"下载出错: {str(e)}")
        raise e


def video_to_audio(video_path, output_audio_path):
    """
    Convert a video to audio and save it to the output path.

    Parameters:
    video_path (str): The path to the video file.
    output_audio_path (str): The path to save the audio to.

    """
    try:
        clip = VideoFileClip(video_path)
        audio = clip.audio
        ffmpeg_params = [
        '-ac', '1',           # Set to mono (1 audio channel)
        '-ar', '16000',       # Set sampling rate to 16kHz
        '-acodec', 'libopus'  # Use opus codec
      ]
        audio.write_audiofile(output_audio_path, fps=16000, codec='libopus', ffmpeg_params=ffmpeg_params)
        audio.close()
        clip.close()
        return True
    except Exception as e:
        print(f"视频转音频出错: {str(e)}")
        raise e
    

def upload_audio_to_oss(audio_path):
    """
    上传音频到oss
    """
    auth = oss2.ProviderAuthV4(EnvironmentVariableCredentialsProvider())
    endpoint = "https://oss-cn-beijing.aliyuncs.com"
    region = "cn-beijing"
    bucket_name = "cow-video"
    bucket = oss2.Bucket(auth, endpoint, bucket_name, region=region)
    object_name = f"audio/{audio_path}"
    url = bucket.sign_url('PUT', object_name, 60*10, slash_safe=True)
    returned_url = bucket.sign_url('GET', object_name, 60*10, slash_safe=True)
    try:
        with open(audio_path, 'rb') as f:
            result = requests.put(url, data=f, verify=False)
        print(f"File uploaded successfully, status code: {result.text}")
        return returned_url
    except oss2.exceptions.OssError as e:
        print(f"Failed to upload file: {e}")
        return None

def remove_audio_from_oss(audio_path):
    auth = oss2.ProviderAuthV4(EnvironmentVariableCredentialsProvider())
    endpoint = "https://oss-cn-beijing.aliyuncs.com"
    region = "cn-beijing"
    bucket_name = "cow-video"
    bucket = oss2.Bucket(auth, endpoint, bucket_name, region=region)
    object_name = f"audio/{audio_path}"
    bucket.delete_object(object_name)


def dashscope_audio_to_text(url):
    """
    使用dashscope进行音频转写
    """
    dashscope.api_key=open_ai_api_key
    task_response=dashscope.audio.asr.Transcription.async_call(
        model='paraformer-v1',
        file_urls=[url]
        )

    transcription_response=dashscope.audio.asr.Transcription.wait(task=task_response.output.task_id)
    print(transcription_response)
    transcription_url=transcription_response.output['results'][0]['transcription_url']
    transcription_results=json.loads(request.urlopen(transcription_url).read().decode('utf8'))
    return transcription_results.get("transcripts")[0].get("text")

def url_to_text(url):
  video_path = f"video_{str(int(time.time()))}.mp4"
  audio_path = f"audio_{str(int(time.time()))}.opus"
  try:
    download_large_file(url, video_path)
    video_to_audio(video_path, audio_path)
    url = upload_audio_to_oss(audio_path)
    if url:
      text = dashscope_audio_to_text(url)
      return text
  except Exception as e:
    print(f"下载视频出错: {str(e)}")
    return None
  finally:
    if os.path.exists(video_path):
        os.remove(video_path)
    if os.path.exists(audio_path):
        os.remove(audio_path)
    remove_audio_from_oss(audio_path)

