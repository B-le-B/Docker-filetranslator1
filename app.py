# app.py
# === 增强版环境变量加载和追踪 ===
# import os

# print("=== 开始加载环境变量 ===")

# 1. 显示加载前状态
# print("加载前环境变量状态:")
# test_vars = ['SILICONFLOW_API_KEY', 'VOLCENGINE_ACCESS_KEY']
# for var in test_vars:
#     value = os.getenv(var)
#     if value:
#         print(f"  {var}: {value[:8]}...")
#     else:
#         print(f"  {var}: 未设置")

# 2. 执行加载
# try:
#     from dotenv import load_dotenv
#     print("正在执行 load_dotenv()...")
#     result = load_dotenv(override=True)
#     print(f"load_dotenv() 返回值: {result}")

# 3. 验证加载结果
# print("加载后环境变量状态:")
# for var in test_vars:
#     value = os.getenv(var)
#     if value:
#         print(f"  ✓ {var}: {value[:8]}...")
#     else:
#         print(f"  ✗ {var}: 仍未设置")

# print("✓ .env 文件加载完成")

# except ImportError:
#     print("⚠️ python-dotenv 未安装，请运行: pip install python-dotenv")
# except Exception as e:
#     print(f"⚠️ 加载 .env 文件失败: {e}")

# print("=== 环境变量加载完成 ===\n")

# # 4. 定期检查函数
# def check_env_status(stage_name):
#     print(f"\n=== 环境变量状态检查: {stage_name} ===")
#     for var in ['SILICONFLOW_API_KEY', 'VOLCENGINE_ACCESS_KEY']:
#         value = os.getenv(var)
#         if value:
#             print(f"  ✓ {var}: {value[:8]}...")
#         else:
#             print(f"  ✗ {var}: 未设置")

# # 在关键位置检查环境变量状态
# check_env_status("初始化后")

import os
import logging
import tempfile
from datetime import datetime, timedelta
from collections import defaultdict

from flask import (
    Flask,
    render_template,
    request,
    jsonify,
    send_from_directory,
    Response,
    session,
    redirect,
    url_for,
)
from werkzeug.utils import secure_filename
import uuid
import threading
import time
import webbrowser
import json
import requests
import shutil
import inspect

logging.basicConfig(
    level=logging.WARNING, format="%(asctime)s - %(levelname)s - %(message)s"
)
logger = logging.getLogger(__name__)


try:
    from translators import (
        SiliconFlowTranslator,
        translate_text_file,
        translate_docx_file,
        translate_docx_file_formatted,
        translate_docx_pythondoc1,
        translate_docx_pythondoc2,
        translate_docx_via_markdown,
        translate_pptx_file_formatted,
        translate_excel_file_formatted,
        VolcEngineNativeTranslator,
        _translate_docx_file_formatted_available,
        _translate_docx_pythondoc1_available,
        _translate_docx_pythondoc2_available,
        _translate_docx_markdown_available,
        _translate_pptx_available,
        _translate_excel_available,
        _volcengine_available,
    )

    # logger.info("All translators imported successfully")
except ImportError as e:
    logger.error(f"Failed to import translators: {e}")
    # 提供最小的回退
    SiliconFlowTranslator = None
    translate_text_file = None
    translate_docx_file = None
    translate_docx_file_formatted = None
    translate_docx_pythondoc1 = None
    translate_docx_pythondoc2 = None
    translate_docx_via_markdown = None
    translate_pptx_file_formatted = None
    translate_excel_file_formatted = None
    VolcEngineNativeTranslator = None
    _translate_docx_file_formatted_available = False
    _translate_docx_pythondoc1_available = False
    _translate_docx_pythondoc2_available = False
    _translate_docx_markdown_available = False
    _translate_pptx_available = False
    _translate_excel_available = False
    _volcengine_available = False


# 新增：导入PPT和Excel翻译器
# try:
#     from translators.pptx_full_translator import translate_pptx_file_formatted
#     _translate_pptx_available = True
#     logger.info("PPTX translator imported successfully")
# except ImportError as e:
#     logger.warning(f"PPTX translator not available: {e}")
#     _translate_pptx_available = False
#     translate_pptx_file_formatted = None

# try:
#     from translators.excel_full_translator import translate_excel_file_formatted
#     _translate_excel_available = True
#     logger.info("Excel translator imported successfully")
# except ImportError as e:
#     logger.warning(f"Excel translator not available: {e}")
#     _translate_excel_available = False
#     translate_excel_file_formatted = None

# else:
#     print("✗ VOLCENGINE_ACCESS_KEY 未找到")


# 新增：导入火山引擎翻译器
# try:
#     from translators.volcengine_native_translator import VolcEngineNativeTranslator

#     _volcengine_available = True
#     logger.info("VolcEngineNativeTranslator imported successfully")
# except ImportError as e:
#     logger.warning(f"VolcEngineNativeTranslator not available: {e}")
#     _volcengine_available = False
#     VolcEngineNativeTranslator = None

# 导入prompt相关模块
try:
    from prompts import get_all_templates, get_template, PromptBuilder, ConfigManager

    _prompt_modules_available = True
except ImportError as e:

    logging.warning(f"Prompt modules not found: {e}. Prompt features will be limited.")
    _prompt_modules_available = False

    # 提供空的fallback
    def get_all_templates():
        return {}

    def get_template(name):
        return None

    class PromptBuilder:
        pass

    class ConfigManager:
        def __init__(self, *args, **kwargs):
            pass


# Import fallback configurations from config.py
try:
    from config import (
        API_KEY as DEFAULT_FALLBACK_API_KEY,
        BASE_URL as DEFAULT_FALLBACK_BASE_URL,
        DEFAULT_MODEL as DEFAULT_FALLBACK_MODEL,
    )
except ImportError:
    logging.warning(
        "config.py not found or missing default fallback constants. Global fallbacks will be None."
    )
    DEFAULT_FALLBACK_API_KEY = None
    DEFAULT_FALLBACK_BASE_URL = None
    DEFAULT_FALLBACK_MODEL = None

app = Flask(__name__)

# 配置密钥和登录密码
app.secret_key = os.environ.get('SECRET_KEY', 'your-secret-key-change-this-in-production')
LOGIN_PASSWORD = os.environ.get('LOGIN_PASSWORD', 'admin123')  # 默认密码，建议通过环境变量设置

# 调试信息：显示当前使用的密码（仅显示前3位）
print(f"🔑 当前登录密码: {LOGIN_PASSWORD[:3]}*** (长度: {len(LOGIN_PASSWORD)})")
print(f"🔐 SECRET_KEY已设置: {'是' if os.environ.get('SECRET_KEY') else '否'}")

# 用户活动跟踪
user_sessions = {}  # 存储用户会话信息
user_activity_log = []  # 存储用户活动日志
translation_stats = defaultdict(int)  # 翻译统计

# 管理员密码（可通过环境变量设置）
ADMIN_PASSWORD = os.environ.get('ADMIN_PASSWORD', 'admin888')

# 获取系统临时目录作为基础路径
TEMP_BASE_DIR = tempfile.gettempdir()

# 配置所有目录为临时目录的子目录
UPLOAD_FOLDER = os.path.join(TEMP_BASE_DIR, "translator_uploads")
TEMPLATE_FOLDER = os.path.join(TEMP_BASE_DIR, "translator_templates")
TRANSLATED_FOLDER = os.path.join(TEMP_BASE_DIR, "translator_translated")

app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER
app.config["TEMPLATE_FOLDER"] = TEMPLATE_FOLDER
app.config["TRANSLATED_FOLDER"] = TRANSLATED_FOLDER
app.config["MAX_CONTENT_LENGTH"] = 50 * 1024 * 1024

# 确保所有目录存在
for folder in [UPLOAD_FOLDER, TEMPLATE_FOLDER, TRANSLATED_FOLDER]:
    os.makedirs(folder, exist_ok=True)

print(f"📁 工作目录:")
print(f"   上传目录: {UPLOAD_FOLDER}")
print(f"   模板目录: {TEMPLATE_FOLDER}")
print(f"   翻译输出目录: {TRANSLATED_FOLDER}")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(TEMPLATE_FOLDER, exist_ok=True)  # 新增：创建模板文件夹
os.makedirs(TRANSLATED_FOLDER, exist_ok=True)

# 初始化配置管理器
if _prompt_modules_available:
    config_manager = ConfigManager()

ALLOWED_EXTENSIONS = {"txt", "docx", "pptx", "xlsx", "xls"}  # 新增：支持更多文件类型


def allowed_file(filename):
    if "." not in filename:
        return False
    try:
        extension = filename.rsplit(".", 1)[1].lower()
        return extension in ALLOWED_EXTENSIONS
    except IndexError:
        return False


def check_function_supports_prompt_config(func):
    """检查函数是否支持prompt_config参数"""
    try:
        sig = inspect.signature(func)
        return "prompt_config" in sig.parameters
    except Exception:
        return False


def apply_prompt_config_to_translator(translator, prompt_config):
    """安全地应用prompt配置到翻译器"""
    if prompt_config and hasattr(translator, "set_prompt_config"):
        try:
            translator.set_prompt_config(prompt_config)
            logger.info("Successfully applied prompt configuration to translator")
            return True
        except Exception as e:
            logger.warning(f"Failed to apply prompt config to translator: {e}")
    return False


# 用户活动跟踪函数
def log_user_activity(action, details=None):
    """记录用户活动"""
    user_ip = request.remote_addr
    user_agent = request.headers.get('User-Agent', 'Unknown')
    session_id = session.get('session_id', 'anonymous')
    
    activity = {
        'timestamp': datetime.now(),
        'session_id': session_id,
        'ip': user_ip,
        'user_agent': user_agent[:100],  # 限制长度
        'action': action,
        'details': details
    }
    
    user_activity_log.append(activity)
    
    # 保持日志大小在合理范围内（最多1000条）
    if len(user_activity_log) > 1000:
        user_activity_log.pop(0)

def update_user_session(action='activity'):
    """更新用户会话信息"""
    session_id = session.get('session_id')
    if not session_id:
        session_id = str(uuid.uuid4())
        session['session_id'] = session_id
    
    user_ip = request.remote_addr
    user_agent = request.headers.get('User-Agent', 'Unknown')
    
    user_sessions[session_id] = {
        'ip': user_ip,
        'user_agent': user_agent[:100],
        'last_activity': datetime.now(),
        'login_time': user_sessions.get(session_id, {}).get('login_time', datetime.now()),
        'status': 'online'
    }

# 登录验证装饰器
def login_required(f):
    """登录验证装饰器"""
    def decorated_function(*args, **kwargs):
        if not session.get('logged_in'):
            return redirect(url_for('login'))
        update_user_session()  # 更新用户活动
        return f(*args, **kwargs)
    decorated_function.__name__ = f.__name__
    return decorated_function

# 管理员验证装饰器
def admin_required(f):
    """管理员验证装饰器"""
    def decorated_function(*args, **kwargs):
        if not session.get('admin_logged_in'):
            return redirect(url_for('admin_login'))
        return f(*args, **kwargs)
    decorated_function.__name__ = f.__name__
    return decorated_function


@app.route("/")
@login_required
def index():
    return render_template("index.html")


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        password = request.form.get("password")
        
        # 检查是否为管理员密码
        if password == ADMIN_PASSWORD:
            session['admin_logged_in'] = True
            session['logged_in'] = True  # 管理员也有普通用户权限
            log_user_activity('管理员登录', {'success': True, 'auto_detect': True})
            return redirect(url_for('admin_dashboard'))
        
        # 检查是否为普通用户密码
        elif password == LOGIN_PASSWORD:
            session['logged_in'] = True
            update_user_session('login')
            log_user_activity('用户登录', {'success': True})
            return redirect(url_for('index'))
        
        else:
            log_user_activity('登录失败', {'reason': '密码错误'})
            return render_template("login.html", error="密码错误，请重试")
    
    return render_template("login.html")


@app.route("/logout")
def logout():
    session_id = session.get('session_id')
    if session_id and session_id in user_sessions:
        user_sessions[session_id]['status'] = 'offline'
        log_user_activity('用户登出')
    session.pop('logged_in', None)
    session.pop('session_id', None)
    return redirect(url_for('login'))


# 管理员登录路由（重定向到统一登录页面）
@app.route("/admin/login")
def admin_login():
    return redirect(url_for('login'))


# 管理员登出
@app.route("/admin/logout")
def admin_logout():
    log_user_activity('管理员登出')
    session.pop('admin_logged_in', None)
    return redirect(url_for('admin_login'))


# 管理员后台首页
@app.route("/admin")
@admin_required
def admin_dashboard():
    # 清理过期会话（超过30分钟无活动）
    current_time = datetime.now()
    expired_sessions = []
    for session_id, session_data in user_sessions.items():
        if current_time - session_data['last_activity'] > timedelta(minutes=30):
            session_data['status'] = 'offline'
            expired_sessions.append(session_id)
    
    # 统计数据
    online_users = sum(1 for s in user_sessions.values() if s['status'] == 'online')
    total_sessions = len(user_sessions)
    recent_activities = sorted(user_activity_log, key=lambda x: x['timestamp'], reverse=True)[:50]
    
    # 翻译统计
    today = datetime.now().date()
    today_translations = sum(1 for activity in user_activity_log 
                           if activity['timestamp'].date() == today and '翻译' in activity['action'])
    
    stats = {
        'online_users': online_users,
        'total_sessions': total_sessions,
        'today_translations': today_translations,
        'total_activities': len(user_activity_log)
    }
    
    return render_template("admin_dashboard.html", 
                         stats=stats, 
                         user_sessions=user_sessions,
                         recent_activities=recent_activities)


# API：获取实时统计数据
@app.route("/admin/api/stats")
@admin_required
def admin_api_stats():
    current_time = datetime.now()
    
    # 更新会话状态
    for session_id, session_data in user_sessions.items():
        if current_time - session_data['last_activity'] > timedelta(minutes=30):
            session_data['status'] = 'offline'
    
    online_users = sum(1 for s in user_sessions.values() if s['status'] == 'online')
    
    # 今日活动统计
    today = current_time.date()
    today_activities = [a for a in user_activity_log if a['timestamp'].date() == today]
    today_logins = sum(1 for a in today_activities if a['action'] == '用户登录')
    today_translations = sum(1 for a in today_activities if '翻译' in a['action'])
    
    return jsonify({
        'online_users': online_users,
        'total_sessions': len(user_sessions),
        'today_logins': today_logins,
        'today_translations': today_translations,
        'total_activities': len(user_activity_log)
    })


# 新增：PPT和Excel文件翻译路由
# 修改 app.py 中的 /translate_file 路由


@app.route("/translate_file", methods=["POST"])
@login_required
def translate_file():
    # 记录翻译活动
    log_user_activity('文件翻译开始', {
        'file_type': request.files.get('file').filename.split('.')[-1] if request.files.get('file') else 'unknown',
        'target_lang': request.form.get('target_lang'),
        'processing_method': request.form.get('processing_method')
    })
    try:
        logger.info("=== Starting translate_file request ===")

        # 获取文件和参数
        file = request.files.get("file")
        if not file or file.filename == "":
            logger.error("No file provided")
            return jsonify({"success": False, "message": "未选择文件"})

        processing_method = request.form.get("processing_method", "")
        target_lang = request.form.get("target_lang")
        source_lang = request.form.get("source_lang")

        logger.info(f"Processing method: {processing_method}")
        logger.info(f"Target language: {target_lang}")
        logger.info(f"Source language: {source_lang}")

        if not target_lang:
            logger.error("Target language not provided")
            return jsonify({"success": False, "message": "目标语言不能为空"})

        # 获取翻译器配置
        api_platform = request.form.get("api_platform", "siliconflow")
        frontend_api_key = request.form.get("api_key", "").strip()
        frontend_base_url = request.form.get("base_url", "").strip()
        frontend_model = request.form.get("model", "").strip()

        logger.info(f"API platform: {api_platform}")

        # ✅ 修复：添加动态平台配置（与/translate_api相同）
        platform_configs = {
            "siliconflow": {
                "api_key_env": "SILICONFLOW_API_KEY",
                "base_url_env": "SILICONFLOW_BASE_URL",
                "model_env": "SILICONFLOW_MODEL",
                "default_base_url": "https://api.siliconflow.cn/v1",
                "default_model": "THUDM/GLM-4-9B-0414",
                # "extra_body": {
                #     "enable_thinking": False,
                # },
            },
            "zhipu": {
                "api_key_env": "ZHIPU_API_KEY",
                "base_url_env": "ZHIPU_BASE_URL",
                "model_env": "ZHIPU_MODEL",
                "default_base_url": "https://open.bigmodel.cn/api/paas/v4",
                "default_model": "GLM-4-Flash-250414",
            },            
            "deepseek": {
                "api_key_env": "DEEPSEEK_API_KEY",
                "base_url_env": "DEEPSEEK_BASE_URL",
                "model_env": "DEEPSEEK_MODEL",
                "default_base_url": "https://api.deepseek.com/v1",
                "default_model": "deepseek-chat",
            },
            "moonshot": {
                "api_key_env": "MOONSHOT_API_KEY",
                "base_url_env": "MOONSHOT_BASE_URL",
                "model_env": "MOONSHOT_MODEL",
                "default_base_url": "https://api.moonshot.cn/v1",
                "default_model": "moonshot-v1-8k",
            },
            "openai": {
                "api_key_env": "OPENAI_API_KEY",
                "base_url_env": "OPENAI_BASE_URL",
                "model_env": "OPENAI_MODEL",
                "default_base_url": "https://api.openai.com/v1",
                "default_model": "gpt-3.5-turbo",
            },
            "ollama": {
                "api_key_env": "OLLAMA_API_KEY",
                "base_url_env": "OLLAMA_BASE_URL",
                "model_env": "OLLAMA_MODEL",
                "default_base_url": "http://localhost:11434/v1",
                "default_model": "llama3",
            },
            "modelscope": {
                "api_key_env": "MODELSCOPE_API_KEY",
                "base_url_env": "MODELSCOPE_BASE_URL",
                "model_env": "MODELSCOPE_MODEL",
                "default_base_url": "https://api-inference.modelscope.cn/v1",
                "default_model": "Qwen/Qwen2.5-72B-Instruct",
                "extra_body": {
                    "enable_thinking": False,
                },
            },
            "groq": {
                "api_key_env": "GROQ_API_KEY",
                "base_url_env": "GROQ_BASE_URL",
                "model_env": "GROQ_MODEL",
                "default_base_url": "https://api.groq.com/openai/v1",
                "default_model": "openai/gpt-oss-120b",
                "extra_body": {
                    "enable_thinking": False,
                },
            },            
            "openrouter": {
                "api_key_env": "OPENROUTER_API_KEY",
                "base_url_env": "OPENROUTER_BASE_URL",
                "model_env": "OPENROUTER_MODEL",
                "default_base_url": "https://openrouter.ai/api/v1",
                "default_model": "google/gemini-2.0-flash-exp:free",
            },
            "gemini": {
                "api_key_env": "GEMINI_API_KEY",
                "base_url_env": "GEMINI_BASE_URL",
                "model_env": "GEMINI_MODEL",
                "default_base_url": "https://generativelanguage.googleapis.com/v1beta/openai/",
                "default_model": "gemini-1.5-flash",
            },
            "volcengine": {
                "api_key_env": "VOLCENGINE_ACCESS_KEY",
                "secret_key_env": "VOLCENGINE_SECRET_KEY",
                "region_env": "VOLCENGINE_REGION",
                "default_region": "cn-beijing",
            },
            "volcengine_sdk": {
                "api_key_env": "VOLCENGINE_ACCESS_KEY",
                "secret_key_env": "VOLCENGINE_SECRET_KEY",
                "region_env": "VOLCENGINE_REGION",
                "default_region": "cn-beijing",
            },
        }

        # ✅ 修复：动态获取配置
        api_key_to_use = frontend_api_key
        base_url_to_use = frontend_base_url
        model_to_use = frontend_model
        selected_platform_config = platform_configs.get(api_platform)

        if selected_platform_config:
            logger.info(f"Processing config for known platform: {api_platform}")
            if not api_key_to_use and selected_platform_config.get("api_key_env"):
                api_key_to_use = os.getenv(
                    selected_platform_config["api_key_env"], api_key_to_use
                )
            if not base_url_to_use and selected_platform_config.get("base_url_env"):
                base_url_to_use = os.getenv(
                    selected_platform_config["base_url_env"], base_url_to_use
                )
            if not model_to_use and selected_platform_config.get("model_env"):
                model_to_use = os.getenv(
                    selected_platform_config["model_env"], model_to_use
                )
            if not base_url_to_use:
                base_url_to_use = selected_platform_config.get("default_base_url")
            if not model_to_use:
                model_to_use = selected_platform_config.get("default_model")

        # 如果是火山引擎SDK模式，也尝试从volc_ak获取API Key
        if api_platform == "volcengine_sdk" and not api_key_to_use:
            api_key_to_use = request.form.get("volc_ak", "").strip()

        if not api_key_to_use:
            api_key_to_use = DEFAULT_FALLBACK_API_KEY
        if not base_url_to_use:
            base_url_to_use = DEFAULT_FALLBACK_BASE_URL
        if not model_to_use:
            model_to_use = DEFAULT_FALLBACK_MODEL

        logger.info(
            f"Final config - Platform: '{api_platform}', API Key: {'SET' if api_key_to_use else 'NOT SET'}, Base URL: {base_url_to_use}, Model: {model_to_use}"
        )

        if not api_key_to_use:
            logger.error("No API key available")
            return jsonify({"success": False, "message": "API密钥未配置"})

        # 修改：火山引擎不需要base_url和model检查
        if api_platform not in ["volcengine", "volcengine_sdk"]:
            if not base_url_to_use:
                logger.error("No base URL available")
                return jsonify({"success": False, "message": "Base URL未配置"})
            if not model_to_use:
                logger.error("No model available")
                return jsonify({"success": False, "message": "模型未配置"})

        # 检查翻译器类是否可用
        if api_platform in ["volcengine", "volcengine_sdk"]:
            if not _volcengine_available or VolcEngineNativeTranslator is None:
                logger.error("VolcEngineNativeTranslator not available")
                return jsonify(
                    {
                        "success": False,
                        "message": "VolcEngineNativeTranslator翻译器不可用",
                    }
                )
        else:
            if SiliconFlowTranslator is None:
                logger.error("SiliconFlowTranslator not available")
                return jsonify(
                    {"success": False, "message": "SiliconFlowTranslator翻译器不可用"}
                )

        # 创建翻译器实例
        try:
            if api_platform in ["volcengine", "volcengine_sdk"]:
                # 火山引擎翻译器需要特殊处理
                secret_key = None
                region = "cn-beijing"  # 默认区域

                # 获取Secret Key
                if selected_platform_config and selected_platform_config.get(
                    "secret_key_env"
                ):
                    secret_key = os.getenv(selected_platform_config["secret_key_env"])

                # 获取区域设置
                if selected_platform_config and selected_platform_config.get(
                    "region_env"
                ):
                    region = os.getenv(selected_platform_config["region_env"], region)
                elif selected_platform_config and selected_platform_config.get(
                    "default_region"
                ):
                    region = selected_platform_config["default_region"]

                # 从前端获取secret key（支持两种字段名）
                frontend_secret_key = request.form.get("secret_key", "").strip()
                if not frontend_secret_key:
                    frontend_secret_key = request.form.get("volc_sk", "").strip()
                if frontend_secret_key:
                    secret_key = frontend_secret_key

                if not secret_key:
                    logger.error("Secret Key for VolcEngine is missing.")
                    return jsonify(
                        {"success": False, "message": "火山引擎Secret Key未配置"}
                    )

                # 创建火山引擎翻译器实例
                translator = VolcEngineNativeTranslator(
                    ak=api_key_to_use, sk=secret_key, region=region
                )
                logger.info(f"Created VolcEngine translator with region: {region}")
            else:
                # 其他平台使用SiliconFlowTranslator
                # 提取extra_body配置
                extra_body = (
                    selected_platform_config.get("extra_body", {})
                    if selected_platform_config
                    else {}
                )

                translator = SiliconFlowTranslator(
                    api_key=api_key_to_use,
                    base_url=base_url_to_use,
                    model=model_to_use,
                    platform_id=api_platform,
                    extra_body=extra_body,
                )
                logger.info("Translator instance created successfully")
        except Exception as e:
            logger.error(f"Failed to create translator: {e}")
            return jsonify({"success": False, "message": f"创建翻译器失败: {str(e)}"})

        # 保存上传的文件
        raw_filename = file.filename
        logger.info(f"Original filename: {raw_filename}")
        
        # 安全地获取文件扩展名（从原始文件名）
        if "." in raw_filename:
            file_extension = raw_filename.rsplit(".", 1)[1].lower()
            filename_without_ext = raw_filename.rsplit(".", 1)[0]
        else:
            logger.error(f"File has no extension: {raw_filename}")
            return jsonify({"success": False, "message": "文件必须有扩展名"})
        
        # 处理纯中文文件名问题
        safe_filename = secure_filename(raw_filename)
        logger.info(f"Secure filename: {safe_filename}")
        
        # 如果secure_filename过滤后文件名为空或没有扩展名，生成新文件名
        if not safe_filename or "." not in safe_filename:
            import time
            new_filename_base = f"file_{int(time.time())}"
            logger.info(f"Chinese filename detected, using generated name: {new_filename_base}")
        else:
            new_filename_base = safe_filename.rsplit(".", 1)[0]
        
        unique_filename_base = str(uuid.uuid4())
        input_filename = unique_filename_base + "." + file_extension
        input_filepath = os.path.join(app.config["UPLOAD_FOLDER"], input_filename)

        try:
            file.save(input_filepath)
            logger.info(f"File saved: {input_filepath}")
        except Exception as e:
            logger.error(f"Failed to save file: {e}")
            return jsonify({"success": False, "message": f"文件保存失败: {str(e)}"})

        # 设置输出目录
        output_dir = app.config["TRANSLATED_FOLDER"]
        logger.info(f"Output directory: {output_dir}")

        # 解析prompt配置
        prompt_config = {}
        try:
            prompt_config_str = request.form.get("prompt_config", "{}")
            if prompt_config_str:
                prompt_config = json.loads(prompt_config_str)
            logger.info(f"Prompt config: {prompt_config}")
        except json.JSONDecodeError as e:
            logger.warning(f"Failed to parse prompt_config: {e}")
            prompt_config = {}

        # 准备通用配置
        common_config = {"prompt_config": prompt_config}

        result_path = None

        # 根据处理方法调用相应的翻译器
        if processing_method == "pptx_full_translator":
            logger.info("Processing PPTX file...")

            if not _translate_pptx_available or translate_pptx_file_formatted is None:
                logger.error("PPTX translator not available")
                # 清理临时文件
                if os.path.exists(input_filepath):
                    os.remove(input_filepath)
                return jsonify({"success": False, "message": "PPTX翻译器不可用"})

            # 获取PPTX配置参数
            ppt_config = {
                "translate_notes": request.form.get("translate_notes", "false")
                == "true",
                "translate_slide_titles": request.form.get(
                    "translate_slide_titles", "true"
                )
                == "true",
            }
            logger.info(f"PPTX config: {ppt_config}")

            try:
                # 检查函数签名，看看需要什么参数
                import inspect

                sig = inspect.signature(translate_pptx_file_formatted)
                logger.info(f"PPTX translator function signature: {sig}")

                # 调用PPTX翻译器
                result_path = translate_pptx_file_formatted(
                    input_filepath=input_filepath,
                    output_dir=output_dir,
                    target_lang=target_lang,
                    source_lang=source_lang,
                    translator=translator,
                    **ppt_config,
                    **common_config,
                )
                logger.info(f"PPTX translation completed: {result_path}")

            except Exception as e:
                logger.error(f"PPTX translation error: {type(e).__name__}: {e}")
                # 清理临时文件
                if os.path.exists(input_filepath):
                    os.remove(input_filepath)
                return jsonify({"success": False, "message": f"PPTX翻译失败: {str(e)}"})

        elif processing_method == "excel_full_translator":
            logger.info("Processing Excel file...")

            if not _translate_excel_available or translate_excel_file_formatted is None:
                logger.error("Excel translator not available")
                # 清理临时文件
                if os.path.exists(input_filepath):
                    os.remove(input_filepath)
                return jsonify({"success": False, "message": "Excel翻译器不可用"})

            # 获取Excel配置参数
            excel_config = {
                "translate_headers": request.form.get("translate_headers", "true")
                == "true",
                "translate_comments": request.form.get("translate_comments", "true")
                == "true",
                "translate_charts": request.form.get("translate_charts", "true")
                == "true",
                "skip_formulas": request.form.get("skip_formulas", "true") == "true",
                "skip_numbers": request.form.get("skip_numbers", "true") == "true",
                "skip_dates": request.form.get("skip_dates", "true") == "true",
                "smart_detection": request.form.get("smart_detection", "true")
                == "true",
                "max_text_length": int(request.form.get("max_text_length", 200)),
                "header_detection_rows": int(
                    request.form.get("header_detection_rows", 5)
                ),
                "min_alpha_ratio": float(request.form.get("min_alpha_ratio", 0.3)),
                "selected_sheets": json.loads(
                    request.form.get("selected_sheets", "null")
                ),
            }
            logger.info(f"Excel config: {excel_config}")

            try:
                # 检查函数签名
                import inspect

                sig = inspect.signature(translate_excel_file_formatted)
                logger.info(f"Excel translator function signature: {sig}")

                # 调用Excel翻译器
                result_path = translate_excel_file_formatted(
                    input_filepath=input_filepath,
                    output_dir=output_dir,
                    target_lang=target_lang,
                    source_lang=source_lang,
                    translator=translator,
                    **excel_config,
                    **common_config,
                )
                logger.info(f"Excel translation completed: {result_path}")

            except Exception as e:
                logger.error(f"Excel translation error: {type(e).__name__}: {e}")
                # 清理临时文件
                if os.path.exists(input_filepath):
                    os.remove(input_filepath)
                return jsonify(
                    {"success": False, "message": f"Excel翻译失败: {str(e)}"}
                )

        elif processing_method == "text_translator":
            logger.info("Processing text file...")

            if file_extension != "txt":
                return jsonify(
                    {"success": False, "message": "文本翻译器只支持.txt文件"}
                )

            encoding = request.form.get("encoding", "utf-8")

            try:
                result_path = translate_text_file(
                    input_filepath=input_filepath,
                    output_dir=output_dir,
                    target_lang=target_lang,
                    translator=translator,
                    source_lang=source_lang,
                    encoding=encoding,
                    unique_filename_base=unique_filename_base,
                    **common_config,
                )
                logger.info(f"Text translation completed: {result_path}")

            except Exception as e:
                logger.error(f"Text translation error: {type(e).__name__}: {e}")
                if os.path.exists(input_filepath):
                    os.remove(input_filepath)
                return jsonify({"success": False, "message": f"文本翻译失败: {str(e)}"})

        elif processing_method == "docx_translator":
            logger.info("Processing DOCX file (basic)...")

            if file_extension != "docx":
                return jsonify(
                    {"success": False, "message": "DOCX翻译器只支持.docx文件"}
                )

            try:
                result_path = translate_docx_file(
                    input_filepath=input_filepath,
                    output_dir=output_dir,
                    target_lang=target_lang,
                    translator=translator,
                    source_lang=source_lang,
                    unique_filename_base=unique_filename_base,
                    **common_config,
                )
                logger.info(f"DOCX translation completed: {result_path}")

            except Exception as e:
                logger.error(f"DOCX translation error: {type(e).__name__}: {e}")
                if os.path.exists(input_filepath):
                    os.remove(input_filepath)
                return jsonify({"success": False, "message": f"DOCX翻译失败: {str(e)}"})

        elif processing_method == "docx_full_translator":
            logger.info("Processing DOCX file (formatted)...")

            if file_extension != "docx":
                return jsonify(
                    {"success": False, "message": "格式化DOCX翻译器只支持.docx文件"}
                )

            try:
                result_path = translate_docx_file_formatted(
                    input_filepath=input_filepath,
                    output_dir=output_dir,
                    target_lang=target_lang,
                    translator=translator,
                    source_lang=source_lang,
                    unique_filename_base=unique_filename_base,
                    **common_config,
                )
                logger.info(f"DOCX formatted translation completed: {result_path}")

            except Exception as e:
                logger.error(
                    f"DOCX formatted translation error: {type(e).__name__}: {e}"
                )
                if os.path.exists(input_filepath):
                    os.remove(input_filepath)
                return jsonify(
                    {"success": False, "message": f"格式化DOCX翻译失败: {str(e)}"}
                )

        elif processing_method in ["docx_pythondoc1_translator", "docx_pythondoc1"]:
            logger.info("Processing DOCX file (pythondoc1)...")

            if file_extension != "docx":
                return jsonify(
                    {"success": False, "message": "PythonDoc1翻译器只支持.docx文件"}
                )

            try:
                result_path = translate_docx_pythondoc1(
                    input_filepath=input_filepath,
                    output_dir=output_dir,
                    target_lang=target_lang,
                    translator=translator,
                    source_lang=source_lang,
                    unique_filename_base=unique_filename_base,
                    **common_config,
                )
                logger.info(f"DOCX pythondoc1 translation completed: {result_path}")

            except Exception as e:
                logger.error(
                    f"DOCX pythondoc1 translation error: {type(e).__name__}: {e}"
                )
                if os.path.exists(input_filepath):
                    os.remove(input_filepath)
                return jsonify(
                    {"success": False, "message": f"PythonDoc1翻译失败: {str(e)}"}
                )

        elif processing_method in ["docx_pythondoc2_translator", "docx_pythondoc2"]:
            logger.info("Processing DOCX file (pythondoc2)...")

            if file_extension != "docx":
                return jsonify(
                    {"success": False, "message": "PythonDoc2翻译器只支持.docx文件"}
                )

            try:
                result_path = translate_docx_pythondoc2(
                    input_filepath=input_filepath,
                    output_dir=output_dir,
                    target_lang=target_lang,
                    translator=translator,
                    source_lang=source_lang,
                    unique_filename_base=unique_filename_base,
                    **common_config,
                )
                logger.info(f"DOCX pythondoc2 translation completed: {result_path}")

            except Exception as e:
                logger.error(
                    f"DOCX pythondoc2 translation error: {type(e).__name__}: {e}"
                )
                if os.path.exists(input_filepath):
                    os.remove(input_filepath)
                return jsonify(
                    {"success": False, "message": f"PythonDoc2翻译失败: {str(e)}"}
                )

        elif processing_method in ["docx_markdown_translator", "docx_markdown"]:
            logger.info("Processing DOCX file (via markdown)...")

            if file_extension != "docx":
                return jsonify(
                    {"success": False, "message": "Markdown翻译器只支持.docx文件"}
                )

            try:
                # 检查是否有模板文件
                template_file_path = None
                if "template_file" in request.files:
                    template_file = request.files["template_file"]
                    if template_file and template_file.filename:
                        # 保存模板文件到临时位置
                        template_filename = secure_filename(template_file.filename)
                        template_file_path = os.path.join(
                            app.config["TEMPLATE_FOLDER"],
                            f"temp_{unique_filename_base}_{template_filename}",
                        )
                        template_file.save(template_file_path)
                        logger.info(f"Using template file: {template_file_path}")

                # 调用markdown翻译器
                kwargs = {"template_path": template_file_path, **common_config}

                result_path = translate_docx_via_markdown(
                    input_filepath=input_filepath,
                    output_dir=output_dir,
                    target_lang=target_lang,
                    translator=translator,
                    source_lang=source_lang,
                    unique_filename_base=unique_filename_base,
                    **kwargs,
                )
                logger.info(f"DOCX markdown translation completed: {result_path}")

                # 清理模板文件
                if template_file_path and os.path.exists(template_file_path):
                    os.remove(template_file_path)
                    logger.info("Template file cleaned up")

            except Exception as e:
                logger.error(
                    f"DOCX markdown translation error: {type(e).__name__}: {e}"
                )
                if os.path.exists(input_filepath):
                    os.remove(input_filepath)
                # 清理模板文件
                if (
                    "template_file_path" in locals()
                    and template_file_path
                    and os.path.exists(template_file_path)
                ):
                    os.remove(template_file_path)
                return jsonify(
                    {"success": False, "message": f"Markdown翻译失败: {str(e)}"}
                )

        else:
            logger.error(f"Unsupported processing method: {processing_method}")
            # 清理临时文件
            if os.path.exists(input_filepath):
                os.remove(input_filepath)
            return jsonify(
                {"success": False, "message": f"不支持的处理方法: {processing_method}"}
            )

        # 清理临时文件
        try:
            if os.path.exists(input_filepath):
                os.remove(input_filepath)
                logger.info("Temporary file cleaned up")
        except Exception as e:
            logger.warning(f"Failed to remove temporary file: {e}")

        # 检查结果
        if result_path and os.path.exists(result_path):
            output_filename = os.path.basename(result_path)
            logger.info(f"Translation successful, output file: {output_filename}")
            return jsonify(
                {
                    "success": True,
                    "message": "翻译完成",
                    "download_url": f"/download/{output_filename}",
                }
            )
        else:
            logger.error(f"Translation failed, result_path: {result_path}")
            return jsonify({"success": False, "message": "翻译失败，未生成输出文件"})

    except Exception as e:
        logger.error(f"Unexpected error in translate_file: {type(e).__name__}: {e}")
        import traceback

        logger.error(f"Traceback: {traceback.format_exc()}")

        # 确保清理临时文件
        try:
            if "input_filepath" in locals() and os.path.exists(input_filepath):
                os.remove(input_filepath)
        except:
            pass

        return jsonify({"success": False, "message": f"翻译过程中出现错误: {str(e)}"})


@app.route("/translate_api", methods=["POST"])
@login_required
def translate_api():
    # 记录API翻译活动
    log_user_activity('API翻译请求', {
        'platform': request.form.get('api_platform', 'unknown'),
        'target_lang': request.form.get('target_lang'),
        'has_text': bool(request.form.get('text_input')),
        'has_file': bool(request.files.get('file'))
    })
    
    data = request.form
    api_platform = data.get("api_platform", "custom")

    logger.info(f"Received translation request for platform: {api_platform}")

    frontend_api_key = data.get("api_key", "").strip()
    frontend_base_url = data.get("base_url", "").strip()
    frontend_model = data.get("model", "").strip()

    target_lang = data.get("target_lang")
    source_lang = data.get("source_lang")

    if not target_lang:
        logger.error("Target language is required.")
        return jsonify({"error": "Target language is required."}), 400

    # 新增：解析prompt配置
    prompt_config = None
    if "prompt_config" in data:
        try:
            prompt_config = json.loads(data.get("prompt_config", "{}"))
            logger.info(
                f"Received prompt config: mode={prompt_config.get('mode')}, template={prompt_config.get('prompt_template')}"
            )
        except json.JSONDecodeError:
            logger.warning("Failed to parse prompt_config, using default")
            prompt_config = None

    # Configuration loading logic
    platform_configs = {
        "siliconflow": {
            "api_key_env": "SILICONFLOW_API_KEY",
            "base_url_env": "SILICONFLOW_BASE_URL",
            "model_env": "SILICONFLOW_MODEL",
            "default_base_url": "https://api.siliconflow.cn/v1",
            "default_model": "THUDM/GLM-4-9B-0414",
            # "extra_body": {
            #     "enable_thinking": False,  # 只对QWEN3生效，其他模型忽略
            # },
        },
        "zhipu": {
            "api_key_env": "ZHIPU_API_KEY",
            "base_url_env": "ZHIPU_BASE_URL",
            "model_env": "ZHIPU_MODEL",
            "default_base_url": "https://open.bigmodel.cn/api/paas/v4",
            "default_model": "GLM-4-Flash-250414",
        },        
        "deepseek": {
            "api_key_env": "DEEPSEEK_API_KEY",
            "base_url_env": "DEEPSEEK_BASE_URL",
            "model_env": "DEEPSEEK_MODEL",
            "default_base_url": "https://api.deepseek.com/v1",
            "default_model": "deepseek-chat",
        },
        "moonshot": {
            "api_key_env": "MOONSHOT_API_KEY",
            "base_url_env": "MOONSHOT_BASE_URL",
            "model_env": "MOONSHOT_MODEL",
            "default_base_url": "https://api.moonshot.cn/v1",
            "default_model": "moonshot-v1-8k",
        },
        "openai": {
            "api_key_env": "OPENAI_API_KEY",
            "base_url_env": "OPENAI_BASE_URL",
            "model_env": "OPENAI_MODEL",
            "default_base_url": "https://api.openai.com/v1",
            "default_model": "gpt-3.5-turbo",
        },
        "ollama": {
            "api_key_env": "OLLAMA_API_KEY",
            "base_url_env": "OLLAMA_BASE_URL",
            "model_env": "OLLAMA_MODEL",
            "default_base_url": "http://localhost:11434/v1",
            "default_model": "llama3",
        },
        "modelscope": {
            "api_key_env": "MODELSCOPE_API_KEY",
            "base_url_env": "MODELSCOPE_BASE_URL",
            "model_env": "MODELSCOPE_MODEL",
            "default_base_url": "https://api-inference.modelscope.cn/v1",
            "default_model": "Qwen/Qwen2.5-72B-Instruct",
            "extra_body": {
                "enable_thinking": False,  # 只对QWEN3生效，其他模型忽略
            },
        },
        "groq": {
            "api_key_env": "GROQ_API_KEY",
            "base_url_env": "GROQ_BASE_URL",
            "model_env": "GROQ_MODEL",
            "default_base_url": "https://api.groq.com/openai/v1",
            "default_model": "openai/gpt-oss-120b",
            # "extra_body": {
            #     "enable_thinking": False,  # 只对QWEN3生效，其他模型忽略
            # },
        },        
        "openrouter": {
            "api_key_env": "OPENROUTER_API_KEY",
            "base_url_env": "OPENROUTER_BASE_URL",
            "model_env": "OPENROUTER_MODEL",
            "default_base_url": "https://openrouter.ai/api/v1",
            "default_model": "google/gemini-2.0-flash-exp:free",
        },
        "gemini": {
            "api_key_env": "GEMINI_API_KEY",
            "base_url_env": "GEMINI_BASE_URL",
            "model_env": "GEMINI_MODEL",
            "default_base_url": "https://generativelanguage.googleapis.com/v1beta/openai/",
            "default_model": "gemini-1.5-flash",
        },

        # 新增：火山引擎配置 - 兼容模式
        "volcengine": {
            "api_key_env": "VOLCENGINE_ACCESS_KEY",  # Access Key
            "secret_key_env": "VOLCENGINE_SECRET_KEY",  # Secret Key
            "region_env": "VOLCENGINE_REGION",
            "default_region": "cn-beijing",
        },
        # 新增：火山引擎配置 - SDK模式
        "volcengine_sdk": {
            "api_key_env": "VOLCENGINE_ACCESS_KEY",  # Access Key
            "secret_key_env": "VOLCENGINE_SECRET_KEY",  # Secret Key
            "region_env": "VOLCENGINE_REGION",
            "default_region": "cn-beijing",
        },
    }

    api_key_to_use = frontend_api_key
    base_url_to_use = frontend_base_url
    model_to_use = frontend_model
    selected_platform_config = platform_configs.get(api_platform)

    if selected_platform_config:
        logger.info(f"Processing config for known platform: {api_platform}")
        if not api_key_to_use and selected_platform_config.get("api_key_env"):
            api_key_to_use = os.getenv(
                selected_platform_config["api_key_env"], api_key_to_use
            )
        if not base_url_to_use and selected_platform_config.get("base_url_env"):
            base_url_to_use = os.getenv(
                selected_platform_config["base_url_env"], base_url_to_use
            )
        if not model_to_use and selected_platform_config.get("model_env"):
            model_to_use = os.getenv(
                selected_platform_config["model_env"], model_to_use
            )
        if not base_url_to_use:
            base_url_to_use = selected_platform_config.get("default_base_url")
        if not model_to_use:
            model_to_use = selected_platform_config.get("default_model")

    # 新增：如果是火山引擎SDK模式，也尝试从volc_ak获取API Key
    if api_platform == "volcengine_sdk" and not api_key_to_use:
        api_key_to_use = data.get("volc_ak", "").strip()

    if not api_key_to_use:
        api_key_to_use = DEFAULT_FALLBACK_API_KEY
    if not base_url_to_use:
        base_url_to_use = DEFAULT_FALLBACK_BASE_URL
    if not model_to_use:
        model_to_use = DEFAULT_FALLBACK_MODEL

    logger.info(
        f"Final config for API call - Platform: '{api_platform}', API Key: {'SET' if api_key_to_use else 'NOT SET'}, Base URL: {base_url_to_use}, Model: {model_to_use}"
    )

    if not api_key_to_use:
        return (
            jsonify({"error": f"API Key for platform '{api_platform}' is missing."}),
            400,
        )

    # 修改：火山引擎不需要base_url和model检查
    if api_platform not in ["volcengine", "volcengine_sdk"]:
        if not base_url_to_use:
            return (
                jsonify(
                    {"error": f"Base URL for platform '{api_platform}' is missing."}
                ),
                400,
            )
        if not model_to_use:
            return (
                jsonify({"error": f"Model for platform '{api_platform}' is missing."}),
                400,
            )

    # 新增：火山引擎SDK特殊处理调试
    if api_platform == "volcengine_sdk" and not api_key_to_use:
        volc_ak_field = data.get("volc_ak", "").strip()
        if volc_ak_field:
            api_key_to_use = volc_ak_field
            logger.info(f"Step 3 - api_key_to_use 从 volc_ak 字段: SET")

    if not api_key_to_use:
        api_key_to_use = DEFAULT_FALLBACK_API_KEY
        logger.info(
            f"Step 4 - api_key_to_use 从 DEFAULT_FALLBACK_API_KEY: {'SET' if api_key_to_use else 'NOT SET'}"
        )
        if api_key_to_use:
            logger.info(f"DEFAULT_FALLBACK_API_KEY 前8位: {api_key_to_use[:8]}...")

    if not base_url_to_use:
        base_url_to_use = DEFAULT_FALLBACK_BASE_URL
    if not model_to_use:
        model_to_use = DEFAULT_FALLBACK_MODEL

    try:
        # 修改：根据平台选择不同的翻译器
        if api_platform in ["volcengine", "volcengine_sdk"]:
            if not _volcengine_available or not VolcEngineNativeTranslator:
                logger.error("VolcEngine translator is not available.")
                return (
                    jsonify(
                        {
                            "error": "VolcEngine translator is not available. Please install volcengine-python-sdk and ensure volcengine_native_translator.py is in the translators directory."
                        }
                    ),
                    500,
                )

            # 火山引擎翻译器需要特殊处理
            secret_key = None
            region = "cn-beijing"  # 默认区域

            # 获取Secret Key
            if selected_platform_config and selected_platform_config.get(
                "secret_key_env"
            ):
                secret_key = os.getenv(selected_platform_config["secret_key_env"])

            # 获取区域设置
            if selected_platform_config and selected_platform_config.get("region_env"):
                region = os.getenv(selected_platform_config["region_env"], region)
            elif selected_platform_config and selected_platform_config.get(
                "default_region"
            ):
                region = selected_platform_config["default_region"]

            # 从前端获取secret key（支持两种字段名）
            frontend_secret_key = data.get("secret_key", "").strip()
            if not frontend_secret_key:
                frontend_secret_key = data.get("volc_sk", "").strip()
            if frontend_secret_key:
                secret_key = frontend_secret_key

            if not secret_key:
                logger.error("Secret Key for VolcEngine is missing.")
                return jsonify({"error": "Secret Key for VolcEngine is required."}), 400

            # 创建火山引擎翻译器实例
            translator_instance = VolcEngineNativeTranslator(
                ak=api_key_to_use, sk=secret_key, region=region
            )

            logger.info(f"Created VolcEngine translator with region: {region}")
        else:
            # 其他平台使用SiliconFlowTranslator
            # 提取extra_body配置
            extra_body = (
                selected_platform_config.get("extra_body", {})
                if selected_platform_config
                else {}
            )

            # 创建翻译器实例
            translator_instance = SiliconFlowTranslator(
                api_key=api_key_to_use,
                base_url=base_url_to_use,
                model=model_to_use,
                platform_id=api_platform,
                extra_body=extra_body,
            )

        # 如果有自定义prompt配置，应用到翻译器
        if prompt_config:
            apply_prompt_config_to_translator(translator_instance, prompt_config)

    except ValueError as e:
        logger.error(f"Translator initialization error: {e}")
        return jsonify({"error": str(e)}), 400
    except Exception as e:
        logger.error(
            f"Unexpected error initializing translator: {type(e).__name__} - {e}"
        )
        return (
            jsonify({"error": f"Server error during translator setup: {str(e)}"}),
            500,
        )

    # --- Text Input Translation ---
    if "text_input" in data and data["text_input"]:
        text_to_translate = data["text_input"]
        if not text_to_translate.strip():
            logger.error("Text to translate cannot be empty.")
            return jsonify({"error": "Text to translate cannot be empty."}), 400

        logger.info(
            f"Processing text translation: '{text_to_translate[:50]}...' to {target_lang} via {api_platform}"
        )

        def generate_translation_stream():
            response_stream = None
            try:
                # 修改：火山引擎不支持流式传输，直接返回翻译结果
                if api_platform in ["volcengine", "volcengine_sdk"]:
                    try:
                        translated_text = translator_instance.translate(
                            text_to_translate, target_lang, source_lang
                        )
                        # 模拟流式返回格式
                        yield f"data: {json.dumps({'text_chunk': translated_text})}\n\n"
                        yield f"data: {json.dumps({'done': True})}\n\n"
                        return
                    except Exception as e:
                        logger.error(f"VolcEngine translation error: {e}")
                        yield f"data: {json.dumps({'error': str(e)})}\n\n"
                        return

                # 传递prompt_config到translate方法（如果翻译器支持）
                kwargs = {"stream": True}
                if prompt_config:
                    kwargs["prompt_config"] = prompt_config

                response_stream = translator_instance.translate(
                    text_to_translate, target_lang, source_lang, **kwargs
                )

                if not isinstance(response_stream, requests.Response):
                    error_msg = f"Translation API for {api_platform} did not return a valid stream. Type: {type(response_stream)}. Content: {str(response_stream)[:200]}"
                    logger.error(error_msg)
                    yield f"data: {json.dumps({'error': error_msg})}\n\n"
                    return

                for chunk in response_stream.iter_lines():
                    if chunk:
                        chunk_str = chunk.decode("utf-8")
                        if chunk_str.startswith("data: "):
                            json_part_str = chunk_str[len("data: ") :].strip()
                        elif chunk_str.strip().startswith(
                            "{"
                        ) and chunk_str.strip().endswith("}"):
                            json_part_str = chunk_str.strip()
                            logger.debug(
                                f"Received non-standard JSON line, processing as data: {json_part_str}"
                            )
                        else:
                            logger.debug(
                                f"Skipping non-data line from stream: {chunk_str}"
                            )
                            continue

                        if json_part_str == "[DONE]":
                            logger.info("Stream finished with [DONE] marker from API.")
                            yield f"data: {json.dumps({'done': True})}\n\n"
                            break

                        try:
                            json_part = json.loads(json_part_str)
                            if (
                                json_part.get("choices")
                                and isinstance(json_part["choices"], list)
                                and len(json_part["choices"]) > 0
                                and json_part["choices"][0].get("delta")
                                and "content" in json_part["choices"][0]["delta"]
                            ):
                                content = json_part["choices"][0]["delta"]["content"]
                                if content is not None:
                                    yield f"data: {json.dumps({'text_chunk': content})}\n\n"
                            elif json_part.get("error"):
                                error_detail = (
                                    json_part["error"].get(
                                        "message", str(json_part["error"])
                                    )
                                    if isinstance(json_part["error"], dict)
                                    else str(json_part["error"])
                                )
                                logger.error(
                                    f"Error in stream from API ({api_platform}): {error_detail}"
                                )
                                yield f"data: {json.dumps({'error': error_detail})}\n\n"
                                break
                            elif json_part.get("done") is True:
                                logger.info(
                                    f"Stream finished with 'done: true' marker from API ({api_platform})."
                                )
                                yield f"data: {json.dumps({'done': True})}\n\n"
                                break
                        except json.JSONDecodeError:
                            logger.warning(
                                f"Could not decode JSON from stream chunk ({api_platform}): {json_part_str}"
                            )
                            continue
            except requests.exceptions.RequestException as e:
                logger.error(
                    f"RequestException during streaming to {api_platform} API: {e}"
                )
                yield f"data: {json.dumps({'error': f'API request error: {e}'})}\n\n"
            except Exception as e:
                import traceback

                logger.error(
                    f"Unexpected error in generate_translation_stream ({api_platform}): {e}\n{traceback.format_exc()}"
                )
                yield f"data: {json.dumps({'error': f'Server error during streaming: {str(e)}'})}\n\n"
            finally:
                if response_stream and hasattr(response_stream, "close"):
                    response_stream.close()
                logger.info(f"Translation stream generation ended for {api_platform}.")

        return Response(generate_translation_stream(), mimetype="text/event-stream")

    # --- File Input Translation ---
    elif "file" in request.files:
        file = request.files["file"]
        if not file or file.filename == "":
            logger.error("No file selected or file object missing.")
            return jsonify({"error": "No file selected."}), 400

        if not allowed_file(file.filename):
            logger.error(f"File type not allowed: {file.filename}")
            return (
                jsonify(
                    {
                        "error": "File type not allowed. Please upload .txt, .docx, .pptx, .xlsx, or .xls files."
                    }
                ),
                400,
            )

        # 新增：处理模板文件上传
        template_file_path = None
        if "template_file" in request.files:
            template_file = request.files["template_file"]
            if (
                template_file
                and template_file.filename != ""
                and allowed_file(template_file.filename)
            ):
                template_filename = secure_filename(template_file.filename)
                if template_filename.endswith(".docx"):  # 只允许docx模板
                    template_unique_name = (
                        f"template_{str(uuid.uuid4())}_{template_filename}"
                    )
                    template_file_path = os.path.join(
                        app.config["TEMPLATE_FOLDER"], template_unique_name
                    )
                    try:
                        template_file.save(template_file_path)
                        logger.info(
                            f"Template file uploaded: {template_filename} to {template_file_path}"
                        )
                    except Exception as e:
                        logger.error(f"Failed to save template file: {e}")
                        return (
                            jsonify(
                                {"error": f"Failed to save template file: {str(e)}"}
                            ),
                            500,
                        )
                else:
                    logger.warning("Template file must be a .docx file, ignoring")
            else:
                logger.info("No valid template file provided")

        raw_filename = file.filename
        logger.info(f"Original filename: {raw_filename}")
        
        # 安全地获取文件扩展名（从原始文件名）
        if "." in raw_filename:
            file_extension = raw_filename.rsplit(".", 1)[1].lower()
            filename_without_ext = raw_filename.rsplit(".", 1)[0]
        else:
            logger.error(f"File has no extension: {raw_filename}")
            return jsonify({"success": False, "message": "文件必须有扩展名"})
        
        # 处理纯中文文件名问题
        safe_filename = secure_filename(raw_filename)
        logger.info(f"Secure filename: {safe_filename}")
        
        # 如果secure_filename过滤后文件名为空或没有扩展名，生成新文件名
        if not safe_filename or "." not in safe_filename:
            import time
            new_filename_base = f"file_{int(time.time())}"
            logger.info(f"Chinese filename detected, using generated name: {new_filename_base}")
        else:
            new_filename_base = safe_filename.rsplit(".", 1)[0]
        unique_filename_base = str(uuid.uuid4())
        input_unique_filename = unique_filename_base + "." + file_extension
        input_filepath = os.path.join(
            app.config["UPLOAD_FOLDER"], input_unique_filename
        )

        try:
            file.save(input_filepath)
            logger.info(f"File uploaded: {original_filename} to {input_filepath}")
        except Exception as e:
            logger.error(f"Failed to save uploaded file {original_filename}: {e}")
            return jsonify({"error": f"Failed to save uploaded file: {str(e)}"}), 500

        user_output_folder = data.get("output_folder_path", "").strip()
        actual_output_dir = app.config["TRANSLATED_FOLDER"]
        if user_output_folder:
            abs_user_output_folder = os.path.abspath(user_output_folder)
            try:
                os.makedirs(abs_user_output_folder, exist_ok=True)
                if os.path.isdir(abs_user_output_folder):
                    actual_output_dir = abs_user_output_folder
                else:
                    logger.warning(
                        f"User output folder '{user_output_folder}' is not a valid directory, using default."
                    )
            except Exception as e:
                logger.warning(
                    f"Could not use/create user output folder '{user_output_folder}': {e}. Using default."
                )

        logger.info(f"Translations will be saved to: {actual_output_dir}")

        translated_filepath_or_error = None

        logger.info(f"Processing file '{original_filename}'")

        if file_extension == "txt":
            encoding = data.get("encoding", "utf-8")
            logger.info(f"Translating TXT file with encoding '{encoding}'.")

            # 检查translate_text_file是否支持prompt_config参数
            if check_function_supports_prompt_config(translate_text_file):
                logger.info("translate_text_file supports prompt_config parameter")
                translated_filepath_or_error = translate_text_file(
                    input_filepath,
                    actual_output_dir,
                    target_lang,
                    translator_instance,
                    source_lang,
                    encoding,
                    unique_filename_base,
                    prompt_config=prompt_config,
                )
            else:
                logger.info(
                    "translate_text_file does not support prompt_config parameter, using translator instance with applied config"
                )
                translated_filepath_or_error = translate_text_file(
                    input_filepath,
                    actual_output_dir,
                    target_lang,
                    translator_instance,
                    source_lang,
                    encoding,
                    unique_filename_base,
                )

        elif file_extension == "docx":
            docx_method = data.get("docx_method")
            if not docx_method:
                logger.error("DOCX method not specified by client for .docx file.")
                if os.path.exists(input_filepath):
                    os.remove(input_filepath)
                if template_file_path and os.path.exists(template_file_path):
                    os.remove(template_file_path)
                return (
                    jsonify(
                        {
                            "error": "DOCX processing method must be selected for .docx files."
                        }
                    ),
                    400,
                )

            logger.info(f"Selected DOCX processing method: {docx_method}")
            if template_file_path:
                logger.info(f"Using template file: {template_file_path}")

            try:
                # 根据不同的DOCX处理方法调用相应的翻译器
                if docx_method == "docx_translator":
                    logger.info(
                        f"Attempting basic DOCX translation (via docx_translator.py)."
                    )

                    # 检查是否支持prompt_config参数
                    if check_function_supports_prompt_config(translate_docx_file):
                        translated_filepath_or_error = translate_docx_file(
                            input_filepath,
                            actual_output_dir,
                            target_lang,
                            translator_instance,
                            source_lang,
                            unique_filename_base,
                            prompt_config=prompt_config,
                        )
                    else:
                        logger.info(
                            "translate_docx_file does not support prompt_config, using configured translator"
                        )
                        translated_filepath_or_error = translate_docx_file(
                            input_filepath,
                            actual_output_dir,
                            target_lang,
                            translator_instance,
                            source_lang,
                            unique_filename_base,
                        )

                elif docx_method == "docx_full_translator":
                    logger.info(
                        f"Attempting formatted DOCX translation (via docx_full_translator.py)."
                    )
                    if not _translate_docx_file_formatted_available:
                        translated_filepath_or_error = "Error: Formatted DOCX translator module (docx_full_translator.py) is not available on server."
                    else:
                        # 检查是否支持prompt_config参数
                        if check_function_supports_prompt_config(
                            translate_docx_file_formatted
                        ):
                            translated_filepath_or_error = (
                                translate_docx_file_formatted(
                                    input_filepath,
                                    actual_output_dir,
                                    target_lang,
                                    translator_instance,
                                    source_lang,
                                    unique_filename_base,
                                    prompt_config=prompt_config,
                                )
                            )
                        else:
                            translated_filepath_or_error = (
                                translate_docx_file_formatted(
                                    input_filepath,
                                    actual_output_dir,
                                    target_lang,
                                    translator_instance,
                                    source_lang,
                                    unique_filename_base,
                                )
                            )

                elif docx_method == "docx_pythondoc1_translator":
                    logger.info(
                        f"Attempting PythonDoc1 DOCX translation (via docx_pythondoc1_translator.py)."
                    )
                    if not _translate_docx_pythondoc1_available:
                        translated_filepath_or_error = "Error: PythonDoc1 translator module (docx_pythondoc1_translator.py) is not available on server."
                    else:
                        # 检查是否支持prompt_config参数
                        if check_function_supports_prompt_config(
                            translate_docx_pythondoc1
                        ):
                            translated_filepath_or_error = translate_docx_pythondoc1(
                                input_filepath,
                                actual_output_dir,
                                target_lang,
                                translator_instance,
                                source_lang,
                                unique_filename_base,
                                reference_doc=template_file_path,
                                prompt_config=prompt_config,
                            )
                        else:
                            translated_filepath_or_error = translate_docx_pythondoc1(
                                input_filepath,
                                actual_output_dir,
                                target_lang,
                                translator_instance,
                                source_lang,
                                unique_filename_base,
                                reference_doc=template_file_path,
                            )

                elif docx_method == "docx_pythondoc2_translator":
                    logger.info(
                        f"Attempting PythonDoc2 DOCX translation (via docx_pythondoc2_translator.py)."
                    )
                    if not _translate_docx_pythondoc2_available:
                        translated_filepath_or_error = "Error: PythonDoc2 translator module (docx_pythondoc2_translator.py) is not available on server."
                    else:
                        # 检查是否支持prompt_config参数
                        if check_function_supports_prompt_config(
                            translate_docx_pythondoc2
                        ):
                            translated_filepath_or_error = translate_docx_pythondoc2(
                                input_filepath,
                                actual_output_dir,
                                target_lang,
                                translator_instance,
                                source_lang,
                                unique_filename_base,
                                reference_doc=template_file_path,
                                prompt_config=prompt_config,
                            )
                        else:
                            translated_filepath_or_error = translate_docx_pythondoc2(
                                input_filepath,
                                actual_output_dir,
                                target_lang,
                                translator_instance,
                                source_lang,
                                unique_filename_base,
                                reference_doc=template_file_path,
                            )

                elif docx_method == "docx_markdown_translator":
                    logger.info(
                        f"Attempting Markdown-based DOCX translation with custom prompt support."
                    )
                    if not _translate_docx_markdown_available:
                        translated_filepath_or_error = "Error: Markdown translator module (docx_markdown_translator.py) is not available on server."
                    else:
                        # 准备参数
                        kwargs = {
                            "mode": (
                                "optimized_with_prompt"
                                if prompt_config
                                else "optimized"
                            ),
                            "max_chunk_size": (
                                prompt_config.get("max_chars_per_chunk", 30000)
                                if prompt_config
                                else 30000
                            ),
                            "max_units_per_chunk": (
                                prompt_config.get("max_units_per_chunk", 50)
                                if prompt_config
                                else 50
                            ),
                            "template_path": template_file_path,  # 传递模板路径
                            "prompt_config": prompt_config,  # 明确传递prompt配置
                        }

                        # 添加prompt相关参数
                        if prompt_config:
                            if prompt_config.get("mode") == "professional":
                                kwargs["prompt_template"] = prompt_config.get(
                                    "prompt_template", "default"
                                )
                            elif prompt_config.get(
                                "mode"
                            ) == "custom" and prompt_config.get("custom_prompt"):
                                kwargs["custom_prompt"] = prompt_config["custom_prompt"]

                            # 添加术语管理
                            if prompt_config.get("preserve_terms"):
                                kwargs["preserve_terms"] = prompt_config[
                                    "preserve_terms"
                                ]
                            if prompt_config.get("glossary"):
                                kwargs["glossary"] = prompt_config["glossary"]
                            if prompt_config.get("additional_context"):
                                kwargs["additional_context"] = prompt_config[
                                    "additional_context"
                                ]

                        translated_filepath_or_error = translate_docx_via_markdown(
                            input_filepath,
                            actual_output_dir,
                            target_lang,
                            translator_instance,
                            source_lang,
                            unique_filename_base,
                            **kwargs,
                        )
                else:
                    logger.error(
                        f"Unknown or unsupported docx_method received: {docx_method}"
                    )
                    translated_filepath_or_error = f"Error: Unknown DOCX processing method '{docx_method}' specified."

            finally:
                # 清理模板文件
                if template_file_path and os.path.exists(template_file_path):
                    try:
                        os.remove(template_file_path)
                        logger.info(
                            f"Temporary template file removed: {template_file_path}"
                        )
                    except Exception as e:
                        logger.warning(
                            f"Failed to remove template file {template_file_path}: {e}"
                        )

        # 清理上传的临时文件
        try:
            if os.path.exists(input_filepath):
                os.remove(input_filepath)
                logger.info(f"Temporary uploaded file removed: {input_filepath}")
        except Exception as e:
            logger.warning(
                f"Failed to remove temporary uploaded file {input_filepath}: {e}"
            )

        # 处理翻译结果
        if isinstance(
            translated_filepath_or_error, str
        ) and not translated_filepath_or_error.lower().startswith("error:"):
            if not os.path.exists(translated_filepath_or_error):
                logger.error(
                    f"Translated file path reported ('{translated_filepath_or_error}') but file not found on server."
                )
                return (
                    jsonify(
                        {
                            "error": "Translated file not found on server after processing."
                        }
                    ),
                    500,
                )

            output_filename = os.path.basename(translated_filepath_or_error)
            final_downloadable_path_in_translated_folder = os.path.join(
                app.config["TRANSLATED_FOLDER"], output_filename
            )

            if os.path.abspath(translated_filepath_or_error) != os.path.abspath(
                final_downloadable_path_in_translated_folder
            ):
                try:
                    shutil.move(
                        translated_filepath_or_error,
                        final_downloadable_path_in_translated_folder,
                    )
                    logger.info(
                        f"Moved translated file from '{translated_filepath_or_error}' to download folder '{final_downloadable_path_in_translated_folder}'"
                    )
                except Exception as e:
                    logger.error(
                        f"Failed to move translated file to download folder: {e}"
                    )
                    return (
                        jsonify(
                            {
                                "error": f"File translated but failed to stage for download: {e}"
                            }
                        ),
                        500,
                    )

            output_file_url = f"/download/{output_filename}"
            logger.info(
                f"File translation successful. URL: {output_file_url}, Path in download folder: {final_downloadable_path_in_translated_folder}"
            )
            return jsonify({"translated_file_url": output_file_url})
        else:
            error_message = (
                str(translated_filepath_or_error)
                if translated_filepath_or_error
                else "Unknown error during file translation."
            )
            logger.error(f"File translation failed: {error_message}")
            return jsonify({"error": error_message}), 500
    else:
        logger.error("No valid text or file input provided to /translate_api.")
        return jsonify({"error": "No valid text or file input provided."}), 400


@app.route("/download/<filename>")
def download_file(filename):
    if ".." in filename or filename.startswith("/"):
        logger.warning(f"Attempted directory traversal in download: {filename}")
        return jsonify({"error": "Invalid filename for download."}), 400

    try:
        logger.info(
            f"Attempting to send file for download: {filename} from {app.config['TRANSLATED_FOLDER']}"
        )
        return send_from_directory(
            app.config["TRANSLATED_FOLDER"], filename, as_attachment=True
        )
    except FileNotFoundError:
        logger.error(f"File not found for download in TRANSLATED_FOLDER: {filename}")
        return jsonify({"error": "File not found for download."}), 404
    except Exception as e:
        logger.error(f"Error during file download ({filename}): {e}")
        return jsonify({"error": "Could not download file due to server error."}), 500


# ==================== Prompt和配置管理API ====================

if _prompt_modules_available:
    # 获取所有prompt模板
    @app.route("/api/prompt-templates", methods=["GET"])
    def get_prompt_templates_api():
        """返回所有可用的prompt模板"""
        templates = get_all_templates()
        return jsonify(templates)

    # 获取特定模板详情
    @app.route("/api/prompt-templates/<template_name>", methods=["GET"])
    def get_prompt_template_detail(template_name):
        """返回特定模板的详细信息"""
        template = get_template(template_name)
        if template:
            return jsonify(template)
        return jsonify({"error": "Template not found"}), 404

    # 保存用户配置
    @app.route("/api/save-config", methods=["POST"])
    @login_required
    def save_user_config():
        """保存用户的翻译配置"""
        data = request.json
        config_name = data.get("name", f"config_{int(time.time())}")

        if config_manager.save_config(config_name, data):
            return jsonify(
                {"success": True, "message": f"Configuration saved as {config_name}"}
            )
        return jsonify({"error": "Failed to save configuration"}), 500

    # 加载用户配置
    @app.route("/api/load-config/<config_name>", methods=["GET"])
    def load_user_config(config_name):
        """加载用户保存的配置"""
        config = config_manager.load_config(config_name)
        if config:
            return jsonify(config)
        return jsonify({"error": "Configuration not found"}), 404

    # 列出所有配置
    @app.route("/api/list-configs", methods=["GET"])
    def list_user_configs():
        """列出所有保存的配置"""
        configs = config_manager.list_configs()
        return jsonify({"configs": configs})

    # 删除配置
    @app.route("/api/delete-config/<config_name>", methods=["DELETE"])
    def delete_user_config(config_name):
        """删除用户配置"""
        if config_manager.delete_config(config_name):
            return jsonify(
                {"success": True, "message": f"Configuration {config_name} deleted"}
            )
        return jsonify({"error": "Failed to delete configuration"}), 500

    # 导出配置
    @app.route("/api/export-config/<config_name>", methods=["GET"])
    def export_user_config(config_name):
        """导出配置为JSON"""
        config_json = config_manager.export_config(config_name)
        if config_json:
            return Response(
                config_json,
                mimetype="application/json",
                headers={
                    "Content-Disposition": f"attachment; filename={config_name}.json"
                },
            )
        return jsonify({"error": "Configuration not found"}), 404

    # 导入配置
    @app.route("/api/import-config", methods=["POST"])
    def import_user_config():
        """导入配置"""
        data = request.json
        config_json = data.get("config_json")
        new_name = data.get("name")

        if not config_json:
            return jsonify({"error": "No configuration data provided"}), 400

        if config_manager.import_config(config_json, new_name):
            return jsonify(
                {"success": True, "message": "Configuration imported successfully"}
            )
        return jsonify({"error": "Failed to import configuration"}), 500

    # 保存自定义prompt
    @app.route("/api/save-prompt", methods=["POST"])
    def save_custom_prompt():
        """保存用户自定义的prompt"""
        data = request.json
        prompt_name = data.get("name")
        prompt = {"system": data.get("system"), "user": data.get("user", "{content}")}

        if not prompt_name or not prompt["system"]:
            return jsonify({"error": "Prompt name and system prompt are required"}), 400

        if config_manager.save_user_prompt(prompt_name, prompt):
            return jsonify(
                {"success": True, "message": f"Prompt saved as {prompt_name}"}
            )
        return jsonify({"error": "Failed to save prompt"}), 500

    # 获取用户自定义prompts
    @app.route("/api/user-prompts", methods=["GET"])
    def get_user_prompts():
        """获取所有用户自定义的prompts"""
        prompts = config_manager.get_user_prompts()
        return jsonify(prompts)

    # 删除用户自定义prompt
    @app.route("/api/delete-prompt/<prompt_name>", methods=["DELETE"])
    def delete_user_prompt(prompt_name):
        """删除用户自定义prompt"""
        if config_manager.delete_user_prompt(prompt_name):
            return jsonify(
                {"success": True, "message": f"Prompt {prompt_name} deleted"}
            )
        return jsonify({"error": "Failed to delete prompt"}), 500

    # 构建自定义prompt
    @app.route("/api/build-prompt", methods=["POST"])
    def build_custom_prompt():
        """使用PromptBuilder构建自定义prompt"""
        data = request.json
        prompt_type = data.get("type", "domain_specific")

        try:
            if prompt_type == "domain_specific":
                prompt = PromptBuilder.create_domain_specific_prompt(
                    domain=data.get("domain", "general"),
                    style=data.get("style", "formal"),
                    preserve_terms=data.get("preserve_terms"),
                    glossary=data.get("glossary"),
                    additional_rules=data.get("additional_rules"),
                    examples=data.get("examples"),
                )
            elif prompt_type == "context_aware":
                prompt = PromptBuilder.create_context_aware_prompt(
                    document_type=data.get("document_type", "document"),
                    target_audience=data.get("target_audience", "general"),
                    tone=data.get("tone", "neutral"),
                    context=data.get("context", ""),
                    cultural_adaptation=data.get("cultural_adaptation", True),
                )
            elif prompt_type == "file_format_specific":
                prompt = PromptBuilder.create_file_format_specific_prompt(
                    file_type=data.get("file_type", "text"),
                    preserve_formatting=data.get("preserve_formatting", True),
                    handle_special_elements=data.get("handle_special_elements"),
                )
            elif prompt_type == "instruction_based":
                prompt = PromptBuilder.create_instruction_based_prompt(
                    instructions=data.get("instructions", []),
                    examples=data.get("examples"),
                    constraints=data.get("constraints"),
                )
            else:
                return jsonify({"error": "Unknown prompt type"}), 400

            return jsonify(prompt)
        except Exception as e:
            logger.error(f"Error building prompt: {e}")
            return jsonify({"error": str(e)}), 500

    # 获取默认配置
    @app.route("/api/default-config", methods=["GET"])
    def get_default_config():
        """获取默认配置"""
        config = config_manager.load_default_config()
        return jsonify(config)

    # 保存默认配置
    @app.route("/api/default-config", methods=["POST"])
    def save_default_config():
        """保存默认配置"""
        data = request.json
        if config_manager.save_default_config(data):
            return jsonify({"success": True, "message": "Default configuration saved"})
        return jsonify({"error": "Failed to save default configuration"}), 500

# 添加获取默认批处理配置的API端点
@app.route("/api/default-batch-config", methods=["GET"])
def get_default_batch_config():
    """获取默认的批处理配置"""
    try:
        # 从翻译器模块导入默认值
        from translators.file_translator import DEFAULT_CHUNK_SIZE, DEFAULT_MAX_CHARS
        
        return jsonify({
            "max_chars_per_chunk": DEFAULT_MAX_CHARS,  # 从翻译器读取
            "max_units_per_chunk": DEFAULT_CHUNK_SIZE   # 从翻译器读取
        })
    except ImportError as e:
        # 如果导入失败，使用备用默认值
        logger.warning(f"Failed to import default values from file_translator: {e}")
        return jsonify({
            "max_chars_per_chunk": 2000,  # 备用默认值
            "max_units_per_chunk": 10     # 备用默认值
        })


    @app.route('/static/<path:filename>')
    def static_files(filename):
        return send_from_directory('static', filename)
    
    @app.route('/manifest.json')
    def manifest():
        return send_from_directory('static', 'manifest.json')


# ==================== 路由添加结束 ====================

import socket
import sys
import os

def is_port_in_use(port, host='127.0.0.1'):
    """检查端口是否被占用"""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        return s.connect_ex((host, port)) == 0

if __name__ == "__main__":
    try:
        # Hugging Face Spaces 部署模式
        port = int(os.environ.get("PORT", 7860))  # Hugging Face 默认使用7860端口
        
        print("=" * 70)
        print(f"🚀 Translator 应用启动中...")
        print(f"🔌 端口: {port}")
        print(f"🌐 Hugging Face Spaces 部署模式")
        print("=" * 70)
        
        # 启动服务器 - Hugging Face Spaces 配置
        app.run(
            debug=False, 
            host="0.0.0.0", 
            port=port, 
            use_reloader=False
        )
        
    except Exception as e:
        print(f"启动失败: {e}")
        sys.exit(1)
    except KeyboardInterrupt:
        print("\n服务器已停止")
        sys.exit(0)
