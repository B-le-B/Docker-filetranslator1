# config.py
import os

# 这个文件现在旨在提供“最后的、全局的”后备默认值。
# 它不应该主动从 .env 加载特定平台的密钥，这个职责由 app.py 统一处理。

# API Key 的全局后备值：
# 通常，对于 API Key，不应该有一个“通用”的有效密钥作为后备。
# 将其设置为 None 是最安全的，这样如果所有其他配置来源（前端、.env、平台特定默认）
# 都未能提供 API Key，app.py 中的最终检查会捕获到 API Key 缺失并报错。
# 如果你确实有一个用于测试或绝对紧急情况的通用测试Key（不推荐用于生产），可以在这里设置，
# 但更好的做法是让它为 None。
API_KEY = None
# 或者，如果你想从一个非常通用的环境变量名中获取最后的后备Key（如果.env中定义了它）：
# API_KEY = os.getenv("ULTIMATE_FALLBACK_API_KEY_IF_ANY")


# API 基本 URL 的全局后备值：
# 这可以是一个非常通用的、你认为合适的默认值，或者也设为 None。
# 例如，如果你没有任何特定平台的配置，可能会尝试这个。
BASE_URL = "https://api.example-default.com/v1" # 只是一个示例，请替换或设为 None
# 或者从一个通用的环境变量获取：
# BASE_URL = os.getenv("ULTIMATE_FALLBACK_BASE_URL", "https://api.example-default.com/v1")


# 默认使用的模型的全局后备值：
# 同上，这可以是一个非常通用的模型名称。
DEFAULT_MODEL = "default-fallback-model" # 只是一个示例，请替换或设为 None
# 或者从一个通用的环境变量获取：
# DEFAULT_MODEL = os.getenv("ULTIMATE_FALLBACK_MODEL_NAME", "default-fallback-model")


# 注意：这里不再调用 load_dotenv()。
# load_dotenv() 应该只在你的主应用程序文件 (app.py) 的顶部调用一次。
# 这里也不再有针对 SILICONFLOW_API_KEY 等特定环境变量的 os.getenv() 调用。

# 如果需要在 app.py 中明确区分这些是从 config.py 来的后备值，
# 你可以在 app.py 中导入时使用别名，例如：
# from config import API_KEY as CONFIG_FALLBACK_API_KEY
# 但你当前的 app.py 逻辑是直接使用这些变量名作为后备，所以保持原样即可。

# 你可以选择在启动时打印这些后备值以供参考（主要用于开发阶段）
# print(f"CONFIG.PY: Fallback API_KEY = {API_KEY}")
# print(f"CONFIG.PY: Fallback BASE_URL = {BASE_URL}")
# print(f"CONFIG.PY: Fallback DEFAULT_MODEL = {DEFAULT_MODEL}")
