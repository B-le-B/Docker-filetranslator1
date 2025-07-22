# translators/volcengine_native_translator.py (Corrected for actual SDK structure)

import os
import logging
from volcenginesdktranslate20250301 import TRANSLATE20250301Api, TranslateTextRequest
from volcenginesdkcore import ApiClient, Configuration

logger = logging.getLogger(__name__)


class VolcEngineNativeTranslator:
    """
    使用火山引擎 volcengine-python-sdk 进行翻译的翻译器。
    """

    # 语言代码映射 - 将常见的语言名称转换为火山引擎支持的代码
    LANGUAGE_MAP = {
        # 英文
        "english": "en",
        "en": "en",
        "eng": "en",
        # 中文
        "chinese": "zh",
        "zh": "zh",
        "zh-cn": "zh",
        "zh-hans": "zh",
        "simplified chinese": "zh",
        "中文": "zh",
        "简体中文": "zh",
        # 日文
        "japanese": "ja",
        "ja": "ja",
        "jp": "ja",
        "jpn": "ja",
        "日语": "ja",
        "日文": "ja",
        # 韩文
        "korean": "ko",
        "ko": "ko",
        "kor": "ko",
        "韩语": "ko",
        "韩文": "ko",
        # 法文
        "french": "fr",
        "fr": "fr",
        "fra": "fr",
        "法语": "fr",
        "法文": "fr",
        # 德文
        "german": "de",
        "de": "de",
        "deu": "de",
        "德语": "de",
        "德文": "de",
        # 西班牙文
        "spanish": "es",
        "es": "es",
        "spa": "es",
        "西班牙语": "es",
        # 俄文
        "russian": "ru",
        "ru": "ru",
        "rus": "ru",
        "俄语": "ru",
        "俄文": "ru",
        # 意大利文
        "italian": "it",
        "it": "it",
        "ita": "it",
        "意大利语": "it",
        # 葡萄牙文
        "portuguese": "pt",
        "pt": "pt",
        "por": "pt",
        "葡萄牙语": "pt",
        # 阿拉伯文
        "arabic": "ar",
        "ar": "ar",
        "ara": "ar",
        "阿拉伯语": "ar",
        # 泰文
        "thai": "th",
        "th": "th",
        "tha": "th",
        "泰语": "th",
        # 越南文
        "vietnamese": "vi",
        "vi": "vi",
        "vie": "vi",
        "越南语": "vi",
        # 自动检测
        "auto": "auto",
        "automatic": "auto",
        "自动": "auto",
        "自动检测": "auto",
    }

    def __init__(self, ak: str, sk: str, region: str = "cn-beijing"):
        if not ak or not sk:
            raise ValueError("Access Key (ak) and Secret Key (sk) cannot be empty.")

        self.ak = ak
        self.sk = sk
        self.region = region

        try:
            # 初始化配置
            configuration = Configuration()
            configuration.ak = self.ak
            configuration.sk = self.sk
            configuration.region = self.region

            # 创建API客户端
            api_client = ApiClient(configuration)
            self.api_instance = TRANSLATE20250301Api(api_client)

            logger.info(f"火山引擎翻译服务初始化成功，区域: {self.region}")
        except Exception as e:
            logger.error(f"Failed to initialize VolcEngine TranslationService: {e}")
            raise ConnectionError(
                f"Could not initialize VolcEngine SDK. Check credentials and SDK version. Error: {e}"
            )

    def _normalize_language_code(self, lang: str) -> str:
        """将语言代码标准化为火山引擎支持的格式"""
        if not lang:
            return "auto"

        # 转换为小写并去除空格
        normalized = lang.lower().strip()

        # 查找映射
        mapped_lang = self.LANGUAGE_MAP.get(normalized)
        if mapped_lang:
            logger.debug(f"Language mapping: '{lang}' -> '{mapped_lang}'")
            return mapped_lang

        # 如果没有找到映射，返回原始值（可能已经是正确的代码）
        logger.warning(f"No language mapping found for '{lang}', using as-is")
        return normalized

    def translate(
        self, text: str, target_lang: str, source_lang: str = "auto", **kwargs
    ) -> str:
        """
        执行翻译。此方法是阻塞的，不支持流式返回。
        """
        if not text:
            return ""

        # 标准化语言代码
        normalized_target = self._normalize_language_code(target_lang)
        normalized_source = (
            self._normalize_language_code(source_lang) if source_lang else "auto"
        )

        logger.info(
            f"Original languages - Source: '{source_lang}' Target: '{target_lang}'"
        )
        logger.info(
            f"Normalized languages - Source: '{normalized_source}' Target: '{normalized_target}'"
        )

        try:
            # 正确的方式：在构造时传入所有必需参数
            request_params = {"target_language": normalized_target, "text_list": [text]}

            # 只有当源语言不是 auto 时才设置
            if normalized_source and normalized_source != "auto":
                request_params["source_language"] = normalized_source

            logger.info(f"Creating request with params: {request_params}")

            # 创建翻译请求 - 传入所有必需参数
            request = TranslateTextRequest(**request_params)

            logger.info(f"Request created successfully")
            logger.info(
                f"Calling VolcEngine native API. TargetLang: {normalized_target}, SourceLang: {normalized_source}"
            )

            # 调用翻译API
            response = self.api_instance.translate_text(request)

            logger.info(f"API response received: {type(response)}")

            # 处理响应
            if (
                response
                and hasattr(response, "translation_list")
                and response.translation_list
            ):
                translated_text = response.translation_list[0].translation
                logger.info("Successfully translated text via VolcEngine native API.")
                return translated_text
            else:
                logger.error(
                    f"Received unexpected response from VolcEngine: {response}"
                )
                raise Exception("Unexpected response structure from VolcEngine API.")

        except Exception as e:
            logger.error(f"Error during VolcEngine native translation: {e}")
            raise

    def set_prompt_config(self, prompt_config):
        logger.warning(
            "VolcEngineNativeTranslator does not support custom prompt configurations."
        )
        pass
