# translator.py
import requests
import json
import logging
from typing import Optional, List, Dict, Any

try:
    from config import API_KEY as FALLBACK_API_KEY_CONFIG, \
                       BASE_URL as FALLBACK_BASE_URL_CONFIG, \
                       DEFAULT_MODEL as FALLBACK_MODEL_CONFIG
except ImportError:
    FALLBACK_API_KEY_CONFIG = None
    FALLBACK_BASE_URL_CONFIG = None
    FALLBACK_MODEL_CONFIG = None

logger = logging.getLogger(__name__)

class SiliconFlowTranslator:
    def __init__(self, api_key: str, base_url: str, model: str, platform_id: Optional[str] = None, extra_body: Optional[Dict[str, Any]] = None):
        if not api_key:
            raise ValueError("API Key must be provided for translator initialization.")
        if not base_url:
            raise ValueError("Base URL must be provided for translator initialization.")
        if not model:
            raise ValueError("Model must be provided for translator initialization.")

        self.api_key = api_key
        self.base_url = base_url
        self.model = model
        self.platform_id = platform_id if platform_id else self._infer_platform_from_url(base_url)
        
        # prompt配置
        self.prompt_config = None
        self.default_merge_strategy = 'merge'

        # 添加这一行：
        self.extra_body = extra_body or {}
        
        # 专业领域模板 - 与前端选项对应
        self.professional_templates = {
            'academic': {
                'system': """You are an expert academic translator specializing in scholarly texts.
Maintain academic tone, preserve citations and references, and use appropriate academic terminology.
Ensure consistency in technical terms throughout the translation. Pay attention to:
- Academic writing style and formal tone
- Proper citation formats
- Technical and disciplinary terminology
- Research methodology descriptions""",
                'user': "Translate this academic text to {target_lang}, maintaining scholarly tone and precision:\n\n{content}"
            },
            'business': {
                'system': """You are a professional business translator with expertise in corporate communications.
Use appropriate business terminology, maintain formal tone, and keep company names/brands unchanged.
Ensure clarity and professionalism in the translation. Focus on:
- Business terminology and corporate language
- Professional communication style
- Financial and commercial concepts
- Maintaining brand consistency""",
                'user': "Translate this business document to {target_lang}, preserving professional tone:\n\n{content}"
            },
            'technical': {
                'system': """You are a technical translator specializing in technical documentation.
Preserve technical accuracy, keep code snippets and commands unchanged, and use industry-standard terminology.
Maintain consistency in technical terms throughout. Pay attention to:
- Technical specifications and procedures
- Software and hardware terminology
- Programming concepts and code examples
- Industry standards and protocols""",
                'user': "Translate this technical content to {target_lang}, preserving all technical elements:\n\n{content}"
            },
            'legal': {
                'system': """You are a certified legal translator with expertise in legal documents.
Use precise legal terminology, maintain legal accuracy and formality, and preserve all legal references.
Ensure no ambiguity in legal terms. Focus on:
- Legal terminology and concepts
- Contractual language and clauses
- Regulatory and compliance terms
- Jurisdictional considerations""",
                'user': "Translate this legal document to {target_lang}, maintaining legal precision:\n\n{content}"
            },
            'medical': {
                'system': """You are a certified medical translator with expertise in medical documents.
Use standard medical terminology, preserve drug names and dosages exactly, and maintain clinical precision.
Follow international medical nomenclature standards. Pay attention to:
- Medical terminology and procedures
- Drug names and dosages
- Clinical protocols and guidelines
- Patient safety considerations""",
                'user': "Translate this medical content to {target_lang}, preserving medical accuracy:\n\n{content}"
            },
            'creative': {
                'system': """You are a creative translator focusing on maintaining style and tone.
Preserve the original style, adapt idioms naturally, and maintain emotional impact.
Focus on readability and flow while being faithful to the original meaning. Consider:
- Literary devices and stylistic elements
- Cultural adaptation of idioms
- Emotional tone and atmosphere
- Creative expression and voice""",
                'user': "Translate this creative text to {target_lang}, preserving style and emotional impact:\n\n{content}"
            }
        }

        # 默认模板
        self.default_templates = {
            'general': {
                'system': "You are a professional translator. Provide accurate, natural translations while preserving the original meaning and tone.Do not include any extra comments in your output.",
                'user': "Translate the following text to {target_lang}:\n\n{content}"
            }
        }

        logger.info(f"Translator initialized for platform: {self.platform_id}, Base URL: {self.base_url}, Model: {self.model}, API Key: {'SET' if self.api_key else 'NOT SET'}")

    def _infer_platform_from_url(self, url: str) -> str:
        if "api.siliconflow.cn" in url: return "siliconflow"
        if "api-inference.modelscope.cn" in url: return "modelscope"
        if "openrouter.ai" in url: return "openrouter"
        if "api.openai.com" in url: return "openai"
        if "api.deepseek.com" in url: return "deepseek"
        if "api.moonshot.cn" in url: return "moonshot"
        if "localhost:11434" in url: return "ollama"
        return "custom"

    def _get_headers(self) -> dict:
        common_headers = {
            "Content-Type": "application/json",
            "Accept": "application/json"
        }

        if self.platform_id == "modelscope":
            common_headers["Authorization"] = f"Bearer {self.api_key}"
        elif self.platform_id == "openrouter":
            common_headers["Authorization"] = f"Bearer {self.api_key}"
        elif self.platform_id == "ollama":
            if self.api_key and self.api_key.lower() != "none" and self.api_key.lower() != "ollama":
                 common_headers["Authorization"] = f"Bearer {self.api_key}"
        elif self.platform_id in ["siliconflow", "openai", "deepseek", "moonshot"]:
            common_headers["Authorization"] = f"Bearer {self.api_key}"
        else:
            if self.api_key:
                common_headers["Authorization"] = f"Bearer {self.api_key}"
            logger.debug(f"Using default Bearer token auth for platform: {self.platform_id}")
        
        return common_headers

    def set_prompt_config(self, prompt_config: Dict[str, Any]):
        """设置自定义prompt配置"""
        self.prompt_config = prompt_config
        mode = prompt_config.get('mode', 'none')
        logger.info(f"Prompt config set: mode={mode}")
        
        if mode == 'professional':
            domain = prompt_config.get('prompt_template', 'academic')
            logger.info(f"Professional mode with domain: {domain}")
        elif mode == 'custom':
            logger.info("Custom prompt mode enabled")
        
        logger.debug(f"Full prompt config: {json.dumps(prompt_config, indent=2, ensure_ascii=False)}")

    def _merge_prompt_configs(self, frontend_config: Optional[Dict[str, Any]], instance_config: Optional[Dict[str, Any]]) -> Optional[Dict[str, Any]]:
        """合并前端配置和实例配置，前端配置优先"""
        if not frontend_config and not instance_config:
            return None
        
        if not frontend_config:
            return instance_config
        
        if not instance_config:
            return frontend_config
        
        # 合并配置，前端配置优先
        merged_config = instance_config.copy()
        merged_config.update(frontend_config)
        
        # 特殊处理需要合并而不是覆盖的字段
        self._merge_preserve_terms(merged_config, frontend_config, instance_config)
        self._merge_glossary(merged_config, frontend_config, instance_config)
        self._merge_additional_context(merged_config, frontend_config, instance_config)
        
        logger.debug(f"Config merged successfully")
        return merged_config

    def _merge_preserve_terms(self, merged_config: Dict[str, Any], frontend_config: Dict[str, Any], instance_config: Dict[str, Any]):
        """合并保留术语"""
        frontend_terms = frontend_config.get('preserve_terms')
        instance_terms = instance_config.get('preserve_terms')
        
        if frontend_terms and instance_terms:
            # 处理字符串格式的术语列表
            if isinstance(frontend_terms, str):
                frontend_terms = [term.strip() for term in frontend_terms.split(',') if term.strip()]
            if isinstance(instance_terms, str):
                instance_terms = [term.strip() for term in instance_terms.split(',') if term.strip()]
            
            # 确保都是列表
            frontend_terms = frontend_terms if isinstance(frontend_terms, list) else [frontend_terms]
            instance_terms = instance_terms if isinstance(instance_terms, list) else [instance_terms]
            
            # 合并并去重
            merged_config['preserve_terms'] = list(set(frontend_terms + instance_terms))
        elif frontend_terms:
            merged_config['preserve_terms'] = frontend_terms
        elif instance_terms:
            merged_config['preserve_terms'] = instance_terms

    def _merge_glossary(self, merged_config: Dict[str, Any], frontend_config: Dict[str, Any], instance_config: Dict[str, Any]):
        """合并术语表"""
        frontend_glossary = frontend_config.get('glossary')
        instance_glossary = instance_config.get('glossary')
        
        if frontend_glossary and instance_glossary:
            if isinstance(frontend_glossary, dict) and isinstance(instance_glossary, dict):
                merged_glossary = instance_glossary.copy()
                merged_glossary.update(frontend_glossary)  # 前端优先
                merged_config['glossary'] = merged_glossary
        elif frontend_glossary:
            merged_config['glossary'] = frontend_glossary
        elif instance_glossary:
            merged_config['glossary'] = instance_glossary

    def _merge_additional_context(self, merged_config: Dict[str, Any], frontend_config: Dict[str, Any], instance_config: Dict[str, Any]):
        """合并额外上下文"""
        frontend_context = frontend_config.get('additional_context', '').strip()
        instance_context = instance_config.get('additional_context', '').strip()
        
        if frontend_context and instance_context:
            merged_config['additional_context'] = f"{instance_context}\n{frontend_context}"
        elif frontend_context:
            merged_config['additional_context'] = frontend_context
        elif instance_context:
            merged_config['additional_context'] = instance_context

    def set_config_merge_strategy(self, strategy: str):
        """设置配置合并策略"""
        if strategy in ['merge', 'override', 'instance_only']:
            self.default_merge_strategy = strategy
            logger.info(f"Config merge strategy set to: {strategy}")
        else:
            logger.warning(f"Invalid merge strategy: {strategy}")

    def get_effective_config(self, frontend_config: Optional[Dict[str, Any]] = None) -> Optional[Dict[str, Any]]:
        """获取当前有效的配置（用于调试）"""
        return self._merge_prompt_configs(frontend_config, self.prompt_config)

    def _build_messages_with_prompt(self, text: str, target_lang: str, source_lang: Optional[str], prompt_config: Optional[Dict[str, Any]] = None) -> List[Dict[str, str]]:
        """根据prompt配置构建消息"""
        # 合并前端配置和实例配置
        config = self._merge_prompt_configs(prompt_config, self.prompt_config)
        
        logger.debug(f"Building messages with merged config")
        
        # 根据模式构建消息
        if not config or config.get('mode') == 'none':
            # 无prompt模式，使用最简单的翻译
            return self._build_simple_messages(text, target_lang, source_lang)
        elif config.get('mode') == 'general':
            # 通用翻译模式
            return self._build_general_messages(text, target_lang, source_lang, config)
        elif config.get('mode') == 'professional':
            # 专业翻译模式
            return self._build_professional_messages(text, target_lang, source_lang, config)
        elif config.get('mode') == 'custom':
            # 自定义prompt模式
            return self._build_custom_messages(text, target_lang, source_lang, config)
        else:
            # 默认使用通用模式
            logger.warning(f"Unknown prompt mode: {config.get('mode')}, using general mode")
            return self._build_general_messages(text, target_lang, source_lang, config)

    def _build_simple_messages(self, text: str, target_lang: str, source_lang: Optional[str]) -> List[Dict[str, str]]:
        """构建简单的翻译消息（无prompt模式）"""
        if source_lang:
            user_content = f"Translate from {source_lang} to {target_lang}:\n\n{text}"
        else:
            user_content = f"Translate to {target_lang}:\n\n{text}"
        
        return [
            {"role": "user", "content": user_content}
        ]

    def _build_general_messages(self, text: str, target_lang: str, source_lang: Optional[str], config: Dict[str, Any]) -> List[Dict[str, str]]:
        """构建通用翻译消息"""
        template = self.default_templates['general']
        system_content = template['system']
        
        # 应用通用增强
        system_content = self._apply_enhancements(system_content, config)
        
        # 构建用户消息
        user_content = template['user'].format(target_lang=target_lang, content=text)
        
        return [
            {"role": "system", "content": system_content},
            {"role": "user", "content": user_content}
        ]

    def _build_professional_messages(self, text: str, target_lang: str, source_lang: Optional[str], config: Dict[str, Any]) -> List[Dict[str, str]]:
        """构建专业翻译消息"""
        # 获取专业领域
        domain = config.get('prompt_template', 'academic')  # 前端使用prompt_template字段存储专业领域
        template = self.professional_templates.get(domain, self.professional_templates['academic'])
        
        system_content = template['system']
        
        # 应用增强配置
        system_content = self._apply_enhancements(system_content, config)
        
        # 构建用户消息
        user_content = template['user'].format(target_lang=target_lang, content=text)
        
        logger.info(f"Using professional template: {domain}")
        return [
            {"role": "system", "content": system_content},
            {"role": "user", "content": user_content}
        ]

    def _build_custom_messages(self, text: str, target_lang: str, source_lang: Optional[str], config: Dict[str, Any]) -> List[Dict[str, str]]:
        """构建自定义prompt消息"""
        custom_prompt = config.get('custom_prompt') or {}
        
        # 获取自定义的系统和用户prompt
        custom_system = custom_prompt.get('system', '')
        custom_user = custom_prompt.get('user', '{content}')
        
        if not custom_system or not custom_user:
            logger.warning("Custom prompt incomplete, falling back to general mode")
            return self._build_general_messages(text, target_lang, source_lang, config)
        
        # 格式化参数
        format_params = {
            'target_lang': target_lang,
            'source_lang': source_lang or 'auto-detect',
            'content': text
        }
        
        try:
            # 格式化自定义prompt
            system_content = custom_system.format(**format_params)
            user_content = custom_user.format(**format_params)
            
            # 应用其他增强配置（保留术语、术语表等）
            system_content = self._apply_enhancements(system_content, config)
            
            logger.info("Successfully applied custom prompt")
            return [
                {"role": "system", "content": system_content},
                {"role": "user", "content": user_content}
            ]
        except KeyError as e:
            logger.warning(f"Custom prompt format error: {e}, falling back to general mode")
            return self._build_general_messages(text, target_lang, source_lang, config)

    def _apply_enhancements(self, system_content: str, config: Dict[str, Any]) -> str:
        """应用增强配置（保留术语、术语表、上下文等）"""
        enhancements = []
        
        # 保留术语
        preserve_terms = config.get('preserve_terms')
        if preserve_terms:
            if isinstance(preserve_terms, str):
                # 前端可能传递逗号分隔的字符串
                terms_list = [term.strip() for term in preserve_terms.split(',') if term.strip()]
                if terms_list:
                    terms = ', '.join(terms_list)
                    enhancements.append(f"Keep these terms unchanged: {terms}")
            elif isinstance(preserve_terms, list):
                if preserve_terms:
                    terms = ', '.join(preserve_terms)
                    enhancements.append(f"Keep these terms unchanged: {terms}")
        
        # 术语表
        glossary = config.get('glossary')
        if glossary and isinstance(glossary, dict):
            glossary_items = [f"{k} → {v}" for k, v in glossary.items() if k and v]
            if glossary_items:
                glossary_text = '\n'.join(glossary_items)
                enhancements.append(f"Use this glossary for translation:\n{glossary_text}")
        
        # 额外上下文
        additional_context = config.get('additional_context')
        if additional_context and additional_context.strip():
            enhancements.append(f"Context: {additional_context.strip()}")
        
        # 批处理相关设置（记录到日志，但不影响prompt）
        max_units = config.get('max_units_per_chunk')
        max_chars = config.get('max_chars_per_chunk')
        if max_units or max_chars:
            logger.debug(f"Batch settings - max_units: {max_units}, max_chars: {max_chars}")
        
        # 将增强内容添加到系统prompt
        if enhancements:
            enhancement_text = '\n\n' + '\n'.join(enhancements)
            system_content += enhancement_text
            logger.debug(f"Applied {len(enhancements)} enhancements to system prompt")
        
        return system_content

    def get_available_professional_domains(self) -> List[str]:
        """获取可用的专业领域列表"""
        return list(self.professional_templates.keys())

    def translate(
        self,
        text: Optional[str] = None,
        target_lang: Optional[str] = None,
        source_lang: Optional[str] = None,
        messages: Optional[List[Dict[str, Any]]] = None,
        stream: bool = False,
        prompt_config: Optional[Dict[str, Any]] = None,
        config_merge_mode: Optional[str] = None
    ):
        """
        翻译方法
        
        Args:
            config_merge_mode: 配置合并模式
                - 'merge': 合并前端和实例配置（默认，前端优先）
                - 'override': 前端配置完全覆盖实例配置
                - 'instance_only': 只使用实例配置，忽略前端配置
        """
        
        # 根据合并模式处理配置
        merge_mode = config_merge_mode or self.default_merge_strategy
        final_prompt_config = None
        
        if merge_mode == 'merge':
            final_prompt_config = self._merge_prompt_configs(prompt_config, self.prompt_config)
        elif merge_mode == 'override':
            final_prompt_config = prompt_config or self.prompt_config
        elif merge_mode == 'instance_only':
            final_prompt_config = self.prompt_config
        else:
            logger.warning(f"Unknown config_merge_mode: {merge_mode}, using 'merge'")
            final_prompt_config = self._merge_prompt_configs(prompt_config, self.prompt_config)
        
        logger.info(f"Using config merge mode: {merge_mode}")
        if final_prompt_config:
            logger.info(f"Final prompt mode: {final_prompt_config.get('mode', 'none')}")
        
        try:
            if messages:
                final_messages = messages
                if not target_lang:
                    logger.warning("Target language is missing even with 'messages' provided.")
            elif text:
                if not target_lang:
                    logger.warning("Attempted translation with empty target language.")
                    return "Error: Target language cannot be empty." if not stream else self._yield_error_stream("Target language cannot be empty.")
                if not text:
                    logger.warning("Attempted translation with empty text.")
                    return "Error: Text to translate cannot be empty." if not stream else self._yield_error_stream("Text to translate cannot be empty.")

                # 使用最终的合并配置构建消息
                final_messages = self._build_messages_with_prompt(text, target_lang, source_lang, final_prompt_config)
            else:
                logger.warning("No text or messages provided for translation.")
                return "Error: No text or messages provided for translation." if not stream else self._yield_error_stream("No text or messages provided for translation.")

            payload = {
                "model": self.model,
                "messages": final_messages,
                "stream": stream
            }
            
            # 添加这一行：
            payload.update(self.extra_body)

            # 根据最终配置调整API参数
            if final_prompt_config:
                temperature = final_prompt_config.get('temperature', 0.7)
                max_tokens = final_prompt_config.get('max_tokens', 4000)
            else:
                temperature = 0.7
                max_tokens = 4000
            
            if self.platform_id not in ["ollama"]:
                payload["temperature"] = temperature
                payload["max_tokens"] = max_tokens

            api_endpoint = "/chat/completions"
            full_api_url = f"{self.base_url.rstrip('/')}{api_endpoint}"
            
            request_headers = self._get_headers()

            try:
                masked_key = f"{self.api_key[:5]}...{self.api_key[-4:]}" if self.api_key and len(self.api_key) > 8 else "Provided (short or empty)"
                logger.info(f"Sending {'STREAM' if stream else 'NON-STREAM'} request to: {full_api_url} for platform {self.platform_id} with model {self.model}")

                if stream:
                    response = requests.post(full_api_url, headers=request_headers, json=payload, stream=True, timeout=(10, 180))
                    if response.status_code >= 400:
                        error_content = response.text
                        response.close()
                        logger.error(f"Initial API HTTP Error {response.status_code} for STREAM request to {full_api_url}. Platform: {self.platform_id}. Response: {error_content[:500]}")
                        return self._yield_error_stream(f"API Error {response.status_code}: {error_content[:200]}")
                    return response
                else: # Non-stream
                    response = requests.post(full_api_url, headers=request_headers, json=payload, timeout=60)
                    response.raise_for_status()
                    data = response.json()
                    logger.debug(f"Received NON-STREAM response from {self.platform_id}")
                    
                    if data and data.get("choices"):
                        if isinstance(data["choices"], list) and len(data["choices"]) > 0 and isinstance(data["choices"][0], dict):
                            if data["choices"][0].get("message") and data["choices"][0]["message"].get("content") is not None:
                                translated_text = data["choices"][0]["message"]["content"].strip()
                                return translated_text
                            else:
                                error_msg = f"No 'message' or 'content' found in choices for non-stream response from {self.platform_id}"
                                logger.error(error_msg)
                                return f"Error: {error_msg}"
                        else:
                            error_msg = f"Unexpected 'choices' structure in non-stream response from {self.platform_id}"
                            logger.error(error_msg)
                            return f"Error: {error_msg}"
                    else:
                        error_msg = f"No 'choices' found in non-stream response or unexpected format from {self.platform_id}"
                        logger.error(error_msg)
                        return f"Error: {error_msg}"

            except requests.exceptions.HTTPError as e:
                status_code = e.response.status_code if e.response is not None else "Unknown"
                response_text = e.response.text if e.response is not None else "No response text"
                try:
                    error_details_json = e.response.json() if e.response is not None else {}
                    message = error_details_json.get("error", {}).get("message", response_text)
                    if not message or message == response_text:
                        message = error_details_json.get("errors", {}).get("message", response_text)
                    if not message or message == response_text:
                        message = error_details_json.get("detail", response_text)
                except json.JSONDecodeError:
                    message = response_text

                error_msg_prefix = f"API HTTP Error {status_code}"
                specific_error_msg = self._get_specific_http_error_message(status_code)
                
                full_error_output = f"{error_msg_prefix}: {specific_error_msg}\nDetails: {message}"
                logger.error(f"HTTPError for {self.platform_id}: {full_error_output}")
                return full_error_output if not stream else self._yield_error_stream(full_error_output)

            except requests.exceptions.RequestException as e:
                error_msg = f"Network/Request Error for {self.platform_id}: {type(e).__name__} - {e}"
                logger.error(error_msg)
                return error_msg if not stream else self._yield_error_stream(error_msg)
            except json.JSONDecodeError as e:
                raw_response_text = response.text if 'response' in locals() and hasattr(response, 'text') else 'No response text available.'
                error_msg = f"JSON Decode Error from {self.platform_id}: Could not decode API response. {e}"
                logger.error(error_msg)
                return error_msg if not stream else self._yield_error_stream(error_msg)
            except Exception as e:
                import traceback
                error_msg = f"Unexpected error in translator for {self.platform_id}: {type(e).__name__} - {e}"
                logger.error(f"{error_msg}\n{traceback.format_exc()}")
                return error_msg if not stream else self._yield_error_stream(error_msg)

        except Exception as e:
            import traceback
            error_msg = f"Unexpected error in translate method: {type(e).__name__} - {e}"
            logger.error(f"{error_msg}\n{traceback.format_exc()}")
            return error_msg if not stream else self._yield_error_stream(error_msg)

    def _get_specific_http_error_message(self, status_code: int) -> str:
        if status_code == 401: return "Unauthorized. API Key is invalid, missing, or expired."
        if status_code == 403: return "Forbidden. API Key may lack permissions for this model/operation."
        if status_code == 404: return f"Not Found. Model '{self.model}' or API endpoint incorrect."
        if status_code == 429: return "Rate Limit Exceeded. Please wait and try again or check your plan."
        if status_code >= 500: return "API Server Error. Please try again later."
        return "A client-side or unexpected API error occurred."

    def _yield_error_stream(self, error_message: str):
        logger.debug(f"Yielding error stream: {error_message}")
        yield f"data: {json.dumps({'error': error_message})}\n\n"