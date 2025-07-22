# config_manager.py
"""
配置管理器，用于保存和加载用户的翻译配置
"""
import json
import os
import tempfile
from typing import Dict, Any, Optional, List
from datetime import datetime

class ConfigManager:
    """管理翻译配置的保存和加载"""
    
    def __init__(self, config_dir: str = None):
        # 如果没有提供config_dir，使用临时目录
        if config_dir is None:
            config_dir = os.path.join(tempfile.gettempdir(), "translator_user_configs")
        
        self.config_dir = config_dir
        os.makedirs(config_dir, exist_ok=True)
        self.default_config_file = os.path.join(config_dir, "default_config.json")
        self.user_prompts_file = os.path.join(config_dir, "user_prompts.json")
        
    def save_config(self, config_name: str, config: Dict[str, Any]) -> bool:
        """保存配置"""
        try:
            config['saved_at'] = datetime.now().isoformat()
            # 确保 config 字典中包含 'name' 键，值为用户输入的配置名称
            config['name'] = config_name 
            
            # 确保文件名安全，只允许字母数字、空格、连字符、下划线，并移除末尾空格
            safe_name = "".join(c for c in config_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
            # 如果安全文件名为空，则使用一个基于时间戳的名称
            if not safe_name:
                safe_name = f"config_{int(datetime.now().timestamp())}"
            filepath = os.path.join(self.config_dir, f"{safe_name}.json")
            
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            
            return True
        except Exception as e:
            print(f"Error saving config '{config_name}': {e}")
            return False
    
    def load_config(self, config_name: str) -> Optional[Dict[str, Any]]:
        """加载配置"""
        try:
            # 确保文件名安全
            safe_name = "".join(c for c in config_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
            filepath = os.path.join(self.config_dir, f"{safe_name}.json")
            
            if os.path.exists(filepath):
                with open(filepath, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return None
        except Exception as e:
            print(f"Error loading config '{config_name}': {e}")
            return None
    
    def list_configs(self) -> List[Dict[str, str]]:
        """列出所有保存的配置，返回包含配置名称的字典列表"""
        configs_list = []
        try:
            for filename in os.listdir(self.config_dir):
                # 排除内部使用的json文件，如 user_prompts.json 和 default_config.json
                if filename.endswith('.json') and \
                   filename not in ['user_prompts.json', 'default_config.json']:
                    
                    filepath = os.path.join(self.config_dir, filename)
                    try:
                        with open(filepath, 'r', encoding='utf-8') as f:
                            config_data = json.load(f)
                            # 从加载的JSON数据中获取真实的配置名称
                            if 'name' in config_data:
                                configs_list.append({"name": config_data['name']})
                            else:
                                # 如果没有'name'字段，则使用文件名作为备用名称
                                configs_list.append({"name": filename[:-5]}) 
                    except json.JSONDecodeError:
                        print(f"Warning: Could not decode JSON from {filename}. Skipping.")
                    except Exception as inner_e:
                        print(f"Warning: Error processing config file {filename}: {inner_e}. Skipping.")
        except Exception as e:
            print(f"Error listing configs: {e}")
        
        # 按照配置名称排序
        return sorted(configs_list, key=lambda x: x.get('name', ''))
    
    def delete_config(self, config_name: str) -> bool:
        """删除配置"""
        try:
            # 确保文件名安全
            safe_name = "".join(c for c in config_name if c.isalnum() or c in (' ', '-', '_')).rstrip()
            filepath = os.path.join(self.config_dir, f"{safe_name}.json")
            
            if os.path.exists(filepath):
                os.remove(filepath)
                return True
            return False
        except Exception as e:
            print(f"Error deleting config '{config_name}': {e}")
            return False
    
    def save_user_prompt(self, prompt_name: str, prompt: Dict[str, str]) -> bool:
        """保存用户自定义prompt"""
        try:
            user_prompts = {}
            if os.path.exists(self.user_prompts_file):
                with open(self.user_prompts_file, 'r', encoding='utf-8') as f:
                    user_prompts = json.load(f)
            
            user_prompts[prompt_name] = {
                'system': prompt.get('system', ''), # Ensure system is handled even if empty
                'user': prompt.get('user', '{content}'),
                'created_at': datetime.now().isoformat(),
                'updated_at': datetime.now().isoformat()
            }
            
            with open(self.user_prompts_file, 'w', encoding='utf-8') as f:
                json.dump(user_prompts, f, ensure_ascii=False, indent=2)
            
            return True
        except Exception as e:
            print(f"Error saving user prompt '{prompt_name}': {e}")
            return False
    
    def get_user_prompts(self) -> Dict[str, Dict[str, str]]:
        """获取所有用户自定义prompts"""
        try:
            if os.path.exists(self.user_prompts_file):
                with open(self.user_prompts_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            return {}
        except Exception as e:
            print(f"Error loading user prompts: {e}")
            return {}
    
    def delete_user_prompt(self, prompt_name: str) -> bool:
        """删除用户自定义prompt"""
        try:
            user_prompts = self.get_user_prompts()
            if prompt_name in user_prompts:
                del user_prompts[prompt_name]
                
                with open(self.user_prompts_file, 'w', encoding='utf-8') as f:
                    json.dump(user_prompts, f, ensure_ascii=False, indent=2)
                return True
            return False
        except Exception as e:
            print(f"Error deleting user prompt '{prompt_name}': {e}")
            return False
    
    def export_config(self, config_name: str) -> Optional[str]:
        """导出配置为JSON字符串"""
        config = self.load_config(config_name)
        if config:
            return json.dumps(config, ensure_ascii=False, indent=2)
        return None
    
    def import_config(self, config_json: str, new_name: Optional[str] = None) -> bool:
        """从JSON字符串导入配置"""
        try:
            config = json.loads(config_json)
            # 如果 new_name 未提供，则尝试使用 config['name']，否则生成一个时间戳名称
            config_name = new_name or config.get('name', f'imported_{int(datetime.now().timestamp())}')
            return self.save_config(config_name, config)
        except Exception as e:
            print(f"Error importing config: {e}")
            return False
    
    def get_default_config(self) -> Dict[str, Any]:
        """提供一个硬编码的默认配置"""
        default_config = {
            'name': 'default', # Added name for default config
            'prompt_config': {
                'mode': 'none',
                'prompt_template': 'default',
                'custom_prompt': { # Ensure default custom_prompt structure is consistent
                    'system': '',
                    'user': '{content}'
                },
                'preserve_terms': [],
                'glossary': {},
                'additional_context': '',
                'max_units_per_chunk': 50,
                'max_chars_per_chunk': 30000
            },
            'api_platform': 'siliconflow',
            'source_lang': '',
            'target_lang': '简体中文',
            'encoding': 'utf-8', # Added default encoding
            'docx_processing_method': 'docx_markdown_translator', # Added default docx method
            'output_folder_path': 'translated_output_web' # Added default output path
        }
        return default_config
    
    def save_default_config(self, config: Dict[str, Any]) -> bool:
        """保存默认配置到 default_config.json"""
        try:
            with open(self.default_config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, ensure_ascii=False, indent=2)
            return True
        except Exception as e:
            print(f"Error saving default config: {e}")
            return False
    
    def load_default_config(self) -> Dict[str, Any]:
        """加载默认配置，如果不存在则返回硬编码的默认配置"""
        try:
            if os.path.exists(self.default_config_file):
                with open(self.default_config_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Error loading default config: {e}")
        
        # 如果加载失败或文件不存在，返回硬编码的默认配置
        return self.get_default_config()
