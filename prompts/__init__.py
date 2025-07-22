# prompts/__init__.py
try:
    from .prompt_templates import get_all_templates, get_template
    from .prompt_builder import PromptBuilder
    from .config_manager import ConfigManager
    _prompt_modules_available = True
except ImportError:
    def get_all_templates():
        return {}
    def get_template(name):
        return None
    class PromptBuilder:
        pass
    class ConfigManager:
        def __init__(self, *args, **kwargs):
            pass
    _prompt_modules_available = False

__all__ = ['get_all_templates', 'get_template', 'PromptBuilder', 'ConfigManager']
