# prompt_builder.py
"""
Prompt构建工具，用于创建自定义prompt
"""
from typing import List, Dict, Optional

class PromptBuilder:
    """帮助构建自定义prompt的工具类"""
    
    @staticmethod
    def create_domain_specific_prompt(
        domain: str,
        style: str = "formal",
        preserve_terms: Optional[List[str]] = None,
        glossary: Optional[Dict[str, str]] = None,
        additional_rules: Optional[List[str]] = None,
        examples: Optional[List[Dict[str, str]]] = None
    ) -> Dict[str, str]:
        """创建特定领域的prompt"""
        
        # 基础系统prompt
        system_prompt = f"""You are a professional {domain} translator. Translate to {{target_lang}}.

Style: {style}

Core Rules:
1. Translate each numbered line separately
2. Keep the exact same number of lines
3. Return ONLY the translations
4. Maintain domain-specific terminology accuracy"""
        
        # 添加需要保留的术语
        if preserve_terms:
            system_prompt += f"\n\nTerms to preserve unchanged:\n"
            system_prompt += "\n".join(f"- {term}" for term in preserve_terms)
        
        # 添加术语表
        if glossary:
            system_prompt += f"\n\nGlossary (use these translations):\n"
            for source, target in glossary.items():
                system_prompt += f"- {source} → {target}\n"
        
        # 添加额外规则
        if additional_rules:
            system_prompt += f"\n\nAdditional rules:\n"
            system_prompt += "\n".join(f"{i+1}. {rule}" for i, rule in enumerate(additional_rules))
        
        # 添加示例
        if examples:
            system_prompt += f"\n\nExamples:\n"
            for i, example in enumerate(examples):
                system_prompt += f"Example {i+1}:\n"
                system_prompt += f"Source: {example['source']}\n"
                system_prompt += f"Target: {example['target']}\n\n"
        
        return {
            'system': system_prompt,
            'user': "{content}"
        }
    
    @staticmethod
    def create_context_aware_prompt(
        document_type: str,
        target_audience: str,
        tone: str,
        context: str,
        cultural_adaptation: bool = True
    ) -> Dict[str, str]:
        """创建上下文感知的prompt"""
        
        system_prompt = f"""You are translating a {document_type} for {target_audience}. 
Translate to {{target_lang}} with a {tone} tone.

Context: {context}

Rules:
1. Adapt the translation to the target audience
2. Maintain the {tone} tone throughout
3. Translate each numbered line separately
4. Keep the exact same number of lines
5. Return ONLY the translations"""
        
        if cultural_adaptation:
            system_prompt += "\n6. Adapt cultural references appropriately for the target audience"
        else:
            system_prompt += "\n6. Preserve cultural references with explanatory notes if needed"
        
        return {
            'system': system_prompt,
            'user': "{content}"
        }
    
    @staticmethod
    def create_file_format_specific_prompt(
        file_type: str,
        preserve_formatting: bool = True,
        handle_special_elements: Optional[Dict[str, str]] = None
    ) -> Dict[str, str]:
        """创建针对特定文件格式的prompt"""
        
        system_prompt = f"""You are translating a {file_type} file. Translate to {{target_lang}}.

File Format Rules:"""
        
        if preserve_formatting:
            system_prompt += "\n1. Preserve all formatting markers and structure"
        else:
            system_prompt += "\n1. Focus on content, formatting will be handled separately"
        
        if handle_special_elements:
            system_prompt += "\n\nSpecial Elements Handling:"
            for element, instruction in handle_special_elements.items():
                system_prompt += f"\n- {element}: {instruction}"
        
        system_prompt += "\n\nGeneral Rules:"
        system_prompt += "\n- Translate each numbered line separately"
        system_prompt += "\n- Keep the exact same number of lines"
        system_prompt += "\n- Return ONLY the translations"
        
        return {
            'system': system_prompt,
            'user': "{content}"
        }
    
    @staticmethod
    def merge_prompts(*prompts: Dict[str, str]) -> Dict[str, str]:
        """合并多个prompt配置"""
        merged_system = ""
        merged_user = "{content}"
        
        for prompt in prompts:
            if prompt.get('system'):
                merged_system += prompt['system'] + "\n\n"
            if prompt.get('user') and prompt['user'] != "{content}":
                merged_user = prompt['user']
        
        return {
            'system': merged_system.strip(),
            'user': merged_user
        }
    
    @staticmethod
    def create_instruction_based_prompt(
        instructions: List[str],
        examples: Optional[List[Dict[str, str]]] = None,
        constraints: Optional[List[str]] = None
    ) -> Dict[str, str]:
        """基于指令列表创建prompt"""
        
        system_prompt = "You are a professional translator. Translate to {target_lang}.\n\n"
        
        # 添加指令
        system_prompt += "Instructions:\n"
        for i, instruction in enumerate(instructions, 1):
            system_prompt += f"{i}. {instruction}\n"
        
        # 添加约束
        if constraints:
            system_prompt += "\nConstraints:\n"
            for constraint in constraints:
                system_prompt += f"- {constraint}\n"
        
        # 添加示例
        if examples:
            system_prompt += "\nExamples:\n"
            for i, example in enumerate(examples, 1):
                system_prompt += f"\nExample {i}:\n"
                system_prompt += f"Input: {example.get('input', '')}\n"
                system_prompt += f"Output: {example.get('output', '')}\n"
        
        return {
            'system': system_prompt,
            'user': "{content}"
        }
