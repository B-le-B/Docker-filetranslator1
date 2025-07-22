# translators/__init__.py

import logging
logger = logging.getLogger(__name__)

# 核心翻译器 - 从 translator.py 导入
try:
    from .translator import SiliconFlowTranslator
    logger.info("SiliconFlowTranslator imported successfully from translator.py")
except ImportError as e:
    logger.error(f"Failed to import SiliconFlowTranslator from translator.py: {e}")
    SiliconFlowTranslator = None

# 文件翻译器 - 从 file_translator.py 导入
try:
    from .file_translator import translate_text_file
    logger.info("translate_text_file imported successfully from file_translator.py")
except ImportError as e:
    logger.warning(f"translate_text_file not available from file_translator.py: {e}")
    translate_text_file = None

# DOCX 翻译器 - 可选导入
try:
    from .docx_translator import translate_docx_file
    _translate_docx_available = True
    logger.info("docx_translator imported successfully")
except ImportError as e:
    logger.warning(f"docx_translator not available: {e}")
    _translate_docx_available = False
    translate_docx_file = None

try:
    from .docx_full_translator import translate_docx_file_formatted
    _translate_docx_file_formatted_available = True
    logger.info("docx_full_translator imported successfully")
except ImportError as e:
    logger.warning(f"docx_full_translator not available: {e}")
    _translate_docx_file_formatted_available = False
    translate_docx_file_formatted = None

try:
    from .docx_pythondoc1_translator import translate_docx_via_markdown as translate_docx_pythondoc1
    _translate_docx_pythondoc1_available = True
    logger.info("docx_pythondoc1_translator imported successfully")
except ImportError as e:
    logger.warning(f"docx_pythondoc1_translator not available: {e}")
    _translate_docx_pythondoc1_available = False
    translate_docx_pythondoc1 = None

try:
    from .docx_pythondoc2_translator import translate_docx_via_markdown as translate_docx_pythondoc2
    _translate_docx_pythondoc2_available = True
    logger.info("docx_pythondoc2_translator imported successfully")
except ImportError as e:
    logger.warning(f"docx_pythondoc2_translator not available: {e}")
    _translate_docx_pythondoc2_available = False
    translate_docx_pythondoc2 = None

try:
    from .docx_markdown_translator import translate_docx_via_markdown
    _translate_docx_markdown_available = True
    logger.info("docx_markdown_translator imported successfully")
except ImportError as e:
    logger.warning(f"docx_markdown_translator not available: {e}")
    _translate_docx_markdown_available = False
    translate_docx_via_markdown = None

# PPTX 翻译器 - 可选导入
try:
    from .pptx_full_translator import translate_pptx_file_formatted
    _translate_pptx_available = True
    logger.info("pptx_full_translator imported successfully")
except ImportError as e:
    logger.warning(f"pptx_full_translator not available: {e}")
    _translate_pptx_available = False
    translate_pptx_file_formatted = None

# Excel 翻译器 - 可选导入
try:
    from .excel_full_translator import translate_excel_file_formatted
    _translate_excel_available = True
    logger.info("excel_full_translator imported successfully")
except ImportError as e:
    logger.warning(f"excel_full_translator not available: {e}")
    _translate_excel_available = False
    translate_excel_file_formatted = None

# 火山引擎翻译器 - 可选导入
try:
    from .volcengine_native_translator import VolcEngineNativeTranslator
    _volcengine_available = True
    logger.info("volcengine_native_translator imported successfully")
except ImportError as e:
    logger.warning(f"volcengine_native_translator not available: {e}")
    _volcengine_available = False
    VolcEngineNativeTranslator = None

# 导出所有内容
__all__ = [
    'SiliconFlowTranslator',
    'translate_text_file',
    'translate_docx_file',
    'translate_docx_file_formatted', 
    'translate_docx_pythondoc1',
    'translate_docx_pythondoc2',
    'translate_docx_via_markdown',
    'translate_pptx_file_formatted',
    'translate_excel_file_formatted',
    'VolcEngineNativeTranslator',
    '_translate_docx_available',
    '_translate_docx_file_formatted_available',
    '_translate_docx_pythondoc1_available',
    '_translate_docx_pythondoc2_available', 
    '_translate_docx_markdown_available',
    '_translate_pptx_available',
    '_translate_excel_available',
    '_volcengine_available',
]