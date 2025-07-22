# file_translator.py
import os
import logging
import time
import hashlib
import threading
from typing import Optional, List, Dict, Any, Tuple
from concurrent.futures import ThreadPoolExecutor, as_completed, TimeoutError
from tqdm import tqdm
from enum import Enum
from dataclasses import dataclass
from .translator import SiliconFlowTranslator

logger = logging.getLogger(__name__)

# 配置常量 - 基于参考代码
DEFAULT_CHUNK_SIZE = 10
DEFAULT_MAX_CHARS = 1000
DEFAULT_THREADS = 5
DEFAULT_TRANSLATION_TIMEOUT = 45
DEFAULT_MAX_RETRIES = 8
DEFAULT_RETRY_DELAY = 2
DEFAULT_RETRY_BATCH_SIZE = 10
DEFAULT_RETRY_THREADS = 5

# 全局缓存
_cache = {}
_cache_lock = threading.Lock()

# 失败原因枚举
class FailureReason(Enum):
    """失败原因枚举"""
    TIMEOUT = "timeout"
    API_ERROR = "api_error"
    PARSE_ERROR = "parse_error"
    NOT_TRANSLATED = "not_translated"
    EMPTY_RESPONSE = "empty_response"
    CONNECTION_ERROR = "connection_error"
    BATCH_FAILURE = "batch_failure"
    FORMAT_ERROR = "format_error"

# 失败任务数据类
@dataclass
class FailedTask:
    """失败任务数据类"""
    original_text: str
    chunk_index: int
    task_id: str
    failure_reason: FailureReason
    retry_count: int = 0
    error_message: str = ""
    is_serious: bool = True
    
    def __post_init__(self):
        self.failure_timestamp = time.time()

class TranslationValidator:
    """翻译完整性验证器"""
    
    ERROR_KEYWORDS = [
        'timeout', 'readtimeout', 'connecttimeout', 'httptimeout',
        'network error', 'connection error', 'api error', 'service error',
        'translation failed', 'service unavailable', 'request failed',
        'server error', 'bad gateway', 'gateway timeout', 'error:',
        '超时', '网络错误', '连接错误', '服务错误', '翻译失败',
        '服务不可用', '请求失败', '服务器错误', '错误:'
    ]
    
    @staticmethod
    def is_error_message(text: str) -> bool:
        """检查是否为错误消息"""
        if not text:
            return True
        text_lower = text.lower()
        return any(keyword.lower() in text_lower for keyword in TranslationValidator.ERROR_KEYWORDS)
    
    @staticmethod
    def is_serious_failure(original_text: str, translated_text: str, 
                          large_text_threshold: int = 50) -> Tuple[bool, str]:
        """检测失败类型"""
        if TranslationValidator.is_error_message(translated_text):
            return True, f"API错误消息: {translated_text[:50]}..."
        
        if not translated_text or translated_text.strip() == "":
            return len(original_text) > 20, f"空翻译结果（原文{len(original_text)}字符）"
        
        is_unchanged = original_text.strip() == translated_text.strip()
        if is_unchanged:
            if len(original_text) > large_text_threshold:
                return True, f"长文本未翻译（{len(original_text)}字符）"
            else:
                return False, f"短文本未变化（{len(original_text)}字符，可能正常）"
        
        return False, "翻译正常"

class FileTranslator:
    """增强的文件翻译器 - 集成并行和重试机制"""
    
    def __init__(self, translator: SiliconFlowTranslator,
                 translation_timeout: int = DEFAULT_TRANSLATION_TIMEOUT,
                 max_retries: int = DEFAULT_MAX_RETRIES,
                 retry_delay: int = DEFAULT_RETRY_DELAY,
                 retry_batch_size: int = DEFAULT_RETRY_BATCH_SIZE,
                 max_retry_workers: int = DEFAULT_RETRY_THREADS,
                 large_text_threshold: int = 50):
        
        self.translator = translator
        self.translation_timeout = translation_timeout
        self.max_retries = max_retries
        self.retry_delay = retry_delay
        self.retry_batch_size = retry_batch_size
        self.max_retry_workers = 5  # 固定为5线程
        self.large_text_threshold = large_text_threshold
        
        # 统计信息
        self.stats = {
            'success': 0, 'failed': 0, 'skipped': 0, 'timeout': 0, 'retried': 0,
            'final_failed': 0, 'serious_failures': 0, 'minor_issues': 0,
            'retry_attempts': 0, 'rescued_tasks': 0, 'retry_success_rate': 0.0,
            'total_retry_tasks': 0, 'parallel_retry_used': 0,
            'cache_hits': 0, 'api_calls': 0, 'total_chunks': 0
        }
        
        # 时间统计
        self.time_stats = {
            'total_start_time': 0, 'main_translation_time': 0,
            'retry_time': 0, 'total_time': 0,
            'file_read_time': 0, 'file_write_time': 0
        }
        
        # 失败任务追踪
        self.failed_tasks = []
        self.failed_tasks_lock = threading.Lock()
        
        # 重试策略配置
        self.retry_batch_sizes = [10, 5, 2, 1, 1, 1, 1, 1]
        self.retry_delays = [1, 2, 4, 8, 12, 16, 20, 25]
        
        # 存储翻译结果
        self.translation_results = {}
        self.translation_results_lock = threading.Lock()
    
    def _get_cache_key(self, text: str, target_lang: str, source_lang: str) -> str:
        """生成缓存键"""
        content = f"{text}_{target_lang}_{source_lang or 'auto'}"
        return hashlib.md5(content.encode('utf-8')).hexdigest()
    
    def _call_translator_with_timeout(self, text: str, target_lang: str, 
                                    source_lang: Optional[str],
                                    prompt_config: Optional[Dict[str, Any]],
                                    timeout: Optional[int] = None) -> str:
        """带超时的翻译调用"""
        timeout = timeout or self.translation_timeout
        
        with ThreadPoolExecutor(max_workers=1) as executor:
            future = executor.submit(self._call_translator, text, target_lang, source_lang, prompt_config)
            try:
                result = future.result(timeout=timeout)
                return result
            except TimeoutError:
                logger.warning(f"Translation timeout after {timeout} seconds")
                self.stats['timeout'] += 1
                return f"Error: Translation timeout after {timeout} seconds"
            except Exception as e:
                logger.error(f"Translation error: {e}")
                return f"Error: {str(e)}"
    
    def _call_translator(self, text: str, target_lang: str, source_lang: Optional[str],
                        prompt_config: Optional[Dict[str, Any]]) -> str:
        """调用翻译器"""
        try:
            return self.translator.translate(
                text=text,
                target_lang=target_lang,
                source_lang=source_lang,
                prompt_config=prompt_config,
                config_merge_mode='merge'
            )
        except Exception as e:
            logger.error(f"Translation failed: {e}")
            return f"Error: {str(e)}"
    
    def _translate_chunk_batch(self, chunk_texts: List[str], target_lang: str,
                              source_lang: Optional[str], prompt_config: Optional[Dict[str, Any]],
                              batch_info: Optional[Dict] = None) -> tuple:
        """翻译一批文本块"""
        if not chunk_texts:
            return [], [], None, []
        
        # 检查缓存
        cached_results = []
        uncached_texts = []
        uncached_indices = []
        cache_flags = [False] * len(chunk_texts)
        
        with _cache_lock:
            for i, text in enumerate(chunk_texts):
                cache_key = self._get_cache_key(text, target_lang, source_lang or 'auto')
                if cache_key in _cache:
                    cached_results.append((i, _cache[cache_key]))
                    cache_flags[i] = True
                    self.stats['cache_hits'] += 1
                else:
                    uncached_texts.append(text)
                    uncached_indices.append(i)
        
        results = [""] * len(chunk_texts)
        failed_indices = []
        error_message = None
        
        # 恢复缓存结果
        for idx, result in cached_results:
            results[idx] = result
        
        # 翻译未缓存的文本
        if uncached_texts:
            for i, (orig_idx, text) in enumerate(zip(uncached_indices, uncached_texts)):
                response = self._call_translator_with_timeout(
                    text, target_lang, source_lang, prompt_config
                )
                self.stats['api_calls'] += 1
                
                if response and not response.startswith("Error:"):
                    if not TranslationValidator.is_error_message(response):
                        results[orig_idx] = response.strip()
                        # 更新缓存
                        with _cache_lock:
                            cache_key = self._get_cache_key(text, target_lang, source_lang or 'auto')
                            _cache[cache_key] = response.strip()
                    else:
                        results[orig_idx] = ""
                        failed_indices.append(orig_idx)
                else:
                    error_message = response if response else "Unknown translation error"
                    failed_indices.append(orig_idx)
                    results[orig_idx] = ""
        
        return results, failed_indices, error_message, cache_flags
    
    def _record_failed_task(self, text: str, chunk_index: int, failure_reason: FailureReason,
                          error_msg: str = "", is_serious: bool = True) -> str:
        """记录失败任务"""
        task_id = f"chunk_{chunk_index}"
        
        failed_task = FailedTask(
            original_text=text,
            chunk_index=chunk_index,
            task_id=task_id,
            failure_reason=failure_reason,
            error_message=error_msg,
            is_serious=is_serious
        )
        
        with self.failed_tasks_lock:
            self.failed_tasks.append(failed_task)
        
        if is_serious:
            self.stats['serious_failures'] += 1
            return "serious"
        else:
            self.stats['minor_issues'] += 1
            return "minor"
    
    def _create_chunk_batches(self, chunks: List[str], batch_size: int = DEFAULT_CHUNK_SIZE) -> List[List[str]]:
        """创建块批次"""
        batches = []
        current_batch = []
        current_chars = 0
        
        for chunk in chunks:
            chunk_len = len(chunk)
            
            if (len(current_batch) >= batch_size or 
                current_chars + chunk_len > DEFAULT_MAX_CHARS) and current_batch:
                batches.append(current_batch)
                current_batch = []
                current_chars = 0
            
            current_batch.append(chunk)
            current_chars += chunk_len
        
        if current_batch:
            batches.append(current_batch)
        
        return batches
    
    def _process_chunk_batch(self, batch_data) -> tuple:
        """处理单个块批次"""
        batch, batch_id, thread_id, target_lang, source_lang, prompt_config = batch_data
        
        try:
            logger.info(f"线程-{thread_id}: 处理批次-{batch_id} (共{len(batch)}块)")
            
            batch_info = {'batch_id': batch_id, 'thread_id': thread_id}
            translations, failed_indices, error_message, cache_flags = self._translate_chunk_batch(
                batch, target_lang, source_lang, prompt_config, batch_info
            )
            
            # 存储成功的翻译结果
            success_count = 0
            serious_failures = 0
            minor_issues = 0
            
            with self.translation_results_lock:
                for i, (chunk_text, translation) in enumerate(zip(batch, translations)):
                    chunk_index = batch_id * len(batch) + i  # 计算全局索引
                    
                    if translation and translation.strip() and not TranslationValidator.is_error_message(translation):
                        is_serious, reason = TranslationValidator.is_serious_failure(
                            chunk_text, translation, self.large_text_threshold
                        )
                        
                        if not is_serious:
                            self.translation_results[chunk_index] = translation
                            success_count += 1
                        else:
                            failure_result = self._record_failed_task(
                                chunk_text, chunk_index, FailureReason.NOT_TRANSLATED, reason, is_serious
                            )
                            serious_failures += 1 if failure_result == "serious" else 0
                            minor_issues += 1 if failure_result == "minor" else 0
                    else:
                        # 翻译失败
                        if error_message and "timeout" in error_message.lower():
                            failure_reason = FailureReason.TIMEOUT
                        else:
                            failure_reason = FailureReason.API_ERROR
                        
                        failure_result = self._record_failed_task(
                            chunk_text, chunk_index, failure_reason, 
                            error_message or "翻译失败或空结果", True
                        )
                        serious_failures += 1 if failure_result == "serious" else 0
                        minor_issues += 1 if failure_result == "minor" else 0
            
            logger.info(f"线程-{thread_id}: 批次-{batch_id} 完成，成功: {success_count}/{len(batch)}")
            return len(batch), success_count, serious_failures + minor_issues, error_message, serious_failures
            
        except Exception as e:
            logger.error(f"线程-{thread_id}: 批次-{batch_id} 处理失败: {e}")
            serious_failures = 0
            for i, chunk_text in enumerate(batch):
                chunk_index = batch_id * len(batch) + i
                failure_result = self._record_failed_task(
                    chunk_text, chunk_index, FailureReason.BATCH_FAILURE, 
                    f"批次处理异常: {str(e)}", True
                )
                serious_failures += 1 if failure_result == "serious" else 0
            
            return len(batch), 0, len(batch), str(e), serious_failures
    
    def _create_smart_retry_batches(self, failed_tasks: List[FailedTask], retry_count: int) -> List[List[FailedTask]]:
        """创建智能重试批次"""
        serious_failed_tasks = [task for task in failed_tasks if task.is_serious]
        
        if not serious_failed_tasks:
            return []
        
        if retry_count < len(self.retry_batch_sizes):
            max_batch_size = self.retry_batch_sizes[retry_count]
        else:
            max_batch_size = 1
        
        logger.info(f"第{retry_count + 1}次重试，严重失败任务: {len(serious_failed_tasks)}, 批次大小: {max_batch_size}")
        
        # 按失败原因分组
        tasks_by_reason = {}
        for task in serious_failed_tasks:
            reason = task.failure_reason
            if reason not in tasks_by_reason:
                tasks_by_reason[reason] = []
            tasks_by_reason[reason].append(task)
        
        chunks = []
        for reason, tasks in tasks_by_reason.items():
            if reason == FailureReason.TIMEOUT:
                batch_size = max(1, max_batch_size // 2)
            elif reason == FailureReason.API_ERROR:
                batch_size = max(1, max_batch_size // 3)
            else:
                batch_size = max_batch_size
            
            # 按文本长度排序
            tasks.sort(key=lambda x: len(x.original_text), reverse=True)
            
            current_chunk = []
            current_chars = 0
            
            for task in tasks:
                text_len = len(task.original_text)
                
                if text_len > DEFAULT_MAX_CHARS // 2:
                    if current_chunk:
                        chunks.append(current_chunk)
                        current_chunk = []
                        current_chars = 0
                    chunks.append([task])
                    continue
                
                if (len(current_chunk) >= batch_size or 
                    current_chars + text_len > DEFAULT_MAX_CHARS) and current_chunk:
                    chunks.append(current_chunk)
                    current_chunk = [task]
                    current_chars = text_len
                else:
                    current_chunk.append(task)
                    current_chars += text_len
            
            if current_chunk:
                chunks.append(current_chunk)
        
        return chunks
    
    def _process_retry_batch(self, retry_batch: List[FailedTask], batch_idx: int, retry_count: int,
                           target_lang: str, source_lang: Optional[str],
                           prompt_config: Optional[Dict[str, Any]]) -> int:
        """处理单个重试批次"""
        logger.info(f"开始重试批次 {batch_idx + 1}，任务数: {len(retry_batch)}")
        
        batch_texts = [task.original_text for task in retry_batch]
        
        translations, failed_indices, error_message, cache_flags = self._translate_chunk_batch(
            batch_texts, target_lang, source_lang, prompt_config
        )
        
        success_count = 0
        
        if not translations:
            batch_reason = error_message if error_message else "重试批次失败"
            logger.warning(f"重试批次失败: {batch_reason}")
            
            for task in retry_batch:
                task.retry_count += 1
                task.error_message = batch_reason
            
            return 0
        
        # 处理重试结果
        for i, task in enumerate(retry_batch):
            if i < len(translations):
                translation = translations[i]
                
                if not translation or TranslationValidator.is_error_message(translation):
                    task.retry_count += 1
                    task.error_message = "翻译超时或错误"
                    continue
                
                is_serious, reason = TranslationValidator.is_serious_failure(
                    task.original_text, translation, self.large_text_threshold
                )
                
                if not is_serious:
                    # 重试成功 - 更新翻译结果
                    with self.translation_results_lock:
                        self.translation_results[task.chunk_index] = translation
                    
                    success_count += 1
                    task.retry_count += 1
                    task.is_serious = False
                    logger.info(f"✅ 重试救援成功: 块{task.chunk_index}")
                else:
                    task.retry_count += 1
                    task.error_message = reason or "重试翻译结果仍有问题"
            else:
                task.retry_count += 1
                task.error_message = "重试结果缺失"
        
        return success_count
    
    def _execute_concurrent_retry(self, retry_batches: List[List[FailedTask]], 
                                target_lang: str, source_lang: Optional[str],
                                prompt_config: Optional[Dict[str, Any]]) -> int:
        """并发重试执行"""
        success_count = 0
        max_workers = 5
        
        try:
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                with tqdm(total=sum(len(batch) for batch in retry_batches), 
                         desc="并行重试", unit="任务") as pbar:
                    
                    future_to_batch = {}
                    for retry_batch_idx, retry_batch in enumerate(retry_batches):
                        future = executor.submit(
                            self._process_retry_batch, retry_batch, retry_batch_idx, 
                            0, target_lang, source_lang, prompt_config
                        )
                        future_to_batch[future] = retry_batch
                    
                    for future in as_completed(future_to_batch):
                        retry_batch = future_to_batch[future]
                        try:
                            batch_success = future.result()
                            success_count += batch_success
                            pbar.update(len(retry_batch))
                        except Exception as e:
                            logger.error(f"并发重试批次处理异常: {e}")
                            for task in retry_batch:
                                task.retry_count += 1
                            pbar.update(len(retry_batch))
            
            self.stats['parallel_retry_used'] += 1
            return success_count
            
        except Exception as e:
            logger.error(f"并发重试执行异常: {e}")
            for batch in retry_batches:
                for task in batch:
                    task.retry_count += 1
            return success_count
    
    def _adaptive_retry_strategy(self, target_lang: str, source_lang: Optional[str],
                               prompt_config: Optional[Dict[str, Any]]) -> int:
        """自适应重试策略"""
        retry_start_time = time.time()
        total_success_count = 0
        
        serious_tasks = [task for task in self.failed_tasks if task.is_serious]
        if not serious_tasks:
            logger.info("没有严重失败任务需要重试")
            return 0
        
        logger.info(f"开始自适应重试，发现 {len(serious_tasks)} 个严重失败任务")
        self.stats['total_retry_tasks'] = len(serious_tasks)
        
        retry_count = 0
        while retry_count < self.max_retries:
            current_retry_tasks = [task for task in self.failed_tasks 
                                 if task.is_serious and task.retry_count <= retry_count]
            
            if not current_retry_tasks:
                logger.info(f"第 {retry_count + 1} 次重试检查：没有(更多)需要重试的任务")
                break
            
            logger.info(f"第 {retry_count + 1} 次重试，处理 {len(current_retry_tasks)} 个严重失败任务")
            
            # 添加重试延迟
            delay = self.retry_delays[min(retry_count, len(self.retry_delays) - 1)]
            if delay > 0:
                logger.info(f"智能延迟: {delay:.1f}秒")
                time.sleep(delay)
            
            # 创建重试批次
            retry_batches = self._create_smart_retry_batches(current_retry_tasks, retry_count)
            
            if not retry_batches:
                logger.info(f"第 {retry_count + 1} 次重试：没有批次需要处理")
                break
            
            # 并行重试
            logger.info(f"启用并发重试，worker数量: 5")
            success_count = self._execute_concurrent_retry(
                retry_batches, target_lang, source_lang, prompt_config
            )
            
            total_success_count += success_count
            self.stats['rescued_tasks'] += success_count
            self.stats['retry_attempts'] += 1
            
            # 更新失败任务列表
            self.failed_tasks = [task for task in self.failed_tasks if task.is_serious]
            
            logger.info(f"第 {retry_count + 1} 次重试完成，恢复 {success_count} 个任务，当前剩余 {len(self.failed_tasks)} 个严重失败")
            
            retry_count += 1
        
        self.time_stats['retry_time'] = time.time() - retry_start_time
        if self.stats['total_retry_tasks'] > 0:
            self.stats['retry_success_rate'] = self.stats['rescued_tasks'] / self.stats['total_retry_tasks']
        
        logger.info(f"重试阶段完成，用时: {self.time_stats['retry_time']:.1f}秒, "
                   f"成功率: {self.stats['retry_success_rate']:.1%}")
        
        return total_success_count

def translate_text_file(
    input_filepath: str,
    output_dir: str,
    target_lang: str,
    translator: SiliconFlowTranslator,
    source_lang: Optional[str] = None,
    encoding: str = 'utf-8',
    unique_filename_base: Optional[str] = None,
    chunk_size: Optional[int] = None,
    prompt_config: Optional[Dict[str, Any]] = None,
    translation_timeout: int = DEFAULT_TRANSLATION_TIMEOUT,
    max_retries: int = DEFAULT_MAX_RETRIES,
    retry_delay: int = DEFAULT_RETRY_DELAY,
    retry_batch_size: int = DEFAULT_RETRY_BATCH_SIZE,
    max_retry_workers: int = DEFAULT_RETRY_THREADS,
    large_text_threshold: int = 50,
    max_workers: int = DEFAULT_THREADS,
    **kwargs
) -> Optional[str]:
    """
    增强的文本文件翻译函数，集成并行和重试机制。
    
    :param input_filepath: 输入文件的完整路径
    :param output_dir: 翻译结果文件存放的目录
    :param target_lang: 目标语言
    :param translator: SiliconFlowTranslator 实例
    :param source_lang: 源语言（可选）
    :param encoding: 文件编码，默认为 'utf-8'
    :param unique_filename_base: 自定义输出文件名基础（可选）
    :param chunk_size: 每次API请求合并的行数（可选）
    :param prompt_config: 前端prompt配置字典
    :param translation_timeout: 翻译超时时间（秒）
    :param max_retries: 最大重试次数
    :param retry_delay: 重试延迟（秒）
    :param retry_batch_size: 重试批次大小
    :param max_retry_workers: 重试线程数
    :param large_text_threshold: 大文本阈值
    :param max_workers: 主翻译线程数
    :param kwargs: 其他参数
    :return: 成功时返回输出文件的路径，失败时返回包含"Error:"前缀的字符串
    """
    
    try:
        # 开始总计时
        start_time = time.time()
        
        # 验证输入文件
        if not os.path.exists(input_filepath):
            logger.error(f"Input file not found: {input_filepath}")
            return f"Error: Input file not found at '{input_filepath}'."

        if not os.path.isfile(input_filepath):
            logger.error(f"Input path is not a file: {input_filepath}")
            return f"Error: Input path '{input_filepath}' is not a file."

        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)

        # 处理prompt配置
        effective_prompt_config = _prepare_prompt_config(prompt_config, kwargs)
        
        # 从prompt配置或kwargs获取chunk_size
        if chunk_size is None:
            chunk_size = _get_chunk_size_from_config(effective_prompt_config, kwargs)
        
        # 构建输出文件路径
        output_filepath = _build_output_filepath(
            input_filepath, output_dir, target_lang, unique_filename_base
        )

        # 记录翻译开始信息
        _log_translation_start(input_filepath, target_lang, effective_prompt_config, chunk_size)

        # 创建文件翻译器实例
        file_translator = FileTranslator(
            translator=translator,
            translation_timeout=translation_timeout,
            max_retries=max_retries,
            retry_delay=retry_delay,
            retry_batch_size=retry_batch_size,
            max_retry_workers=max_retry_workers,
            large_text_threshold=large_text_threshold
        )
        
        file_translator.time_stats['total_start_time'] = start_time
        
        # 读取文件
        read_start = time.time()
        try:
            with open(input_filepath, 'r', encoding=encoding) as infile:
                content = infile.read()
        except UnicodeDecodeError as e:
            logger.error(f"Error decoding file '{input_filepath}' with encoding '{encoding}': {e}")
            return f"Error: Encoding issue with input file. Try specifying a different encoding. Details: {e}"
        except Exception as e:
            logger.error(f"Error reading file '{input_filepath}': {e}")
            return f"Error: File I/O issue: {e}"
        
        file_translator.time_stats['file_read_time'] = time.time() - read_start
        
        # 将内容分割成块
        chunks = _split_content_into_chunks(content, chunk_size)
        total_chunks = len(chunks)
        
        if not chunks:
            logger.info("No content to translate")
            # 直接复制原文件
            with open(output_filepath, 'w', encoding=encoding) as outfile:
                outfile.write(content)
            return output_filepath
        
        logger.info(f"发现 {total_chunks} 个文本块需要翻译")
        
        # 重置统计
        file_translator.stats.update({
            'success': 0, 'failed': 0, 'timeout': 0, 'retried': 0, 'final_failed': 0,
            'serious_failures': 0, 'minor_issues': 0, 'retry_attempts': 0,
            'rescued_tasks': 0, 'retry_success_rate': 0.0, 'total_retry_tasks': 0,
            'parallel_retry_used': 0, 'cache_hits': 0, 'api_calls': 0, 'total_chunks': total_chunks
        })
        file_translator.failed_tasks = []
        file_translator.translation_results = {}
        
        # 清理缓存
        with _cache_lock:
            _cache.clear()
        
        # === 第一阶段：多线程批量翻译 ===
        logger.info("=== 阶段1：初始批量翻译 ===")
        main_translation_start = time.time()
        
        # 创建批次
        batches = file_translator._create_chunk_batches(chunks, chunk_size)
        logger.info(f"创建了 {len(batches)} 个批次")
        
        with ThreadPoolExecutor(max_workers=min(max_workers, len(batches))) as executor:
            # 准备批次数据
            batch_data_list = [
                (batch, i, i+1, target_lang, source_lang, effective_prompt_config)
                for i, batch in enumerate(batches)
            ]
            
            # 提交任务
            future_to_batch = {
                executor.submit(file_translator._process_chunk_batch, batch_data): batch_data[0]
                for batch_data in batch_data_list
            }
            
            # 处理结果
            with tqdm(total=total_chunks, desc="翻译进度", unit="块") as pbar:
                for future in as_completed(future_to_batch):
                    total, success, failed, error, serious_failures = future.result()
                    file_translator.stats['success'] += success
                    file_translator.stats['failed'] += failed
                    pbar.update(total)
                    
                    if error:
                        logger.warning(f"批次错误: {error}")
        
        file_translator.time_stats['main_translation_time'] = time.time() - main_translation_start
        logger.info(f"主翻译完成，耗时: {file_translator.time_stats['main_translation_time']:.1f}秒")
        
        # === 第二阶段：重试失败的任务 ===
        serious_failures = len([task for task in file_translator.failed_tasks if task.is_serious])
        if serious_failures > 0:
            logger.info(f"=== 阶段2：重试 {serious_failures} 个严重失败任务 ===")
            
            retry_success_count = file_translator._adaptive_retry_strategy(
                target_lang, source_lang, effective_prompt_config
            )
            
            logger.info(f"重试阶段完成，恢复 {retry_success_count} 个任务")
            
            # 更新统计
            file_translator.stats['success'] += retry_success_count
            file_translator.stats['failed'] -= retry_success_count
            file_translator.stats['retried'] += retry_success_count
            
            # 统计最终失败任务
            final_failures = len([task for task in file_translator.failed_tasks if task.is_serious])
            file_translator.stats['final_failed'] = final_failures
            
            if final_failures > 0:
                logger.warning(f"重试后仍有 {final_failures} 个严重失败任务保持原文")
        else:
            logger.info("=== 阶段2：没有需要重试的严重失败任务 ===")
        
        # === 第三阶段：重建完整内容 ===
        logger.info("=== 阶段3：重建文件内容 ===")
        
        final_content = _rebuild_content_from_results(
            chunks, file_translator.translation_results, file_translator.failed_tasks
        )
        
        # 写入文件
        write_start = time.time()
        try:
            with open(output_filepath, 'w', encoding=encoding) as outfile:
                outfile.write(final_content)
        except Exception as e:
            logger.error(f"Error writing output file '{output_filepath}': {e}")
            return f"Error: Failed to write output file: {e}"
        
        file_translator.time_stats['file_write_time'] = time.time() - write_start
        
        # 计算总时间
        file_translator.time_stats['total_time'] = time.time() - file_translator.time_stats['total_start_time']
        
        # 输出统计
        _log_enhanced_completion(output_filepath, file_translator, effective_prompt_config)
        
        return output_filepath
        
    except Exception as e:
        logger.exception(f"Unexpected error during file translation for '{input_filepath}': {e}")
        return f"Error: Unexpected error during file translation: {e}"


def _split_content_into_chunks(content: str, chunk_size: int) -> List[str]:
    """将内容分割成块"""
    lines = content.split('\n')
    chunks = []
    
    current_chunk_lines = []
    for line in lines:
        current_chunk_lines.append(line)
        
        if len(current_chunk_lines) >= chunk_size:
            chunks.append('\n'.join(current_chunk_lines))
            current_chunk_lines = []
    
    # 添加最后一个块
    if current_chunk_lines:
        chunks.append('\n'.join(current_chunk_lines))
    
    return chunks


def _rebuild_content_from_results(chunks: List[str], 
                                 translation_results: Dict[int, str],
                                 failed_tasks: List[FailedTask]) -> str:
    """从翻译结果重建完整内容"""
    final_chunks = []
    
    for i, original_chunk in enumerate(chunks):
        if i in translation_results:
            # 使用翻译结果
            final_chunks.append(translation_results[i])
        else:
            # 使用原文（失败的情况）
            final_chunks.append(original_chunk)
    
    return '\n'.join(final_chunks)


def _log_enhanced_completion(output_filepath: str, 
                           file_translator: FileTranslator,
                           prompt_config: Optional[Dict[str, Any]]):
    """记录增强的完成信息"""
    stats = file_translator.stats
    time_stats = file_translator.time_stats
    
    success_rate = (stats['success'] / stats['total_chunks'] * 100) if stats['total_chunks'] > 0 else 0
    
    logger.info(f"文件翻译完成: 成功率 {success_rate:.1f}%")
    logger.info(f"总耗时: {time_stats['total_time']:.1f}秒 "
               f"(主翻译: {time_stats['main_translation_time']:.1f}s, "
               f"重试: {time_stats['retry_time']:.1f}s)")
    
    print(f"\n✅ 文本文件翻译完成!")
    print(f"📄 输出: {output_filepath}")
    print(f"📊 翻译统计:")
    print(f"  ✨ 成功翻译: {stats['success']} 块")
    print(f"  🔄 重试恢复: {stats['retried']} 块")
    print(f"  ⏰ 超时失败: {stats['timeout']} 块")
    print(f"  ❌ 最终失败: {stats['final_failed']} 块")
    print(f"  🎯 成功率: {success_rate:.1f}%")
    print(f"  ⏱️ 总用时: {time_stats['total_time']:.1f}秒")
    print(f"     - 主翻译: {time_stats['main_translation_time']:.1f}秒")
    print(f"     - 重试: {time_stats['retry_time']:.1f}秒")
    print(f"     - 文件读取: {time_stats['file_read_time']:.1f}秒")
    print(f"     - 文件写入: {time_stats['file_write_time']:.1f}秒")
    
    if stats['retry_attempts'] > 0:
        print(f"  🔁 重试统计: {stats['retry_attempts']}轮，成功率{stats['retry_success_rate']:.1%}")
    
    if stats['final_failed'] > 0:
        print(f"  ⚠️ 注意: {stats['final_failed']} 个块保持原文")
    
    if prompt_config:
        mode = prompt_config.get('mode', 'none')
        print(f"  🎭 Prompt模式: {mode}")


# 保持原有的辅助函数
def _prepare_prompt_config(prompt_config: Optional[Dict[str, Any]], kwargs: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """准备和标准化prompt配置"""
    if not prompt_config and not any(k in kwargs for k in ['preserve_terms', 'glossary', 'additional_context']):
        return None
    
    config = prompt_config.copy() if prompt_config else {}
    
    if 'preserve_terms' in kwargs:
        config['preserve_terms'] = kwargs['preserve_terms']
    if 'glossary' in kwargs:
        config['glossary'] = kwargs['glossary']
    if 'additional_context' in kwargs:
        config['additional_context'] = kwargs['additional_context']
    
    config = _normalize_prompt_config(config)
    
    logger.debug(f"Prepared prompt config: {config}")
    return config


def _normalize_prompt_config(config: Dict[str, Any]) -> Dict[str, Any]:
    """标准化prompt配置格式，确保与前端格式兼容"""
    normalized = config.copy()
    
    if 'mode' not in normalized:
        if 'custom_prompt' in normalized:
            normalized['mode'] = 'custom'
        elif 'prompt_template' in normalized:
            normalized['mode'] = 'professional'
        elif any(k in normalized for k in ['preserve_terms', 'glossary', 'additional_context']):
            normalized['mode'] = 'general'
        else:
            normalized['mode'] = 'none'
    
    preserve_terms = normalized.get('preserve_terms')
    if preserve_terms:
        if isinstance(preserve_terms, str):
            terms_list = [term.strip() for term in preserve_terms.split(',') if term.strip()]
            normalized['preserve_terms'] = terms_list
        elif isinstance(preserve_terms, list):
            normalized['preserve_terms'] = [str(term).strip() for term in preserve_terms if str(term).strip()]
    
    glossary = normalized.get('glossary')
    if glossary and not isinstance(glossary, dict):
        logger.warning(f"Glossary should be a dictionary, got {type(glossary)}, ignoring")
        normalized.pop('glossary', None)
    
    if normalized.get('mode') == 'custom':
        custom_prompt = normalized.get('custom_prompt', {})
        if not custom_prompt or not isinstance(custom_prompt, dict):
            system_prompt = normalized.get('custom_system_prompt', normalized.get('system'))
            user_prompt = normalized.get('custom_user_prompt', normalized.get('user'))
            
            if system_prompt:
                normalized['custom_prompt'] = {
                    'system': system_prompt,
                    'user': user_prompt or 'Please translate the following content to {target_lang}:\n\n{content}'
                }
            else:
                logger.warning("Custom mode selected but no valid custom prompt provided, falling back to general mode")
                normalized['mode'] = 'general'
    
    if normalized.get('mode') == 'professional':
        domain = normalized.get('professional_domain', normalized.get('prompt_template', 'academic'))
        normalized['prompt_template'] = domain
    
    return normalized


def _get_chunk_size_from_config(prompt_config: Optional[Dict[str, Any]], kwargs: Dict[str, Any]) -> int:
    """从配置中获取chunk_size"""
    if 'chunk_size' in kwargs:
        return kwargs['chunk_size']
    
    if prompt_config:
        max_chars = prompt_config.get('max_chars_per_chunk')
        if max_chars:
            estimated_lines = max(10, max_chars // 80)
            return min(estimated_lines, 2000)
        
        max_units = prompt_config.get('max_units_per_chunk')
        if max_units:
            return max(10, min(max_units, 2000))
    
    return DEFAULT_CHUNK_SIZE


def _build_output_filepath(input_filepath: str, output_dir: str, target_lang: str, unique_filename_base: Optional[str]) -> str:
    """构建输出文件路径"""
    base_name = os.path.basename(input_filepath)
    name_without_ext, ext = os.path.splitext(base_name)
    
    if unique_filename_base:
        output_filename = f"{unique_filename_base}_translated_{target_lang.replace(' ', '_').lower()}{ext}"
    else:
        output_filename = f"{name_without_ext}_translated_{target_lang.replace(' ', '_').lower()}{ext}"
    
    return os.path.join(output_dir, output_filename)


def _log_translation_start(input_filepath: str, target_lang: str, prompt_config: Optional[Dict[str, Any]], chunk_size: int):
    """记录翻译开始的详细信息"""
    logger.info(f"Starting enhanced text file translation: '{input_filepath}' to '{target_lang}'")
    logger.info(f"Chunk size: {chunk_size} lines")
    
    if prompt_config:
        mode = prompt_config.get('mode', 'none')
        logger.info(f"Prompt mode: {mode}")
        
        if mode == 'professional':
            domain = prompt_config.get('prompt_template', 'academic')
            logger.info(f"Professional domain: {domain}")
        elif mode == 'custom':
            logger.info("Using custom prompt configuration")
        
        enhance_features = []
        if prompt_config.get('preserve_terms'):
            terms_count = len(prompt_config['preserve_terms'])
            enhance_features.append(f"preserve {terms_count} terms")
        if prompt_config.get('glossary'):
            glossary_count = len(prompt_config['glossary'])
            enhance_features.append(f"apply {glossary_count} glossary entries")
        if prompt_config.get('additional_context'):
            enhance_features.append("use additional context")
        
        if enhance_features:
            logger.info(f"Translation enhancements: {', '.join(enhance_features)}")
    else:
        logger.info("Using default translation mode (no prompt)")


# 高级翻译函数（保持向后兼容）
def translate_text_file_advanced(
    input_filepath: str,
    output_dir: str,
    target_lang: str,
    translator: SiliconFlowTranslator,
    source_lang: Optional[str] = None,
    encoding: str = 'utf-8',
    unique_filename_base: Optional[str] = None,
    chunk_size: Optional[int] = None,
    prompt_config: Optional[Dict[str, Any]] = None,
    preserve_line_structure: bool = True,
    quality_check: bool = False,
    **kwargs
) -> Optional[str]:
    """
    高级文本文件翻译函数，支持更精细的控制。
    现在默认集成了并行和重试机制。
    """
    
    logger.info("Using advanced text file translation with enhanced parallel and retry features")
    
    if prompt_config:
        advanced_config = prompt_config.copy()
        
        if prompt_config.get('mode') == 'advanced':
            logger.info("Advanced prompt mode detected")
        
        if quality_check:
            advanced_config['ensure_consistency'] = True
            advanced_config['quality_level'] = 'high'
            logger.info("Quality check enabled")
        
        if preserve_line_structure:
            advanced_config['preserve_formatting'] = True
            logger.info("Line structure preservation enabled")
    else:
        advanced_config = {'mode': 'general'} if quality_check or preserve_line_structure else None
    
    return translate_text_file(
        input_filepath=input_filepath,
        output_dir=output_dir,
        target_lang=target_lang,
        translator=translator,
        source_lang=source_lang,
        encoding=encoding,
        unique_filename_base=unique_filename_base,
        chunk_size=chunk_size,
        prompt_config=advanced_config,
        **kwargs
    )


# 批量翻译函数（增强版）
def translate_multiple_text_files(
    input_filepaths: List[str],
    output_dir: str,
    target_lang: str,
    translator: SiliconFlowTranslator,
    source_lang: Optional[str] = None,
    encoding: str = 'utf-8',
    chunk_size: Optional[int] = None,
    prompt_config: Optional[Dict[str, Any]] = None,
    **kwargs
) -> Dict[str, List[str]]:
    """
    批量翻译多个文本文件，集成并行和重试机制。
    """
    results = {
        'successful': [],
        'failed': [],
        'errors': []
    }
    
    total_files = len(input_filepaths)
    logger.info(f"Starting enhanced batch translation of {total_files} files")
    
    if prompt_config:
        mode = prompt_config.get('mode', 'none')
        logger.info(f"Batch translation using prompt mode: {mode}")
    
    with tqdm(total=total_files, desc="批量翻译", unit="files") as pbar:
        for i, input_filepath in enumerate(input_filepaths):
            try:
                base_name = os.path.splitext(os.path.basename(input_filepath))[0]
                unique_base = f"{base_name}_batch_{i+1:03d}"
                
                result = translate_text_file(
                    input_filepath=input_filepath,
                    output_dir=output_dir,
                    target_lang=target_lang,
                    translator=translator,
                    source_lang=source_lang,
                    encoding=encoding,
                    unique_filename_base=unique_base,
                    chunk_size=chunk_size,
                    prompt_config=prompt_config,
                    **kwargs
                )
                
                if result and not result.startswith("Error:"):
                    results['successful'].append(result)
                    logger.info(f"✓ Successfully translated: {os.path.basename(input_filepath)}")
                else:
                    results['failed'].append(input_filepath)
                    results['errors'].append(result or "Unknown error")
                    logger.error(f"✗ Failed to translate: {os.path.basename(input_filepath)}")
                    
            except Exception as e:
                results['failed'].append(input_filepath)
                results['errors'].append(str(e))
                logger.error(f"✗ Exception translating {os.path.basename(input_filepath)}: {e}")
            
            pbar.update(1)
    
    # 输出批量翻译结果
    success_count = len(results['successful'])
    failed_count = len(results['failed'])
    success_rate = (success_count / total_files * 100) if total_files > 0 else 0
    
    logger.info(f"Enhanced batch translation completed: {success_count}/{total_files} files successful ({success_rate:.1f}%)")
    
    print(f"\n{'='*50}")
    print(f"增强批量翻译完成！")
    print(f"{'='*50}")
    print(f"总文件数: {total_files}")
    print(f"成功翻译: {success_count} ({success_rate:.1f}%)")
    print(f"翻译失败: {failed_count}")
    print(f"✨ 特性: 并行处理 + 智能重试 + 缓存优化")
    
    if prompt_config:
        mode = prompt_config.get('mode', 'none')
        print(f"Prompt模式: {mode}")
        if mode == 'professional':
            print(f"专业领域: {prompt_config.get('prompt_template', 'academic')}")
    
    if failed_count > 0:
        print(f"\n失败文件列表:")
        for i, failed_file in enumerate(results['failed']):
            error_msg = results['errors'][i] if i < len(results['errors']) else "Unknown error"
            print(f"  {i+1}. {os.path.basename(failed_file)}")
            print(f"     错误: {error_msg[:100]}...")
    
    return results


if __name__ == '__main__':
    import sys
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s'
    )
    
    print("=== 增强文本文件翻译器（并行 + 重试 + 缓存）===")
    print("🚀 新特性：")
    print("✅ 多线程并行处理 - 大幅提高翻译速度")
    print("✅ 智能重试机制 - 自动恢复失败任务")
    print("✅ 失败检测分类 - 区分严重/轻微失败")
    print("✅ 缓存机制 - 避免重复翻译")
    print("✅ 超时处理 - 防止长时间阻塞")
    print("✅ 详细统计 - 完整的时间和成功率追踪")
    print("✅ 完整Prompt支持 - 保持所有原有功能")
    print("✅ 分阶段处理 - 主翻译 → 重试 → 结果重建")
    
    print(f"\n核心改进:")
    print("• 两阶段翻译：初始并行翻译 + 智能重试")
    print("• 自适应批次：根据失败类型调整批次大小")
    print("• 并行重试：失败任务也使用5线程并发处理")
    print("• 结果重建：从翻译结果和失败任务重建完整文件")
    print("• 时间统计：分别统计读取、翻译、重试、写入时间")
    
    print(f"\n支持的Prompt模式:")
    print("• none: 无prompt，简单翻译")
    print("• general: 通用翻译模式")
    print("• professional: 专业领域翻译")
    print("• custom: 完全自定义prompt")
    
    print(f"\n配置参数:")
    print(f"• 默认线程数: {DEFAULT_THREADS}")
    print(f"• 重试线程数: {DEFAULT_RETRY_THREADS}（固定）")
    print(f"• 翻译超时: {DEFAULT_TRANSLATION_TIMEOUT}秒")
    print(f"• 最大重试: {DEFAULT_MAX_RETRIES}次")
    print(f"• 重试批次: {[10, 5, 2, 1, 1, 1, 1, 1]}")
    
    if len(sys.argv) > 1:
        print(f"\n命令行参数: {sys.argv[1:]}")
    else:
        print(f"\n使用示例:")
        print("# 基础增强翻译")
        print("translate_text_file('input.txt', 'output/', 'Chinese', translator)")
        print("")
        print("# 高并发翻译")
        print("translate_text_file('large.txt', 'output/', 'Chinese', translator,")
        print("                   max_workers=8, chunk_size=20)")
        print("")
        print("# 专业翻译 + 重试优化")
        print("config = {'mode': 'professional', 'prompt_template': 'academic'}")
        print("translate_text_file('paper.txt', 'output/', 'Chinese', translator,")
        print("                   prompt_config=config, max_retries=15)")
        
    print(f"\n✅ 增强版文本翻译器就绪！支持并行处理和智能重试！")
