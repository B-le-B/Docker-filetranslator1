document.addEventListener('DOMContentLoaded', function () {
    // --- 元素获取 ---
    const sourceLangSelect = document.getElementById('source_lang');
    const targetLangSelect = document.getElementById('target_lang');
    const swapLangButton = document.getElementById('swap_lang_button');

    const modeTextButton = document.getElementById('mode_text_button');
    const modeFileButton = document.getElementById('mode_file_button');
    const textInputSection = document.getElementById('text_input_section');
    const fileInputSection = document.getElementById('file_input_section');

    const textInput = document.getElementById('text_input');
    const fileInput = document.getElementById('file_input');
    const fileNameDisplay = document.getElementById('file_name_display');
    const encodingGroup = document.getElementById('encoding_group');
    const docxProcessingMethodGroup = document.getElementById('docx_processing_method_group');
    const docxProcessingMethodSelect = document.getElementById('docx_processing_method');

    // 新增：模板文件相关元素
    const templateUploadArea = document.getElementById('template_upload_area');
    const templateFileInput = document.getElementById('template_file_input');
    const templateFileNameDisplay = document.getElementById('template_file_name_display');
    const clearTemplateButton = document.getElementById('clear_template_button');

    const outputFolderPathGroup = document.querySelector('.file-output-path');
    const outputFolderPathInput = document.getElementById('output_folder_path');
    const copyDefaultPathButton = document.getElementById('copy_default_path_button');

    const translateButton = document.getElementById('translate_button');
    const statusMessage = document.getElementById('status_message');
    const translatedTextDisplay = document.getElementById('translated_text_display');
    const downloadArea = document.getElementById('download_area');
    const downloadLink = document.getElementById('download_link');
    const logOutputPre = document.getElementById('log_output');

    const apiPlatformSelect = document.getElementById('api_platform');
    const apiKeyInput = document.getElementById('api_key');
    const baseUrlInput = document.getElementById('base_url');
    const modelInput = document.getElementById('model');
    const apiKeyHint = document.querySelector('#api_key + .hint');
    const copyTranslatedTextButton = document.getElementById('copy_translated_text_button');

    const copySourceTextButton = document.getElementById('copy_source_text_button');
    const cleanPdfLineBreaksButton = document.getElementById('clean_pdf_line_breaks_button'); // 新增
    const clearSourceTextButton = document.getElementById('clear_source_text_button');

    // --- 新增：Prompt相关元素 ---
    const translationModeRadios = document.querySelectorAll('input[name="translation_mode"]');
    const professionalDomainGroup = document.getElementById('professional_domain_group');
    const professionalDomainSelect = document.getElementById('professional_domain');
    const customPromptGroup = document.getElementById('custom_prompt_group');
    const customSystemPrompt = document.getElementById('custom_system_prompt');
    const customUserPrompt = document.getElementById('custom_user_prompt');
    const preserveTermsInput = document.getElementById('preserve_terms');
    const glossaryContainer = document.getElementById('glossary_container');
    const addGlossaryBtn = document.getElementById('add_glossary_btn');
    const additionalContextTextarea = document.getElementById('additional_context');
    const maxUnitsPerChunkInput = document.getElementById('max_units_per_chunk');
    const maxCharsPerChunkInput = document.getElementById('max_chars_per_chunk');

    // --- 新增：获取需要动态显示/隐藏的 Prompt 设置组 ---
    const preserveTermsGroup = document.getElementById('preserve_terms_group');
    const glossaryGroup = document.getElementById('glossary_group');
    const additionalContextGroup = document.getElementById('additional_context_group');

    // --- 新增：配置管理相关元素 ---
    const saveConfigBtn = document.getElementById('save_config_btn');
    const savedConfigsDropdown = document.getElementById('saved_configs');
    const loadConfigBtn = document.getElementById('load_config_btn');
    const deleteConfigBtn = document.getElementById('delete_config_btn');
    const exportConfigBtn = document.getElementById('export_config_btn');
    const importConfigFileInput = document.getElementById('import_config_file');

    const languageSelectionBar = document.getElementById('language_selection_bar'); // For auto-scrolling

    // --- 新增：处理方法映射函数 ---
    function getProcessingMethodForFile(file, docxMethod) {
        const fileExtension = file.name.split('.').pop().toLowerCase();
        
        switch (fileExtension) {
            case 'txt':
                return 'text_translator';
            case 'docx':
                return docxMethod || 'docx_full_translator';
            case 'pptx':
                return 'pptx_full_translator';
            case 'xlsx':
            case 'xls':
                return 'excel_full_translator';
            default:
                return null;
        }
    }

    // --- 新增：模板文件处理函数 ---

    // 检查当前选择的处理方法是否支持模板
    function isTemplateSupported(docxMethod) {
        const templateSupportedMethods = [

            'docx_markdown_translator'
            // 可以根据需要添加其他支持模板的方法



        ];
        return templateSupportedMethods.includes(docxMethod);
    }


    // 更新后的显示/隐藏函数
    function updateTemplateUploadVisibility() {
        if (!templateUploadArea || !docxProcessingMethodSelect) return;

        const selectedMethod = docxProcessingMethodSelect.value;
        const shouldShow = isTemplateSupported(selectedMethod);

        templateUploadArea.style.display = shouldShow ? 'block' : 'none';

        if (!shouldShow) {
            clearTemplateFile();
        }
    }

    // 清除模板文件函数保持不变，但文案更简洁
    function clearTemplateFile() {
        if (templateFileInput) templateFileInput.value = '';
        if (templateFileNameDisplay) templateFileNameDisplay.textContent = '未选择';
        if (clearTemplateButton) clearTemplateButton.style.display = 'none';
    }

    // 处理模板文件选择
    function handleTemplateFileSelection() {
        const file = templateFileInput.files[0];
        if (file) {
            if (file.name.toLowerCase().endsWith('.docx')) {
                templateFileNameDisplay.textContent = file.name;
                clearTemplateButton.style.display = 'inline-block';
                appendLog(`已选择模板文件: ${file.name}`);
            } else {
                alert('模板文件必须是 .docx 格式');
                clearTemplateFile();
            }
        } else {
            clearTemplateFile();
        }
    }

    // 清除模板文件
    function clearTemplateFile() {
        if (templateFileInput) templateFileInput.value = '';
        if (templateFileNameDisplay) templateFileNameDisplay.textContent = '未选择模板文件';
        if (clearTemplateButton) clearTemplateButton.style.display = 'none';
        appendLog('已清除模板文件选择');
    }

    // --- 语言列表 ---
    const LANGUAGES = [
        { value: "", text: "自动检测" },
        { value: "简体中文", text: "简体中文" },
        { value: "繁體中文", text: "繁體中文" },
        { value: "English", text: "English" },
        { value: "日本語", text: "日本語 (Japanese)" },
        { value: "Français", text: "Français (French)" },
        { value: "Deutsch", text: "Deutsch (German)" },
        { value: "Español", text: "Español (Spanish)" },
        { value: "Português", text: "Português (Portuguese)" },
        { value: "Русский", text: "Русский (Russian)" },
        { value: "العربية", text: "العربية (Arabic)" },
    ];

    // --- API平台配置保持不变 ---
    const API_PLATFORM_CONFIGS = {
        "siliconflow": {
            baseUrl: "https://api.siliconflow.cn/v1",
            model: "THUDM/GLM-4-9B-0414",
            apiKeyEnvHint: "SILICONFLOW_API_KEY"
        },
        "zhipu": {
            baseUrl: "https://open.bigmodel.cn/api/paas/v4",
            model: "GLM-4-Flash-250414",
            apiKeyEnvHint: "ZHIPU_API_KEY"
        },            
        "modelscope": {
            baseUrl: "https://api-inference.modelscope.cn/v1",
            model: "Qwen/Qwen2.5-72B-Instruct",
            apiKeyEnvHint: "MODELSCOPE_API_KEY"
        },
        "openrouter": {
            baseUrl: "https://openrouter.ai/api/v1",
            model: "google/gemini-2.0-flash-exp:free",
            apiKeyEnvHint: "OPENROUTER_API_KEY"
        },
        "openai": {
            baseUrl: "https://api.openai.com/v1",
            model: "gpt-3.5-turbo",
            apiKeyEnvHint: "OPENAI_API_KEY"
        },
        "ollama": {
            baseUrl: "http://localhost:11434/v1",
            model: "llama3",
            apiKeyEnvHint: "OLLAMA_API_KEY (通常不需要或任意字符串)"
        },
        "custom": {
            baseUrl: "",
            model: "",
            apiKeyEnvHint: "自定义平台的环境变量"
        },
        "volcengine": {
            baseUrl: "translate.volcengineapi.com",
            model: "",
            apiKeyEnvHint: "自定义平台的环境变量"
        }
    };

    // --- 新增：配置管理相关函数 ---

    // 保存当前配置
    async function saveCurrentConfig() {
        const configName = prompt('请输入配置名称：');
        if (!configName) return;

        const config = {
            name: configName,
            prompt_config: await getPromptConfig(),
            api_platform: apiPlatformSelect?.value,
            source_lang: sourceLangSelect?.value,
            target_lang: targetLangSelect?.value,
            api_key: apiKeyInput?.value,
            base_url: baseUrlInput?.value,
            model: modelInput?.value,
            encoding: document.getElementById('encoding') ? document.getElementById('encoding').value : 'utf-8',
            docx_processing_method: docxProcessingMethodSelect?.value || 'docx_markdown_translator',
            output_folder_path: outputFolderPathInput?.value || '',
            max_units_per_chunk: parseInt(maxUnitsPerChunkInput?.value) || 10,
            max_chars_per_chunk: parseInt(maxCharsPerChunkInput?.value) || 2000
        };

        fetch('/api/save-config', {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify(config)
        })
            .then(response => response.json())
            .then(data => {
                if (data.success) {
                    appendLog(`配置已保存：${configName}`);
                    alert('配置保存成功！');
                    loadConfigList();
                } else {
                    alert('配置保存失败：' + (data.error || '未知错误'));
                }
            })
            .catch(error => {
                console.error('Error saving config:', error);
                alert('保存配置时出错');
            });
    }

    // 加载配置列表
    function loadConfigList() {
        fetch('/api/list-configs')
            .then(response => {
                if (!response.ok) {
                    throw new Error(`Failed to load config list: ${response.status} ${response.statusText}`);
                }
                return response.json();
            })
            .then(data => {
                appendLog(`已加载配置列表 (${data.configs.length} 个配置)。`);
                updateConfigDropdown(data.configs);
            })
            .catch(error => {
                console.error('Error loading configs:', error);
                appendLog(`加载配置列表失败: ${error.message}。请检查后端是否运行正常并提供了 /api/list-configs 接口。`);
            });
    }

    // 更新配置下拉框
    function updateConfigDropdown(configs) {
        if (!savedConfigsDropdown) return;

        savedConfigsDropdown.innerHTML = '';
        const defaultOption = document.createElement('option');
        defaultOption.value = '';
        defaultOption.textContent = '-- 选择配置 --';
        savedConfigsDropdown.appendChild(defaultOption);

        if (configs && configs.length > 0) {
            configs.forEach(config => {
                const option = document.createElement('option');
                option.value = config.name;
                option.textContent = config.name;
                savedConfigsDropdown.appendChild(option);
            });
        } else {
            appendLog("未找到任何保存的配置。");
        }
    }

    // 加载选定的配置
    function loadSelectedConfig(configName) {
        if (!configName) {
            alert('请选择一个要加载的配置。');
            return;
        }

        fetch(`/api/load-config/${encodeURIComponent(configName)}`)
            .then(response => {
                if (!response.ok) {
                    throw new Error(`Failed to load config: ${response.status} ${response.statusText}`);
                }
                return response.json();
            })
            .then(config => {
                if (config.prompt_config) {
                    applyPromptConfig(config.prompt_config);
                }
                if (config.api_platform && apiPlatformSelect) {
                    apiPlatformSelect.value = config.api_platform;
                    updateApiFieldsForPlatform(config.api_platform);
                }
                if (config.api_key !== undefined && apiKeyInput) apiKeyInput.value = config.api_key;
                if (config.base_url !== undefined && baseUrlInput) baseUrlInput.value = config.base_url;
                if (config.model !== undefined && modelInput) modelInput.value = config.model;
                if (config.source_lang !== undefined && sourceLangSelect) {
                    sourceLangSelect.value = config.source_lang;
                }
                if (config.target_lang !== undefined && targetLangSelect) {
                    targetLangSelect.value = config.target_lang;
                }
                if (document.getElementById('encoding') && config.encoding !== undefined) {
                    document.getElementById('encoding').value = config.encoding;
                }
                if (docxProcessingMethodSelect && config.docx_processing_method !== undefined) {
                    docxProcessingMethodSelect.value = config.docx_processing_method;
                    updateTemplateUploadVisibility(); // 新增：更新模板上传区域显示
                }
                if (outputFolderPathInput && config.output_folder_path !== undefined) {
                    outputFolderPathInput.value = config.output_folder_path;
                }
                if (maxUnitsPerChunkInput && config.max_units_per_chunk !== undefined) {
                    maxUnitsPerChunkInput.value = config.max_units_per_chunk;
                }
                if (maxCharsPerChunkInput && config.max_chars_per_chunk !== undefined) {
                    maxCharsPerChunkInput.value = config.max_chars_per_chunk;
                }

                appendLog(`已加载配置：${configName}`);
                alert('配置加载成功！');
            })
            .catch(error => {
                console.error('Error loading config:', error);
                alert(`加载配置失败：${error.message}。请检查后端是否运行正常并提供了 /api/load-config 接口。`);
            });
    }

    // 删除选定的配置
    function deleteSelectedConfig() {
        if (!savedConfigsDropdown) return;
        const configName = savedConfigsDropdown.value;
        if (!configName) {
            alert('请选择一个要删除的配置。');
            return;
        }

        if (!confirm(`确定要删除配置 "${configName}" 吗？`)) {
            return;
        }

        fetch(`/api/delete-config/${encodeURIComponent(configName)}`, {
            method: 'DELETE'
        })
            .then(response => {
                if (!response.ok) {
                    throw new Error(`Failed to delete config: ${response.status} ${response.statusText}`);
                }
                return response.json();
            })
            .then(data => {
                if (data.success) {
                    appendLog(`配置已删除：${configName}`);
                    alert('配置删除成功！');
                    loadConfigList();
                } else {
                    alert('配置删除失败：' + (data.error || '未知错误'));
                }
            })
            .catch(error => {
                console.error('Error deleting config:', error);
                alert(`删除配置时出错：${error.message}。请检查后端是否运行正常并提供了 /api/delete-config 接口。`);
            });
    }

    // 导出选定的配置 (已修改，支持重命名)
    function exportSelectedConfig() {
        if (!savedConfigsDropdown) return;
        const configName = savedConfigsDropdown.value;
        if (!configName) {
            alert('请选择一个要导出的配置。');
            return;
        }

        // 弹出prompt让用户输入文件名，默认是配置名称+.json
        let downloadFilename = prompt('请输入导出文件名:', `${configName}.json`);

        if (downloadFilename === null) { // 用户点击了取消
            appendLog('导出操作已取消。');
            return;
        }

        // 确保文件名有.json扩展名
        if (!downloadFilename.toLowerCase().endsWith('.json')) {
            downloadFilename += '.json';
        }

        fetch(`/api/load-config/${encodeURIComponent(configName)}`)
            .then(response => {
                if (!response.ok) {
                    throw new Error(`Failed to load config for export: ${response.status} ${response.statusText}`);
                }
                return response.json();
            })
            .then(config => {
                const configJson = JSON.stringify(config, null, 2);
                const blob = new Blob([configJson], { type: 'application/json' });
                const url = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = url;
                a.download = downloadFilename; // 使用用户指定的文件名
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(url);
                appendLog(`配置已导出为：${downloadFilename}`);
                alert('配置导出成功！');
            })
            .catch(error => {
                console.error('Error exporting config:', error);
                alert(`导出配置失败：${error.message}。请检查后端是否运行正常并提供了 /api/load-config 接口。`);
            });
    }

    // 导入配置
    function importConfig(event) {
        const file = event.target.files[0];
        if (!file) {
            return;
        }

        const reader = new FileReader();
        reader.onload = function (e) {
            try {
                const config = JSON.parse(e.target.result);
                const configName = config.name || prompt('请输入导入配置的名称：', file.name.replace('.json', ''));
                if (!configName) {
                    alert('导入配置需要一个名称。');
                    return;
                }
                config.name = configName;

                fetch('/api/save-config', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify(config)
                })
                    .then(response => {
                        if (!response.ok) {
                            throw new Error(`Failed to save imported config: ${response.status} ${response.statusText}`);
                        }
                        return response.json();
                    })
                    .then(data => {
                        if (data.success) {
                            appendLog(`配置已导入并保存为：${configName}`);
                            alert('配置导入成功！');
                            loadConfigList();
                        } else {
                            alert('配置导入失败：' + (data.error || '未知错误'));
                        }
                    })
                    .catch(error => {
                        console.error('Error importing and saving config:', error);
                        alert(`导入配置时出错：${error.message}。请检查后端是否运行正常并提供了 /api/save-config 接口。`);
                    });

            } catch (error) {
                alert('文件内容不是有效的JSON格式或配置格式不正确。');
                appendLog('导入文件解析失败: ' + error.message);
            }
        };
        reader.readAsText(file);
        event.target.value = '';
    }

    // 应用prompt配置
    function applyPromptConfig(config) {
        const modeRadio = document.querySelector(`input[name="translation_mode"][value="${config.mode}"]`);
        if (modeRadio) {
            modeRadio.checked = true;
            updatePromptUI();
        }

        if (config.prompt_template && professionalDomainSelect) {
            professionalDomainSelect.value = config.prompt_template;
        }
        if (config.custom_prompt) {
            if (customSystemPrompt) customSystemPrompt.value = config.custom_prompt.system || '';
            if (customUserPrompt) customUserPrompt.value = config.custom_prompt.user || '';
        } else {
            if (customSystemPrompt) customSystemPrompt.value = '';
            if (customUserPrompt) customUserPrompt.value = '{content}';
        }

        if (config.preserve_terms && preserveTermsInput) {
            preserveTermsInput.value = config.preserve_terms.join(', ');
        } else {
            if (preserveTermsInput) preserveTermsInput.value = '';
        }

        if (config.glossary) {
            glossaryContainer.innerHTML = '';
            Object.entries(config.glossary).forEach(([source, target]) => {
                addGlossaryItem(source, target);
            });
            if (Object.keys(config.glossary).length === 0) {
                addGlossaryItem();
            }
        } else {
            glossaryContainer.innerHTML = '';
            addGlossaryItem();
        }

        if (config.additional_context !== undefined && additionalContextTextarea) {
            additionalContextTextarea.value = config.additional_context;
        } else {
            if (additionalContextTextarea) additionalContextTextarea.value = '';
        }

        if (config.max_units_per_chunk !== undefined && maxUnitsPerChunkInput) {
            maxUnitsPerChunkInput.value = config.max_units_per_chunk;
        }
        if (config.max_chars_per_chunk !== undefined && maxCharsPerChunkInput) {
            maxCharsPerChunkInput.value = config.max_chars_per_chunk;
        }
    }

    // 获取默认批处理配置（不设置到输入框，只用于程序运行时）
    function getDefaultBatchConfigFromBackend() {
        return fetch('/api/default-batch-config')
            .then(response => {
                if (!response.ok) {
                    throw new Error(`Failed to load default batch config: ${response.status} ${response.statusText}`);
                }
                return response.json();
            })
            .catch(error => {
                console.error('Error loading default batch config:', error);
                // 如果后端获取失败，返回翻译器模块的默认值作为备用
                return {
                    max_units_per_chunk: 10,   // DEFAULT_CHUNK_SIZE
                    max_chars_per_chunk: 2000  // DEFAULT_MAX_CHARS
                };
            });
    }

    // 加载prompt模板列表
    function loadPromptTemplates() {
        fetch('/api/prompt-templates')
            .then(response => {
                if (!response.ok) {
                    throw new Error(`Failed to load prompt templates: ${response.status} ${response.statusText}`);
                }
                return response.json();
            })
            .then(templates => {
                if (professionalDomainSelect) {
                    professionalDomainSelect.innerHTML = '';
                    Object.entries(templates).forEach(([key, template]) => {
                        const option = document.createElement('option');
                        option.value = key;
                        option.textContent = template.name;
                        option.title = template.description;
                        professionalDomainSelect.appendChild(option);
                    });
                    if (!professionalDomainSelect.value && templates.default) {
                        professionalDomainSelect.value = 'default';
                    } else if (professionalDomainSelect.options.length > 0) {
                        professionalDomainSelect.value = professionalDomainSelect.options[0].value;
                    }
                }
                appendLog("Prompt模板已加载。");
            })
            .catch(error => {
                console.error('Error loading prompt templates:', error);
                appendLog(`加载Prompt模板失败: ${error.message}。`);
            });
    }

    function updateApiFieldsForPlatform(platform) {
        const config = API_PLATFORM_CONFIGS[platform] || API_PLATFORM_CONFIGS["custom"];
        if (baseUrlInput) baseUrlInput.value = config.baseUrl;
        if (modelInput) modelInput.value = config.model;
        if (apiKeyHint) {
            apiKeyHint.textContent = `请输入各平台的API KEY (例如 ${config.apiKeyEnvHint})。`;
        }
        if (apiPlatformSelect) {
            appendLog(`API 平台切换为: ${apiPlatformSelect.options[apiPlatformSelect.selectedIndex].text}. Base URL 和 Model 已更新。`);
        }
    }

    // --- 辅助函数 ---
    function populateLanguageSelect(selectElement, languages, includeAutoDetect = false) {
        if (!selectElement) return;
        selectElement.innerHTML = '';
        const filteredLanguages = includeAutoDetect ? languages : languages.filter(lang => lang.value !== "");

        filteredLanguages.forEach(lang => {
            let option = document.createElement('option');
            option.value = lang.value;
            option.textContent = lang.text;
            selectElement.appendChild(option);
        });
    }

    function appendLog(message) {
        if (!logOutputPre) return;
        const now = new Date();
        const timestamp = now.toLocaleTimeString([], { hour: '2-digit', minute: '2-digit', second: '2-digit' });
        logOutputPre.textContent += `[${timestamp}] ${message}\n`;
        logOutputPre.scrollTop = logOutputPre.scrollHeight;
    }

    function resetOutputArea() {
        if (translatedTextDisplay) translatedTextDisplay.value = '';
        if (downloadArea) downloadArea.style.display = 'none';
        if (downloadLink) {
            downloadLink.href = '#';
            downloadLink.textContent = '';
            downloadLink.style.display = 'none';
        }
        if (statusMessage) {
            statusMessage.textContent = '';
            statusMessage.className = 'status-message';
        }
    }

    function switchToTextMode() {
        if (modeTextButton) modeTextButton.classList.add('active');
        if (modeFileButton) modeFileButton.classList.remove('active');
        if (textInputSection) textInputSection.classList.add('active');
        if (fileInputSection) fileInputSection.classList.remove('active');

        if (outputFolderPathGroup) outputFolderPathGroup.style.display = 'none';
        if (encodingGroup) encodingGroup.style.display = 'none';
        if (docxProcessingMethodGroup) docxProcessingMethodGroup.style.display = 'none';
        if (templateUploadArea) templateUploadArea.style.display = 'none'; // 新增：隐藏模板上传区域
        appendLog("切换到文字翻译模式。");
    }

    function switchToFileMode() {
        if (modeFileButton) modeFileButton.classList.add('active');
        if (modeTextButton) modeTextButton.classList.remove('active');
        if (textInputSection) textInputSection.classList.remove('active');
        if (fileInputSection) fileInputSection.classList.add('active');

        if (outputFolderPathGroup) outputFolderPathGroup.style.display = 'block';
        updateFileSpecificOptions();
        appendLog("切换到文档翻译模式。");
    }

    function updateFileSpecificOptions() {
        const currentFile = fileInput && fileInput.files[0];
        if (currentFile && modeFileButton && modeFileButton.classList.contains('active')) {
            const fileExtension = currentFile.name.split('.').pop().toLowerCase();
            if (fileExtension === 'txt') {
                if (encodingGroup) encodingGroup.style.display = 'block';
                if (docxProcessingMethodGroup) docxProcessingMethodGroup.style.display = 'none';
                if (templateUploadArea) templateUploadArea.style.display = 'none'; // 新增：txt文件不显示模板区域
                appendLog(`已选择 .txt 文档: ${currentFile.name}，显示编码选项。`);
            } else if (fileExtension === 'docx') {
                if (encodingGroup) encodingGroup.style.display = 'none';
                if (docxProcessingMethodGroup) docxProcessingMethodGroup.style.display = 'block';
                updateTemplateUploadVisibility(); // 新增：更新模板上传区域显示
                appendLog(`已选择 .docx 文档: ${currentFile.name}，显示 DOCX 处理方法选项。`);
            } else {
                if (encodingGroup) encodingGroup.style.display = 'none';
                if (docxProcessingMethodGroup) docxProcessingMethodGroup.style.display = 'none';
                if (templateUploadArea) templateUploadArea.style.display = 'none'; // 新增：其他文件不显示模板区域
                appendLog(`已选择 .${fileExtension} 文档: ${currentFile.name}，隐藏特定文件选项。`);
            }
        } else {
            if (encodingGroup) encodingGroup.style.display = 'none';
            if (docxProcessingMethodGroup) docxProcessingMethodGroup.style.display = 'none';
            if (templateUploadArea) templateUploadArea.style.display = 'none'; // 新增：没有文件时隐藏模板区域
        }
    }

    // --- 新增：Prompt相关函数 (已修改) ---
    function updatePromptUI() {
        const selectedMode = document.querySelector('input[name="translation_mode"]:checked')?.value || 'none';

        // 确定是否显示通用的Prompt设置
        const showCommonPromptSettings = (selectedMode !== 'none');

        // 根据模式更新日志
        switch (selectedMode) {
            case 'none':
                appendLog("切换到无 Prompt 模式。");
                break;
            case 'general':
                appendLog("切换到通用翻译模式。");
                break;
            case 'professional':
                appendLog("切换到专业翻译模式。");
                break;
            case 'custom':
                appendLog("切换到自定义Prompt模式。");
                break;
        }

        // 更新通用Prompt设置的可见性
        if (preserveTermsGroup) preserveTermsGroup.style.display = showCommonPromptSettings ? 'block' : 'none';
        if (glossaryGroup) glossaryGroup.style.display = showCommonPromptSettings ? 'block' : 'none';
        if (additionalContextGroup) additionalContextGroup.style.display = showCommonPromptSettings ? 'block' : 'none';

        // 更新特定模式的设置可见性
        if (professionalDomainGroup) professionalDomainGroup.style.display = (selectedMode === 'professional') ? 'block' : 'none';
        if (customPromptGroup) customPromptGroup.style.display = (selectedMode === 'custom') ? 'block' : 'none';
    }

    function addGlossaryItem(source = '', target = '') {
        const glossaryItem = document.createElement('div');
        glossaryItem.className = 'glossary-item';
        glossaryItem.innerHTML = `
            <input type="text" class="glossary-source" placeholder="原文术语" value="${source}">
            <span>→</span>
            <input type="text" class="glossary-target" placeholder="目标翻译" value="${target}">
            <button type="button" class="remove-glossary-btn" title="删除">×</button>
        `;

        const removeBtn = glossaryItem.querySelector('.remove-glossary-btn');
        removeBtn.addEventListener('click', function () {
            glossaryItem.remove();
            appendLog("删除了一个术语表项");
        });

        glossaryContainer.appendChild(glossaryItem);
    }

    function getGlossary() {
        const glossary = {};
        const items = glossaryContainer.querySelectorAll('.glossary-item');
        items.forEach(item => {
            const source = item.querySelector('.glossary-source').value.trim();
            const target = item.querySelector('.glossary-target').value.trim();
            if (source && target) {
                glossary[source] = target;
            }
        });
        return glossary;
    }

    async function getPromptConfig() {
        const mode = document.querySelector('input[name="translation_mode"]:checked')?.value || 'none';

        // 获取批处理设置
        let maxUnits = parseInt(maxUnitsPerChunkInput?.value);
        let maxChars = parseInt(maxCharsPerChunkInput?.value);
        
        // 如果前端输入框为空或无效，则从后端获取翻译器的默认值
        if (!maxUnits || isNaN(maxUnits) || !maxChars || isNaN(maxChars)) {
            const defaultConfig = await getDefaultBatchConfigFromBackend();
            
            if (!maxUnits || isNaN(maxUnits)) {
                maxUnits = defaultConfig.max_units_per_chunk;
            }
            if (!maxChars || isNaN(maxChars)) {
                maxChars = defaultConfig.max_chars_per_chunk;
            }
        }

        // 总是包含批处理设置
        const baseConfig = {
            max_units_per_chunk: maxUnits,
            max_chars_per_chunk: maxChars
        };

        if (mode === 'none') {
            // 对于 "无 Prompt" 模式，只发送模式和批处理设置
            return { mode: 'none', ...baseConfig };
        }

        // 对于其他模式，收集所有相关设置
        const promptConfig = {
            ...baseConfig,
            mode: mode,
            prompt_template: null,
            custom_prompt: null,
            preserve_terms: preserveTermsInput?.value.split(',').map(t => t.trim()).filter(t => t) || [],
            glossary: getGlossary(),
            additional_context: additionalContextTextarea?.value.trim() || ''
        };

        // 根据具体模式添加特定设置
        if (mode === 'general') {
            promptConfig.prompt_template = 'default';
        } else if (mode === 'professional') {
            promptConfig.prompt_template = professionalDomainSelect?.value || 'default';
        } else if (mode === 'custom') {
            const systemPrompt = customSystemPrompt?.value.trim();
            const userPrompt = customUserPrompt?.value.trim();
            if (systemPrompt || (userPrompt && userPrompt !== '{content}')) {
                promptConfig.custom_prompt = {
                    system: systemPrompt || '',
                    user: userPrompt || '{content}'
                };
            }
        }

        return promptConfig;
    }

    // --- 新增：PDF换行符清理函数 ---
    function cleanPdfLineBreaks(text) {
        if (!text || typeof text !== 'string') {
            return text;
        }

        // 第一步：标记和保护真正的段落分隔符
        // 将多个连续换行符（段落分隔）替换为特殊标记
        let processed = text.replace(/\n{2,}/g, '§PARAGRAPH_BREAK§');

        // 第二步：处理特殊情况 - 标点符号后的单换行可能是段落结束
        // 句号、问号、感叹号、冒号后的换行，如果下一行是大写字母或数字开头，保留为段落分隔
        processed = processed.replace(/([.!?:])\n([A-Z0-9\u4e00-\u9fff])/g, '$1§PARAGRAPH_BREAK§$2');

        // 第三步：处理列表项 - 数字或字母开头的列表项
        processed = processed.replace(/\n(\s*(?:\d+[.)]\s*|[a-zA-Z][.)]\s*|[•·\-\*]\s*))/g, '§PARAGRAPH_BREAK§$1');

        // 第四步：处理可能的标题 - 单独一行且较短的文本（通常少于50个字符）
        processed = processed.replace(/\n([^\n]{1,50})\n/g, function (match, title) {
            // 如果这一行没有句号结尾，可能是标题
            if (!/[.!?]$/.test(title.trim())) {
                return '§PARAGRAPH_BREAK§' + title + '§PARAGRAPH_BREAK§';
            }
            return match;
        });

        // 第五步：将剩余的单换行符替换为空格（这些是PDF的软换行）
        processed = processed.replace(/\n/g, ' ');

        // 第六步：清理多余的空格
        processed = processed.replace(/\s+/g, ' ');

        // 第七步：恢复段落分隔符为双换行
        processed = processed.replace(/§PARAGRAPH_BREAK§/g, '\n\n');

        // 第八步：清理首尾空白和多余的换行
        processed = processed.trim();

        // 第九步：确保段落间只有双换行（清理可能的三个或更多换行）
        processed = processed.replace(/\n{3,}/g, '\n\n');

        return processed;
    }

    // --- 初始状态设置 ---
    populateLanguageSelect(sourceLangSelect, LANGUAGES, true);
    populateLanguageSelect(targetLangSelect, LANGUAGES, false);

    if (sourceLangSelect) sourceLangSelect.value = '';
    if (targetLangSelect) targetLangSelect.value = '简体中文';

    if (apiPlatformSelect) updateApiFieldsForPlatform(apiPlatformSelect.value);
    switchToTextMode();
    updatePromptUI();

    // 加载配置和模板
    loadConfigList();
    loadPromptTemplates();

    // --- 事件监听器 ---
    if (apiPlatformSelect) {
        apiPlatformSelect.addEventListener('change', function () {
            updateApiFieldsForPlatform(this.value);
        });
    }

    // --- 新增：配置管理事件监听器 ---
    if (saveConfigBtn) {
        saveConfigBtn.addEventListener('click', saveCurrentConfig);
    }

    if (loadConfigBtn && savedConfigsDropdown) {
        loadConfigBtn.addEventListener('click', function () {
            loadSelectedConfig(savedConfigsDropdown.value);
        });
    }

    if (deleteConfigBtn) {
        deleteConfigBtn.addEventListener('click', deleteSelectedConfig);
    }
    if (exportConfigBtn) {
        exportConfigBtn.addEventListener('click', exportSelectedConfig);
    }
    if (importConfigFileInput) {
        importConfigFileInput.addEventListener('change', importConfig);
    }

    // --- 新增：模板文件相关事件监听器 ---
    if (templateFileInput) {
        templateFileInput.addEventListener('change', handleTemplateFileSelection);
    }

    if (clearTemplateButton) {
        clearTemplateButton.addEventListener('click', clearTemplateFile);
    }

    if (docxProcessingMethodSelect) {
        docxProcessingMethodSelect.addEventListener('change', updateTemplateUploadVisibility);
    }

    // --- 新增：Prompt相关事件监听器 ---
    translationModeRadios.forEach(radio => {
        radio.addEventListener('change', updatePromptUI);
    });

    if (addGlossaryBtn) {
        addGlossaryBtn.addEventListener('click', function () {
            addGlossaryItem();
            appendLog("添加了新的术语表项");
        });
    }

    if (glossaryContainer && glossaryContainer.children.length === 0) {
        addGlossaryItem();
    }

    if (swapLangButton) {
        swapLangButton.addEventListener('click', function () {
            if (!sourceLangSelect || !targetLangSelect) return;
            let sourceValue = sourceLangSelect.value;
            let targetValue = targetLangSelect.value;

            if (sourceValue === "") {
                sourceLangSelect.value = targetValue;
                if (targetValue === "简体中文") {
                    targetLangSelect.value = "English";
                } else if (targetValue === "English") {
                    targetLangSelect.value = "简体中文";
                } else {
                    targetLangSelect.value = "English";
                }
            } else {
                sourceLangSelect.value = targetValue;
                targetLangSelect.value = sourceValue;
            }
            appendLog(`语言互换：源语言变为 ${sourceLangSelect.options[sourceLangSelect.selectedIndex].textContent}，目标语言变为 ${targetLangSelect.options[targetLangSelect.selectedIndex].textContent}`);
        });
    }

    if (modeTextButton) {
        modeTextButton.addEventListener('click', function () {
            switchToTextMode();
            resetOutputArea();
        });
    }

    if (modeFileButton) {
        modeFileButton.addEventListener('click', function () {
            switchToFileMode();
            resetOutputArea();
        });
    }

    if (fileInput) {
        fileInput.addEventListener('change', function () {
            const file = this.files[0];
            if (file) {
                if (fileNameDisplay) fileNameDisplay.textContent = file.name;
                if (textInput) textInput.value = '';
                switchToFileMode();
                resetOutputArea();
                updateFileSpecificOptions();
            } else {
                if (fileNameDisplay) fileNameDisplay.textContent = '未选择任何文档';
                updateFileSpecificOptions();
                appendLog("未选择文档。");
            }
        });
    }

    if (copyDefaultPathButton && outputFolderPathInput) {
        copyDefaultPathButton.addEventListener('click', function () {
            outputFolderPathInput.value = 'translated_output_web';
            appendLog("已复制默认下载目录路径到输入框。");
        });
    }

    if (copySourceTextButton && textInput) {
        copySourceTextButton.addEventListener('click', function () {
            const textToCopy = textInput.value;
            if (textToCopy) {
                navigator.clipboard.writeText(textToCopy).then(function () {
                    appendLog('原文已复制到剪贴板。');
                    if (statusMessage) {
                        const originalStatusText = statusMessage.textContent;
                        const originalStatusClass = statusMessage.className;
                        statusMessage.textContent = '原文已复制!';
                        statusMessage.className = 'status-message success short-lived';
                        setTimeout(() => {
                            if (statusMessage.textContent === '原文已复制!') {
                                statusMessage.textContent = originalStatusText;
                                statusMessage.className = originalStatusClass;
                            }
                        }, 2000);
                    }
                }).catch(function (err) {
                    appendLog('复制原文到剪贴板失败: ' + err);
                    try {
                        textInput.select();
                        textInput.setSelectionRange(0, 99999);
                        document.execCommand('copy');
                        appendLog('尝试使用旧版API复制原文成功。');
                        if (statusMessage) {
                            const originalStatusText = statusMessage.textContent;
                            const originalStatusClass = statusMessage.className;
                            statusMessage.textContent = '原文已复制 (旧版)!';
                            statusMessage.className = 'status-message success short-lived';
                            setTimeout(() => { if (statusMessage.textContent === '原文已复制 (旧版)!') { statusMessage.textContent = originalStatusText; statusMessage.className = originalStatusClass; } }, 2000);
                        }
                    } catch (execErr) {
                        appendLog('旧版API复制原文也失败: ' + execErr);
                        alert('无法自动复制原文。请手动复制。');
                    }
                });
            } else {
                appendLog('原文区域无内容可复制。');
                if (statusMessage) {
                    const originalStatusText = statusMessage.textContent;
                    const originalStatusClass = statusMessage.className;
                    statusMessage.textContent = '原文无内容';
                    statusMessage.className = 'status-message info short-lived';
                    setTimeout(() => { if (statusMessage.textContent === '原文无内容') { statusMessage.textContent = originalStatusText; statusMessage.className = originalStatusClass; } }, 2000);
                }
            }
        });
    }

    // 新增：清除PDF换行符按钮事件监听器
    if (cleanPdfLineBreaksButton && textInput) {
        cleanPdfLineBreaksButton.addEventListener('click', function () {
            const originalText = textInput.value;
            if (originalText.trim()) {
                const cleanedText = cleanPdfLineBreaks(originalText);
                textInput.value = cleanedText;

                // 记录操作日志
                const lineBreaksRemoved = (originalText.match(/\n/g) || []).length - (cleanedText.match(/\n/g) || []).length;
                appendLog(`PDF换行符已清理，移除了 ${lineBreaksRemoved} 个软换行符。`);

                // 显示状态消息
                if (statusMessage) {
                    const originalStatusText = statusMessage.textContent;
                    const originalStatusClass = statusMessage.className;
                    statusMessage.textContent = `已清理PDF换行符 (移除${lineBreaksRemoved}个)`;
                    statusMessage.className = 'status-message success short-lived';
                    setTimeout(() => {
                        if (statusMessage.textContent.includes('已清理PDF换行符')) {
                            statusMessage.textContent = originalStatusText;
                            statusMessage.className = originalStatusClass;
                        }
                    }, 2000);
                }

                // 聚焦到文本框
                if (textInput.focus) textInput.focus();
            } else {
                appendLog('文本区域无内容，无需清理PDF换行符。');
                if (statusMessage) {
                    const originalStatusText = statusMessage.textContent;
                    const originalStatusClass = statusMessage.className;
                    statusMessage.textContent = '文本无内容';
                    statusMessage.className = 'status-message info short-lived';
                    setTimeout(() => {
                        if (statusMessage.textContent === '文本无内容') {
                            statusMessage.textContent = originalStatusText;
                            statusMessage.className = originalStatusClass;
                        }
                    }, 1500);
                }
            }
        });
    }

    if (clearSourceTextButton && textInput) {
        clearSourceTextButton.addEventListener('click', function () {
            if (textInput.value) {
                textInput.value = '';
                appendLog('原文已清除。');
                if (textInput.focus) textInput.focus();
                if (statusMessage) {
                    const originalStatusText = statusMessage.textContent;
                    const originalStatusClass = statusMessage.className;
                    statusMessage.textContent = '原文已清除';
                    statusMessage.className = 'status-message info short-lived';
                    setTimeout(() => {
                        if (statusMessage.textContent === '原文已清除') {
                            statusMessage.textContent = originalStatusText;
                            statusMessage.className = originalStatusClass;
                        }
                    }, 1500);
                }
            } else {
                appendLog('原文区域已为空，无需清除。');
            }
        });
    }

    // 复制翻译文本按钮保持不变
    if (copyTranslatedTextButton && translatedTextDisplay) {
        copyTranslatedTextButton.addEventListener('click', function () {
            const textToCopy = translatedTextDisplay.value;
            if (textToCopy) {
                navigator.clipboard.writeText(textToCopy).then(function () {
                    appendLog('翻译结果已复制到剪贴板。');
                    if (statusMessage) {
                        const originalStatusText = statusMessage.textContent;
                        const originalStatusClass = statusMessage.className;
                        statusMessage.textContent = '已复制到剪贴板!';
                        statusMessage.className = 'status-message success short-lived';
                        setTimeout(() => {
                            if (statusMessage.textContent === '已复制到剪贴板!') {
                                statusMessage.textContent = originalStatusText;
                                statusMessage.className = originalStatusClass;
                            }
                        }, 2000);
                    }
                }).catch(function (err) {
                    appendLog('复制到剪贴板失败 (navigator.clipboard): ' + err);
                    try {
                        translatedTextDisplay.select();
                        translatedTextDisplay.setSelectionRange(0, 99999);
                        document.execCommand('copy');
                        appendLog('尝试使用旧版API复制成功。');
                        if (statusMessage) {
                            const originalStatusText = statusMessage.textContent;
                            const originalStatusClass = statusMessage.className;
                            statusMessage.textContent = '已复制 (旧版)!';
                            statusMessage.className = 'status-message success short-lived';
                            setTimeout(() => {
                                if (statusMessage.textContent === '已复制 (旧版)!') {
                                    statusMessage.textContent = originalStatusText;
                                    statusMessage.className = originalStatusClass;
                                }
                            }, 2000);
                        }
                    } catch (execErr) {
                        appendLog('旧版API复制也失败: ' + execErr);
                        alert('无法自动复制文本。请手动选择并复制。');
                        if (statusMessage) {
                            statusMessage.textContent = '自动复制失败，请手动操作。';
                            statusMessage.className = 'status-message error';
                        }
                    }
                });
            } else {
                appendLog('没有可复制的翻译结果。');
                if (statusMessage) {
                    const originalStatusText = statusMessage.textContent;
                    const originalStatusClass = statusMessage.className;
                    statusMessage.textContent = '无内容可复制';
                    statusMessage.className = 'status-message info short-lived';
                    setTimeout(() => {
                        if (statusMessage.textContent === '无内容可复制') {
                            statusMessage.textContent = originalStatusText;
                            statusMessage.className = originalStatusClass;
                        }
                    }, 2000);
                }
            }
        });
    }

    // 修改翻译按钮事件处理器 - 关键修改部分！
    if (translateButton) {
        translateButton.addEventListener('click', async function () {
            resetOutputArea();

            if (translateButton) translateButton.disabled = true;
            if (statusMessage) {
                statusMessage.textContent = '正在翻译中...请稍候。';
                statusMessage.className = 'status-message info';
            }
            appendLog("开始翻译...");

            const selectedPlatform = apiPlatformSelect ? apiPlatformSelect.value : 'custom';
            const apiKey = apiKeyInput ? apiKeyInput.value : '';
            const baseUrl = baseUrlInput ? baseUrlInput.value : '';
            const model = modelInput ? modelInput.value : '';
            const sourceLang = sourceLangSelect ? sourceLangSelect.value : '';
            const targetLang = targetLangSelect ? targetLangSelect.value : '';
            const encodingValue = document.getElementById('encoding') ? document.getElementById('encoding').value : 'utf-8';
            const outputFolderPathValue = outputFolderPathInput ? outputFolderPathInput.value : '';

            if (!targetLang) {
                if (statusMessage) {
                    statusMessage.textContent = '错误：目标语言是必填项。';
                    statusMessage.className = 'status-message error';
                }
                if (translateButton) translateButton.disabled = false;
                appendLog("错误：目标语言是必填项。");
                return;
            }

            const promptConfig = await getPromptConfig();
            appendLog(`Prompt配置: 模式=${promptConfig.mode}, 模板=${promptConfig.prompt_template}`);

            const formData = new FormData();
            formData.append('api_platform', selectedPlatform);
            formData.append('api_key', apiKey);
            formData.append('base_url', baseUrl);
            formData.append('model', model);
            formData.append('target_lang', targetLang);
            formData.append('source_lang', sourceLang);

            formData.append('prompt_config', JSON.stringify(promptConfig));

            // --- 针对火山引擎SDK模式，显式添加AK/SK ---
            const volcAkInput = document.getElementById('volc_ak');
            const volcSkInput = document.getElementById('volc_sk');
            if (selectedPlatform === 'volcengine' || selectedPlatform === 'volcengine_sdk') {
                if (volcAkInput && volcAkInput.value.trim()) {
                    formData.append('volc_ak', volcAkInput.value.trim());
                    appendLog("Frontend: volc_ak added to FormData.");
                }
                if (volcSkInput && volcSkInput.value.trim()) {
                    formData.append('volc_sk', volcSkInput.value.trim());
                    appendLog("Frontend: volc_sk added to FormData.");
                }
            }
            // --- 火山引擎SDK模式AK/SK处理结束 ---            

            let isFileTranslation = false;
            let apiEndpoint = '/translate_api'; // 默认端点

            if (modeFileButton && modeFileButton.classList.contains('active') && fileInput && fileInput.files.length > 0) {
                const file = fileInput.files[0];
                formData.append('file', file);
                isFileTranslation = true;
                apiEndpoint = '/translate_file'; // 文件翻译使用新端点
                appendLog(`翻译模式：文档翻译。文档: ${file.name}`);

                // 新增：处理模板文件
                if (templateFileInput && templateFileInput.files.length > 0) {
                    const templateFile = templateFileInput.files[0];
                    formData.append('template_file', templateFile);
                    appendLog(`使用模板文件: ${templateFile.name}`);
                } else {
                    appendLog('未选择模板文件，将使用原文档格式');
                }

                const fileExtension = file.name.split('.').pop().toLowerCase();
                
                // 获取处理方法
                let processingMethod;
                if (fileExtension === 'txt') {
                    processingMethod = 'text_translator';
                    formData.append('encoding', encodingValue);
                } else if (fileExtension === 'docx') {
                    processingMethod = docxProcessingMethodSelect ? docxProcessingMethodSelect.value : 'docx_full_translator';
                    appendLog(`.docx 处理方法: ${docxProcessingMethodSelect ? docxProcessingMethodSelect.options[docxProcessingMethodSelect.selectedIndex].text : 'default'} (value: ${processingMethod})`);
                } else if (fileExtension === 'pptx') {
                    processingMethod = 'pptx_full_translator';
                } else if (fileExtension === 'xlsx' || fileExtension === 'xls') {
                    processingMethod = 'excel_full_translator';
                } else {
                    // 不支持的文件类型
                    if (statusMessage) {
                        statusMessage.textContent = `错误：不支持的文件类型 .${fileExtension}`;
                        statusMessage.className = 'status-message error';
                    }
                    if (translateButton) translateButton.disabled = false;
                    appendLog(`错误：不支持的文件类型 .${fileExtension}`);
                    return;
                }

                // 添加处理方法参数
                formData.append('processing_method', processingMethod);
                appendLog(`使用处理方法: ${processingMethod}`);
                
                formData.append('output_folder_path', outputFolderPathValue);

            } else if (modeTextButton && modeTextButton.classList.contains('active')) {
                const textToTranslate = textInput ? textInput.value : '';
                if (!textToTranslate.trim()) {
                    if (statusMessage) {
                        statusMessage.textContent = '错误：待翻译文字不能为空。';
                        statusMessage.className = 'status-message error';
                    }
                    if (translateButton) translateButton.disabled = false;
                    appendLog("错误：待翻译文字不能为空。");
                    return;
                }
                formData.append('text_input', textToTranslate);
                appendLog("翻译模式：文字翻译 (流式)。");
            } else {
                if (statusMessage) {
                    statusMessage.textContent = '错误：请选择文件或输入文字进行翻译。';
                    statusMessage.className = 'status-message error';
                }
                if (translateButton) translateButton.disabled = false;
                appendLog("错误：未指定翻译内容。");
                return;
            }

            try {
                appendLog(`发送请求到: ${apiEndpoint}`);
                const response = await fetch(apiEndpoint, {
                    method: 'POST',
                    body: formData
                });

                if (!response.ok) {
                    const errData = await response.json().catch(() => ({ error: `HTTP error! status: ${response.status} - ${response.statusText}` }));
                    throw new Error(errData.error || `HTTP error! status: ${response.status} - ${response.statusText}`);
                }

                if (!isFileTranslation && response.headers.get("content-type")?.includes("text/event-stream")) {
                    if (translatedTextDisplay) translatedTextDisplay.value = '';
                    appendLog("开始接收流式响应...");
                    const reader = response.body.getReader();
                    const decoder = new TextDecoder();
                    let accumulatedText = "";
                    let receivedAnyChunk = false;
                    let streamCancelledInternally = false;

                    while (true) {
                        try {
                            const { done, value } = await reader.read();
                            if (done) {
                                if (statusMessage) {
                                    statusMessage.textContent = receivedAnyChunk ? '翻译完成!' : '翻译完成 (无内容)。';
                                    statusMessage.className = 'status-message success';
                                }
                                appendLog(receivedAnyChunk ? "文字翻译流结束 (done=true)。" : "文字翻译流结束 (done=true, 未收到有效文本块)。");
                                break;
                            }
                            accumulatedText += decoder.decode(value, { stream: true });
                            let parts = accumulatedText.split('\n\n');
                            accumulatedText = parts.pop() || "";
                            for (const part of parts) {
                                if (part.startsWith('data: ')) {
                                    const jsonDataString = part.substring(6);
                                    if (jsonDataString.trim() === "[DONE]") {
                                        if (statusMessage) { statusMessage.textContent = receivedAnyChunk ? '翻译完成!' : '翻译完成 (API标记[DONE])。'; statusMessage.className = 'status-message success'; }
                                        appendLog(receivedAnyChunk ? "文字翻译流由 API [DONE] 标记结束。" : "文字翻译流由 API [DONE] 标记结束 (未收到有效文本块)。");
                                        if (!reader.closed) await reader.cancel();
                                        streamCancelledInternally = true;
                                        break;
                                    }
                                    try {
                                        const data = JSON.parse(jsonDataString);
                                        if (data.text_chunk && translatedTextDisplay) {
                                            translatedTextDisplay.value += data.text_chunk;
                                            translatedTextDisplay.scrollTop = translatedTextDisplay.scrollHeight;
                                            receivedAnyChunk = true;
                                        } else if (data.error) {
                                            if (statusMessage) { statusMessage.textContent = `错误：${data.error}`; statusMessage.className = 'status-message error'; }
                                            appendLog(`流中错误: ${data.error}`);
                                            if (!reader.closed) await reader.cancel();
                                            streamCancelledInternally = true;
                                            break;
                                        } else if (data.done) {
                                            if (statusMessage) { statusMessage.textContent = receivedAnyChunk ? '翻译完成!' : '翻译完成 (API标记done:true)。'; statusMessage.className = 'status-message success'; }
                                            appendLog(receivedAnyChunk ? "文字翻译流由 API JSON 'done:true' 标记结束。" : "文字翻译流由 API JSON 'done:true' 标记结束 (未收到有效文本块)。");
                                            if (!reader.closed) await reader.cancel();
                                            streamCancelledInternally = true;
                                            break;
                                        } else if (Object.keys(data).length > 0 && !data.text_chunk) {
                                            appendLog(`流中收到意外JSON结构: ${JSON.stringify(data)}`);
                                        }
                                    } catch (e) {
                                        appendLog(`无法解析流数据块: ${jsonDataString} - ${e}`);
                                    }
                                }
                            }
                            if (streamCancelledInternally) { break; }
                        } catch (readError) {
                            appendLog(`读取流时发生错误: ${readError.message}`);
                            if (statusMessage) { statusMessage.textContent = `读取流错误: ${readError.message}`; statusMessage.className = 'status-message error'; }
                            break;
                        }
                    }
                } else if (isFileTranslation) {
                    const result = await response.json();
                    
                    if (result.success) {
                        if (statusMessage) { 
                            statusMessage.textContent = '翻译成功！'; 
                            statusMessage.className = 'status-message success'; 
                        }
                        
                        if (result.download_url && downloadLink && downloadArea) {
                            downloadLink.href = result.download_url;
                            downloadLink.textContent = `下载翻译文件 (${result.download_url.split('/').pop()})`;
                            downloadArea.style.display = 'block';
                            downloadLink.style.display = 'inline-block';
                            appendLog(`文档翻译完成。下载链接: ${result.download_url}`);
                        }
                        
                        if (result.message) {
                            appendLog(`翻译消息: ${result.message}`);
                        }
                    } else {
                        // 翻译失败
                        if (statusMessage) {
                            statusMessage.textContent = `错误：${result.message || '翻译失败'}`;
                            statusMessage.className = 'status-message error';
                        }
                        appendLog(`文件翻译失败: ${result.message || '未知错误'}`);
                    }
                } else {
                    const result = await response.json().catch(() => ({ "error": "非流式响应且无法解析JSON" }));
                    if (result.error && statusMessage) {
                        statusMessage.textContent = `错误：${result.error}`;
                        statusMessage.className = 'status-message error';
                        appendLog(`非流式响应错误: ${JSON.stringify(result)}`);
                    } else if (result.translated_text && translatedTextDisplay) {
                        translatedTextDisplay.value = result.translated_text;
                        if (statusMessage) { statusMessage.textContent = '翻译完成!'; statusMessage.className = 'status-message success'; }
                    } else if (statusMessage) {
                        statusMessage.textContent = `错误：收到意外的响应格式。`;
                        statusMessage.className = 'status-message error';
                        appendLog(`意外响应格式: ${JSON.stringify(result)}`);
                    }
                }
            } catch (error) {
                if (statusMessage) { statusMessage.textContent = `翻译请求失败：${error.message}。`; statusMessage.className = 'status-message error'; }
                appendLog(`网络或JS错误: ${error.message}`);
            } finally {
                if (translateButton) translateButton.disabled = false;
                appendLog(isFileTranslation ? "文档翻译过程结束。" : "文字翻译过程结束。");
            }
        });
    }

    // Auto-scroll to language selection bar
    if (languageSelectionBar) {
        // 使用 setTimeout 确保页面布局稳定后再滚动，提升平滑度
        setTimeout(() => {
            languageSelectionBar.scrollIntoView({ behavior: 'smooth', block: 'start' });
        }, 100);
    }
});


// =================================
// PWA安装功能
// =================================

let deferredPrompt;
const installBanner = document.getElementById('installBanner');
const installBtn = document.getElementById('installBtn');
const dismissBtn = document.getElementById('dismissBtn');

// 页面加载完成后初始化PWA功能
document.addEventListener('DOMContentLoaded', function() {
    initializePWA();
});

function initializePWA() {
    console.log('🚀 初始化PWA功能');
    
    // 检查PWA支持
    if (!('serviceWorker' in navigator)) {
        console.log('❌ 浏览器不支持PWA');
        return;
    }
    
    // 检查是否已经是PWA模式运行
    if (window.matchMedia('(display-mode: standalone)').matches) {
        console.log('✅ 当前运行在PWA模式');
        return;
    }
    
    // 检查是否已经显示过安装提示
    if (localStorage.getItem('pwa-dismissed') === 'true') {
        console.log('ℹ️ 用户已关闭过安装提示');
        return;
    }
    
    // 延迟显示安装提示（让用户先体验应用）
    setTimeout(() => {
        checkAndShowInstallBanner();
    }, 5000); // 5秒后显示
}

// 监听 beforeinstallprompt 事件
window.addEventListener('beforeinstallprompt', (e) => {
    console.log('📲 PWA安装提示事件触发');
    
    // 阻止浏览器默认的安装提示
    e.preventDefault();
    
    // 保存事件供后续使用
    deferredPrompt = e;
    
    // 显示自定义安装提示
    showInstallBanner();
});

// 检查并显示安装横幅
function checkAndShowInstallBanner() {
    // 如果有保存的安装提示事件，或者是移动设备，则显示横幅
    if (deferredPrompt || isMobileDevice()) {
        showInstallBanner();
    }
}

// 显示安装横幅
function showInstallBanner() {
    if (!installBanner) return;
    
    console.log('📱 显示PWA安装横幅');
    installBanner.style.display = 'block';
    
    // 平滑显示动画
    setTimeout(() => {
        installBanner.classList.add('show');
    }, 100);
}

// 隐藏安装横幅
function hideInstallBanner() {
    if (!installBanner) return;
    
    installBanner.classList.remove('show');
    setTimeout(() => {
        installBanner.style.display = 'none';
    }, 300);
}

// 安装按钮点击事件
if (installBtn) {
    installBtn.addEventListener('click', async () => {
        console.log('🔥 用户点击安装按钮');
        
        if (deferredPrompt) {
            // 使用保存的安装提示
            try {
                deferredPrompt.prompt();
                const { outcome } = await deferredPrompt.userChoice;
                
                console.log(`用户选择: ${outcome}`);
                
                if (outcome === 'accepted') {
                    console.log('✅ 用户接受了安装');
                    hideInstallBanner();
                } else {
                    console.log('❌ 用户拒绝了安装');
                }
                
                deferredPrompt = null;
            } catch (error) {
                console.error('安装过程出错:', error);
                showManualInstallGuide();
            }
        } else {
            // 显示手动安装指导
            showManualInstallGuide();
        }
    });
}

// 关闭按钮点击事件
if (dismissBtn) {
    dismissBtn.addEventListener('click', () => {
        console.log('❌ 用户关闭了安装提示');
        hideInstallBanner();
        
        // 记录用户已关闭，24小时内不再显示
        const dismissTime = Date.now() + (24 * 60 * 60 * 1000); // 24小时后过期
        localStorage.setItem('pwa-dismissed', 'true');
        localStorage.setItem('pwa-dismiss-time', dismissTime.toString());
    });
}

// 检查关闭状态是否过期
function checkDismissExpired() {
    const dismissTime = localStorage.getItem('pwa-dismiss-time');
    if (dismissTime && Date.now() > parseInt(dismissTime)) {
        localStorage.removeItem('pwa-dismissed');
        localStorage.removeItem('pwa-dismiss-time');
    }
}

// 手动安装指导
function showManualInstallGuide() {
    const userAgent = navigator.userAgent.toLowerCase();
    let title = '安装应用到桌面';
    let message = '';
    
    if (userAgent.includes('chrome') && !userAgent.includes('mobile')) {
        message = '1. 点击地址栏右侧的安装图标 📥\n2. 或者点击浏览器菜单中的"安装翻译器"';
    } else if (userAgent.includes('safari') && userAgent.includes('mobile')) {
        message = '1. 点击底部分享按钮 📤\n2. 滑动找到"添加到主屏幕"\n3. 点击"添加"';
    } else if (userAgent.includes('mobile')) {
        message = '1. 点击浏览器菜单按钮\n2. 查找"添加到主屏幕"或"安装应用"\n3. 确认安装';
    } else {
        message = '1. 点击浏览器菜单\n2. 查找"安装应用"或"创建快捷方式"\n3. 确认安装';
    }
    
    alert(title + '\n\n' + message);
}

// 检测是否是移动设备
function isMobileDevice() {
    return /Android|webOS|iPhone|iPad|iPod|BlackBerry|IEMobile|Opera Mini/i.test(navigator.userAgent);
}

// 监听应用安装成功事件
window.addEventListener('appinstalled', (evt) => {
    console.log('🎉 PWA安装成功!');
    hideInstallBanner();
    showInstallSuccessMessage();
    
    // 清除关闭状态
    localStorage.removeItem('pwa-dismissed');
    localStorage.removeItem('pwa-dismiss-time');
});

// 显示安装成功消息
function showInstallSuccessMessage() {
    const successMsg = document.createElement('div');
    successMsg.className = 'install-success';
    successMsg.innerHTML = '🎉 应用已成功安装到桌面！';
    
    document.body.appendChild(successMsg);
    
    // 3秒后自动移除
    setTimeout(() => {
        successMsg.remove();
    }, 3000);
}

// 初始化时检查关闭状态是否过期
checkDismissExpired();

console.log('✅ PWA脚本加载完成');

// ================================================== 
// 咖啡打赏功能
// ==================================================
document.addEventListener('DOMContentLoaded', function() {
    const coffeeBtn = document.getElementById('coffee-btn');
    const coffeeModal = document.getElementById('coffee-modal');
    const coffeeClose = document.querySelector('.coffee-close');

    // 打开弹窗
    if (coffeeBtn) {
        coffeeBtn.addEventListener('click', function() {
            coffeeModal.style.display = 'block';
            document.body.style.overflow = 'hidden'; // 防止背景滚动
        });
    }

    // 关闭弹窗
    if (coffeeClose) {
        coffeeClose.addEventListener('click', function() {
            coffeeModal.style.display = 'none';
            document.body.style.overflow = 'auto'; // 恢复滚动
        });
    }

    // 点击弹窗外部关闭
    if (coffeeModal) {
        coffeeModal.addEventListener('click', function(event) {
            if (event.target === coffeeModal) {
                coffeeModal.style.display = 'none';
                document.body.style.overflow = 'auto'; // 恢复滚动
            }
        });
    }

    // ESC键关闭弹窗
    document.addEventListener('keydown', function(event) {
        if (event.key === 'Escape' && coffeeModal && coffeeModal.style.display === 'block') {
            coffeeModal.style.display = 'none';
            document.body.style.overflow = 'auto'; // 恢复滚动
        }
    });

    console.log('✅ 咖啡打赏功能加载完成');
});
