# main.py
from dotenv import load_dotenv
load_dotenv()

import argparse
import logging
import os
from translator import SiliconFlowTranslator # 不修改
from file_translator import translate_text_file # 不修改
from docx_translator import translate_docx_file # 不修改

# Logger configuration
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
)
logger = logging.getLogger(__name__)

# --- Standard DOCX Translators ---
_translate_docx_file_formatted_is_available = False
_translate_docx_file_formatted_placeholder_message = "Error: Formatted DOCX translator (docx_full_translator.py) module or function is not configured."
try:
    from docx_full_translator import translate_docx_file_formatted # 不修改
    logger.info("Successfully imported translate_docx_file_formatted from docx_full_translator.")
    _translate_docx_file_formatted_is_available = True
except ImportError:
    logger.warning("docx_full_translator.py or translate_docx_file_formatted function not found. Formatted DOCX translation (--docx_full_translator) will not be available.")
    def translate_docx_file_formatted(*args, **kwargs): # 不修改 (Placeholder)
        return _translate_docx_file_formatted_placeholder_message

_translate_docx_via_markdown_is_available = False # For the original docx_markdown_translator
_translate_docx_via_markdown_placeholder_message = "Error: Original Markdown-based DOCX translator (docx_markdown_translator.py) module or function is not configured."
try:
    from docx_markdown_translator import translate_docx_via_markdown
    logger.info("Successfully imported translate_docx_via_markdown from docx_markdown_translator.")
    _translate_docx_via_markdown_is_available = True
except ImportError:
    logger.warning("docx_markdown_translator.py or translate_docx_via_markdown function not found. Original Markdown-based DOCX translation (--docx_markdown_translate) will not be available.")
    def translate_docx_via_markdown(*args, **kwargs):
        return _translate_docx_via_markdown_placeholder_message


_translate_docx_via_html_is_available = False
_translate_docx_via_html_placeholder_message = "Error: HTML-based DOCX translator (docx_html_translator.py) module or function is not configured."
try:
    from docx_html_translator import translate_docx_via_html
    logger.info("Successfully imported translate_docx_via_html from docx_html_translator.")
    _translate_docx_via_html_is_available = True
except ImportError:
    logger.warning("docx_html_translator.py or translate_docx_via_html function not found. HTML-based DOCX translation (--docx_html_translate) will not be available.")
    def translate_docx_via_html(*args, **kwargs):
        return _translate_docx_via_html_placeholder_message

# --- NEW DOCX Translators (PythonDoc1 & PythonDoc2) ---
_translate_docx_pythondoc1_is_available = False
_translate_docx_pythondoc1_placeholder_message = "Error: PythonDoc1 translator (docx_pythondoc1_translator.py) module or function is not configured."
try:
    from docx_pythondoc1_translator import translate_docx_via_markdown as translate_docx_pythondoc1
    logger.info("Successfully imported translate_docx_via_markdown as translate_docx_pythondoc1 from docx_pythondoc1_translator.")
    _translate_docx_pythondoc1_is_available = True
except ImportError:
    logger.warning("docx_pythondoc1_translator.py or its translate_docx_via_markdown function not found. PythonDoc1 translation (--docx_pythondoc1_translator) will not be available.")
    def translate_docx_pythondoc1(*args, **kwargs): # Placeholder
        return _translate_docx_pythondoc1_placeholder_message

_translate_docx_pythondoc2_is_available = False
_translate_docx_pythondoc2_placeholder_message = "Error: PythonDoc2 translator (docx_pythondoc2_translator.py) module or function is not configured."
try:
    from docx_pythondoc2_translator import translate_docx_via_markdown as translate_docx_pythondoc2
    logger.info("Successfully imported translate_docx_via_markdown as translate_docx_pythondoc2 from docx_pythondoc2_translator.")
    _translate_docx_pythondoc2_is_available = True
except ImportError:
    logger.warning("docx_pythondoc2_translator.py or its translate_docx_via_markdown function not found. PythonDoc2 translation (--docx_pythondoc2_translator) will not be available.")
    def translate_docx_pythondoc2(*args, **kwargs): # Placeholder
        return _translate_docx_pythondoc2_placeholder_message


from config import API_KEY, DEFAULT_MODEL, BASE_URL

CLI_EXAMPLES = """
Usage Examples:
  Translate text:
    python main.py -t "你好，世界！" -l "English"

  Translate a plain text file:
    python main.py -i input.txt -o output_dir -l "English"

  Translate a Word document (.docx) using various methods:
    python main.py -id doc.docx -od out -l "简体中文" --docx_translator (from docx_translator.py)
    python main.py -id doc.docx -od out -l "简体中文" --docx_full_translator (from docx_full_translator.py)
    python main.py -id doc.docx -od out -l "简体中文" --docx_markdown_translate (from docx_markdown_translator.py)
    python main.py -id doc.docx -od out -l "简体中文" --docx_html_translate (from docx_html_translator.py)
    python main.py -id doc.docx -od out -l "简体中文" --docx_pythondoc1_translator (from docx_pythondoc1_translator.py)
    python main.py -id doc.docx -od out -l "简体中文" --docx_pythondoc2_translator (from docx_pythondoc2_translator.py)
""" # <<< MODIFIED: CLI Examples updated

def main():
    parser = argparse.ArgumentParser(
        description="A simple command-line text translator using SiliconFlow-like API.",
        formatter_class=argparse.RawTextHelpFormatter,
        epilog=CLI_EXAMPLES
    )

    parser.add_argument("-t", "--text", help="The text to be translated.")
    parser.add_argument("-l", "--target_lang", required=True, help="The target language.")
    parser.add_argument("-s", "--source_lang", help="The source language (optional).")
    parser.add_argument("-i", "--input_file", help="Path to the input plain text file (.txt).")
    parser.add_argument("-o", "--output_dir", help="Directory for translated plain text file.")
    parser.add_argument("--encoding", default="utf-8", help="Encoding for plain text files (default: utf-8).")
    parser.add_argument("-id", "--input_docx", help="Path to the input Word document (.docx).")
    parser.add_argument("-od", "--output_docx_dir", help="Directory for translated Word document.")

    # DOCX translation method group
    docx_method_group = parser.add_mutually_exclusive_group()
    docx_method_group.add_argument("--docx_translator", action="store_true", help="Use translation from docx_translator.py (basic).")
    docx_method_group.add_argument("--docx_full_translator", action="store_true", help="Use formatted translation from docx_full_translator.py.")
    docx_method_group.add_argument("--docx_markdown_translate", action="store_true", help="Use Markdown conversion from docx_markdown_translator.py.")
    docx_method_group.add_argument("--docx_html_translate", action="store_true", help="Use HTML conversion from docx_html_translator.py.")
    # <<< RENAMED: Arguments renamed to match filenames
    docx_method_group.add_argument("--docx_pythondoc1_translator", action="store_true", help="Use translation from docx_pythondoc1_translator.py.")
    docx_method_group.add_argument("--docx_pythondoc2_translator", action="store_true", help="Use translation from docx_pythondoc2_translator.py.")


    parser.add_argument("-m", "--model", help=f"Model to use (overrides '{DEFAULT_MODEL}').")
    parser.add_argument("-u", "--base_url", help=f"Base URL for API (overrides '{BASE_URL}').")
    parser.add_argument("-k", "--api_key", help="API Key (overrides environment variable). USE WITH CAUTION.")

    args = parser.parse_args()

    input_modes = sum(1 for mode in [args.text, args.input_file, args.input_docx] if mode is not None)
    if input_modes > 1:
        parser.error("Only one input type (-t, -i, or -id) can be provided.")
    if input_modes == 0:
        parser.error("At least one input type (-t, -i, or -id) must be provided.")
    
    if args.input_file and not args.output_dir:
        parser.error("-o/--output_dir is required with -i/--input_file.")
    if args.input_docx and not args.output_docx_dir:
        parser.error("-od/--output_docx_dir is required with -id/--input_docx.")
    
    if args.input_docx:
        # <<< RENAMED: Flags in this check updated
        if not (args.docx_translator or args.docx_full_translator or 
                args.docx_markdown_translate or args.docx_html_translate or
                args.docx_pythondoc1_translator or args.docx_pythondoc2_translator):
            parser.error("When using -id/--input_docx, you must specify a DOCX translation method: "
                         "--docx_translator, --docx_full_translator, --docx_markdown_translate, "
                         "--docx_html_translate, --docx_pythondoc1_translator, or --docx_pythondoc2_translator.")

    # <<< RENAMED: Flags in this check updated
    docx_specific_flag_used = (args.docx_translator or args.docx_full_translator or 
                               args.docx_markdown_translate or args.docx_html_translate or
                               args.docx_pythondoc1_translator or args.docx_pythondoc2_translator)
    
    if docx_specific_flag_used and not args.input_docx:
         parser.error("DOCX-specific translation flags require -id/--input_docx.")

    if docx_specific_flag_used and (args.text or args.input_file):
        parser.error("DOCX-specific translation flags cannot be used with -t/--text or -i/--input_file.")

    try:
        logger.info("Translator application started.")
        
        resolved_api_key = args.api_key or os.getenv("SILICONFLOW_API_KEY") or API_KEY
        resolved_base_url = args.base_url or os.getenv("SILICONFLOW_BASE_URL") or BASE_URL
        resolved_model = args.model or os.getenv("SILICONFLOW_MODEL") or DEFAULT_MODEL
        
        if not all([resolved_api_key, resolved_base_url, resolved_model]):
            missing = [name for name, val in [("API Key", resolved_api_key), ("Base URL", resolved_base_url), ("Model", resolved_model)] if not val]
            raise ValueError(f"{', '.join(missing)} is missing. Configure via CLI, .env, or config.py.")

        translator_for_api = SiliconFlowTranslator(
            api_key=resolved_api_key,
            base_url=resolved_base_url,
            model=resolved_model
        )
        logger.info(f"Translator for API calls initialized with model: {resolved_model}.")

        if args.input_docx:
            logger.info(f"Processing DOCX file: {args.input_docx}")
            if not os.path.exists(args.input_docx):
                logger.error(f"Input Word document not found: {args.input_docx}")
                print(f"\nError: Input Word document not found at '{args.input_docx}'")
                return
            
            if not os.path.exists(args.output_docx_dir):
                try:
                    os.makedirs(args.output_docx_dir, exist_ok=True)
                    logger.info(f"Created output directory for DOCX: {args.output_docx_dir}")
                except Exception as e:
                    logger.error(f"Could not create output directory '{args.output_docx_dir}': {e}")
                    print(f"\nError: Could not create output directory '{args.output_docx_dir}': {e}")
                    return

            output_filepath_or_error = ""

            if args.docx_translator:
                logger.info(f"Using DOCX translation from docx_translator.py for: {args.input_docx}")
                output_filepath_or_error = translate_docx_file(
                    input_filepath=args.input_docx,
                    output_dir=args.output_docx_dir,
                    target_lang=args.target_lang,
                    translator=translator_for_api,
                    source_lang=args.source_lang,
                )
            elif args.docx_full_translator:
                logger.info(f"Using formatted DOCX translation with docx_full_translator.py for: {args.input_docx}")
                if not _translate_docx_file_formatted_is_available:
                    output_filepath_or_error = _translate_docx_file_formatted_placeholder_message
                    logger.error(f"Attempted to use --docx_full_translator, but the module is not available. {_translate_docx_file_formatted_placeholder_message}")
                else:
                    try:
                        output_filepath_or_error = translate_docx_file_formatted(
                            input_filepath=args.input_docx, output_dir=args.output_docx_dir, target_lang=args.target_lang,
                            translator=translator_for_api, source_lang=args.source_lang,
                        )
                    except Exception as e_fmt: 
                        logger.error(f"Formatted DOCX translation (docx_full_translator.py) failed: {e_fmt}", exc_info=True)
                        output_filepath_or_error = f"Error: Formatted DOCX translation (docx_full_translator.py) failed: {e_fmt}"

            elif args.docx_markdown_translate: 
                logger.info(f"Using DOCX translation via Markdown (docx_markdown_translator.py) for: {args.input_docx}")
                if not _translate_docx_via_markdown_is_available:
                    output_filepath_or_error = _translate_docx_via_markdown_placeholder_message
                    logger.error(f"Attempted to use --docx_markdown_translate, but the module is not available. {_translate_docx_via_markdown_placeholder_message}")
                else:
                    try:
                        output_filepath_or_error = translate_docx_via_markdown( 
                            input_filepath=args.input_docx,
                            output_dir=args.output_docx_dir,
                            target_lang=args.target_lang,
                            translator=translator_for_api,
                            source_lang=args.source_lang,
                        )
                    except Exception as e_md_orig:
                        logger.error(f"Markdown-based DOCX translation (docx_markdown_translator.py) failed: {e_md_orig}", exc_info=True)
                        output_filepath_or_error = f"Error: Markdown-based DOCX translation (docx_markdown_translator.py) failed: {e_md_orig}"

            elif args.docx_html_translate:
                logger.info(f"Using DOCX translation via HTML (docx_html_translator.py) for: {args.input_docx}")
                if not _translate_docx_via_html_is_available:
                    output_filepath_or_error = _translate_docx_via_html_placeholder_message
                    logger.error(f"Attempted to use --docx_html_translate, but the module is not available. {_translate_docx_via_html_placeholder_message}")
                else:
                    try:
                        output_filepath_or_error = translate_docx_via_html(
                            input_filepath=args.input_docx,
                            output_dir=args.output_docx_dir,
                            target_lang=args.target_lang,
                            translator=translator_for_api,
                            source_lang=args.source_lang,
                        )
                    except Exception as e_html:
                        logger.error(f"HTML-based DOCX translation failed: {e_html}", exc_info=True)
                        output_filepath_or_error = f"Error: HTML-based DOCX translation failed: {e_html}"
            
            # <<< RENAMED: Handling for docx_pythondoc1_translator
            elif args.docx_pythondoc1_translator:
                logger.info(f"Using DOCX translation from docx_pythondoc1_translator.py for: {args.input_docx}")
                if not _translate_docx_pythondoc1_is_available:
                    output_filepath_or_error = _translate_docx_pythondoc1_placeholder_message
                    logger.error(f"Attempted to use --docx_pythondoc1_translator, but the module is not available. {_translate_docx_pythondoc1_placeholder_message}")
                else:
                    try:
                        output_filepath_or_error = translate_docx_pythondoc1( 
                            input_filepath=args.input_docx,
                            output_dir=args.output_docx_dir,
                            target_lang=args.target_lang,
                            translator=translator_for_api,
                            source_lang=args.source_lang,
                        )
                    except Exception as e_pydoc1:
                        logger.error(f"DOCX translation (docx_pythondoc1_translator.py) failed: {e_pydoc1}", exc_info=True)
                        output_filepath_or_error = f"Error: DOCX translation (docx_pythondoc1_translator.py) failed: {e_pydoc1}"
            
            # <<< RENAMED: Handling for docx_pythondoc2_translator
            elif args.docx_pythondoc2_translator:
                logger.info(f"Using DOCX translation from docx_pythondoc2_translator.py for: {args.input_docx}")
                if not _translate_docx_pythondoc2_is_available:
                    output_filepath_or_error = _translate_docx_pythondoc2_placeholder_message
                    logger.error(f"Attempted to use --docx_pythondoc2_translator, but the module is not available. {_translate_docx_pythondoc2_placeholder_message}")
                else:
                    try:
                        output_filepath_or_error = translate_docx_pythondoc2( 
                            input_filepath=args.input_docx,
                            output_dir=args.output_docx_dir,
                            target_lang=args.target_lang,
                            translator=translator_for_api,
                            source_lang=args.source_lang,
                        )
                    except Exception as e_pydoc2:
                        logger.error(f"DOCX translation (docx_pythondoc2_translator.py) failed: {e_pydoc2}", exc_info=True)
                        output_filepath_or_error = f"Error: DOCX translation (docx_pythondoc2_translator.py) failed: {e_pydoc2}"
            
            # Result reporting
            if isinstance(output_filepath_or_error, str) and output_filepath_or_error.lower().startswith("error:"):
                logger.error(f"Word document translation failed for {args.input_docx}: {output_filepath_or_error}")
                print(f"\n--- Word Document Translation Failed ---\nError details: {output_filepath_or_error}\n----------------------------------------")
            elif isinstance(output_filepath_or_error, str) and os.path.exists(output_filepath_or_error):
                logger.info(f"Word document translation successful. Output: {output_filepath_or_error}")
                print(f"\n--- Word Document Translation Complete ---\nTranslated document saved to: {output_filepath_or_error}\n------------------------------------------")
            else:
                logger.warning(f"Word document translation for {args.input_docx} ended with undetermined state. Returned: '{output_filepath_or_error}'.")
                print(f"\n--- Word Document Translation Ended with Undetermined State ---\nFunction call returned: {output_filepath_or_error}\n-------------------------------------------------------------")
                
        elif args.input_file:
            logger.info(f"Starting plain text file translation for: {args.input_file}")
            if not os.path.exists(args.input_file):
                logger.error(f"Input file not found: {args.input_file}")
                print(f"\nError: Input file not found at '{args.input_file}'")
                return
            
            output_filepath_or_error = translate_text_file(
                input_filepath=args.input_file, output_dir=args.output_dir, target_lang=args.target_lang,
                translator=translator_for_api, source_lang=args.source_lang, encoding=args.encoding
            )
            if isinstance(output_filepath_or_error, str) and output_filepath_or_error.lower().startswith("error:"):
                logger.error(f"Plain text file translation failed for {args.input_file}: {output_filepath_or_error}")
                print(f"\n--- Plain Text File Translation Failed ---\n{output_filepath_or_error}\n------------------------------------------")
            elif isinstance(output_filepath_or_error, str) and os.path.exists(output_filepath_or_error):
                logger.info(f"Plain text file translation successful. Output: {output_filepath_or_error}")
                print(f"\n--- Plain Text File Translation Complete ---\nTranslated file saved to: {output_filepath_or_error}\n------------------------------------------")
            else:
                logger.warning(f"Plain text file translation for {args.input_file} ended with undetermined state. Returned: '{output_filepath_or_error}'.")
                print(f"\n--- Plain Text File Translation Ended with Undetermined State ---\nFunction call returned: {output_filepath_or_error}\n-------------------------------------------------------------")


        elif args.text:
            logger.info(f"Starting text translation for: '{args.text[:50]}...'")
            translated_text = translator_for_api.translate(text=args.text, target_lang=args.target_lang, source_lang=args.source_lang)
            print(f"\n--- Text Translation Result ---\n{translated_text}\n--------------------------")
            if not (isinstance(translated_text, str) and translated_text.lower().startswith("error:")) :
                logger.info("Text translation successful.")
            else:
                logger.error(f"Text translation failed: {translated_text}")
        
    except ValueError as ve:
        logger.error(f"Configuration Error: {ve}", exc_info=False) 
        print(f"\nConfiguration Error: {ve}\nCheck API Key, Base URL, Model settings.")
    except Exception as e:
        logger.exception(f"An unexpected error occurred: {e}")
        print(f"\nAn unexpected error occurred: {e}")

if __name__ == "__main__":
    main()
