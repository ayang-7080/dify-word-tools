import io
import json
import tempfile
from collections.abc import Generator
from io import BytesIO
from typing import Any

import mammoth
import pypandoc
from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.file.file import File

from tools.tools import no_proxy


class DocxMarkdownPandocTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:

        print("DocxMarkdownPandoc")

        print(tool_parameters)

        docx_file: File = tool_parameters.get("file")
        host = tool_parameters.get("host")

        in_no_proxy = tool_parameters.get("no_proxy")
        output_format = tool_parameters.get("output_format","markdown")
        if host is not None:
            docx_file.url = f"{host}{docx_file.url}"
        print(docx_file)

        try:
            # 只在没有 Pandoc 时再下载
            pypandoc.get_pandoc_path()

            print( pypandoc.get_pandoc_formats() )
        except OSError:
            print("Pandoc not found. Downloading...")
            pypandoc.download_pandoc()

        if in_no_proxy == "close":
            with no_proxy():
                file_data = docx_file.blob
        else:
            file_data = docx_file.blob

        def convert_docx_to_md(file_blob: bytes):
            with tempfile.NamedTemporaryFile(delete=False, suffix=".docx") as tmp:
                tmp.write(file_blob)
                tmp_path = tmp.name

            md_text = pypandoc.convert_file(tmp_path, output_format,extra_args=["--wrap=none"])
            return md_text

        markdown_data = convert_docx_to_md(file_data)
        # print(markdown_data)

        # doc_stream = BytesIO(file_data)
        #
        # document = docx.Document(doc_stream)

        yield self.create_text_message(markdown_data)





