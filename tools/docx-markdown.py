import io
import json
from collections.abc import Generator
from io import BytesIO
from typing import Any

import mammoth
from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.file.file import File

from tools.tools import no_proxy


class DocxMarkdownTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        print("to markdown .... ")
        # print(tool_parameters)

        docx_file: File = tool_parameters.get("file")
        host = tool_parameters.get("host")
        in_no_proxy = tool_parameters.get("no_proxy")

        if host is not None:
            docx_file.url = f"{host}{docx_file.url}"
        #print(docx_file)

        if in_no_proxy == "close":
            with no_proxy():
                file_data = docx_file.blob
        else:
            file_data = docx_file.blob

        def docx_to_markdown(file_blob: bytes) -> str:
            """使用 mammoth 转换 DOCX -> Markdown"""
            result = mammoth.convert_to_markdown(io.BytesIO(file_blob))
            return result.value  # 返回 Markdown 文本

        markdown_data = docx_to_markdown(file_data)

        ## 分块
        def chunk_text(text, chunk_size=3000):
            for i in range(0, len(text), chunk_size):
                yield text[i:i + chunk_size]

        for chunk in chunk_text(markdown_data):
            yield self.create_text_message(chunk)

        #yield self.create_text_message(markdown_data)





