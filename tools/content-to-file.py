import json

from collections.abc import Generator

from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from tools.tools import generate_timestamp_random, DocxGenerator


class ContentToFileTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:

        # print(tool_parameters)
        data = tool_parameters.get("output_data", "你没有输入data")
        filename = tool_parameters.get("filename", None)
        file_type = tool_parameters.get("file_type", "txt")

        if filename is None:
            filename = generate_timestamp_random()
        if file_type == "txt":
            # 将文本编码成字节
            text_bytes = data.encode('utf-8')
            # 返回 blob
            yield self.create_blob_message(
                blob=text_bytes,
                meta={
                    "mime_type": "text/plain",
                    "filename": f"{filename}.{file_type}"  # 可选：指定文件名
                }
            )

        else:
            yield self.create_text_message("未知的格式")





