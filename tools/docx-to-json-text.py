import base64
import json
from collections import Counter
from collections.abc import Generator
from io import BytesIO
from itertools import islice

from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.file.file import File

from tools.tools import no_proxy, extract_docx_to_json


class DocxToJsonTextTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:

        #print("转Json")
        # print(tool_parameters)

        docx_file: File = tool_parameters.get("file")
        table_markdown = tool_parameters.get("table_markdown", False)
        if table_markdown == "":
            table_markdown = False
        host = tool_parameters.get("host")
        in_no_proxy = tool_parameters.get("no_proxy")

        if host is not None:
            docx_file.url = f"{host}{docx_file.url}"

        if in_no_proxy == "close":
            with no_proxy():
                file_data = docx_file.blob
        else:
            file_data = docx_file.blob

        datas = extract_docx_to_json(BytesIO(file_data),table_markdown=True)
        #print(datas)

        image_index = 0

        image_list = []
        for idx, data in enumerate(datas):
            if data.get("type") == "image":
                image_list.append({
                    "type":"image",
                    "content":data["content"] ,
                    "image_index": image_index,
                })
                data["content"] = image_index
                image_index += 1
            yield self.create_json_message(data)

        yield self.create_text_message(f"读取成功, 共 {len(datas)} 条数据")






