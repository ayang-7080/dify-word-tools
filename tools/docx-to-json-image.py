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


class DocxToJsonImageTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:


        docx_file: File = tool_parameters.get("file")
        host = tool_parameters.get("host")
        in_no_proxy = tool_parameters.get("no_proxy")

        in_image_index = tool_parameters.get("image_index",None)


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
                    "filename": data.get("filename", f"image_{image_index}.png"),
                    "mime_type": data.get("mime_type", "application/octet-stream"),
                    "image_index": image_index,
                })
                data["content"] = image_index
                image_index += 1
            #yield self.create_json_message(data)

        yield self.create_text_message(f"共 {len(image_list)} 个图片")

        if in_image_index is not None and int(in_image_index) >= len(image_list):
            yield self.create_text_message(f",图片不存在(0~{len(image_list)-1})")
        elif in_image_index is not None:
            in_image_index = int(in_image_index)
            image = image_list[in_image_index]
            image_data = base64.b64decode(image.get("content"))  # 假设你保留了原始 base64
            yield self.create_blob_message(image_data, {
                'mime_type': image.get("mime_type", "application/octet-stream"),
                'filename': image.get("filename", f'image_{in_image_index}.bin')
            })
        else:
            for image in image_list:
                if image.get("type") == "image":
                    image_index = image.get("image_index")
                    image_data = base64.b64decode(image.get("content"))  # 假设你保留了原始 base64
                    yield self.create_blob_message(image_data, {
                        'mime_type': image.get("mime_type", "application/octet-stream"),
                        'filename': image.get("filename", f'image_{image_index}.bin')
                    })

        yield self.create_text_message(f"！")





