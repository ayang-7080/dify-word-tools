import json
from collections.abc import Generator
from io import BytesIO
from typing import Any

import requests
from PIL import Image
from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from docx.shared import Mm, Cm
from docxtpl import DocxTemplate, InlineImage

from tools.tools import generate_timestamp_random, no_proxy


class AyToolsTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:
        template_file = tool_parameters.get('template_file')
        in_data = tool_parameters.get("in_data", "你没有输入data")

        filename = tool_parameters.get("filename", "你没有输入data")
        if filename is None or filename == "":
            filename = generate_timestamp_random()

        try:
            data = json.loads(in_data)
        except json.JSONDecodeError as e:
            data = []
            raise Exception(f"JSON解析错误 {str(e)}")

        # 下载模板文件
        with no_proxy():
            file_data = template_file.blob
        file_stream = BytesIO(file_data)

        doc = DocxTemplate(file_stream)

        # 处理数据
        word_context = {}
        if isinstance(data, dict):
            word_context = build_word_context(data, doc)
        else:
            raise ValueError("数据有误")

        #print(word_context)
        # 渲染
        doc.render( word_context )

        # 输出为 bytes（不落盘）
        output_stream = BytesIO()
        doc.save(output_stream)

        # 获取最终 Word bytes
        result_bytes = output_stream.getvalue()

        # 返回 blob
        yield self.create_blob_message(
            blob=result_bytes,
            meta={
                "mime_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                "filename": f"{filename}.docx"
            }
        )




def build_word_context(data: dict, doc, image_width_cm=17):
    def normalize_image_bytes(raw_bytes, out_format="PNG"):
      img = Image.open(BytesIO(raw_bytes))
      img = img.convert("RGB")
      out = BytesIO()
      img.save(out, format=out_format)
      out.seek(0)
      return out
    def handle_value(v):
        # ========== 图片对象 ==========
        if isinstance(v, dict) and v.get("type") == "image":
            # ---------- 获取原始 bytes ----------
            if "bytes" in v:
              raw_bytes = v["bytes"]
            elif "url" in v:
              resp = requests.get(v["url"], timeout=50)
              resp.raise_for_status()
              raw_bytes = resp.content
            elif "file" in v:
              with open(v["file"], "rb") as f:
                raw_bytes = f.read()
            else:
              raise ValueError(f"image 类型缺少数据源: {v}")
            # ---------- 🔥 关键步骤：图像标准化 ----------
            try:
              img_stream = normalize_image_bytes(raw_bytes, out_format="PNG")
            except Exception as e:
              raise ValueError(f"图像规范化失败（非法图像结构）: {e}")

            # ---------- 插入 Word ----------
            return InlineImage(doc, img_stream, width=Cm(image_width_cm))

        # ========== 普通 dict ==========
        elif isinstance(v, dict):
            return {kk: handle_value(vv) for kk, vv in v.items()}

        # ========== list ==========
        elif isinstance(v, list):
            return [handle_value(i) for i in v]
        else:
            return v
    if not isinstance(data, dict):
        raise ValueError("data 必须是 dict 类型")

    word_context = {}
    for k, v in data.items():
        word_context[k] = handle_value(v)

    return word_context
