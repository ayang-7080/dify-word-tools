
from collections.abc import Generator
from io import BytesIO
from itertools import islice

from docx import Document
from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage
from dify_plugin.file.file import File

from tools.tools import no_proxy


class DocxToJsonTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:

        #print(tool_parameters)

        docx_file: File = tool_parameters.get("file")
        host = tool_parameters.get("host")
        in_no_proxy = tool_parameters.get("no_proxy")
        output_format = tool_parameters.get("output_format", "text")

        if host is not None:
            docx_file.url = f"{host}{docx_file.url}"
        #print(docx_file)

        if in_no_proxy == "close":
            with no_proxy():
                file_data = docx_file.blob
        else:
            file_data = docx_file.blob

        def extract_paragraphs_from_bytes(file_blob: bytes):
            doc = Document(BytesIO(file_blob))
            paragraphs = [p.text.strip() for p in doc.paragraphs if p.text.strip()]
            return paragraphs

        datas = extract_paragraphs_from_bytes(file_data)

        if output_format == "json":

            for data_str in datas:
                yield self.create_json_message({"content":data_str})

            # yield self.create_json_message({
            #     "content" : datas
            # })
        else:
            datas_text = "\n\n".join(datas)
            ## 分块
            def chunk_text(text, chunk_size=3000):
                for i in range(0, len(text), chunk_size):
                    yield text[i:i + chunk_size]

            for chunk in chunk_text(datas_text):
                yield self.create_text_message(chunk)

            # yield self.create_text_message("\n\n".join(datas))





