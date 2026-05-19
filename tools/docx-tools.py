import json

from collections.abc import Generator

from typing import Any

from dify_plugin import Tool
from dify_plugin.entities.tool import ToolInvokeMessage

from tools.tools import generate_timestamp_random, DocxGenerator


class DocxToolsTool(Tool):
    def _invoke(self, tool_parameters: dict[str, Any]) -> Generator[ToolInvokeMessage]:

        #print("生成docx")
        data = tool_parameters.get("output_data","你没有输入data")
        filename = tool_parameters.get("filename",None)

        page_padding = int(tool_parameters.get("page_padding",2))

        enable_line = tool_parameters.get("enable_line", False)
        if enable_line == "":
            enable_line = False
        enable_page = tool_parameters.get("enable_page", False)
        if enable_page == "":
            enable_page = False

        # 将output_data字符串转换为Python对象
        try:
            items = json.loads(data)
            #print("转换成功！")
        except json.JSONDecodeError as e:
            items = []
            raise Exception(f"JSON解析错误 {str(e)}")
        

        if filename is None or filename == "":
            filename = generate_timestamp_random()

        generator_doc = DocxGenerator()
        if enable_line:
            generator_doc.enable_line_numbering()

        for item in items:
            #print(item)
            text = item.get("data")
            in_type = item.get("type","paragraph")
            font_name = item.get("font_name","宋体")
            font_size = item.get("font_size",12)
            bold = item.get("bold",False)
            alignment = item.get("alignment","left")
            line_spacing  = item.get("line_spacing",1)
            first_line_indent = item.get("first_line_indent",0)
            break_type = item.get("break_type", None)
            color = item.get("color", "#000000")
            bg_color = item.get("bg_color", None)
            space_after = item.get("space_after", 0)
            para = None
            if in_type == "table":
                style = item.get("style","")
                cell_width = item.get("cell_width",0)
                generator_doc.add_table_from_markdown(text, style=style, font_name=font_name, font_size=font_size,alignment=alignment,cell_width=cell_width)
            elif in_type == "table2":
                style = item.get("style", "")
                cell_width = item.get("cell_width", 0)
                column_alignments = item.get("column_alignments", [])
                bold_content = item.get("bold_content", False)
                generator_doc.add_table_from_markdown2(text, style=style, font_name=font_name, font_size=font_size,
                                                      alignment=alignment, cell_width=cell_width,column_alignments=column_alignments,bold_content=bold_content,text_color=color)
            elif in_type == "image":
                width = item.get("width",None)
                height = item.get("height",None)
                para = generator_doc.add_image(text , width = width, height = height, alignment = alignment)
                comment = item.get("comment", None)
                if comment and isinstance(comment, dict) and comment.get("text"):
                    comment_text = comment.get("text")
                    author = comment.get("author")
                    initials = comment.get("initials")
                    generator_doc.add_paragraph_comment(para=para, comment_text=comment_text, author=author,
                                                        initials=initials)
                if comment and isinstance(comment, list) and len(comment) > 0:
                    for com in comment:
                        comment_text = com.get("text")
                        author = com.get("author")
                        initials = com.get("initials")
                        generator_doc.add_paragraph_comment(para=para, comment_text=comment_text, author=author,
                                                            initials=initials)
            elif in_type == "float_image":
                width = item.get("width",None)
                height = item.get("height",None)
                pos_x = item.get("pos_x",0)
                pos_y = item.get("pos_y",0)
                wrap_type = item.get("wrap_type","topBottom")
                generator_doc.add_float_image_absolute(text , width = width, height = height,  pos_x=pos_x, pos_y=pos_y, wrap_type=wrap_type)
            else:
                comment = item.get("comment", None)
                texts = []
                if type(text) == str:
                    text = text.replace("\n\n", "\n")
                    texts = text.split("\n")
                elif type(text) == list:
                    texts = text

                for text_item in texts:
                    para = generator_doc.add_paragraph(text=text_item, font_name=font_name, font_size=font_size,
                                                bold=bold, color=color,alignment=alignment, line_spacing=line_spacing,
                                                space_after=space_after,first_line_indent=first_line_indent,bg_color=bg_color)
                    if comment and isinstance(comment,dict) and comment.get("text"):
                        comment_text = comment.get("text")
                        author = comment.get("author")
                        initials = comment.get("initials")
                        generator_doc.add_paragraph_comment(para=para,comment_text = comment_text,author=author,initials=initials)
                    if comment and isinstance(comment,list) and len(comment) > 0:
                        for com in comment:
                            comment_text = com.get("text")
                            author = com.get("author")
                            initials = com.get("initials")
                            generator_doc.add_paragraph_comment(para=para,comment_text = comment_text,author=author,initials=initials)

            if break_type == "page":
                generator_doc.add_page_break()
            if break_type == "line":
                generator_doc.add_line_break(para=para)

        generator_doc.set_page_margins(page_padding)
        if enable_page:
            generator_doc.add_page_number()
        docx_bytes = generator_doc.buildBytes()

        # 返回 blob
        yield self.create_blob_message(
            blob=docx_bytes,
            meta={
                "mime_type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document", # application/vnd.openxmlformats-officedocument.wordprocessingml.document
                "filename": f"{filename}.docx"
            }
        )



