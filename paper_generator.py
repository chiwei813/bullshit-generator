# -*- coding: utf-8 -*-
import requests
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING, WD_TAB_LEADER
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import docx.opc.constants
import os
import json
import time
from urllib.parse import quote
from scholarly import scholarly
import time
import random

class PaperGenerator:
    def __init__(self):
        # 從 api.txt 讀取 API key
        self.api_config = {}
        with open('api.txt', 'r') as f:
            for line in f:
                if '=' in line and not line.strip().startswith('#'):
                    key, value = line.strip().split('=')
                    self.api_config[key] = value
        
        # 初始化進度追蹤
        self._progress = 0
        self._progress_steps = {
            'search': 20,     # 搜尋資料完成時的進度
            'abstract': 40,   # 摘要生成完成時的進度
            'content': 80,    # 內容生成完成時的進度
            'format': 100     # 檔案格式化完成時的進度
        }

    @property
    def progress(self):
        return self._progress

    def convert_to_traditional(self, text):
        """將簡體中文轉換為繁體中文"""
        try:
            from opencc import OpenCC
            cc = OpenCC('s2t')  # 簡體到繁體的轉換
            return cc.convert(text)
        except ImportError:
            print("Warning: opencc-python-reimplemented package is not installed. Using original text.")
            return text

    def generate_content(self, topic, sources):
        """使用 Gemini API 生成論文內容"""
        url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent"
        
        # 準備提示詞
        prompt = f"""作為一位資深的學術研究者，請根據以下主題撰寫一篇完整的學術論文。使用繁體中文撰寫，確保內容專業、嚴謹且具有深度。

主題：{topic}

===寫作規範與要求===

1. 整體寫作風格：
- 使用正式的學術用語，避免口語化表達
- 保持客觀中立的論述立場
- 適當引用參考文獻支持論點
- 使用精確的專業術語
- 每個論點都需要具體的研究數據或文獻支持
- 確保邏輯推理清晰，段落轉折自然
- 適當使用學術寫作常用詞彙，如「此外」、「因此」、「基於上述分析」等

2. 具體章節要求：

[摘要]
字數：2000字
結構要求：
- 研究背景：說明研究領域現況和問題的重要性
- 研究目的：闡明研究要解決的具體問題
- 研究方法：詳述採用的研究方法和工具
- 研究結果：列舉主要研究發現和數據
- 研究結論：總結研究貢獻和應用價值
- 研究限制：指出研究的局限性
關鍵提示：
- 使用準確的研究術語
- 提供具體的數據和發現
- 清楚說明研究的創新點

[緒論]
字數：4000字
章節架構：
1. 研究背景（1200字）
   - 領域現況分析
   - 國內外研究綜述
   - 現有問題和挑戰
2. 研究動機（1000字）
   - 問題的重要性
   - 現有研究的不足
   - 待解決的理論或實務問題
3. 研究目的（800字）
   - 具體研究目標
   - 預期研究貢獻
4. 研究問題（500字）
   - 主要研究問題
   - 次要研究問題
5. 研究假設（500字）
   - 理論基礎
   - 具體假設內容

[研究方法]
字數：4000字
章節架構：
1. 研究架構（1000字）
   - 理論框架
   - 變項關係圖
   - 研究流程
2. 研究設計（1000字）
   - 研究方法選擇理由
   - 研究對象界定
   - 抽樣方法說明
3. 研究工具（1000字）
   - 測量工具說明
   - 問卷或實驗設計
   - 信效度分析
4. 資料收集（500字）
   - 收集程序
   - 回收情況
5. 資料分析（500字）
   - 統計方法說明
   - 分析工具說明

[研究結果]
字數：4000字
章節架構：
1. 描述性統計（1000字）
   - 樣本特性分析
   - 主要變項分布
   - 集中與離散趨勢
2. 假設檢定（1500字）
   - 各項假設的檢定結果
   - 統計顯著性分析
   - 效果量分析
3. 進階分析（1500字）
   - 調節效果分析
   - 中介效果分析
   - 補充分析結果

[討論與建議]
字數：4000字
章節架構：
1. 研究發現討論（1500字）
   - 主要發現詮釋
   - 與既有研究比較
   - 創新發現說明
2. 理論意涵（1000字）
   - 理論貢獻
   - 概念模型的修正與擴充
   - 研究方法的創新
3. 實務建議（1000字）
   - 針對實務界的具體建議
   - 應用價值說明
   - 政策建議
4. 研究限制與未來研究建議（500字）
   - 研究限制說明
   - 後續研究方向建議

===內容生成要求===

1. 每個章節都要有明確的結構和層次
2. 各小節之間要有適當的轉折語
3. 要引用參考文獻支持論述
4. 適當使用圖表說明研究結果
5. 確保前後論述邏輯一致
6. 使用學術用語和專業術語
7. 提供具體的研究數據
8. 確保每個章節都達到指定字數

請根據以上主題和要求，生成一篇完整的學術論文。請特別注意專業性、邏輯性和學術規範。請參考以下文獻：

"""
        # 添加參考文獻
        for source in sources[:3]:
            prompt += f"- {source['title']}\n"
            prompt += f"摘要：{source['content'][:300]}...\n\n"

        # 設定生成參數以取得更長和更詳細的內容
        payload = {
            "contents": [{
                "parts":[{"text": prompt}]
            }],
            "generationConfig": {
                "temperature": 0.7,  # 保持適度的創造力
                "topK": 40,
                "topP": 0.95,
                "maxOutputTokens": 16384  # 提高輸出長度上限
            }
        }

        try:
            response = requests.post(
                f"{url}?key={self.api_config['GEMINI_API_KEY']}", 
                headers={'Content-Type': 'application/json'},
                json=payload,
                timeout=180  # 增加超時時間到3分鐘
            )
            
            if response.status_code == 200:
                generated_text = response.json()['candidates'][0]['content']['parts'][0]['text']
                # 將生成的文本轉換為繁體中文
                generated_text = self.convert_to_traditional(generated_text)
                
                # 初始化章節內容
                sections = {
                    'abstract': '',
                    'introduction': '',
                    'methods': '',
                    'results': '',
                    'discussion': '',
                    'references': [{'title': self.convert_to_traditional(s['title']), 'link': s['link']} for s in sources]
                }
                
                # 改進章節解析邏輯
                current_section = None
                section_text = []
                
                for line in generated_text.split('\n'):
                    line = line.strip()
                    if not line:
                        continue
                        
                    # 檢查章節標記並處理子標題
                    if '[摘要]' in line:
                        if current_section and section_text:
                            sections[current_section] = '\n'.join(section_text)
                        current_section = 'abstract'
                        section_text = []
                        continue
                    elif '[緒論]' in line:
                        if current_section and section_text:
                            sections[current_section] = '\n'.join(section_text)
                        current_section = 'introduction'
                        section_text = []
                        continue
                    elif '[研究方法]' in line:
                        if current_section and section_text:
                            sections[current_section] = '\n'.join(section_text)
                        current_section = 'methods'
                        section_text = []
                        continue
                    elif '[研究結果]' in line:
                        if current_section and section_text:
                            sections[current_section] = '\n'.join(section_text)
                        current_section = 'results'
                        section_text = []
                        continue
                    elif '[討論與建議]' in line or '[討論]' in line:
                        if current_section and section_text:
                            sections[current_section] = '\n'.join(section_text)
                        current_section = 'discussion'
                        section_text = []
                        continue
                    
                    # 如果有當前章節且行不是空的，則添加到當前章節
                    if current_section and line:
                        # 處理子標題的格式
                        if line.startswith(('一、', '二、', '三、', '四、', '五、')):
                            section_text.append(f'\n{line}\n')
                        elif line.startswith(('1.', '2.', '3.', '4.', '5.')):
                            section_text.append(f'\n{line}\n')
                        else:
                            section_text.append(line)
                
                # 確保最後一個章節的內容也被保存
                if current_section and section_text:
                    sections[current_section] = '\n'.join(section_text)
                
                # 內容品質檢查
                for section_name, content in sections.items():
                    if section_name != 'references':
                        # 檢查字數是否足夠
                        min_words = {
                            'abstract': 2000,
                            'introduction': 4000,
                            'methods': 4000,
                            'results': 4000,
                            'discussion': 4000
                        }
                        
                        if len(content) < min_words[section_name]:
                            # 如果內容不足，生成補充內容
                            additional_prompt = f"""請為論文的{section_name}章節生成補充內容，
                            目前內容字數不足，需要補充至少{min_words[section_name]-len(content)}字。
                            主題是：{topic}
                            請確保補充內容與原有內容保持一致，並符合學術寫作規範。
                            """
                            
                            additional_payload = {
                                "contents": [{
                                    "parts":[{"text": additional_prompt}]
                                }],
                                "generationConfig": {
                                    "temperature": 0.7,
                                    "maxOutputTokens": 8192
                                }
                            }
                            
                            try:
                                additional_response = requests.post(
                                    f"{url}?key={self.api_config['GEMINI_API_KEY']}", 
                                    headers={'Content-Type': 'application/json'},
                                    json=additional_payload,
                                    timeout=60
                                )
                                
                                if additional_response.status_code == 200:
                                    additional_text = additional_response.json()['candidates'][0]['content']['parts'][0]['text']
                                    # 將補充內容也轉換為繁體中文
                                    additional_text = self.convert_to_traditional(additional_text)
                                    sections[section_name] = content + '\n\n' + additional_text
                            except Exception as e:
                                print(f"Error generating additional content: {str(e)}")
                
                return sections
                
            else:
                print(f"Error in AI generation: {response.status_code} - {response.text}")
                return None
                
        except Exception as e:
            print(f"Error in AI generation: {str(e)}")
            return None

    def iterate_content(self, topic, sections, iteration_focus):
        """對指定章節進行內容迭代優化
        
        Args:
            topic: 論文主題
            sections: 當前的論文內容
            iteration_focus: 需要優化的章節名稱
        """
        url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent"
        
        if iteration_focus not in sections:
            return sections
            
        current_content = sections[iteration_focus]
        
        # 準備迭代優化的提示詞
        prompt = f"""作為資深學術研究者，請對以下{iteration_focus}章節內容進行優化改進。主題是：{topic}

當前內容：
{current_content}

請基於以下方向進行優化：
1. 增強論述的邏輯性和連貫性
2. 補充更多具體的研究數據或案例
3. 加強學術用語的專業性
4. 確保每個論點都有充分的支持證據
5. 優化段落之間的轉折
6. 補充必要的文獻引用

請提供優化後的完整內容。"""

        payload = {
            "contents": [{
                "parts":[{"text": prompt}]
            }],
            "generationConfig": {
                "temperature": 0.7,
                "maxOutputTokens": 8192
            }
        }

        try:
            response = requests.post(
                f"{url}?key={self.api_config['GEMINI_API_KEY']}", 
                headers={'Content-Type': 'application/json'},
                json=payload,
                timeout=120
            )
            
            if response.status_code == 200:
                improved_text = response.json()['candidates'][0]['content']['parts'][0]['text']
                # 將優化後的內容轉換為繁體中文
                improved_text = self.convert_to_traditional(improved_text)
                
                # 更新指定章節的內容
                sections[iteration_focus] = improved_text
                
            return sections
                
        except Exception as e:
            print(f"迭代優化過程中發生錯誤: {str(e)}")
            return sections

    def add_page_number(self, paragraph):
        """添加頁碼"""
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        page_num = OxmlElement('w:fldSimple')
        page_num.set(qn('w:instr'), 'PAGE')
        paragraph._p.append(page_num)

    def create_hyperlink(self, paragraph, text, url):
        """為 docx 文件創建超連結"""
        # 這個方法會正確創建帶有超連結的文字
        part = paragraph.part
        r_id = part.relate_to(url, docx.opc.constants.RELATIONSHIP_TYPE.HYPERLINK, is_external=True)

        # 創建超連結元素
        hyperlink = OxmlElement('w:hyperlink')
        hyperlink.set(qn('r:id'), r_id)

        # 創建文字運行
        new_run = OxmlElement('w:r')
        rPr = OxmlElement('w:rPr')

        # 設置樣式
        color = OxmlElement('w:color')
        color.set(qn('w:val'), '0000FF')
        rPr.append(color)

        # 添加下劃線
        u = OxmlElement('w:u')
        u.set(qn('w:val'), 'single')
        rPr.append(u)

        new_run.append(rPr)
        t = OxmlElement('w:t')
        t.text = text
        new_run.append(t)
        hyperlink.append(new_run)
        
        paragraph._p.append(hyperlink)
        return hyperlink

    def generate_docx(self, topic, content):
        """生成 DOCX 格式的論文"""
        doc = Document()
        
        # 設置頁面邊距（符合 APA 格式）
        sections = doc.sections
        for section in sections:
            section.top_margin = Inches(1)
            section.bottom_margin = Inches(1)
            section.left_margin = Inches(1.25)
            section.right_margin = Inches(1.25)
        
        # 添加封面頁
        doc.add_heading('', level=0)  # 空白開頭
        doc.add_paragraph().alignment = WD_ALIGN_PARAGRAPH.CENTER  # 空行
        
        # 添加標題（置中，16pt，粗體）
        title_paragraph = doc.add_paragraph()
        title_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
        title_run = title_paragraph.add_run(topic)
        title_run.font.size = Pt(16)
        title_run.font.bold = True
        title_run.font.name = '新細明體'  # 使用標準中文字體
        
        # 添加學校和系所資訊（預留，可由使用者填寫）
        school_para = doc.add_paragraph()
        school_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        school_para.add_run('\n\n國立某某大學\n某某研究所\n碩士論文').font.size = Pt(14)
        
        # 添加作者資訊
        author_para = doc.add_paragraph()
        author_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        author_para.add_run('\n\n研究生：（姓名）\n指導教授：（姓名）\n\n\n').font.size = Pt(12)
        
        # 添加日期（使用民國年）
        date_para = doc.add_paragraph()
        date_para.alignment = WD_ALIGN_PARAGRAPH.CENTER
        date_para.add_run('\n中華民國 112 年 12 月').font.size = Pt(12)
        
        # 添加分頁符
        doc.add_page_break()
        
        # 添加摘要頁
        abstract_title = doc.add_heading('摘要', level=1)
        abstract_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # 添加摘要內容（縮排兩個中文字）
        abstract_para = doc.add_paragraph(content['abstract'])
        abstract_para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
        abstract_para.paragraph_format.space_after = Pt(12)
        abstract_para.paragraph_format.first_line_indent = Inches(0.5)
        
        # 添加關鍵詞（粗體標題）
        key_words_para = doc.add_paragraph()
        key_words_para.add_run('關鍵詞：').bold = True
        key_words_para.add_run(topic)
        key_words_para.paragraph_format.space_after = Pt(24)
        
        doc.add_page_break()
        
        # 添加目錄
        toc_heading = doc.add_heading('目錄', level=1)
        toc_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # 添加目錄項目（包含章節編號和頁碼）
        chapter_num = 1
        for section in ['緒論', '研究方法', '研究結果', '討論與建議', '參考文獻']:
            toc_para = doc.add_paragraph()
            if section != '參考文獻':
                toc_para.add_run(f'第{chapter_num}章 ')
                chapter_num += 1
            toc_para.add_run(section)
            toc_para.add_run('...').font.name = 'Times New Roman'
            toc_para.paragraph_format.left_indent = Inches(0.5)
            toc_para.paragraph_format.tab_stops.add_tab_stop(
                Inches(6), 
                WD_ALIGN_PARAGRAPH.RIGHT, 
                WD_TAB_LEADER.DOTS
            )
        
        doc.add_page_break()
        
        # 添加正文章節（包含章節編號）
        chapter_num = 1
        sections_map = {
            'introduction': ('緒論', content['introduction']),
            'methods': ('研究方法', content['methods']),
            'results': ('研究結果', content['results']),
            'discussion': ('討論與建議', content['discussion'])
        }
        
        for section_id, (title, content_text) in sections_map.items():
            # 添加章節標題（包含編號）
            heading = doc.add_heading(f'第{chapter_num}章 {title}', level=1)
            heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
            
            # 分段處理內容
            paragraphs = content_text.split('\n')
            for para_text in paragraphs:
                if para_text.strip():
                    # 創建新段落
                    para = doc.add_paragraph()
                    para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
                    para.paragraph_format.first_line_indent = Inches(0.5)
                    
                    # 檢查是否包含引用
                    if section_id == 'introduction':
                        # 在引言中處理參考文獻連結
                        for source in content['references']:
                            if source['title'].lower() in para_text.lower():
                                # 找到引用的位置
                                start_idx = para_text.lower().find(source['title'].lower())
                                if start_idx >= 0:
                                    # 添加前面的文字
                                    if start_idx > 0:
                                        para.add_run(para_text[:start_idx])
                                    # 添加帶連結的引用
                                    self.create_hyperlink(para, source['title'], source['link'])
                                    # 添加剩餘的文字
                                    remaining_text = para_text[start_idx + len(source['title']):]
                                    if remaining_text:
                                        para.add_run(remaining_text)
                                    continue
                            
                    # 如果沒有找到引用或不是引言部分，直接添加文字
                    para.add_run(para_text)
            
            chapter_num += 1
            
            if section_id != 'discussion':  # 如果不是最後一章，添加分頁符
                doc.add_page_break()
        
        # 添加參考文獻
        heading = doc.add_heading('參考文獻', level=1)
        heading.alignment = WD_ALIGN_PARAGRAPH.LEFT
        
        # 添加參考文獻列表
        for source in content['references']:
            ref_para = doc.add_paragraph()
            ref_para.paragraph_format.first_line_indent = Inches(0.5)
            ref_para.paragraph_format.line_spacing_rule = WD_LINE_SPACING.ONE_POINT_FIVE
            self.create_hyperlink(ref_para, source['title'], source['link'])
        
        # 添加頁碼（置中，從摘要頁開始）
        for section in doc.sections:
            footer = section.footer
            footer_para = footer.paragraphs[0]
            self.add_page_number(footer_para)
        
        output_path = 'generated_paper.docx'
        doc.save(output_path)
        return output_path

    def generate_odt(self, topic, content):
        """生成 ODT 格式的論文"""
        doc = OpenDocumentText()
        
        # 創建樣式
        # 標題樣式
        title_style = Style(name="Title", family="paragraph")
        title_style.addElement(TextProperties(fontsize="16pt", fontweight="bold"))
        title_style.addElement(ParagraphProperties(textalign="center", linespacing="150%"))
        doc.styles.addElement(title_style)
        
        # 正文樣式
        text_style = Style(name="Text", family="paragraph")
        text_style.addElement(TextProperties(fontsize="12pt"))
        text_style.addElement(ParagraphProperties(textalign="justify", linespacing="150%", firstlineindent="0.5in"))
        doc.styles.addElement(text_style)
        
        # 章節標題樣式
        heading_style = Style(name="Heading", family="paragraph")
        heading_style.addElement(TextProperties(fontsize="14pt", fontweight="bold"))
        heading_style.addElement(ParagraphProperties(margintop="24pt", marginbottom="12pt"))
        doc.styles.addElement(heading_style)
        
        # 添加標題頁
        title = H(stylename="Title", outlinelevel=1, text=topic)
        doc.text.addElement(title)
        
        # 添加作者資訊（預留空間）
        doc.text.addElement(P(text=""))
        
        # 添加摘要
        doc.text.addElement(H(stylename="Heading", outlinelevel=1, text="摘要"))
        p = P(stylename="Text", text=content['abstract'])
        doc.text.addElement(p)
        
        # 添加關鍵詞
        p = P(stylename="Text", text="關鍵詞：" + topic)
        doc.text.addElement(p)
        
        # 添加目錄標題
        doc.text.addElement(H(stylename="Heading", outlinelevel=1, text="目錄"))
        
        # 添加目錄項目
        for section in ['緒論', '研究方法', '研究結果', '討論與建議', '參考文獻']:
            p = P(text=section)
            doc.text.addElement(p)
        
        # 添加正文章節
        sections_map = {
            'introduction': ('緒論', content['introduction']),
            'methods': ('研究方法', content['methods']),
            'results': ('研究結果', content['results']),
            'discussion': ('討論與建議', content['discussion'])
        }
        
        for section_id, (title, content_text) in sections_map.items():
            # 添加章節標題
            doc.text.addElement(H(stylename="Heading", outlinelevel=1, text=title))
            # 添加章節內容
            p = P(stylename="Text", text=content_text)
            doc.text.addElement(p)
        
        # 添加參考文獻
        doc.text.addElement(H(stylename="Heading", outlinelevel=1, text="參考文獻"))
        
        for ref in content['references']:
            p = P(stylename="Text")
            if isinstance(ref, dict) and 'link' in ref and ref['link'] != '#':
                href = A(href=ref['link'])
                span = Span(text=ref['title'])
                href.addElement(span)
                p.addElement(href)
                p.addText(' - ')
                href = A(href=ref['link'])
                span = Span(text=ref['link'])
                href.addElement(span)
                p.addElement(href)
            else:
                title = ref if isinstance(ref, str) else ref.get('title', 'Unknown Reference')
                p.addText(title)
            doc.text.addElement(p)
        
        output_path = 'generated_paper.odt'
        doc.save(output_path)
        return output_path
    
    def search_references(self, topic, num_refs=15):
        """
        根據主題搜尋相關的學術文獻
        """
        references = []
        search_query = scholarly.search_pubs(topic)
        
        try:
            for i in range(num_refs):
                pub = next(search_query)
                # 確保有完整的引用資訊
                if hasattr(pub, 'bib') and 'title' in pub.bib and 'author' in pub.bib and 'year' in pub.bib:
                    ref = {
                        'title': pub.bib.get('title', ''),
                        'authors': pub.bib.get('author', ''),
                        'year': pub.bib.get('year', ''),
                        'venue': pub.bib.get('journal', pub.bib.get('venue', '')),
                    }
                    references.append(ref)
                    # 避免過快請求被封鎖
                    time.sleep(random.uniform(1, 3))
        except Exception as e:
            print(f"搜尋文獻時發生錯誤: {str(e)}")
            
        return references
        
    def generate_references(self, topic):
        """
        產生參考文獻列表
        """
        refs = self.search_references(topic)
        formatted_refs = []
        
        for i, ref in enumerate(refs, 1):
            authors = ref['authors'].split(' and ')[0] + ' et al.' if ' and ' in ref['authors'] else ref['authors']
            formatted_ref = f"[{i}] {authors}. {ref['title']}. {ref['venue']}, {ref['year']}."
            formatted_refs.append(formatted_ref)
            
        return formatted_refs
    
    def search_latest_papers(self, topic, limit=10):
        """搜尋最新的學術文獻"""
        url = "https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent"
        search_prompt = f"""請幫我搜尋關於「{topic}」的最新學術研究，要求：
        1. 發表年份必須在2023年後
        2. 必須來自知名期刊或會議
        3. 需要包含研究方法和結論
        4. 特別注重最新的研究發現和趨勢
        請用JSON格式返回，包含：標題、作者、發表年份、期刊名稱、摘要、DOI"""
        
        try:
            response = requests.post(
                f"{url}?key={self.api_config['GEMINI_API_KEY']}", 
                headers={'Content-Type': 'application/json'},
                json={
                    "contents": [{"parts":[{"text": search_prompt}]}],
                    "generationConfig": {
                        "temperature": 0.3,
                        "topK": 20,
                        "topP": 0.8,
                        "maxOutputTokens": 8192
                    }
                }
            )
            
            if response.status_code == 200:
                result_text = response.json()['candidates'][0]['content']['parts'][0]['text']
                try:
                    # 嘗試解析返回的 JSON 字符串
                    papers = json.loads(result_text)
                    # 確保結果是列表格式
                    if not isinstance(papers, list):
                        papers = [papers]
                    
                    # 轉換為所需的格式
                    formatted_papers = []
                    for paper in papers[:limit]:
                        formatted_paper = {
                            'title': paper.get('標題', ''),
                            'content': paper.get('摘要', ''),
                            'link': f"https://doi.org/{paper.get('DOI', '')}" if paper.get('DOI') else '#'
                        }
                        formatted_papers.append(formatted_paper)
                    return formatted_papers
                except json.JSONDecodeError:
                    print("無法解析搜尋結果為 JSON 格式")
                    # 返回空列表作為備用
                    return []
        except Exception as e:
            print(f"搜尋時發生錯誤: {e}")
            return []
            
    def generate_paper(self, topic):
        """生成完整論文"""
        try:
            # 1. 首先搜索最新文獻
            print("正在搜尋最新文獻...")
            self._progress = 10
            latest_papers = self.search_latest_papers(topic)
            
            # 2. 生成論文內容
            print("正在生成論文內容...")
            self._progress = 30
            content = self.generate_content(topic, latest_papers)
            
            if content is None:
                raise Exception("生成論文內容失敗")
            
            # 3. 生成 DOCX 格式文件
            print("正在格式化論文...")
            self._progress = 70
            
            # 直接返回生成的文件路徑
            filepath = self.generate_docx(topic, content)
            
            self._progress = 100
            return filepath
            
        except Exception as e:
            print(f"生成論文時發生錯誤: {str(e)}")
            raise