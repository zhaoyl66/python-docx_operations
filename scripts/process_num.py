import re

from docx import Document
from docx.oxml.ns import qn


class WithNumberDocxReader:
    ideographTraditional = "甲乙丙丁戊己庚辛壬癸"
    ideographZodiac = "子丑寅卯辰巳午未申酉戌亥"

    def __init__(self, docx, gap_text="\t"):
        self.docx = Document(docx)
        try:
            self.numId2style = self.get_style_data()
        except NotImplementedError:
            self.numId2style = {}  # fallback to empty if no numbering styles
        self.gap_text = gap_text
        self.cnt = {}
        self.cache = {}
        self.result = []

    @property
    def texts(self):
        if self.result:
            return self.result.copy()
        self.cnt.clear()
        self.cache.clear()
        for paragraph in self.docx.paragraphs:
            number_text = self.get_number_text(paragraph._element.pPr.numPr)
            self.result.append(number_text + paragraph.text)
        return self.result.copy()

    def get_style_data(self):
        numbering_part = self.docx.part.numbering_part._element
        abstractId2numId = {num.abstractNumId.val: num.numId for num in numbering_part.num_lst}
        numId2style = {}
        for abstractNumIdTag in numbering_part.findall(qn("w:abstractNum")):
            abstractNumId = abstractNumIdTag.get(qn("w:abstractNumId"))
            numId = abstractId2numId[int(abstractNumId)]
            for lvlTag in abstractNumIdTag.findall(qn("w:lvl")):
                ilvl = lvlTag.get(qn("w:ilvl"))
                style = {tag.tag[tag.tag.rfind("}") + 1:]: tag.get(qn("w:val"))
                         for tag in lvlTag.xpath("./*[@w:val]", namespaces=numbering_part.nsmap)}
                if "numFmt" not in style:
                    numFmtVal = lvlTag.xpath("./mc:AlternateContent/mc:Fallback/w:numFmt/@w:val",
                                             namespaces=numbering_part.nsmap)
                    if numFmtVal and numFmtVal[0] == "decimal":
                        numFmt_format = lvlTag.xpath("./mc:AlternateContent/mc:Choice/w:numFmt/@w:format",
                                                     namespaces=numbering_part.nsmap)
                        if numFmt_format:
                            style["numFmt"] = "decimal" + numFmt_format[0].split(",")[0]
                if style.get("numFmt") == "decimalZero":
                    style["numFmt"] = "decimal01"
                numId2style[(numId, int(ilvl))] = style
        return numId2style

    @staticmethod
    def int2upperLetter(num):
        result = []
        while num > 0:
            num -= 1
            remainder = num % 26
            result.append(chr(remainder + ord('A')))
            num //= 26
        return "".join(reversed(result))

    @staticmethod
    def int2upperRoman(num):
        t = [
            (1000, 'M'), (900, 'CM'), (500, 'D'),
            (400, 'CD'), (100, 'C'), (90, 'XC'),
            (50, 'L'), (40, 'XL'), (10, 'X'),
            (9, 'IX'), (5, 'V'), (4, 'IV'), (1, 'I')
        ]
        roman_num = ''
        i = 0
        while num > 0:
            val, syb = t[i]
            for _ in range(num // val):
                roman_num += syb
                num -= val
            i += 1
        return roman_num

    @staticmethod
    def int2cardinalText(num):
        if not isinstance(num, int) or num < 0 or num > 999999999:
            raise ValueError(
                "Invalid number: must be a positive integer within four digits")
        base = ["Zero", "One", "Two", "Three", "Four", "Five", "Six",
                "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen",
                "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"]
        tens = ["", "", "Twenty", "Thirty", "Fourty",
                "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]
        thousands = ["", "Thousand", "Million", "Billion"]

        def two_digits(n):
            if n < 20:
                return base[n]
            ten, unit = divmod(n, 10)
            if unit == 0:
                return f"{tens[ten]}"
            else:
                return f"{tens[ten]}-{base[unit]}"

        def three_digits(n):
            hundred, rest = divmod(n, 100)
            if hundred == 0:
                return two_digits(rest)
            result = f"{base[hundred]} hundred "
            if rest > 0:
                result += two_digits(rest)
            return result.strip()

        if num < 99:
            return two_digits(num)
        chunks = []
        while num > 0:
            num, remainder = divmod(num, 1000)
            chunks.append(remainder)
        words = []
        for i in range(len(chunks) - 1, -1, -1):
            if chunks[i] == 0:
                continue
            chunk_word = three_digits(chunks[i])
            if thousands[i]:
                chunk_word += f" {thousands[i]}"
            words.append(chunk_word)
        words = " ".join(words).lower()
        return words[0].upper() + words[1:]

    @staticmethod
    def int2ordinalText(num):
        if not isinstance(num, int) or num < 0 or num > 999999:
            raise ValueError(
                "Invalid number: must be a positive integer within four digits")
        base = ["Zero", "One", "Two", "Three", "Four", "Five", "Six",
                "Seven", "Eight", "Nine", "Ten", "Eleven", "Twelve", "Thirteen", "Fourteen",
                "Fifteen", "Sixteen", "Seventeen", "Eighteen", "Nineteen"]
        baseth = ['Zeroth', 'First', 'Second', 'Third', 'Fourth', 'Fifth', 'Sixth', 'Seventh',
                  'Eighth', 'Ninth', 'Tenth', 'Eleventh', 'Twelfth', 'Thirteenth', 'Fourteenth',
                  'Fifteenth', 'Sixteenth', 'Seventeenth', 'Eighteenth', 'Nineteenth', 'Twentieth']
        tens = ["", "", "Twenty", "Thirty", "Fourty",
                "Fifty", "Sixty", "Seventy", "Eighty", "Ninety"]
        tensth = ["", "", "Twentieth", "Thirtieth", "Fortieth",
                  "Fiftieth", "Sixtieth", "Seventieth", "Eightieth", "Ninetieth"]

        def two_digits(n):
            if n <= 20:
                return baseth[n]
            ten, unit = divmod(n, 10)
            result = tensth[ten]
            if unit != 0:
                result = f"{tens[ten]}-{baseth[unit]}"
            return result

        thousand, num = divmod(num, 1000)
        result = []
        if thousand > 0:
            if num == 0:
                return f"{WithNumberDocxReader.int2cardinalText(thousand)} thousandth"
            result.append(f"{WithNumberDocxReader.int2cardinalText(thousand)} thousand")
        hundred, num = divmod(num, 100)
        if hundred > 0:
            if num == 0:
                result.append(f"{base[hundred]} hundredth")
                return " ".join(result)
            result.append(f"{base[hundred]} hundred")
        result.append(two_digits(num))
        result = " ".join(result).lower()
        return result[0].upper() + result[1:]

    @staticmethod
    def int2Chinese(num, ch_num, units):
        if not (0 <= num <= 99999999):
            raise ValueError("仅支持小于一亿以内的正整数")

        def int2Chinese_in(num, ch_num, units):
            if not (0 <= num <= 9999):
                raise ValueError("仅支持小于一万以内的正整数")
            result = [ch_num[int(i)] + unit for i, unit in zip(reversed(str(num).zfill(4)), units)]
            result = "".join(reversed(result))
            zero_char = ch_num[0]
            result = re.sub(f"(?:{zero_char}[{units}])+", zero_char, result)
            result = result.rstrip(units[0])
            if result != zero_char:
                result = result.rstrip(zero_char)
            if result.lstrip(zero_char).startswith("一十"):
                result = result.replace("一", "")
            return result

        if num < 10000:
            result = int2Chinese_in(num, ch_num, units)
        else:
            left = num // 10000
            right = num % 10000
            result = int2Chinese_in(left, ch_num, units) + "万" + int2Chinese_in(right, ch_num, units)
        if result != ch_num[0]:
            result = result.strip(ch_num[0])
        return result

    @staticmethod
    def int2ChineseCounting(num):
        return WithNumberDocxReader.int2Chinese(num, ch_num='〇一二三四五六七八九', units='个十百千')

    @staticmethod
    def int2ChineseLegalSimplified(num):
        return WithNumberDocxReader.int2Chinese(num, ch_num='零壹贰叁肆伍陆柒捌玖', units='个拾佰仟')

    def get_number_text(self, numpr):
        if numpr is None or numpr.numId.val == 0 or not self.numId2style:
            return ""
        numId = numpr.numId.val
        ilvl = numpr.ilvl.val
        style = self.numId2style[(numId, ilvl)]
        numFmt: str = style.get("numFmt")
        lvlText = style.get("lvlText")
        if (numId, ilvl) in self.cnt:
            self.cnt[(numId, ilvl)] += 1
        else:
            self.cnt[(numId, ilvl)] = int(style["start"])
        pos = self.cnt[(numId, ilvl)]
        num_text = str(pos)
        if numFmt.startswith('decimal'):
            num_text = num_text.zfill(numFmt.count("0") + 1)
        elif numFmt == 'upperRoman':
            num_text = self.int2upperRoman(pos)
        elif numFmt == 'lowerRoman':
            num_text = self.int2upperRoman(pos).lower()
        elif numFmt == 'upperLetter':
            num_text = self.int2upperLetter(pos)
        elif numFmt == 'lowerLetter':
            num_text = self.int2upperLetter(pos).lower()
        elif numFmt == 'ordinal':
            num_text = f"{pos}{'th' if 11 <= pos <= 13 else {1: 'st', 2: 'nd', 3: 'rd'}.get(pos % 10, 'th')}"
        elif numFmt == 'cardinalText':
            num_text = self.int2cardinalText(pos)
        elif numFmt == 'ordinalText':
            num_text = self.int2ordinalText(pos)
        elif numFmt == 'ideographTraditional':
            if 1 <= pos <= 10:
                num_text = self.ideographTraditional[pos - 1]
        elif numFmt == 'ideographZodiac':
            if 1 <= pos <= 12:
                num_text = self.ideographZodiac[pos - 1]
        elif numFmt == 'chineseCounting':
            num_text = self.int2ChineseCounting(pos)
        elif numFmt == 'chineseLegalSimplified':
            num_text = self.int2ChineseLegalSimplified(pos)
        elif numFmt == 'decimalEnclosedCircleChinese':
            pass
        self.cache[(numId, ilvl)] = num_text
        for i in range(0, ilvl + 1):
            lvlText = lvlText.replace(f'%{i + 1}', self.cache.get((numId, i), ""))
        suff_text = {"space": " ", "nothing": ""}.get(style.get("suff"), self.gap_text)
        lvlText += suff_text
        return lvlText
