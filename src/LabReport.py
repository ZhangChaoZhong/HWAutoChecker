import re
import docx2txt
import os.path
import glob
import hashlib
import shutil
import jieba.analyse


def md5(fname):
    hash_md5 = hashlib.md5()
    with open(fname, "rb") as f:
        for chunk in iter(lambda: f.read(4096), b""):
            hash_md5.update(chunk)
    return hash_md5.hexdigest()


class LabReport(object):
    def __init__(self, filename):
        self.studentID = '000000000000'
        self.studentName = 'Unknown'
        self.filename = filename
        self.keywords = {}
        self.similarityReport = {}
        tag1 = '实验步骤'
        tag2 = '实验总结'
        self.keywords[tag1] = {}
        self.keywords[tag2] = {}
        studentInfo = self.getStudentInfo(filename)
        if studentInfo:
            self.studentID, self.studentName = studentInfo
            reportData = self.getReportData(filename)
            if reportData:
                self.rawtext, self.imgfileMd5 = reportData
                self.keywords = self.parseReportText(self.rawtext)
            else:
                print('不能解析文件', self.filename)
        else:
            print("文件名命名不符合规范，不能正确获取学生信息")

    def getStudentInfo(self, filename):
        matchObj = re.match(r'(\d{12})(.*)\.docx', os.path.basename(filename))
        if matchObj:
            id = matchObj.group(1)
            name = matchObj.group(2)
            name = re.sub('-|实验.*|－|（.*）|_|\s|\(.*\)|\+', '', name)
            print(id, name)
            return id, name
        else:
            print('文件格式不符合命名规则，无法自动识别')
            return None

    def getReportData(self, filename):
        imgdir = os.path.join(os.path.dirname(filename), 'tmp')
        print(filename)
        if not os.path.exists(imgdir):
            os.mkdir(imgdir)
        try:
            text = docx2txt.process(filename, imgdir).replace('\n+', '\n')
        except:
            print('文件格式不正确，无法获取数据', filename)
            return None
        else:
            imgfilesMd5 = [(os.path.basename(f), md5(f)) for f in glob.glob(imgdir + "/*", recursive=True)]
            text = re.sub('[\n]+', '\n', text)  # remove all empty lines
            shutil.rmtree(imgdir)

            return text, imgfilesMd5

    def parseReportText(self, text):
        tag1 = '实验步骤'
        tag2 = '实验总结'
        parsedText = {}
        keywords = {}
        parsedText[tag1] = text[text.find(tag1):text.find(tag2)]
        parsedText[tag2] = text[text.find(tag2):]
        parsedText[tag1] = re.sub(tag1, '', parsedText[tag1])
        parsedText[tag2] = re.sub(tag2, '', parsedText[tag2])
        keywords[tag1] = self.getKeywordFromText(parsedText[tag1])
        keywords[tag2] = self.getKeywordFromText(parsedText[tag2])
        return keywords

    def getKeywordFromText(self, text):
        jieba.analyse.set_stop_words('stopwords.txt')
        seg = [word for word in jieba.cut(text, cut_all=False)]
        keyWord = jieba.analyse.extract_tags('|'.join(seg), topK=100, withWeight=True, allowPOS=())

        if keyWord:
            return [one[0] for one in keyWord]
        else:
            return ["无"]

    def appendSimilarityReport(self, tag, peer):
        if not self.similarityReport.keys().__contains__(tag):
            self.similarityReport[tag] = []

        self.similarityReport[tag].append(peer)
