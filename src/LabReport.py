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
    def __init__(self, filename, textTags):
        self.studentID = '000000000000'
        self.studentName = 'Unknown'
        self.filename = filename
        self.keywords = {}
        self.similarityText = {}
        self.similarityImage = []
        self.statistics = {}
        self.scores = {}
        self.imgfileMd5 = set()
        studentInfo = self.getStudentInfo(filename)
        if studentInfo:
            self.studentID, self.studentName = studentInfo
            reportData = self.getReportData(filename)
            if reportData:
                self.rawtext, self.imgfileMd5 = reportData
                self.keywords = self.parseReportText(self.rawtext, textTags)
                self.status = 'ok'
            else:
                self.status = '文件损坏或格式不符合，不能正确解析'
                print(self.filename, self.status)
        else:
            self.status = '文件名称命名不规范,不能正确获取学生信息'
            print(self.filename, self.status)

    def getStudentInfo(self, filename):
        matchObj = re.match(r'(\d{12})(.*)\.docx', os.path.basename(filename))
        if matchObj:
            id = matchObj.group(1)
            name = matchObj.group(2)
            name = re.sub('-|实验.*|－|（.*）|_|\s|\(.*\)|\+', '', name)
            return id, name
        else:
            print('文件格式不符合命名规则，无法自动识别')
            return None

    def getReportData(self, filename):
        imgdir = os.path.join(os.path.dirname(filename), 'tmp')
        if not os.path.exists(imgdir):
            os.mkdir(imgdir)
        try:
            text = docx2txt.process(filename, imgdir).replace('\n+', '\n')
        except:
            print('文件格式不正确，无法获取数据', filename)
            return None
        else:
            imgfilesMd5 = {(os.path.basename(f), md5(f)) for f in glob.glob(imgdir + "/*", recursive=True)}
            text = re.sub('[\n]+', '\n', text)  # remove all empty lines
            shutil.rmtree(imgdir)
            self.statistics['nImages'] = len(imgfilesMd5)

            return text, imgfilesMd5

    def parseReportText(self, text, tags):
        self.parsedText = {}
        keywords = {}
        totalTags = len(tags)
        for i in range(totalTags):
            if i < totalTags - 1:
                self.parsedText[tags[i]] = text[text.find(tags[i]):text.find(tags[i + 1])]
            elif i == totalTags - 1:
                self.parsedText[tags[i]] = text[text.find(tags[i]):]

        self.statistics['nKeywords'] = {}
        for tag in tags:
            self.parsedText[tag] = re.sub(tag, '', self.parsedText[tag])
            keywords[tag] = self.getKeywordFromText(self.parsedText[tag])
            self.statistics['nKeywords'][tag] = len(keywords[tag])
        return keywords

    def getKeywordFromText(self, text):
        jieba.analyse.set_stop_words('stopwords.txt')
        seg = [word for word in jieba.cut(text, cut_all=False)]
        keyWord = jieba.analyse.extract_tags('|'.join(seg), topK=100, withWeight=True, allowPOS=())

        if keyWord:
            return [one[0] for one in keyWord]
        else:
            return ["无"]

    def appendSimilarityText(self, tag, peer):
        if not self.similarityText.keys().__contains__(tag):
            self.similarityText[tag] = []

        self.similarityText[tag].append(peer)

    def appendSimilarityImage(self, peer):
        self.similarityImage.append(peer)

    def print(self):
        print(self.studentID, self.studentName, ":")
        for tag in self.keywords:
            print("\t total keywords:", self.statistics['nKeywords'][tag.title()], tag.title(), ":", self.keywords[tag])
        # for word in self.parsedText:
        #     print("\t total words:", len(self.parsedText[word]), word.title())
        print("\t total images:", self.statistics['nImages'])
        if self.similarityImage:
            print("\t same pictures in other reports:", max([img['similarity'] for img in self.similarityImage]))
        for tag in self.similarityText:
            print("\t same text", tag, "in other reports:",
                  max([txt['similarity'] for txt in self.similarityText[tag]]))

    def evaluate(self, grade, weight, remarks, textTags):
        self.scores['final'] = 0
        self.scores['remark'] = ""
        penalty = 0
        if self.similarityImage:
            penalty = max([img['similarity'] for img in self.similarityImage])
        self.scores['final'] += weight['nImages'] * grade['nImages'][self.scores["nImages"]] * (1 - penalty)
        self.scores['remark'] += remarks['nImages'][self.scores["nImages"]]
        for tag in textTags:
            penalty = 0
            if tag in self.similarityText:
                penalty = max([txt['similarity'] for txt in self.similarityText[tag]])

            self.scores['final'] += weight['nKeywords'][tag] * grade['nKeywords'][tag][
                self.scores["nKeywords"][tag]] * (1 - penalty)
            self.scores['remark'] += remarks['nKeywords'][tag][self.scores["nKeywords"][tag]]

        if self.similarityImage:
            self.scores['remark'] += '报告中图片和以下同学的报告雷同：'
            for entity in self.similarityImage:
                self.scores['remark'] += entity["studentName"]+","

        if tag in self.similarityText:
            self.scores['remark'] += '报告中'+tag+'和以下同学的报告雷同：'
            for entity in self.similarityText[tag]:
                self.scores['remark'] += entity["studentName"] + ","
