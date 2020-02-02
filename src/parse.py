import docx2txt
import os.path
import glob
import re
import hashlib
import LabReport
from gensim import corpora, models, similarities
import pandas as pd
import numpy as np
from scipy import stats
from collections import Counter
import matplotlib.pyplot as plt
import openpyxl as pyxl

datafile = '/Users/wliu/PycharmProjects/HWAutoChecker/data/云平台管理技术/实验1'


def calculateTextSimilarity(reports, tag, threshold=0.8):
    corpus = []
    for report in reports.values():
        corpus.append(report.keywords[tag])

    # 以下使用doc2bow制作语料库 使用TF-IDF模型对语料库建模
    dictionary = corpora.Dictionary(corpus)
    bow_corpus = [dictionary.doc2bow(doc) for doc in corpus]
    tfidf = models.TfidfModel(bow_corpus)
    corpus_tfidf = tfidf[bow_corpus]
    similar_matrix = similarities.MatrixSimilarity(corpus_tfidf)

    # compare
    for id, report in reports.items():
        corpus = dictionary.doc2bow(report.keywords[tag])
        test_tfidf = tfidf[corpus]
        result = similar_matrix[test_tfidf]
        for value,(_id, _report) in zip(result, reports.items()):
            if value > threshold and id != _id:
                report.appendSimilarityText(tag, {'studentID': _report.studentID, 'studentName': _report.studentName,
                                                    'similarity': value})
                # print(report.studentID, report.studentName, "-- 相似度 {:.2f}".format(value), _report.studentID, _report.studentName)

def calculateImageSimilarity(reports, threshold=0.2):
    for id, report in reports.items():
        imgfileMd5 = report.imgfileMd5
        imgSimilarity = [len(imgfileMd5.intersection(_report.imgfileMd5))/len(report.imgfileMd5) if len(report.imgfileMd5)>0 and id != _id else 0 for _id, _report in reports.items() ]
        #get similarity info
        for value, (_id, _report) in zip(imgSimilarity, reports.items()):
            if value > threshold and id != _id:
                report.appendSimilarityImage( {'studentID': _report.studentID, 'studentName': _report.studentName,
                                                  'similarity': value})


def drawHistgram(data):
    data_pd=pd.Series(data)
    l_unique_data = list(data_pd.value_counts().index)  # 该数据中的唯一值
    l_num = list(data_pd.value_counts())  # 唯一值出现的次数
    plt.bar(l_unique_data, l_num, width=0.1)
    plt.show()


def getSectionPoint(dataArray):
    x=np.array(dataArray)
    mean, std = x.mean(), x.std(ddof=1)
    a90 = stats.norm.interval(0.8, loc=mean, scale=std)
    a80 = stats.norm.interval(0.5, loc=mean, scale=std)
    return (0, a90[0], a80[0], a80[1], a90[1], x.max())


def statisticData2Section(statisticData):
    sections = getSectionPoint(statisticData)
    result = pd.cut(statisticData, sections, labels=['E', 'D', 'C', 'B', 'A'])
    return result


def traversal(path):
    files = [f for f in glob.glob(path + "**/*.docx", recursive=True)]
    reports = {}
    reportsFailed = {}
    textTags = ['实验步骤', '实验总结']
    # get raw text from files
    for f in files:
        report = LabReport.LabReport(f, textTags)
        if report.status == 'ok':
            reports[report.studentID]=report
        else:
            reportsFailed[report.studentID]=report

    # calculate similarity
    calculateImageSimilarity(reports)
    for tag in textTags:
        calculateTextSimilarity(reports, tag)

    print("total number of files:", len(files))
    print("total number of correctly parsed records:", len(reports))
    print("total number of failed to parse records:", len(reportsFailed))
    for report in reportsFailed.values():
        print(report.filename, report.status)

    statisticImages = [report.statistics['nImages'] for report in reports.values()]

    statisticText={}
    for tag in textTags:
        statisticText[tag] = [report.statistics['nKeywords'][tag] for report in reports.values()]

    sections = getSectionPoint(statisticImages)

    result={}
    result["images"] = pd.cut(statisticImages, sections, labels=['E','D','C','B','A'])


    for tag in textTags:
        sections = getSectionPoint(statisticText[tag])
        print(sections)
        # drawHistgram(statisticText[tag])
        result[tag] = pd.cut(statisticText[tag], sections, labels=['E','D','C','B','A'])
        print(result)

    for report, img, text1, text2 in zip(reports.values(), result["images"], result["实验步骤"], result['实验总结']):
        print (report.studentID, report.studentName, report.statistics["nImages"],report.statistics['nKeywords']["实验步骤"], report.statistics['nKeywords']["实验总结"],\
               img, text1, text2)
        report.scores['nImages'] = img
        report.scores['nKeywords']= dict()
        report.scores['nKeywords']['实验步骤'] = text1
        report.scores['nKeywords']['实验总结'] = text2

    # drawHistgram(statisticText[textTags[1]])
    return reports

def saveReportInfo(filename, sheetname, reports):
    wb = pyxl.load_workbook(filename)
    if sheetname in wb.sheetnames:
        ws = wb[sheetname]
    else:
        ws = wb.copy_worksheet(wb[wb.sheetnames[0]])
    ws.title = sheetname
    columnStart = 3
    ws['D1'] = '图片数量'
    ws['E1'] = '实验步骤'
    ws['F1'] = '实验总结'
    ws['G1'] = '图片数量'
    ws['H1'] = '实验步骤'
    ws['I1'] = '实验总结'
    for row in ws.iter_rows(min_row=2):
        studentID = row[0].value
        if studentID in reports:
            report = reports[row[0].value]
            reports[row[0].value].print()
            row[columnStart].value = report.statistics["nImages"]
            row[columnStart+1].value = report.statistics["nKeywords"]['实验步骤']
            row[columnStart+2].value = report.statistics["nKeywords"]['实验总结']
            row[columnStart+ 3].value = report.scores["nImages"]
            row[columnStart + 4].value = report.scores["nKeywords"]['实验步骤']
            row[columnStart + 5].value = report.scores["nKeywords"]['实验总结']
        for i in row:
            print(i.value,end="\t")

        print("")

    wb.save(filename)
    wb.close()

def main():
    reports = traversal(datafile)
    xlsFileName = "/Users/wliu/PycharmProjects/HWAutoChecker/data/云平台管理技术/成绩登记表.xlsx"
    saveReportInfo(xlsFileName, '实验1', reports)


if __name__ == "__main__":
    main()

