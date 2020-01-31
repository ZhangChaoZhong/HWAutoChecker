import docx2txt
import os.path
import glob
import re
import hashlib
import LabReport
from gensim import corpora, models, similarities

datafile = '/Users/wliu/PycharmProjects/HWAutoChecker/data/云平台管理技术/实验1'


def calculateSimilarity(reports, tag, threshold=0.8):
    corpus = []
    for report in reports:
        corpus.append(report.keywords[tag])

    # 以下使用doc2bow制作语料库 使用TF-IDF模型对语料库建模
    dictionary = corpora.Dictionary(corpus)
    bow_corpus = [dictionary.doc2bow(doc) for doc in corpus]
    tfidf = models.TfidfModel(bow_corpus)
    corpus_tfidf = tfidf[bow_corpus]
    similar_matrix = similarities.MatrixSimilarity(corpus_tfidf)

    # compare
    for report in reports:
        corpus = dictionary.doc2bow(report.keywords[tag])
        test_tfidf = tfidf[corpus]
        result = similar_matrix[test_tfidf]
        for value, _report in zip(result, reports):
            if value > threshold and report != _report:
                report.appendSimilarityText(tag, {'studentID': _report.studentID, 'studentName': _report.studentName,
                                                    'similarity': value})
                # print(report.studentID, report.studentName, "-- 相似度 {:.2f}".format(value), _report.studentID, _report.studentName)

def calculateImageSimilarity(reports, threshold=0.2):
    for report in reports:
        imgfileMd5 = report.imgfileMd5
        imgSimilarity = [len(imgfileMd5.intersection(_report.imgfileMd5))/len(report.imgfileMd5) if len(report.imgfileMd5)>0 and report!=_report else 0 for _report in reports ]
        #get similarity info
        for value, _report in zip(imgSimilarity, reports):
            if value > threshold and report != _report:
                report.appendSimilarityImage( {'studentID': _report.studentID, 'studentName': _report.studentName,
                                                  'similarity': value})

def traversal(path):
    files = [f for f in glob.glob(path + "**/*.docx", recursive=True)]
    reports = []
    reportsFailed=[]
    for f in files:
        report = LabReport.LabReport(f)
        if report.status == 'ok':
            reports.append(report)
        else:
            reportsFailed.append(report)
    calculateSimilarity(reports, "实验总结")
    # calculateSimilarity(reports, "实验步骤")
    calculateImageSimilarity(reports)

    print("total number of files:", len(files))
    print("total number of correctly parsed records:", len(reports))
    print("total number of failed to parse records:", len(reportsFailed))
    for report in reportsFailed:
        print(report.filename, report.status)

    for report in reports:
        report.print()
    #     if len(report.similarityText):
    #         print(report.studentID, '-', report.studentName, ':')
    #         print(report.similarityText)
    #     if len(report.similarityImage):
    #         print(report.studentID, '-', report.studentName, ':')
    #         print(report.similarityImage)


traversal(datafile)
