import docx2txt
import os.path
import glob
import re
import hashlib
import LabReport
from gensim import corpora, models, similarities

datafile = '/Users/wliu/PycharmProjects/HWAutoChecker/data/云平台管理技术/实验1'


def calculateSimilarity(reports, tag):
    corpus = []
    for report in reports:
        print(report.studentName, tag, report.keywords[tag])
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
            if value > 0.8 and report != _report:
                report.appendSimilarityReport(tag, {'studentID': _report.studentID, 'studentName': _report.studentName,
                                                    'similarity': value})
                # print(report.studentID, report.studentName, "-- 相似度 {:.2f}".format(value), _report.studentID, _report.studentName)


def traversal(path):
    files = [f for f in glob.glob(path + "**/*.docx", recursive=True)]
    reports = []
    for f in files:
        reports.append(LabReport.LabReport(f))
    calculateSimilarity(reports, "实验总结")
    # calculateSimilarity(reports, "实验步骤")

    for report in reports:
        if report.similarityReport:
            print(report.studentID, '-', report.studentName, ':')
            print(report.similarityReport)


traversal(datafile)
