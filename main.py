import pandas as pd
import openpyxl
import requests
import os
import spacy
from spacy.lang.ru import Russian
import networkx as nx
import matplotlib.pyplot as plt


class News:
    sentences = []

    def __init__(self, date, text):
        self.date = date
        self.text = text


def parse_file():
    path = 'news_dataset.xlsx'
    xls = pd.ExcelFile(path)
    sheet_names = xls.sheet_names
    # column_names = xls.parse(sheet_name=sheet_names[0], skiprows=0, nrows=1)
    df = xls.parse(sheet_name=sheet_names[0], skiprows=995, index_col=None)
    rows = df.values.tolist()
    result = []
    for row in rows:
        result.append(News(date=row[0], text=row[1]))

    return result


def getSentences(text):
    nlp = Russian()
    nlp.add_pipe('sentencizer')
    document = nlp(text)
    nlp_text = [sent.text.strip() for sent in document.sents]
    return nlp_text


def printToken(token):
    print(token.text, "->", token.dep_)


def mergeChunk(original, chunk):
    return original + ' ' + chunk


def isRelationCandidate(token):
    deps = ["ROOT", "adj", "attr", "agent", "amod"]
    return any(subs in token.dep_ for subs in deps)


def isConstructionCandidate(token):
    deps = ["compound", "prep", "conj", "mod"]
    return any(subs in token.dep_ for subs in deps)


def processSubject(tokens):
    subject = ''
    object = ''
    relation = ''
    subjectConstruction = ''
    objectConstruction = ''
    for token in tokens:
        printToken(token)
        if "punct" in token.dep_:
            continue
        if isRelationCandidate(token):
            relation = mergeChunk(relation, token.lemma_)
        if isConstructionCandidate(token):
            if subjectConstruction:
                subjectConstruction = mergeChunk(subjectConstruction, token.text)
            if objectConstruction:
                objectConstruction = mergeChunk(objectConstruction, token.text)
        if "subj" in token.dep_:
            subject = mergeChunk(subject, token.text)
            subject = mergeChunk(subjectConstruction, subject)
            subjectConstruction = ''
        if "obj" in token.dep_:
            object = mergeChunk(object, token.text)
            object = mergeChunk(objectConstruction, object)
            objectConstruction = ''

    print(subject.strip(), ",", relation.strip(), ",", object.strip())
    return (subject.strip(), relation.strip(), object.strip())


def processSentence(sentence):
    tokens = nlp_model(sentence)
    return processSubject(tokens)


def printGraph(triples):
    G = nx.Graph()
    for triple in triples:
        G.add_node(triple[0])
        G.add_node(triple[1])
        G.add_node(triple[2])
        G.add_edge(triple[0], triple[1])
        G.add_edge(triple[1], triple[2])

    pos = nx.spring_layout(G)
    plt.figure()
    nx.draw(G, pos, edge_color='black', width=1, linewidths=1,
            node_size=500, node_color='seagreen', alpha=0.9,
            labels={node: node for node in G.nodes()})
    plt.axis('off')
    plt.show()


if __name__ == "__main__":

    news_list = parse_file()

    nlp_model = spacy.load('ru_core_news_sm')
    while True:
        print(f"Найдено {len(news_list)} новостей")
        print("Введите номер новости начиная от 1. 0 - выход")
        num = input()
        if num == 0:
            break
        else:
            news = news_list[int(num) - 1]
            news.sentences = getSentences(news.text)
            triples = []
            for sentence in news.sentences:
                triples.append(processSentence(sentence))

            printGraph(triples)
