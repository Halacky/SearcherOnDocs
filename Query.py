import pickle
import pandas as pd
import datetime
import Build_Index
import re
import math
import dill
import pymorphy2
import ast


class Query:
    def __init__(self):
        with open("index_test.dill", "rb") as index_file:
            self.index = dill.load(index_file)
        #self.index = Build_Index(self.files)
        self.inverted_index = self.index.total_index
        self.regular_index = self.index.regdex
        self.morph_analyzer = pymorphy2.MorphAnalyzer()

    def one_word_query(self, word):
        # вызов обработки слова и приведение к нормально форме
        if word in self.inverted_index.keys():
            return self.rank_results([filename for filename in self.inverted_index[word].keys()], word)
        else:
            return []


    def free_text_query(self, words):
        # вызов обработки строки и приведение её к нормальной форме
        result = []
        for word in words.split():
            result += self.one_word_query(word)
        return self.rank_results(list(set(result)), words)


    def phrase_query(self, words):
        # вызов обработки фразы и приведение её к нормальной форме
        list_of_lists, result = [], []
        for word in words.split():
            list_of_lists.append(self.one_word_query(word))
        setted = set(list_of_lists[0]).intersection(*list_of_lists)
        for filename in setted:
            temp = []
            for word in words.split():
                temp.append(self.inverted_index[word][filename][:])
            for i in range(len(temp)):
                for ind in range(len(temp[i])):
                    temp[i][ind] -= i
            if set(temp[0]).intersection(*temp):
                result.append(filename)
        return self.rank_results(result, words)


    def not_excact_match_query(self, words, discr_coef):
        list_of_list, result = [], []
        tokens = words.split()
        for word in tokens:
            list_of_list.append(self.one_word_query(word))
        setted = set(list_of_list[0]).intersection(*list_of_list)
        for filename in setted:
            temp = []
            mas_check = []
            for word in tokens:
                temp.append(self.inverted_index[word][filename][:])
            for i in range(0, len(temp) - 1):
                for elem_i in temp[i]:
                    for elem_j in temp[i + 1]:
                        if elem_j > elem_i and 1 < elem_j - elem_i <= discr_coef:
                            mas_check.append(True)
                            break
                    if len(temp[i + 1]) != 1:
                        break
            if len(mas_check) + 1 == len(tokens):
                result.append(filename)
        return self.rank_results(result, words)


    def search_entity(self, text_query, df):
        tokens = [self.morph_analyzer.parse(word.rstrip())[0].normal_form for word in text_query.split()]
        print(tokens)
        list_of_list, result = [], {}
        for word in tokens:
            list_of_list.append(self.one_word_query(word))
        setted = set(list_of_list[0]).intersection(*list_of_list)
        for filename in setted:
            temp = []
            mas_check = []
            for word in tokens:
                temp.append(self.inverted_index[word][filename][:])
            df_ = df[df["path"] == filename]
            print(df_)
            #tokens_in_df = list(filter(None ,map(str.strip, df_["tokens"].values[0].strip('][').replace("\'", " ").split(","))))
            #tokens_in_df = df_["text_lem"].values[0].split(" ")
            tokens_in_df = self.index.file_to_terms[filename]
            print(tokens_in_df)
            mas_entity = []
            for index in temp:
                for ind in index:
                    qqq = " ".join(tokens_in_df[ind:ind+15])
                    print(qqq)
                    index_entity = re.search(word, qqq)
                    unit = re.search(r"(кв[\s\.]м|м[\s\.]кв|м2)", qqq)
                    print(f"Юнит: {unit}")
                    check_values = {}
                    if unit != None:
                        vals = re.findall(r"(\s\d+\.\d+|\s\d+)", qqq)
                        if vals != []:
                            for elem in vals:
                                res = re.search(elem, qqq)
                                if index_entity.end() <= res.start() <= unit.start():
                                    mas_entity.append(f"Площадь = {res.group(0)}")
                                    continue
                                else:
                                    print(f"Результат поиска: {res}")
                                    print(unit.start() - res.end())
                                    check_values[res.group(0)] = abs(unit.start() - res.end())
                            if check_values != {}:
                                key_with_min = min(check_values, key=lambda num: check_values[num])
                                print(f"Площадь = {key_with_min}")
                                mas_entity.append(f"Площадь = {key_with_min}")
            if mas_entity != []:
                result[filename] = mas_entity
        return result


    def make_vectors(self, documents):
        vecs = {}
        for doc in documents:
            docVec = [0]*len(self.index.get_Uniques())
            for ind, term in enumerate(self.index.get_Uniques()):
                docVec[ind] = self.index.generateScore(term, doc)
            vecs[doc] = docVec
        return vecs


    def query_vec(self, query):
        queryls = query.split()
        queryVec = [0]*len(queryls)
        index = 0
        for ind, word in enumerate(queryls):
            queryVec[index] = self.query_freq(word, query)
            index += 1
        queryidf = [self.index.idf[word] for word in self.index.get_Uniques()]
        magnitude = pow(sum(map(lambda x: x ** 2, queryVec)), .5)
        freq = self.term_freq(self.index.get_Uniques(), query)
        tf = [x/magnitude for x in freq]
        final = [tf[i]*queryidf[i] for i in range(len(self.index.get_Uniques()))]
        return final


    def query_freq(self, term, query):
        count = 0
        for word in query.split():
            if word == term:
                count += 1
        return count


    def term_freq(self, terms, query):
        temp = [0]*len(terms)
        for i, term in enumerate(terms):
            temp[i] = self.query_freq(term, query)
        return temp


    def dot_product(self, doc1, doc2):
        if len(doc1) != len(doc2):
            return 0
        return sum([x*y for x,y in zip(doc1, doc2)])


    def rank_results(self, result_docs, query):
        vectors = self.make_vectors(result_docs)
        queryVec = self.query_vec(query)
        results = [[self.dot_product(vectors[result], queryVec), result] for result in result_docs]
        results.sort(key=lambda x: x[0])
        results = [x[1] for x in results]
        return results


if __name__ == "__main__":

    obj_query = Query()

    path_to_doc = r"C:\Users\Oshchepkov-VA\PycharmProjects\NLPAnalyzer\File_Hundler\output\Res_test.xlsx"
    df = pd.read_excel(path_to_doc)

    writer = pd.ExcelWriter(r"C:\Users\Oshchepkov-VA\PycharmProjects\NLPAnalyzer\File_Hundler\output\query_test.xlsx")
    print(obj_query.inverted_index["площадь"])
    start_time1 = datetime.datetime.now()
    text_query = "площадь"
    result = obj_query.search_entity(text_query, df)
    print(result)
    df_query_search_entity = df.loc[df["path"].isin(result.keys())]
    df_query_search_entity["entity"] = df_query_search_entity["path"].apply(lambda x: result[x])
    df_query_search_entity.to_excel(writer, sheet_name=text_query)
    # result_query_one_word = obj_query.one_word_query(text_query_one_word)
    # print(result_query_one_word)
    # df_query_one_word = df.loc[df["path"].isin(result_query_one_word)]
    # df_query_one_word.to_excel(writer, sheet_name=text_query_one_word)
    print(str(datetime.datetime.now() - start_time1))

    # text_query_one_word = "площадь"
    # result_query_one_word = obj_query.one_word_query(text_query_one_word)
    # print(result_query_one_word)
    # # df_query_one_word = df.loc[df["path"].isin(result_query_one_word)]
    # # df_query_one_word.to_excel(writer, sheet_name=text_query_one_word)
    # print(str(datetime.datetime.now() - start_time1))
    #
    # start_time2 = datetime.datetime.now()
    # #text_query_phrase = "г Нефтекамск ул Строителей д 51А"
    # text_query_phrase = "мыла машину"
    # result_query_phrase = obj_query.not_excact_match_query(text_query_phrase, 5)
    # print(result_query_phrase)
    # # df_query_phrase = df.loc[df["path"].isin(result_query_phrase)]
    # # df_query_phrase.to_excel(writer, sheet_name=text_query_phrase[:30])
    # print(str(datetime.datetime.now() - start_time2))

    # start_time3 = datetime.datetime.now()
    # text_query_free_text = ""
    # result_query_free_text = obj_query.free_text_query(text_query_free_text)
    # #print(result_query_free_text)
    # df_query_free_text = df.loc[df["path"].isin(result_query_free_text)]
    # df_query_free_text.to_excel(writer, sheet_name=text_query_free_text[:30])
    # print(str(datetime.datetime.now() - start_time3))

    writer.save()
    writer.close()




