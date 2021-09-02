import re
import math
import pandas as pd
import pickle
import datetime
import dill


class Build_Index:
    def __init__(self, files):
        self.tf = {}
        self.idf ={}
        self.df = {}
        self.files = files
        self.filenames = list(files.keys())
        self.file_to_terms = self.process_files()
        self.regdex = self.reg_index()
        self.total_index = self.execute()
        self.vectors = self.vectorize()
        self.mags = self.magnitudes(self.filenames)
        self.populate_scores()

    def process_files(self):
        files_to_term = {}
        for file in self.filenames:
            files_to_term[file] = self.files[file].split()
        return files_to_term

    def index_one_file(self, termlist):
        file_Index = {}
        for index, word in enumerate(termlist):
            if word in file_Index.keys():
                file_Index[word].append(index)
            else:
                file_Index[word] = [index]
        return file_Index


    def make_indices(self, termlist):
        total = {}
        for filename in termlist.keys():
            total[filename] = self.index_one_file(termlist[filename])
        return total


    def full_index(self):
        total_index = {}
        indie_indices = self.regdex
        for filename in indie_indices.keys():
            self.tf[filename] = {}
            for word in indie_indices[filename].keys():
                self.tf[filename][word] = len(indie_indices[filename][word])
                if word in self.df.keys():
                    self.df[word] += 1
                else:
                    self.df[word] = 1
                if word in total_index.keys():
                    if filename in total_index[word].keys():
                        total_index[word][filename].append(indie_indices[filename][word][:])
                    else:
                        total_index[word][filename] = indie_indices[filename][word]
                else:
                    total_index[word] = {filename: indie_indices[filename][word]}
        return total_index


    def vectorize(self):
        vectors = {}
        for filename in self.filenames:
            vectors[filename] = [len(self.regdex[filename][word]) for word in self.regdex[filename].keys()]
        return vectors


    def document_frequency(self, term):
        if term in self.total_index.keys():
            return len(self.total_index[term].keys())
        else:
            return 0


    def collection_size(self):
        return len(self.filenames)


    def magnitudes(self, documents):
        mags = {}
        for document in documents:
            mags[document] = pow(sum(map(lambda x: x**2, self.vectors[document])), .5)
        return mags


    def term_frequency(self, term, document):
        return self.tf[document][term]/self.mags[document] if term in self.tf[document].keys() else 0


    def populate_scores(self):
        for filename in self.filenames:
            for term in self.get_Uniques():
                self.tf[filename][term] = self.term_frequency(term, filename)
                if term in self.df.keys():
                    self.idf[term] = self.idf_func(self.collection_size(), self.df[term])
                else:
                    self.df[term] = 0

        return self.df, self.tf, self.idf


    def idf_func(self, N, N_t):
        if N_t != 0:
            return math.log(N/N_t)
        else:
            return 0


    def generateScore(self, term, document):
        return self.tf[document][term] * self.idf[term]


    def execute(self):
        return self.full_index()


    def reg_index(self):
        return self.make_indices(self.file_to_terms)


    def get_Uniques(self):
        return self.total_index.keys()


if __name__ == "__main__":

    path_to_doc = r"C:\Users\Oshchepkov-VA\PycharmProjects\NLPAnalyzer\File_Hundler\output\Res_test.xlsx"
    docs = pd.read_excel(path_to_doc)

    paths = docs["path"]
    texts = docs["text_lem"]

    dict_docs = {}
    for ind, path in enumerate(paths):
        dict_docs[path] = texts[ind]

    start_time = datetime.datetime.now()

    search = Build_Index(dict_docs)

    # connect = sqlite3.connect("indexed_doc.db")
    #
    # create_some_table = '''create table '''
    #
    # cursor = connect.cursor(create_some_table)
    #
    print(search)
    with open("index_test.dill", "wb") as index_file:
        dill.dump(search, index_file)

    print("The end!")
    print(str(datetime.datetime.now()-start_time))
    # print(search)
    #
    # print(search.free_text_query("войнолович владимирович"))
