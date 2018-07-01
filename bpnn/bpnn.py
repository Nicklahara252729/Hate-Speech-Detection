# -*- coding: utf-8 -*-
"""
Created on Fri Dec 01 21:07:19 2017

@author: Infinity#munir
"""
import math
import xlrd
import xlsxwriter
import re
import nltk
import random

# import Sastrawi package
from Sastrawi.Stemmer.StemmerFactory import StemmerFactory
# create stemmer
factory = StemmerFactory()
stemmer = factory.create_stemmer()

class Preprocessing:
    
    attribute = 0
    data_train_size = 0
    data_test_size = 0
    data_train = []
    data_test = []
    emoticons = []
    stopwords = []
    sentistrength = []
    tokens = []
    documents = []
    documents_test = []
    bpnn_bow = []
    bpnn_lbf = []
    bpnn_bow_test = []
    bpnn_lbf_test = []
    bpnn_data_train = []
    bpnn_data_test = []
    bpnn_output = []
    bpnn_output_test = []
    
    # inisialisasi data latih dan data uji dan ukuran masing-masing data    
    def __init__(self, dtrain_size, dtest_size, attr):
        self.attribute = attr
        self.data_train_size = dtrain_size
        self.data_test_size = dtest_size
        self.data_train = [[0 for m in range(self.attribute)] for n in range(self.data_train_size)]
        self.data_test = [[0 for m in range(self.attribute)] for n in range(self.data_test_size)]        
    
    # membuka dan menyimpan dataset
    def open_dataset(self, location, filename, sheet_name):
        # membuka workbook
        workbook = xlrd.open_workbook(location+filename)
        worksheet = workbook.sheet_by_name(sheet_name)
        
        # menyimpan data latih dan uji dari excel
        counter = 0
        for i in range(self.data_train_size+self.data_test_size):
            for j in range(self.attribute):
                if i < self.data_train_size:
                    # menyimpan data latih
                    self.data_train[counter][j] = worksheet.cell(i, j).value
                else:
                    # menyimpan data uji
                    self.data_test[counter][j] = worksheet.cell(i, j).value    
                                  
            # mereset counter jika data latih sudah disimpan semua
            if counter == self.data_train_size-1:
                counter = 0
            else:
                counter += 1        

    # mengambil emoticon dari database
    def get_emoticons(self):
        loc = 'D:/K/S\'\'\'\'\'\'\'/Skripsi saya/Data/sentistrength_id-master/'
        self.emoticons = open(loc+'emoticon_id.txt', 'r').read().splitlines()

        temp_emoticons = []
        for sw in self.emoticons:
            temp_emoticons.append(sw.split(' | '))
            self.emoticons = temp_emoticons

        #print self.emoticons
    
    # mengambil kata sentimen dari database (lexicon)
    def get_sentistrength(self):
        loc = 'D:/K/S\'\'\'\'\'\'\'/Skripsi saya/Data/sentistrength_id-master/'
        self.sentistrength = open(loc+'sentiwords_id.txt', 'r').read().splitlines()
#        self.sentistrength = open(loc+'sentiwords_id - edit.txt', 'r').read().splitlines()

        temp_sentistrength = []
        for ss in self.sentistrength:
            temp_sentistrength.append(ss.split(':'))
            self.sentistrength = temp_sentistrength

#        print self.sentistrength
    
    # mengambil stopwords dari database
    def get_stopwords(self):
        loc = 'D:/K/S\'\'\'\'\'\'\'/Skripsi saya/Data/stopwords-id-master/'
        self.stopwords = open(loc+'stopwords-id.txt', 'r').read().splitlines()

#        print self.stopwords

    # membersihkan tweet
    def tweets_cleaning(self):
        # membersihkan data latih
        train_tweets_clean = [[0 for m in range(self.attribute)] for n in range(self.data_train_size)]
        for x in range(self.data_train_size):
            train_tweets_clean[x][0] = self.data_train[x][0]
            train_tweets_clean[x][1] = ' '.join(re.sub("(@[A-Za-z0-9_]+)|([#\t])|(\w+:\/\/\S+)|([0-9]+)","",self.data_train[x][1]).split())
        self.data_train = train_tweets_clean
        
        # membersihkan data uji
        test_tweets_clean = [[0 for m in range(self.attribute)] for n in range(self.data_test_size)]        
        for x in range(self.data_test_size):
            train_tweets_clean[x][0] = self.data_test[x][0]
            test_tweets_clean[x][1] = ' '.join(re.sub("(@[A-Za-z0-9_]+)|([#\t])|(\w+:\/\/\S+)|([0-9]+)","",self.data_test[x][1]).split())
        self.data_test = test_tweets_clean
        
#        print self.data_train
#        print self.data_test
    
    # melakukan case folding
    def case_folding(self):
        # case folding data latih
        for x in range(self.data_train_size):
            self.data_train[x][1] = self.data_train[x][1].lower()
        
        # case folding data uji
        for x in range(self.data_test_size):
            self.data_test[x][1] = self.data_test[x][1].lower()
        
#        print self.data_train
#        print self.data_test
       
    # memisahkan emoticon dari tweet
    def separate_emoticons(self):
        temp_data_train = []
        temp_data_test = []
        
        for x in range(self.data_train_size):
            temp_emoticons = ''
            for i in range(len(self.emoticons)):
                if self.emoticons[i][0] in self.data_train[x][1] and self.emoticons[i][0] not in temp_emoticons:
                    temp_emoticons += self.emoticons[i][0]+' '
                    
                    temp_string = self.data_train[x][1].replace(''+self.emoticons[i][0]+'','')
                    self.data_train[x][1] = temp_string
                                   
            temp_data_train.append([self.data_train[x][0],self.data_train[x][1],temp_emoticons.strip()])
        
        for x in range(self.data_test_size):
            temp_emoticons = ''
            for i in range(len(self.emoticons)):
                if self.emoticons[i][0] in self.data_test[x][1] and self.emoticons[i][0] not in temp_emoticons:
                    temp_emoticons += self.emoticons[i][0]+' '
                    
                    temp_string = self.data_test[x][1].replace(''+self.emoticons[i][0]+'','')
                    self.data_test[x][1] = temp_string                                                    
            temp_data_test.append([self.data_test[x][0],self.data_test[x][1],temp_emoticons.strip()])
         
        self.data_train = temp_data_train
        self.data_test = temp_data_test
#        print self.data_train
#        print self.data_test

    # melakukan proses filtering
    def filtering(self):        
        for x in range(self.data_train_size):
            temp_string = ' '.join(re.sub("([^0-9A-Za-z \t])|([0-9]+)","",self.data_train[x][1]).split())
            self.data_train[x][0] = self.data_train[x][0]
            self.data_train[x][1] = temp_string
            self.data_train[x][2] = self.data_train[x][2]
        
        for x in range(self.data_test_size):
            temp_string = ' '.join(re.sub("([^0-9A-Za-z \t])|([0-9]+)","",self.data_test[x][1]).split())
            self.data_test[x][0] = self.data_test[x][0]
            self.data_test[x][1] = temp_string
            self.data_test[x][2] = self.data_test[x][2]
        
#        print self.data_train
#        print self.data_test
    
    # tokenizing dan stemming data latih
    def sub_preprocessing(self):
        tokens_word = []
        tokens_emot = []
        
        # memasukkan emoticon ke token                   
        for row in self.data_train:
            if row[2]:
                if ' ' in row[2]:
                    temp = row[2].split(' ')
                    for x in temp:
                        tokens_emot.append(x)
                else:
                    tokens_emot.append(row[2])
        tokens_emot = list(set(tokens_emot))
        
        # memasukkan kata ke token dan ke masing-masing dokumen
        for row in self.data_train:
            word = nltk.word_tokenize(row[1])
            
            for x in range(len(word)):
                word[x] = stemmer.stem(str(word[x]))
            
            tokens_word.extend(word)
            
            if row[2]:
                if ' ' in row[2]:
                    temp = row[2].split(' ')
                    for x in temp:
                        word.append(x)
                else:
                    word.append(row[2])
            self.documents.append((word, row[0]))
        
        temp_tokens = []
        for t in tokens_word:
            if not t in self.stopwords:
                temp_tokens.append(t)

        tokens_word = temp_tokens
        tokens_word = list(set(tokens_word))
        
        tokens_word.sort()
        self.tokens = tokens_emot + tokens_word
        
        #print self.tokens
    
    # tokenizing dan stemming data uji
    def sub_preprocessing_test(self):        
        # memasukkan kata ke token dan ke masing-masing dokumen
        for row in self.data_test:
            word = nltk.word_tokenize(row[1])
            
            for x in range(len(word)):
                word[x] = stemmer.stem(str(word[x]))
            
            if row[2]:
                if ' ' in row[2]:
                    temp = row[2].split(' ')
                    for x in temp:
                        word.append(x)
                else:
                    word.append(row[2])
            self.documents_test.append((word, row[0]))
        
#        print self.documents_test
                
    # membuat data latih untuk bpnn dengan metode bag of words
    def bags_of_words(self):
        for doc in self.documents:
            bow = []
            tokens_word = doc[0]
            for t in self.tokens:
                bow.append(1) if t in tokens_word else bow.append(0)

            self.bpnn_bow.append(bow)
        
#        print self.documents
#        print self.tokens
#        print self.bpnn_data_train

    # membuat data uji untuk bpnn dengan metode bag of words
    def bags_of_words_test(self):
        for doc in self.documents_test:
            bow = []
            tokens_word = doc[0]
            for t in self.tokens:
                bow.append(1) if t in tokens_word else bow.append(0)

            self.bpnn_bow_test.append(bow)
        
    
    # membuat data latih untuk bpnn dengan metode lexicon based features
    def lexicon_based_features(self):
        for doc in self.documents:
            lbf = [0 for i in range(4)]
            tokens_word = doc[0]
            for i in range(len(tokens_word)):
                for j in range(len(self.sentistrength)):
                    if tokens_word[i] == self.sentistrength[j][0]:
                        if int(self.sentistrength[j][1]) > 0:
                            lbf[0] = 1
                            continue
                        if int(self.sentistrength[j][1]) < 0:
                            lbf[1] = 1
                            continue
                    if tokens_word[i] != self.sentistrength[j][0]:
                        lbf[2] = 1
                        continue
                for j in range(len(self.emoticons)):
                    if tokens_word[i] in self.emoticons[j][0]:
                        lbf[3] = 1
                        continue
            self.bpnn_lbf.append(lbf)
        
#        print self.bpnn_lbf

    # membuat data uji untuk bpnn dengan metode lexicon based features
    def lexicon_based_features_test(self):
        for doc in self.documents_test:
            lbf = [0 for i in range(4)]
            tokens_word = doc[0]
            for i in range(len(tokens_word)):
                for j in range(len(self.sentistrength)):
                    if tokens_word[i] == self.sentistrength[j][0]:
                        if int(self.sentistrength[j][1]) > 0:
                            lbf[0] = 1
                            continue
                        if int(self.sentistrength[j][1]) < 0:
                            lbf[1] = 1
                            continue
                    if tokens_word[i] != self.sentistrength[j][0]:
                        lbf[2] = 1
                        continue
                for j in range(len(self.emoticons)):
                    if tokens_word[i] in self.emoticons[j][0]:
                        lbf[3] = 1
                        continue
            self.bpnn_lbf_test.append(lbf)
        
    
    # membuat data latih untuk bpnn dari bow dan lbf
    def join_data_train(self):
        for x in range(self.data_train_size):
            self.bpnn_data_train.append(self.bpnn_bow[x]+self.bpnn_lbf[x])
        return self.bpnn_data_train
#        print self.bpnn_data_train[0]

    # menyimpan data latih untuk pembelajaran bpnn ke excel
    def export_data_train(self):
        workbook = xlsxwriter.Workbook('DataLatihBPNN.xlsx')
        worksheet = workbook.add_worksheet('Data Latih')
        
        row = 0
        col = 0

        for i in range(len(self.bpnn_data_train)):
            for j in range(len(self.bpnn_data_train[i])):
                worksheet.write_number(row, col, self.bpnn_data_train[i][j])
                col += 1        
            col = 0
            row += 1

        workbook.close()
    
    # menyimpan data uji untuk pengujian bpnn ke excel
    def export_data_test(self):
        workbook = xlsxwriter.Workbook('DataUjiBPNN.xlsx')
        worksheet = workbook.add_worksheet('Data Uji')
        
        row = 0
        col = 0

        for i in range(len(self.bpnn_data_test)):
            for j in range(len(self.bpnn_data_test[i])):
                worksheet.write_number(row, col, self.bpnn_data_test[i][j])
                col += 1        
            col = 0
            row += 1

        workbook.close()

    # menyimpan data uji untuk pengujian bpnn ke excel
    def export_fitur_input(self):
        workbook = xlsxwriter.Workbook('FiturInputBPNN.xlsx')
        worksheet = workbook.add_worksheet('Input')
        
        row = 1
        
        worksheet.write_string(0, 0, 'No')
        worksheet.write_string(0, 1, 'Token')
        
        for i in range(len(self.tokens)):
            worksheet.write_number(row, 0, i)
            worksheet.write_string(row, 1, self.tokens[i])
            row += 1                

        workbook.close()
        
    # membuat data uji untuk bpnn dari bow dan lbf
    def join_data_test(self):
        for x in range(self.data_test_size):
            self.bpnn_data_test.append(self.bpnn_bow_test[x]+self.bpnn_lbf_test[x])        
        return self.bpnn_data_test

    # menyimpan output data latih dan data uji untuk pengujian bpnn ke excel
    def export_output_train_test(self):
        workbook = xlsxwriter.Workbook('OutputBPNN.xlsx')
        worksheet = workbook.add_worksheet('Target Output')
        
        row = 0
        col = 0

        for i in range(self.data_train_size):
            worksheet.write_number(row, col, int(self.data_train[i][0]))
            row += 1        

        row = self.data_train_size
        for i in range(self.data_test_size):
            worksheet.write_number(row, col, int(self.data_test[i][0]))
            row += 1        
        
        workbook.close()
    
    # membuat target keluaran untuk bpnn
    def get_bpnn_output(self):
        for x in range(self.data_train_size):
            self.bpnn_output.append(int(self.data_train[x][0]))
        
        return self.bpnn_output
#        print self.bpnn_output

    # membuat target keluaran uji untuk bpnn
    def get_bpnn_output_test(self):
        for x in range(self.data_test_size):
            self.bpnn_output_test.append(int(self.data_test[x][0]))        
        return self.bpnn_output_test
    
    # mencetak data latih
    def print_data_train(self):
        print self.data_train
        
    # mencetak data uji
    def print_data_test(self):
        print self.data_test
        

class BPNN:
    # input (data latih)
    x = []
    # target (data latih)
    t = []
    # test (data uji)
    test = []
    # target (data uji)
    tt = []
    
    # jumlah data latih
    train_size = 0
    # jumlah data uji
    test_size = 0
    
    # jumlah masing-masing lapisan
    # jumlah lapisan masukan
    input_layer_size = 0
    # jumlah lapisan tersembunyi
    hidden_layer_size = 0
    # jumlah lapisan keluaran
    output_layer_size = 0
    
    # bobot masing-masing lapisan
    # bobot bias pada lapisan tersembunyi
    v0 = []
    # bobot pada lapisan tersembunyi
    v = []
    # bobot bias pada lapisan keluaran
    w0 = []
    # bobot pada lapisan keluaran
    w = []
    
    # learning rate
    alfa = 0.0
    # maksimum epoch
    maxEpoch = 0
    # target Error
    ERR = 0.0
    # Mean Square Error pada sistem
    MSE = 0.0

    # inisialisasi variabel bpnn
    def initialize(self, x, t, inputlsize, hiddenlsize, outputlsize, 
                       v0, v, w0, w, alfa, maxEpoch, err, train_size, test_size):
        self.x = x
        self.t = t
        self.input_layer_size = inputlsize
        self.hidden_layer_size = hiddenlsize
        self.output_layer_size = outputlsize
        self.v0 = v0
        self.v = v
        self.w0 = w0
        self.w = w
        self.alfa = alfa
        self.maxEpoch = maxEpoch
        self.ERR = err
        self.train_size = train_size
        self.test_size = test_size
    
    # inisialisasi variabel bpnn
    def initialize2(self, inputlsize, hiddenlsize, outputlsize, 
                   alfa, maxEpoch, err, train_size, test_size):        
        self.input_layer_size = inputlsize
        self.hidden_layer_size = hiddenlsize
        self.output_layer_size = outputlsize
        self.alfa = alfa
        self.maxEpoch = maxEpoch
        self.ERR = err
        self.train_size = train_size
        self.test_size = test_size
    
    # inisialisasi bobot dan bias
    def initialize_weight_bias(self):
        self.w0 = [0 for m in range(self.output_layer_size)]
        for k in range(self.hidden_layer_size):
                self.w0[k] = random.uniform(-0.5, 0.5)        
                
        self.w = [[0 for m in range(self.hidden_layer_size)] for n in range(self.output_layer_size)]
        for j in range(self.output_layer_size):
            for k in range(self.hidden_layer_size):
                self.w[j][k] = random.uniform(-0.5, 0.5)
        
        self.v0 = [0 for m in range(self.hidden_layer_size)]
        for j in range(self.hidden_layer_size):
                self.v0[j] = random.uniform(-0.5, 0.5)        
        
        self.v = [[0 for p in range(self.input_layer_size)] for q in range(self.hidden_layer_size)]
        for i in range(self.hidden_layer_size):
            for j in range(self.input_layer_size):
                self.v[i][j] = random.uniform(-0.5, 0.5)
        
#        print self.w0, self.w, self.v0, self.v

    # membuka dan menyimpan data latih dan data uji
    def open_dataset(self, location, train_filename, train_sheet_name, 
                     test_filename, test_sheet_name, t_filename, t_sheet_name):
        
        # membuka workbook
        train_workbook = xlrd.open_workbook(location+train_filename)
        train_worksheet = train_workbook.sheet_by_name(train_sheet_name)
        
        test_workbook = xlrd.open_workbook(location+test_filename)
        test_worksheet = test_workbook.sheet_by_name(test_sheet_name)
        
        t_workbook = xlrd.open_workbook(location+t_filename)
        t_worksheet = t_workbook.sheet_by_name(t_sheet_name)
        
        # menyimpan data latih dan uji dari excel
        counter = 0
        self.x = [[0 for m in range(self.input_layer_size)] for n in range(self.train_size)]
        self.test = [[0 for m in range(self.input_layer_size)] for n in range(self.test_size)]
        for i in range(self.train_size+self.test_size):
            for j in range(self.input_layer_size):
                if i < self.train_size:
                    # menyimpan data latih
                    self.x[counter][j] = train_worksheet.cell(i, j).value
                else:
                    # menyimpan data uji
                    self.test[counter][j] = test_worksheet.cell(counter, j).value    
                                  
            # mereset counter jika data latih sudah disimpan semua
            if counter == self.train_size-1:
                counter = 0
            else:
                counter += 1     
        
        # menyimpan target data latih dan uji dari excel
        counter = 0
        self.t = [0 for m in range(self.train_size)]
        self.tt = [0 for m in range(self.test_size)]
        for i in range(self.train_size+self.test_size):
            if i < self.train_size:
                # menyimpan data latih
                self.t[counter] = t_worksheet.cell(i, 0).value
            else:
                # menyimpan data uji
                self.tt[counter] = t_worksheet.cell(i, 0).value    
                                  
            # mereset counter jika data latih sudah disimpan semua
            if counter == self.train_size-1:
                counter = 0
            else:
                counter += 1
    
    # mencetak data latih, data uji, target
    def print_dataset(self):
        print self.x, self.test, self.t, self.tt
    
    # menjalankan pembelajaran bpnn
    def learn(self):
        epoch = 0
        stopping_condition = False
        while stopping_condition == False:
            # untuk setiap data latih
            for l in range(len(self.x)):
                # umpan maju (feedforward)
                # menghitung z_in dan z (lapisan tersembunyi)
                # inisialisasi z
                z = []
                for j in range(self.hidden_layer_size):
                    # inisialisasi z_in
                    z_in = []
                    # inisialisasi sigma xv
                    xv = 0
                    for i in range(self.input_layer_size):
                        # menghitung xv per-indeks
                        xv_temp = self.x[l][i]*self.v[j][i]
                        # menjumlahkan ke sigma xv
                        xv += xv_temp
                    # menghitung z_in
                    z_in.append(self.v0[j]+xv)
                    # menghitung z
                    z.append(1/(1+math.exp(-z_in[j])))
                
                
                # menghitung y_in dan y (lapisan keluaran)
                # inisialisasi y
                y = []
                for k in range(self.output_layer_size):
                    # inisialisasi y_in
                    y_in = []
                    # inisialisasi sigma zw
                    zw = 0
                    for j in range(self.hidden_layer_size):
                        # menghitung zw per-indeks
                        zw_temp = z[j]*self.w[j][k]
                        # menjumlahkan ke sigma zw
                        zw += zw_temp
                    # menghitung y_in
                    y_in.append(self.w0[k]+zw)
                    # menghitung y
                    y.append(1/(1+math.exp(-y_in[k])))
                    
                    # menghitung sigma y untuk MSE
                    self.MSE += math.pow(y[k],2)

                    
                # propagasi error (backpropagation of error)
                # menghitung kesalahan keluaran dan delta bobot dan delta bias
                # inisialisasi delta output
                delta_y = []
                # inisialisasi Delta w (bobot)
                Delta_w = [[0 for m in range(self.hidden_layer_size)] for n in range(self.output_layer_size)]
                # inisialisasi Delta w0 (bias)
                Delta_w0 = []
                for k in range(self.output_layer_size):
                    # menghitung kesalahan keluaran (output)
                    temp = (self.t[l]-y[k])*y[k]*(1-y[k])                    
                    delta_y.append(temp)
                    
                    for j in range(self.hidden_layer_size):
                        # menghitung koreksi kesalahan (Delta w)
                        Delta_w[j][k] = self.alfa*delta_y[k]*z[j]                        
                    
                    # menghitung koreksi bias (Delta w0)
                    Delta_w0.append(self.alfa*delta_y[k])
                
                
                # menghitung kesalahan tersembunyi dan delta bobot dan delta bias
                # inisialisasi delta input
                delta_in = []
                # inisialisasi delta hidden
                delta_z = []
                # inisialisasi Delta v (bobot)
                Delta_v = [[0 for p in range(self.input_layer_size)] for q in range(self.hidden_layer_size)]
                # inisialisasi Delta v0 (bias)
                Delta_v0 = []
                for j in range(self.hidden_layer_size):
                    temp = 0
                    for k in range(self.output_layer_size):
                        temp += delta_y[k]*self.w[j][k]
                    
                    # delta input
                    delta_in.append(temp)
                    # delta hidden
                    delta_z.append(delta_in[j]*z[j]*(1-z[j]))
                    
                    for i in range(self.input_layer_size):
                        # menghitung koreksi kesalahan (Delta v)
                        Delta_v[j][i] = self.alfa*delta_z[j]*self.x[l][i]
                        
                    # menghitung koreksi bias (Delta v0)
                    Delta_v0.append(self.alfa*delta_z[j])
                
                # memperbarui bobot dan bias
                # memparbarui bobot bias keluaran
                for j in range(self.hidden_layer_size):
                    for k in range(self.output_layer_size):
                        self.w[j][k] += Delta_w[j][k]
                    
                for k in range(self.output_layer_size):
                    self.w0[k] += Delta_w0[k]
                
                # memperbarui bobot bias tersembunyi
                for i in range(self.input_layer_size):
                    for j in range(self.hidden_layer_size):
                        self.v[j][i] += Delta_v[j][i]
                    
                for j in range(self.hidden_layer_size):
                    self.v0[j] += Delta_v0[j]
            
            # menambahkan epoch
            epoch += 1
            
            # menghitung Mean Square Error
            self.MSE = self.MSE/self.input_layer_size
#            print 'epoch: '+str(epoch)+' MSE: '+str(self.MSE)
            # mengecek stopping condition
#            if epoch == self.maxEpoch and self.MSE > self.ERR:
            if epoch == self.maxEpoch:
                stopping_condition = True
#                print self.maxEpoch
    
    # menyimpan bobot dan bias ke excel
    def export_weight(self, filename):
        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet('Bobot')
        
        row = 0
        col = 3
        
        worksheet.write_number(row, 0, self.w0[0])
        worksheet.write_number(row, 1, self.w[0][0])
        worksheet.write_number(row, 2, self.v0[0])
        
        for i in range(self.input_layer_size):
            worksheet.write_number(0, col, self.v[0][i])            
            col += 1                

        workbook.close()
    
    def open_model(self, location, filename, sheet_name):
        # membuka workbook
        workbook = xlrd.open_workbook(location+filename)
        worksheet = workbook.sheet_by_name(sheet_name)
        
        # menyimpan w0, w, v0, v dari excel
        self.w0[0] = worksheet.cell(0, 0).value
        self.w[0][0] = worksheet.cell(0, 1).value
        self.v0[0] = worksheet.cell(0, 2).value       
               
        for i in range(self.input_layer_size):
            self.v[0][i] = worksheet.cell(0, i+3).value    
        
#        print self.w0,self.w,self.v0,self.v
        
    def testing(self):
        counter = 0
        a = 0.0
        b = 0.0
        c = 0.0
        d = 0.0
        for x in range(self.test_size):
            predict = classify2(self.test[x],self.w0,self.w,self.v0,self.v,self.input_layer_size)
#            print 'data '+str(x+1)+' '+str(predict)+' '+str(self.tt[x])
            if predict == self.tt[x]:
                counter += 1
            if (predict==1) and (self.tt[x]==1):
                a += 1
            if (predict==1) and (self.tt[x]==0):
                b += 1
            if (predict==0) and (self.tt[x]==1):
                c += 1
            if (predict==0) and (self.tt[x]==0):
                d += 1
#        print counter
#        print a
#        print b
#        print c
#        print d
        if a == 0:
            precision = 0
            recall = 0
            fmeasure = 0
        else:                
            precision = a/(a+b)
            recall = a/(a+c)
            fmeasure = (2*precision*recall)/(precision+recall)
#        print 'klasifikasi benar: '+str(counter)
#        print 'precision: '+str(precision)
#        print 'recall: '+str(recall)
#        print 'f-measure: '+str(fmeasure)
        print str(counter)+'\t'+str(precision)+'\t'+str(recall)+'\t'+str(fmeasure)
    
    # menjalankan proses klasifikasi
    def classify(self, data):    
        # umpan maju (feedforward)
        # menghitung z_in dan z (lapisan tersembunyi)
        # inisialisasi z
        z = []
        for j in range(self.hidden_layer_size):
            # inisialisasi z_in
            z_in = []
            # inisialisasi sigma xv
            xv = 0
            for i in range(self.input_layer_size):
                # menghitung xv per-indeks
                xv_temp = data[i]*self.v[j][i]
                # menjumlahkan ke sigma xv
                xv += xv_temp
            # menghitung z_in
            z_in.append(self.v0[j]+xv)
            # menghitung z
            z.append(1/(1+math.exp(-z_in[j])))
                
                
        # menghitung y_in dan y (lapisan keluaran)
        # inisialisasi y
        y = []
        for k in range(self.output_layer_size):
            # inisialisasi y_in
            y_in = []
            # inisialisasi sigma zw
            zw = 0
            for j in range(self.hidden_layer_size):
                # menghitung zw per-indeks
                zw_temp = z[j]*self.w[j][k]
                # menjumlahkan ke sigma zw
                zw += zw_temp
            # menghitung y_in
            y_in.append(self.w0[k]+zw)
            # menghitung y
            y.append(1/(1+math.exp(-y_in[k])))
           
            # membulatkan hasil keluaran (y)
            if y[k] > 0.5:
                y[k] = 1
            else:
                y[k] = 0
            
##            # mencetak hasil klasifikasi
            print "klasifikasi: "+str(y[k])
#            # mengembalikan hasil klasifikasi
#            return str(y[k])
                               
def classify2(data,w0,w,v0,v,isize):    
    # umpan maju (feedforward)
    # menghitung z_in dan z (lapisan tersembunyi)
    # inisialisasi z
    z = []
    for j in range(1):
        # inisialisasi z_in
        z_in = []
        # inisialisasi sigma xv
        xv = 0
        for i in range(isize):
            # menghitung xv per-indeks
            xv_temp = data[i]*v[j][i]
            # menjumlahkan ke sigma xv
            xv += xv_temp
        # menghitung z_in
        z_in.append(v0[j]+xv)
        # menghitung z
        z.append(1/(1+math.exp(-z_in[j])))
                
                
    # menghitung y_in dan y (lapisan keluaran)
    # inisialisasi y
    y = []
    for k in range(1):
        # inisialisasi y_in
        y_in = []
        # inisialisasi sigma zw
        zw = 0
        for j in range(1):
            # menghitung zw per-indeks
            zw_temp = z[j]*w[j][k]
            # menjumlahkan ke sigma zw
            zw += zw_temp
        # menghitung y_in
        y_in.append(w0[k]+zw)
        # menghitung y
        y.append(1/(1+math.exp(-y_in[k])))
        
        # membulatkan hasil keluaran (y)
        if y[k] > 0.5:
            y[k] = 1
        else:
            y[k] = 0
            
##      # mencetak hasil klasifikasi
#       print "klasifikasi: "+str(y[k])
#       mengembalikan hasil klasifikasi
        return y[k]

# untuk pengujian manualisasi
#pp = Preprocessing(4,1,2)
#pp.open_dataset('D:/K/S\'\'\'\'\'\'\'/S/Data/','Dataset_sample.xlsx', 'Sample')
#
#
## untuk pembelajaran seseungguhnya
#
## instansiasi objek dan parameter
#pp = Preprocessing(400,100,2)
#pp.open_dataset('D:/K/S\'\'\'\'\'\'\'/Skripsi saya/Data/Dataset/','Dataset - KFold Cross Validation.xlsx', 'Dataset 2')
#
##pp.print_data_train()
##pp.print_data_test()
#pp.export_output_train_test()
#
## mengambil database emoticon, stopword, dan sentimen strength
#pp.get_emoticons()
#pp.get_stopwords()
#pp.get_sentistrength()
#
## proses Preprocessing
#pp.tweets_cleaning()
#pp.case_folding()
#pp.separate_emoticons()
#pp.filtering()
#pp.sub_preprocessing()
#
## mendapatkan data latih untuk BPNN setelah diektraksi
#pp.bags_of_words()
#pp.lexicon_based_features()
#pp.join_data_train()
#pp.get_bpnn_output()
#
## mendapatkan data uji untuk BPNN setelah diektraksi
#pp.sub_preprocessing_test()
#pp.bags_of_words_test()
#pp.lexicon_based_features_test()
#pp.join_data_test()
#pp.get_bpnn_output_test()
#
## export hasil ekstraksi
#pp.export_fitur_input()
#pp.export_data_train()
#pp.export_data_test()


# proses BPNN
bpnn = BPNN()
# BoW dan LBF
isize = 1535
## LBF
#isize = 4
## BoW
#isize = 1516
hsize = 1
osize = 1
err = 0.01
alfa = 0.4
maxE = 350

bpnn.initialize2(isize,hsize,osize,alfa,maxE,err,400,100)
bpnn.open_dataset('D:/K/S\'\'\'\'\'\'\'/Skripsi saya/Data/Dataset/2/','DataLatihBPNN.xlsx',
                 'Data Latih','DataUjiBPNN.xlsx','Data Uji',
                 'OutputBPNN.xlsx','Target Output')


## mengetes hasil sistem setelah diketahui parameter optimal, pengujian terakhir
#for y in range(5):
#    bpnn.initialize_weight_bias()
#    bpnn.open_model('D:/K/S\'\'\'\'\'\'\'/Skripsi saya/Data/Dataset/2/Learning Rate/','Bobot Learning Rate 0.4 rand '+str(y+1)+'.xlsx','Bobot')
#    bpnn.testing()

## mengetes hasil sistem dengan contoh satu tweet
#bpnn.initialize_weight_bias()
#bpnn.open_model('D:/K/S\'\'\'\'\'\'\'/Skripsi saya/Data/Learning Rate/','Bobot Learning Rate 0.4 rand 1.xlsx','Bobot')
#test = [0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,0]
#bpnn.classify(test)



## pengujian maxEpoch
#maxE = 5
#for x in range(10):
#    for y in range(5):
#        bpnn.initialize_weight_bias()
#        bpnn.open_model('D:/K/S\'\'\'\'\'\'\'/Skripsi saya/Data/Dataset/1/Max Epoch/','Bobot MaxEpoch '+str(maxE)+' rand '+str(y+1)+'.xlsx','Bobot')
#        bpnn.testing()
##        print 'Max Epoch '+str(maxE)+' rand '+str(y+1)
#    maxE += 5
    
## pengujian Learning Rate
#lr = 0.1
#for x in range(10):
#    for y in range(5):
#        bpnn.initialize_weight_bias()
#        bpnn.open_model('D:/K/S\'\'\'\'\'\'\'/Skripsi saya/Data/Dataset/1/Learning Rate/','Bobot Learning Rate '+str(lr)+' rand '+str(y+1)+'.xlsx','Bobot')
#        bpnn.testing()
##        print 'Max Epoch '+str(maxE)+' rand '+str(y+1)
#    lr += 0.1


## mendapatkan nilai bobot dengan nilai max epoch dan learning rate
#maxE = 5
#alfa = 0.4
#kelompok = 1
#fitur = 'BoW'
#bpnn.initialize2(isize,hsize,osize,alfa,maxE,err,400,100)
##bpnn.open_dataset('D:/K/S\'\'\'\'\'\'\'/Skripsi saya/Data/Dataset/'+str(kelompok)+'/','DataLatihBPNN - '+fitur+'.xlsx',
##             'Data Latih','DataUjiBPNN - '+fitur+'.xlsx','Data Uji',
##             'OutputBPNN.xlsx','Target Output')
#bpnn.open_dataset('D:/K/S\'\'\'\'\'\'\'/Skripsi saya/Data/Dataset/'+str(kelompok)+'/','DataLatihBPNN.xlsx',
#             'Data Latih','DataUjiBPNN.xlsx','Data Uji',
#             'OutputBPNN.xlsx','Target Output')
#for y in range(5):
#    bpnn.initialize_weight_bias()
#    bpnn.learn()
#    bpnn.export_weight('Bobot Max Epoch '+str(maxE)+' Learning Rate '+str(alfa)+' rand '+str(y+1)+'.xlsx')    
#    print 'Bobot Max Epoch '+str(maxE)+' Learning Rate '+str(alfa)+' rand '+str(y+1)    

## mengetes hasil sistem setelah diketahui parameter optimal, pengujian terakhir
#for x in range(5):
#    bpnn.initialize_weight_bias()
#    bpnn.open_model('D:/K/S\'\'\'\'\'\'\'/Skripsi saya/','Bobot Max Epoch '+str(maxE)+' Learning Rate '+str(alfa)+' rand '+str(x+1)+'.xlsx','Bobot')
#    bpnn.testing()

maxE = 5
bpnn.initialize2(isize,hsize,osize,alfa,maxE,err,400,100)
bpnn.open_dataset('D:/K/S\'\'\'\'\'\'\'/Skripsi saya/Data/Dataset/1/','DataLatihBPNN.xlsx',
                 'Data Latih','DataUjiBPNN.xlsx','Data Uji',
                 'OutputBPNN.xlsx','Target Output')
bpnn.initialize_weight_bias()
bpnn.learn()
bpnn.export_weight('Bobot MaxEpoch '+str(maxE)+' rand 1.xlsx')    
print 'Max Epoch '+str(maxE)+' rand '+str(y+1)

## mendapatkan nilai bobot dengan nilai max epoch
#for x in range(10):
#    bpnn.initialize2(isize,hsize,osize,alfa,maxE,err,400,100)
#    bpnn.open_dataset('D:/K/S\'\'\'\'\'\'\'/Skripsi saya/Data/Dataset/4/','DataLatihBPNN.xlsx',
#                 'Data Latih','DataUjiBPNN.xlsx','Data Uji',
#                 'OutputBPNN.xlsx','Target Output')
#    for y in range(5):
#        bpnn.initialize_weight_bias()
#        bpnn.learn()
#        bpnn.export_weight('Bobot MaxEpoch '+str(maxE)+' rand '+str(y+1)+'.xlsx')    
#        print 'Max Epoch '+str(maxE)+' rand '+str(y+1)
#    maxE += 50


## mendapatkan nilai bobot dengan nilai learning rate
#alfa = 0.1
#for x in range(10):
#    bpnn.initialize2(isize,hsize,osize,alfa,maxE,err,400,100)
#    bpnn.open_dataset('D:/K/S\'\'\'\'\'\'\'/Skripsi saya/Data/Dataset/5/','DataLatihBPNN.xlsx',
#                 'Data Latih','DataUjiBPNN.xlsx','Data Uji',
#                 'OutputBPNN.xlsx','Target Output')
#    for y in range(5):
#        bpnn.initialize_weight_bias()
#        bpnn.learn()
#        bpnn.export_weight('Bobot Learning Rate '+str(alfa)+' rand '+str(y+1)+'.xlsx')    
#        print 'Learning Rate '+str(alfa)+' rand '+str(y+1)
#    alfa += 0.1



# mendapatkan nilai bobot
#for x in range(5):
#    bpnn.initialize_weight_bias()
#    bpnn.learn()
#    bpnn.export_weight('Bobot Learning Rate '+str(alfa)+' rand '+str(x+1)+'.xlsx')    
#    print "sukses random ke-"+str(x+1)


## pengujian learning rate
#lr = 0.1
#for x in range(10):
#    for y in range(5):
#        bpnn.initialize_weight_bias()
#        print 'Learning Rate '+str(lr)+' rand '+str(y+1)
#        bpnn.open_model('D:/K/S\'\'\'\'\'\'\'/Skripsi saya/Data/Learning Rate/','Bobot Learning Rate '+str(lr)+' rand '+str(y+1)+'.xlsx','Bobot')
#        bpnn.testing()
#    lr += 0.1







#x = [[0,1,1,0,1,0,0,1,0,0,0,0,1,1,0,1,0,0,1,0,0,1,0,0,1,0,1,1,1,1,0],
#     [0,0,0,0,0,1,0,0,0,1,0,1,0,0,1,0,0,1,0,0,0,0,0,1,0,0,0,0,1,1,0],     
#     [1,1,0,0,0,0,1,0,1,0,0,1,0,0,0,0,0,0,0,1,1,0,1,0,0,1,0,0,1,1,1],
#     [1,0,0,1,0,0,0,0,0,0,1,0,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,1,0,1,1]]
#t = [1,1,0,0]
#test = [0,1,1,0,1,1,0,0,0,0,0,0,1,0,0,0,0,0,0,0,0,0,0,0,0,0,0,0,1,1,0]
#v0 = [0.2]
#v = [[0.2,0.3,0.3,0.1,0.1,0.2,0.2,0.1,0.1,0.3,0.3,0.1,0.1,0.2,0.2,
#      0.1,0.2,0.2,0.1,0.1,0.3,0.3,0.1,0.1,0.2,0.2,0.1,0.1,0.3,0.3,0.1]]
#w0 = [0.1]
#w = [[0.3]]
#bpnn.initialize(x,t,isize,hsize,osize,v0,v,w0,w,alfa,maxE,err,x,test)
#bpnn.learn()
#bpnn.classify(test)
#bpnn.print_dataset()
