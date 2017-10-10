# -*- coding: utf-8 -*-
# author:Haochun Wang

import xlrd, time


class Apriori:
    def __init__(self, add, min_support=0.02):
        '''
        This is the main function of apriori
        :param dataset: the dataset input
        :param min_support: the minimum support
        :return: null
        '''
        dataset = self.load_data_set(add)
        C1 = self.createC1(dataset)  # build C1
        # d_set: [set([1, 3, 4]), set([2, 3, 5]), set([1, 2, 3, 5]), set([2, 5])]
        d_set = map(set,dataset)
        L1, supportdata = self.scanD(d_set, C1, min_support)
        L = [L1]
        k = 2
        while len(L[k-2]) > 0:
            Ck = self.apriori_gen(L[k-2], k)
            Lk, supk = self.scanD(d_set, Ck, min_support)
            supportdata.update(supk)
            L.append(Lk)
            k += 1

        self.output(L, add, min_support)


    def load_data_set(self, s_addr):
        '''
        This function is to load the excel file and store all transactions in memory as a list
        , with a format of list[[transaction1],[transaction2],,,,,[transaction n]]
        :param s_addr: the soure address of the data file
        :return: a list that includes all transactions
        '''
        items = xlrd.open_workbook(s_addr).sheet_by_index(4)  # open the excel file
        trans_list = []
        for i in xrange(1, items.nrows):
            tp = []
            j = 1
            try:
                # put the products in one transaction in a list
                while isinstance(items.row_values(i)[j], float):
                    tp.append(items.row_values(i)[j])
                    j += 1
            except:
                pass
            trans_list.append(tp)
        return trans_list
        #return [[1, 3, 4], [2, 3, 5], [1, 2, 3, 5], [2, 5]]

    def createC1(self, dataset):
        '''
        This function is to create the initial C1 set for the data set
        :param dataset: the list that includes all the transactions
        :return: a list that
        '''
        C1 = []   # the set of the items with the number of no less than 1
        for transaction in dataset:
            for item in transaction:
                if not [item] in C1:
                    C1.append([int(item)])
        C1.sort()
        return map(frozenset,C1)    # frozenset is a kind of set that can not be modified

    def scanD(self, d_set,Ck,min_support):
        '''
        This function is to generate L1 from C1. L1 is the elements set whose supports are no less than the min_support
        :param d_set: the lists of the candidates
        :param Ck: the data set
        :param min_support: the minimum of the support
        :return:
        '''
        s_dic = {}
        for i in d_set:
            for candidate in Ck:
                # if each element of candidate is in i, return true
                if candidate.issubset(i):
                    # count the times that each items set appear and store that in the s_dic dictionary
                    # the key of dictionary is the items set
                    # the value of the dictionary is the times that items set appear
                    if not s_dic.has_key(candidate):
                        s_dic[candidate] = 1
                    else:
                        s_dic[candidate] += 1
        num_items = float(len(d_set))
        retlist = []
        supportdata = {}
        for key in s_dic:
            # calculate the support of each items set
            # if that is no less than the min_support, put it in the retlist
            support = s_dic[key]/num_items
            if support >= min_support:
                retlist.insert(0, key)
            # build the dictionary of support data
            supportdata[key] = support
        return retlist,supportdata

    def apriori_gen(self, Lk,k):
        '''
        This function is to generate Lk in a loop
        :param Lk: the list of frequent items set
        :param k: the number of items
        :return: Ck
        '''
        retList = []
        lenLk = len(Lk)
        for i in range(lenLk):
            for j in range(i+1,lenLk):
                # union the two sets when the first k-2 items are equal
                L1 = list(Lk[i])[:k-2]
                L2 = list(Lk[j])[:k-2]
                L1.sort()
                L2.sort()
                if L1 == L2:
                    retList.append(Lk[i] | Lk[j])
        return retList

    def output(self, listL, addr, min_support):
        i = 0
        products = xlrd.open_workbook(addr).sheet_by_index(2)
        for one in listL[:-1]:
            #print type(one[0])
            print "The frequent items set with %s of the number of items:" % (i + 1)\
            #, one, "\n"
            for j in one:
                tp = []
                for k in j:
                    name = str(products.row_values(k)[0])
                    tp.append(name)
                print tp
            i += 1
        i = 0
        with open('result-minu_support%s.txt' % min_support, 'w') as resfile:
            for one in listL[:-1]:
                resfile.write("The frequent items set with %s of the number of items:" % (i + 1))
                resfile.write('\n')
                for j in one:
                    tp = []
                    for k in j:
                        name = str(products.row_values(k)[0])
                        tp.append(name)
                    resfile.write(str(tp))
                    resfile.write('\t')
                    #print tp
                resfile.write("\n")

                i += 1



    #def visualize(self):


if __name__ == "__main__":
    start = time.clock()
    # Apriori('/Users/hcwang/OneDrive/1-UQ/2017s2/DM/UQ_DataMining/processed_data/1.xls', min_support=0.01)
    # min_list = [0.1, 0.09, 0.08, 0.07, 0.06, 0.05, 0.04, 0.03, 0.02, 0.009, 0.008, 0.007]
    min_list = [0.006, 0.005, 0.004, 0.003, 0.002, 0.001]
    for i in min_list:
        Apriori('/Users/hcwang/OneDrive/1-UQ/2017s2/DM/UQ_DataMining/processed_data/1.xls', min_support=i)
    end = time.clock()
    print "Done! With a time consumption of %0.2f seconds" % (end - start)