# -*- coding: utf-8 -*-
# author:Haochun Wang
import xlrd, xlwt

'''
    This class is to find the purchasing pattern for each customer.
    Including regular shopping day from whole week,
              regular commodities,
'''

'''
class Dic_day:
    def __init__(self):
        self.map_dic = {3: 'Fri', 4: 'Sat', 5: 'Sun', 6: 'Mon', 0: 'Tue', 1: 'Wed', 2: 'Thu'}
        self.day_dic = {'Mon': 0, 'Tue': 0, 'Wed': 0, 'Thu': 0, 'Fri': 0, 'Sat': 0, 'Sun': 0}
'''


class Pattern:
    def __init__(self, s_addr, d_addr, sheet_index=0, len_dic=10):
        '''
        :param s_addr: the source address of excel file, in order by customer and time
        :param d_addr: the destination address of excel file that includes the result patterns
        :param sheet_index: the index of the target sheet, with the default of 0
        :param len_dic: the length of the dictionary that contains the purchasing pattern
        '''
        self.s_addr = s_addr
        self.d_addr = d_addr
        self.sheet_index = sheet_index
        self.len_dic = len_dic
        pass

    def __load__(self, index):
        return xlrd.open_workbook(self.s_addr).sheet_by_index(index)

    def __new_sheet__(self):
        self.newbook = xlwt.Workbook(encoding='utf-8', style_compression=0)
        return self.newbook.add_sheet('sheet1', cell_overwrite_ok=True)

    def pattern(self):
        items = self.__load__(self.sheet_index)  # load the execl file
        map_items = self.__load__(2)
        map_com_dic ={}
        for i in xrange(0, map_items.nrows):
            map_com_dic[int(map_items.row_values(i)[1])] = str(map_items.row_values(i)[0])
        print map_com_dic
        newsheet = self.__new_sheet__()  # build a new excel file
        newsheet.write(0, 0, 'keyC')  # title info
        newsheet.write(0, 1, 'pt-comm')
        newsheet.write(0, 2, 'pt-day')

        # {3:Fri, 4:Sat, 5:Sun, 6: Mon, 0:Tue, 1: Wed, 2:Thu}
        map_dic = {3: 'Fri', 4: 'Sat', 5: 'Sun', 6: 'Mon', 0: 'Tue', 1: 'Wed', 2: 'Thu'}
        day_dic = {'Mon': 0, 'Tue': 0, 'Wed': 0, 'Thu': 0, 'Fri': 0, 'Sat': 0, 'Sun': 0}
        comm_dic ={}
        res_comm_dic = {}
        res_comm_dic_final = {}
        res_day_dic = {}
        tpC = 0
        write_row = 1
        for i in xrange(1, items.nrows - 1):
        # for i in xrange(1, 20):
            current_row = items.row_values(i)[0]
            j = 2
            if current_row != tpC:
                if tpC != 0:
                    newsheet.write(write_row, 0, tpC)
                    #while len(comm_dic) > self.len_dic:

                    res_comm_dic = comm_dic
                    l = 1
                    while len(res_comm_dic) > 10:
                        res_comm_dic = {}
                        for k in comm_dic:
                            if comm_dic[k] > l:
                                res_comm_dic[k] = comm_dic[k]
                                l += 1
                                #comm_dic.pop(k)
                    for k in res_comm_dic:
                        res_comm_dic_final[map_com_dic[k]] = res_comm_dic[k]

                    for m in day_dic:
                        if day_dic[m] > 0:
                            res_day_dic[m] = day_dic[m]
                    sortlist_day = sorted(day_dic.iteritems(), key=lambda asd:asd[1], reverse=True)
                    sortlist_com = sorted(res_comm_dic_final.iteritems(), key=lambda asd:asd[1], reverse=True)
                    #print sortlist
                    newsheet.write(write_row, 1, str(res_comm_dic_final))
                    newsheet.write(write_row, 2, str(res_day_dic))
                    newsheet.write(write_row, 3, str(sortlist_day[0]))
                    newsheet.write(write_row, 4, str(sortlist_com[0:4]))
                    write_row += 1
                tpC = current_row
                comm_dic = {}
                day_dic = {'Mon': 0, 'Tue': 0, 'Wed': 0, 'Thu': 0, 'Fri': 0, 'Sat': 0, 'Sun': 0}

            day_dic[map_dic[items.row_values(i)[1] % 7]] += 1
            try:
                # put the products in one transaction in a list
                while isinstance(items.row_values(i)[j], float):
                    if comm_dic.has_key(int(items.row_values(i)[j])):
                        comm_dic[int(items.row_values(i)[j])] += 1
                    else:
                        comm_dic[int(items.row_values(i)[j])] = 1
                    j += 1
            except:
                pass
        self.newbook.save(self.d_addr)
pat = Pattern("/Users/hcwang/OneDrive/1-UQ/2017s2/DM/UQ_DataMining/processed_data/sales_m_t_m.xls"
              , "/Users/hcwang/OneDrive/1-UQ/2017s2/DM/UQ_DataMining/processed_data/sales_pattern.xls")
pat.pattern()














